"""
Izracun pokazatelja nad historijom.

Sve se racuna u Pandasu i u Excel odlazi kao gotova vrijednost.
Excel formule se koriste samo na Market Overview sheetu (vidi excel_writer.py).
"""

from __future__ import annotations

import logging

import numpy as np
import pandas as pd

import config

logger = logging.getLogger(__name__)


# --------------------------------------------------------------------------
# Pojedinacni indikatori
# --------------------------------------------------------------------------
def rsi(close: pd.Series, period: int = config.RSI_PERIOD) -> pd.Series:
    """
    RSI po Wilderovom izgladjivanju (isto sto koriste TradingView i vecina platformi).
    Wilder koristi EMA sa alpha = 1/period, ne obican rolling prosjek.
    """
    delta = close.diff()
    gain = delta.clip(lower=0.0)
    loss = -delta.clip(upper=0.0)

    avg_gain = gain.ewm(alpha=1 / period, adjust=False, min_periods=period).mean()
    avg_loss = loss.ewm(alpha=1 / period, adjust=False, min_periods=period).mean()

    rs = avg_gain / avg_loss.replace(0, np.nan)
    result = 100 - (100 / (1 + rs))
    # Kad nema gubitaka uopste -> RSI je 100 po definiciji.
    result = result.where(avg_loss != 0, 100.0)
    return result


def macd(close: pd.Series, fast: int = 12, slow: int = 26, signal: int = 9):
    ema_fast = close.ewm(span=fast, adjust=False).mean()
    ema_slow = close.ewm(span=slow, adjust=False).mean()
    macd_line = ema_fast - ema_slow
    signal_line = macd_line.ewm(span=signal, adjust=False).mean()
    return macd_line, signal_line, macd_line - signal_line


def max_drawdown(close: pd.Series) -> pd.Series:
    """Trenutni pad od dosadasnjeg vrha, kao negativan udio (-0.15 = -15%)."""
    running_max = close.cummax()
    return close / running_max - 1.0


# --------------------------------------------------------------------------
# Glavni izracun
# --------------------------------------------------------------------------
def _enrich_one(df: pd.DataFrame) -> pd.DataFrame:
    """Doda sve pokazatelje za jedan instrument (df je vec sortiran po datumu)."""
    df = df.sort_values("date").copy()
    close = df["close"]

    # Prinosi - kao udio (0.015 = 1.5%), jer Excel procente cuva kao udjele.
    df["return_1d"] = close.pct_change(1)
    df["return_5d"] = close.pct_change(5)
    df["return_20d"] = close.pct_change(20)
    df["return_ytd"] = _ytd_return(df)
    df["return_1y"] = close.pct_change(config.WEEK_52_WINDOW)

    # Moving averages
    for window in config.MA_WINDOWS:
        df[f"ma_{window}"] = close.rolling(window, min_periods=window).mean()

    ma_short, ma_long = f"ma_{config.MA_WINDOWS[1]}", f"ma_{config.MA_WINDOWS[2]}"
    df["trend"] = np.where(
        df[ma_short].isna() | df[ma_long].isna(), "n/a",
        np.where(df[ma_short] > df[ma_long], "Uzlazni", "Silazni"),
    )
    df["price_vs_ma200"] = close / df["ma_200"] - 1.0

    # RSI i MACD
    df["rsi"] = rsi(close)
    df["macd"], df["macd_signal"], df["macd_hist"] = macd(close)

    # Volatilnost - anualizovana standardna devijacija dnevnih prinosa
    df["volatility"] = (
        df["return_1d"].rolling(config.VOLATILITY_WINDOW, min_periods=config.VOLATILITY_WINDOW)
        .std() * np.sqrt(config.TRADING_DAYS_PER_YEAR)
    )

    # Volume (samo tamo gdje postoji)
    if df["volume"].notna().any():
        df["volume_avg"] = df["volume"].rolling(
            config.VOLUME_AVG_WINDOW, min_periods=config.VOLUME_AVG_WINDOW
        ).mean()
        df["volume_ratio"] = df["volume"] / df["volume_avg"].replace(0, np.nan)
    else:
        df["volume_avg"] = pd.NA
        df["volume_ratio"] = pd.NA

    # 52-week raspon
    window = min(config.WEEK_52_WINDOW, len(df))
    df["high_52w"] = df["high"].rolling(window, min_periods=1).max()
    df["low_52w"] = df["low"].rolling(window, min_periods=1).min()
    df["pct_from_52w_high"] = close / df["high_52w"] - 1.0
    df["pct_from_52w_low"] = close / df["low_52w"] - 1.0

    # Drawdown od svih vremena u nasoj historiji
    df["drawdown"] = max_drawdown(close)

    return df


def _ytd_return(df: pd.DataFrame) -> pd.Series:
    """Prinos od prvog trgovackog dana tekuce godine."""
    years = df["date"].dt.year
    first_close_of_year = df.groupby(years)["close"].transform("first")
    return df["close"] / first_close_of_year - 1.0


def enrich(history: pd.DataFrame) -> pd.DataFrame:
    """Doda pokazatelje za sve instrumente."""
    if history.empty:
        return history

    history = history.copy()
    history["date"] = pd.to_datetime(history["date"])

    # Namjerno petlja umjesto groupby.apply - ponasanje groupby.apply se
    # mijenjalo izmedju verzija pandasa, a ovako radi svugdje isto.
    frames = [
        _enrich_one(group)
        for _, group in history.sort_values(["symbol", "date"]).groupby("symbol", sort=True)
    ]
    enriched = pd.concat(frames, ignore_index=True)
    logger.info("Pokazatelji izracunati za %d instrumenata.", enriched["symbol"].nunique())
    return enriched


def latest_snapshot(enriched: pd.DataFrame) -> pd.DataFrame:
    """Zadnji red po instrumentu - ovo ide u Excel sheetove."""
    if enriched.empty:
        return enriched
    return (
        enriched.sort_values("date")
        .groupby("symbol", as_index=False)
        .tail(1)
        .sort_values(["category", "name"])
        .reset_index(drop=True)
    )


# --------------------------------------------------------------------------
# Medju-instrumentne mjere
# --------------------------------------------------------------------------
def correlation_matrix(enriched: pd.DataFrame, window: int | None = None) -> pd.DataFrame:
    """Korelacija dnevnih prinosa izmedju svih instrumenata."""
    window = window or config.CORRELATION_WINDOW
    if enriched.empty:
        return pd.DataFrame()

    wide = enriched.pivot_table(index="date", columns="name", values="return_1d")
    wide = wide.tail(window)
    return wide.corr().round(3)


def beta_vs_benchmark(enriched: pd.DataFrame) -> pd.Series:
    """
    Beta svakog instrumenta naspram benchmarka (default SPY).
    Beta = kovarijansa(instrument, benchmark) / varijansa(benchmark).
    """
    if enriched.empty:
        return pd.Series(dtype=float)

    wide = enriched.pivot_table(index="date", columns="symbol", values="return_1d")
    wide = wide.tail(config.CORRELATION_WINDOW)

    benchmark = config.BENCHMARK_SYMBOL
    if benchmark not in wide.columns:
        logger.warning("Benchmark %s nije u podacima - beta se preskace.", benchmark)
        return pd.Series(dtype=float)

    bench = wide[benchmark]
    bench_var = bench.var()
    if not bench_var or np.isnan(bench_var):
        return pd.Series(dtype=float)

    betas = {}
    for symbol in wide.columns:
        joined = pd.concat([wide[symbol], bench], axis=1).dropna()
        if len(joined) < 20:
            continue
        betas[symbol] = joined.cov().iloc[0, 1] / bench_var

    return pd.Series(betas).round(3)


def price_history_wide(enriched: pd.DataFrame, days: int | None = None) -> pd.DataFrame:
    """Cijene u sirokom formatu (datum x instrument) - za grafikon u Excelu."""
    days = days or config.CHART_HISTORY_DAYS
    if enriched.empty:
        return pd.DataFrame()
    wide = enriched.pivot_table(index="date", columns="name", values="close")
    return wide.tail(days)


def normalized_history_wide(enriched: pd.DataFrame, days: int | None = None) -> pd.DataFrame:
    """
    Isto kao gore, ali sve svedeno na 100 na prvi dan.
    Tako se BTC od 60.000 i EUR/USD od 1,08 mogu porediti na istom grafikonu.
    """
    wide = price_history_wide(enriched, days)
    if wide.empty:
        return wide
    return (wide / wide.iloc[0] * 100).round(2)
