"""
Detekcija zanimljivih kretanja.

Namjerno jednostavan rule-based sistem: svaki uslov je jedno pravilo,
pragovi dolaze iz config.py, a izlaz je tabela spremna za Excel.
"""

from __future__ import annotations

import logging

import pandas as pd

import config

logger = logging.getLogger(__name__)

ALERT_COLUMNS = ["Datum", "Instrument", "Kategorija", "Tip", "Vrijednost", "Opis", "Smjer"]


def _row(date, name, category, kind, value, description, direction) -> dict:
    return {
        "Datum": date,
        "Instrument": name,
        "Kategorija": category,
        "Tip": kind,
        "Vrijednost": value,
        "Opis": description,
        "Smjer": direction,
    }


def detect(snapshot: pd.DataFrame) -> pd.DataFrame:
    """Prodje kroz zadnje stanje svakog instrumenta i vrati sve pogodjene alerte."""
    if snapshot.empty:
        return pd.DataFrame(columns=ALERT_COLUMNS)

    found: list[dict] = []

    for _, r in snapshot.iterrows():
        date, name, category = r["date"], r["name"], r["category"]

        # --- 1. Velika dnevna promjena cijene ---
        change = r.get("return_1d")
        if pd.notna(change) and abs(change) * 100 >= config.ALERT_PRICE_CHANGE_PCT:
            direction = "Rast" if change > 0 else "Pad"
            found.append(_row(
                date, name, category, "Promjena cijene", change,
                f"Dnevna promjena {change * 100:+.2f}% "
                f"(prag {config.ALERT_PRICE_CHANGE_PCT:.1f}%)",
                direction,
            ))

        # --- 2. RSI ekstremi ---
        rsi_value = r.get("rsi")
        if pd.notna(rsi_value):
            if rsi_value >= config.ALERT_RSI_OVERBOUGHT:
                found.append(_row(
                    date, name, category, "RSI overbought", rsi_value,
                    f"RSI {rsi_value:.1f} - moguce prekupljeno stanje", "Rast",
                ))
            elif rsi_value <= config.ALERT_RSI_OVERSOLD:
                found.append(_row(
                    date, name, category, "RSI oversold", rsi_value,
                    f"RSI {rsi_value:.1f} - moguce preprodano stanje", "Pad",
                ))

        # --- 3. Neuobicajen volume ---
        ratio = r.get("volume_ratio")
        if pd.notna(ratio) and ratio >= config.ALERT_VOLUME_MULTIPLIER:
            found.append(_row(
                date, name, category, "Volume anomalija", ratio,
                f"Volume {ratio:.1f}x veci od {config.VOLUME_AVG_WINDOW}-dnevnog prosjeka",
                "Neutralno",
            ))

        # --- 4. Blizina 52-nedjeljnog maksimuma / minimuma ---
        from_high = r.get("pct_from_52w_high")
        if pd.notna(from_high) and abs(from_high) * 100 <= config.ALERT_NEAR_52W_HIGH_PCT:
            found.append(_row(
                date, name, category, "Blizu 52w maksimuma", from_high,
                f"Cijena je {abs(from_high) * 100:.1f}% ispod godisnjeg maksimuma", "Rast",
            ))

        from_low = r.get("pct_from_52w_low")
        if pd.notna(from_low) and from_low * 100 <= config.ALERT_NEAR_52W_LOW_PCT:
            found.append(_row(
                date, name, category, "Blizu 52w minimuma", from_low,
                f"Cijena je {from_low * 100:.1f}% iznad godisnjeg minimuma", "Pad",
            ))

        # --- 5. Veliki drawdown ---
        dd = r.get("drawdown")
        if pd.notna(dd) and abs(dd) * 100 >= config.ALERT_DRAWDOWN_PCT:
            found.append(_row(
                date, name, category, "Drawdown", dd,
                f"Pad od vrha {dd * 100:.1f}%", "Pad",
            ))

        # --- 6. Cijena presjekla MA200 ---
        price_vs_ma = r.get("price_vs_ma200")
        if pd.notna(price_vs_ma) and abs(price_vs_ma) * 100 <= 1.0:
            found.append(_row(
                date, name, category, "Test MA200", price_vs_ma,
                "Cijena je unutar 1% od 200-dnevnog prosjeka", "Neutralno",
            ))

    if not found:
        logger.info("Nijedan alert nije pogodjen.")
        return pd.DataFrame(columns=ALERT_COLUMNS)

    df = pd.DataFrame(found)[ALERT_COLUMNS]
    df = df.sort_values(["Kategorija", "Instrument", "Tip"]).reset_index(drop=True)
    logger.info("Pronadjeno %d alerta.", len(df))
    return df
