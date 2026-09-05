"""Generise sinteticku historiju za testiranje bez interneta/API kljuca."""
import numpy as np
import pandas as pd

import config

rng = np.random.default_rng(7)
dates = pd.bdate_range(end="2026-09-03", periods=700)

base_price = {
    "AAPL": 180, "MSFT": 410, "NVDA": 120, "AMZN": 185, "JPM": 205,
    "SPY": 540, "QQQ": 470, "EUR/USD": 1.08, "GBP/USD": 1.27, "USD/JPY": 150,
    "USD/CHF": 0.88, "GLD": 215, "SLV": 26, "USO": 78,
    "BTC/USD": 62000, "ETH/USD": 3000,
}

market = rng.normal(0.0004, 0.009, len(dates))
rows = []

for inst in config.INSTRUMENTS:
    symbol = inst["symbol"]
    p0 = base_price[symbol]
    beta = 1.2 if inst["category"] == "Stocks" else 0.2
    vol = 0.018 if symbol in ("NVDA", "BTC/USD", "ETH/USD", "USO") else (
        0.004 if inst["category"] == "FX" else 0.011
    )
    r = beta * market + rng.normal(0, vol, len(dates))
    close = p0 * np.exp(np.cumsum(r))
    high = close * (1 + abs(rng.normal(0, 0.004, len(dates))))
    low = close * (1 - abs(rng.normal(0, 0.004, len(dates))))
    op = close * (1 + rng.normal(0, 0.003, len(dates)))
    vols = rng.lognormal(16, 0.4, len(dates)).round() if inst["has_volume"] else [None] * len(dates)

    rows.append(pd.DataFrame({
        "date": dates, "symbol": symbol, "name": inst["name"], "category": inst["category"],
        "open": op, "high": high, "low": low, "close": close, "volume": vols,
    }))

df = pd.concat(rows, ignore_index=True)
df.to_csv(config.HISTORY_FILE, index=False)
print(f"Sinteticka historija: {len(df)} redova, {df.symbol.nunique()} instrumenata")
