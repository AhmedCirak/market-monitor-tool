# Financial Market Monitoring & Reporting Tool

An automated system that **collects → analyzes → stores → reports** on the state of
financial markets. Pulls OHLCV data from the Twelve Data API, calculates technical and
statistical indicators in Pandas, detects interesting price movements, and generates a
formatted Excel report with a dashboard, tables, correlation matrix, and chart.

Uses a **free Twelve Data API key** (no credit card, 800 calls/day) —
sign-up takes a minute at [twelvedata.com](https://twelvedata.com).

---

## What it tracks

15 instruments across three categories:

| Category | Instruments |
|---|---|
| Stocks / ETF | Apple, Microsoft, NVIDIA, Amazon, JPMorgan, S&P 500 ETF, Nasdaq 100 ETF |
| FX | EUR/USD, GBP/USD, USD/JPY, USD/CHF |
| Commodities / Crypto | Gold, Silver, Bitcoin, Ethereum |

The list is configured in `config.py` — adding an instrument is a single line.

## What it calculates

- Returns: 1d, 5d, 20d, YTD, 1Y
- Moving averages 20 / 50 / 200 + trend (MA50 vs MA200) + price vs MA200
- RSI(14) using Wilder's smoothing
- MACD (12/26/9) and histogram
- Annualized volatility (20-day)
- Volume and ratio vs 20-day average
- 52-week high/low and distance from them
- Max drawdown from peak
- Beta vs the S&P 500 ETF
- Correlation matrix of daily returns

## Alerts

Rule-based detection, thresholds configured in `config.py`:

| Type | Condition (default) |
|---|---|
| Price change | \|daily change\| ≥ 3% |
| RSI overbought / oversold | RSI ≥ 70 / RSI ≤ 30 |
| Volume anomaly | volume ≥ 2× 20-day average |
| Near 52w high / low | within 2% of the yearly extreme |
| Drawdown | drop from peak ≥ 20% |
| MA200 test | price within 1% of MA200 |

## Excel report

| Sheet | Content |
|---|---|
| **Market Overview** | Dashboard with **live Excel formulas** — market summary, alerts, breakdown by category, top 5 gainers/losers, alerts by type. Thresholds sit in blue input cells you can edit, and the formulas recalculate automatically. |
| **Stocks / FX / Commodities** | Latest snapshot per instrument: OHLC, returns, volume, 52w range, volatility, RSI, trend |
| **Technical Analysis** | All indicators in one place, including beta |
| **Alerts** | What was flagged, the value, and why |
| **Correlation** | Correlation matrix with color-scale formatting |
| **Price History** | Normalized price movement (day 1 = 100) + line chart |
| **Pivot Data** | Long-format data (with Year/Month columns) ready as a source for your own pivot tables |

All data sheets are proper **Excel Tables** (`tbl_Stocks`, `tbl_Technical`, …), so they
expand automatically as the script adds new rows — pivot tables and formulas stay accurate.

---

## Running it

```bash
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt

cp .env.example .env              # Windows: copy .env.example .env
# open .env and paste in your key from https://twelvedata.com

python main.py
```

The first run pulls deep history (260 data points per instrument) so MA200 and every
other indicator work right away. Every following run only pulls a few fresh points.
The report is saved to `reports/market_report.xlsx`.

### Options

```bash
python main.py --no-fetch              # run against the existing history.csv, no network
python main.py --schedule              # keep running, refresh every 12h
python main.py --schedule --interval-hours 6
python main.py -v                      # more detailed logging
```

### Demo without internet / without a key

```bash
python demo_data.py               # generates synthetic history
python main.py --no-fetch         # builds the report from it
```

---

## Architecture

```
main.py            orchestrator + CLI + scheduling
config.py          instruments, thresholds, periods, paths
data_fetcher.py     Twelve Data API (behind the MarketDataProvider interface)
storage.py          history.csv - write, read, deduplication
analytics.py        all indicators (Pandas)
alerts.py            rule-based detection
excel_writer.py       report generation (openpyxl)
demo_data.py         synthetic data for testing without network/key
```

Two decisions worth explaining:

**Analytics run in Pandas, not in Excel formulas.** RSI and rolling windows in Excel
need helper columns and are hard to test; in Pandas they're a few lines and change with
one parameter. Excel formulas exist only on the Overview sheet, where their value is
in keeping the dashboard live.

**History is kept separate from the report.** `history.csv` is the single source of
truth; the Excel file is an output that can be deleted and regenerated any time.
Deduplication is by `(symbol, date)`, with newer data taking precedence.

**The data source is abstracted** (`MarketDataProvider`) — `TwelveDataProvider` is the
current implementation, but storage/analytics/alerts/excel_writer don't know or care
where the data comes from. Swapping or adding another source doesn't touch the rest
of the system.

---

## Limitations

- Twelve Data free tier: 8 calls/minute, 800/day. For 15 instruments refreshed every
  12h (2 runs a day), that's ~30 calls a day — well under the limit.
- FX pairs typically don't carry volume with most providers, so volume alerts don't
  fire for them.

## Technologies

Python · Pandas · NumPy · Requests · openpyxl · Twelve Data API
