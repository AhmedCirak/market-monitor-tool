"""
Centralna konfiguracija projekta.

Sve sto se mijenja - instrumenti, pragovi, periodi, putanje - zivi ovdje,
tako da ostali moduli ne moraju biti dirani.
"""

from pathlib import Path
import os

from dotenv import load_dotenv

load_dotenv()

# --------------------------------------------------------------------------
# API - Twelve Data
# --------------------------------------------------------------------------
TWELVE_DATA_API_KEY = os.getenv("TWELVE_DATA_API_KEY", "")
TWELVE_DATA_BASE_URL = "https://api.twelvedata.com"

# Free tier: 8 poziva/minut, 800/dan.
API_CALLS_PER_MINUTE = 8
API_REQUEST_TIMEOUT = 20
API_MAX_RETRIES = 3

# --------------------------------------------------------------------------
# Putanje
# --------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
REPORTS_DIR = BASE_DIR / "reports"
LOGS_DIR = BASE_DIR / "logs"

HISTORY_FILE = DATA_DIR / "history.csv"
REPORT_FILE = REPORTS_DIR / "market_report.xlsx"
LOG_FILE = LOGS_DIR / "run.log"

for _folder in (DATA_DIR, REPORTS_DIR, LOGS_DIR):
    _folder.mkdir(parents=True, exist_ok=True)

# --------------------------------------------------------------------------
# Instrumenti
# --------------------------------------------------------------------------
# "symbol"   -> kako ga Twelve Data zove
# "name"     -> citljivo ime u reportu
# "category" -> odredjuje u koji Excel sheet ide
# "has_volume" -> FX i plemeniti metali obicno nemaju volume
INSTRUMENTS = [
    # --- Dionice ---
    {"symbol": "AAPL", "name": "Apple",           "category": "Stocks", "has_volume": True},
    {"symbol": "MSFT", "name": "Microsoft",       "category": "Stocks", "has_volume": True},
    {"symbol": "NVDA", "name": "NVIDIA",          "category": "Stocks", "has_volume": True},
    {"symbol": "AMZN", "name": "Amazon",          "category": "Stocks", "has_volume": True},
    {"symbol": "JPM",  "name": "JPMorgan Chase",  "category": "Stocks", "has_volume": True},
    # --- ETF / indeksi ---
    {"symbol": "SPY",  "name": "S&P 500 ETF",     "category": "Stocks", "has_volume": True},
    {"symbol": "QQQ",  "name": "Nasdaq 100 ETF",  "category": "Stocks", "has_volume": True},
    # --- Valute ---
    {"symbol": "EUR/USD", "name": "EUR / USD",    "category": "FX", "has_volume": False},
    {"symbol": "GBP/USD", "name": "GBP / USD",    "category": "FX", "has_volume": False},
    {"symbol": "USD/JPY", "name": "USD / JPY",    "category": "FX", "has_volume": False},
    {"symbol": "USD/CHF", "name": "USD / CHF",    "category": "FX", "has_volume": False},
    # --- Roba (preko ETF-ova) ---
    # Napomena: XAU/USD i XAG/USD (spot zlato/srebro) nisu dostupni na besplatnom
    # Twelve Data planu i vracaju HTTP 404. GLD i SLV su ETF-ovi koji drze fizicko
    # zlato/srebro, pa prate istu cijenu, a dolaze kao obicne americke dionice.
    {"symbol": "GLD",  "name": "Zlato (GLD ETF)",  "category": "Commodities", "has_volume": True},
    {"symbol": "SLV",  "name": "Srebro (SLV ETF)", "category": "Commodities", "has_volume": True},
    {"symbol": "USO",  "name": "Nafta (USO ETF)",  "category": "Commodities", "has_volume": True},
    # --- Kripto ---
    {"symbol": "BTC/USD", "name": "Bitcoin",  "category": "Commodities", "has_volume": True},
    {"symbol": "ETH/USD", "name": "Ethereum", "category": "Commodities", "has_volume": True},
]

BENCHMARK_SYMBOL = "SPY"

# --------------------------------------------------------------------------
# Parametri analitike
# --------------------------------------------------------------------------
MA_WINDOWS = (20, 50, 200)
RSI_PERIOD = 14
VOLATILITY_WINDOW = 20
VOLUME_AVG_WINDOW = 20
WEEK_52_WINDOW = 252
CORRELATION_WINDOW = 120
TRADING_DAYS_PER_YEAR = 252

CHART_HISTORY_DAYS = 180
PIVOT_HISTORY_DAYS = 365

# --------------------------------------------------------------------------
# Pragovi za alerte
# --------------------------------------------------------------------------
ALERT_PRICE_CHANGE_PCT = 3.0
ALERT_RSI_OVERBOUGHT = 70.0
ALERT_RSI_OVERSOLD = 30.0
ALERT_VOLUME_MULTIPLIER = 2.0
ALERT_NEAR_52W_HIGH_PCT = 2.0
ALERT_NEAR_52W_LOW_PCT = 2.0
ALERT_DRAWDOWN_PCT = 20.0

# --------------------------------------------------------------------------
# Preuzimanje
# --------------------------------------------------------------------------
DEFAULT_INTERVAL = "1day"
BACKFILL_SIZE = 260           # koliko tacaka povuci pri PRVOM runu (treba nam 200+ za MA200)
INCREMENTAL_SIZE = 5          # koliko povlacimo pri svakom sljedecem runu
KEEP_HISTORY_DAYS = 1500      # koliko dana historije cuvamo u CSV-u
