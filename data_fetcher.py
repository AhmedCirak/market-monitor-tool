"""
Povlacenje trzisnih podataka sa Twelve Data API-ja.

Provider je apstrahovan (MarketDataProvider), tako da se izvor podataka
kasnije moze zamijeniti bez diranja ostatka projekta (storage, analytics,
alerts i excel_writer ne znaju niti mari im odakle podaci dolaze).
"""

from __future__ import annotations

import logging
import time
from abc import ABC, abstractmethod
from typing import Any

import pandas as pd
import requests

import config

logger = logging.getLogger(__name__)

# "Ugovor" prema ostatku sistema - svaki provider vraca tacno ove kolone.
SCHEMA = ["date", "symbol", "name", "category", "open", "high", "low", "close", "volume"]


class FetchError(Exception):
    """Podaci za jedan instrument nisu mogli biti povuceni."""


def _redact(value: object, secret: str) -> str:
    """
    Izbaci API kljuc iz teksta greske.

    requests u poruku greske ubaci cijeli URL, ukljucujuci ?apikey=... —
    a te poruke zavrsavaju i u logs/run.log. Bez ovoga bi kljuc bio na disku
    u citljivom obliku.
    """
    text = str(value)
    if secret and secret in text:
        text = text.replace(secret, "***REDACTED***")
    return text


class RateLimiter:
    """Garantuje minimalni razmak izmedju poziva, da ne udarimo u minutni limit."""

    def __init__(self, calls_per_minute: int):
        self.min_interval = 60.0 / max(calls_per_minute, 1)
        self._last_call = 0.0

    def wait(self) -> None:
        elapsed = time.monotonic() - self._last_call
        if self._last_call and elapsed < self.min_interval:
            time.sleep(self.min_interval - elapsed)
        self._last_call = time.monotonic()


class MarketDataProvider(ABC):
    @abstractmethod
    def fetch_instrument(self, instrument: dict, output_size: int) -> pd.DataFrame:
        """Vrati DataFrame sa kolonama iz SCHEMA za jedan instrument."""


class TwelveDataProvider(MarketDataProvider):
    def __init__(self, api_key: str | None = None):
        self.api_key = api_key or config.TWELVE_DATA_API_KEY
        if not self.api_key:
            raise ValueError(
                "Nedostaje TWELVE_DATA_API_KEY. Kopiraj .env.example u .env "
                "i upisi svoj kljuc sa twelvedata.com."
            )
        self.base_url = config.TWELVE_DATA_BASE_URL
        self.limiter = RateLimiter(config.API_CALLS_PER_MINUTE)
        self.session = requests.Session()

    def _request(self, endpoint: str, params: dict[str, Any]) -> dict:
        params = {**params, "apikey": self.api_key}
        url = f"{self.base_url}/{endpoint}"

        last_error: Exception | None = None
        for attempt in range(1, config.API_MAX_RETRIES + 1):
            self.limiter.wait()
            try:
                response = self.session.get(url, params=params, timeout=config.API_REQUEST_TIMEOUT)
                response.raise_for_status()
                payload = response.json()
            except requests.HTTPError as exc:
                status = exc.response.status_code if exc.response is not None else None
                # 4xx (osim 429) se nece promijeniti ponavljanjem - odustajemo odmah.
                if status and 400 <= status < 500 and status != 429:
                    raise FetchError(
                        f"HTTP {status} - simbol vjerovatno nije dostupan na tvom planu "
                        f"ili ne postoji pod ovim nazivom."
                    ) from None
                last_error = exc
                wait = 2 ** attempt
                logger.warning("Pokusaj %d/%d nije uspio (HTTP %s). Cekam %ds.",
                               attempt, config.API_MAX_RETRIES, status, wait)
                time.sleep(wait)
                continue
            except (requests.RequestException, ValueError) as exc:
                last_error = exc
                wait = 2 ** attempt
                # _redact: da API kljuc nikad ne zavrsi u logu ni u run.log fajlu
                logger.warning("Pokusaj %d/%d nije uspio (%s). Cekam %ds.",
                               attempt, config.API_MAX_RETRIES, _redact(exc, self.api_key), wait)
                time.sleep(wait)
                continue

            # Twelve Data greske dolaze kao HTTP 200 sa status=error u tijelu.
            if isinstance(payload, dict) and payload.get("status") == "error":
                message = payload.get("message", "nepoznata greska")
                code = payload.get("code")
                if code == 429:
                    wait = 60
                    logger.warning("API limit dostignut. Cekam %ds.", wait)
                    time.sleep(wait)
                    last_error = FetchError(message)
                    continue
                raise FetchError(f"API greska ({code}): {message}")

            return payload

        raise FetchError(
            f"Neuspjelo nakon {config.API_MAX_RETRIES} pokusaja: "
            f"{_redact(last_error, self.api_key)}"
        )

    def fetch_instrument(self, instrument: dict, output_size: int) -> pd.DataFrame:
        symbol = instrument["symbol"]
        payload = self._request(
            "time_series",
            {
                "symbol": symbol,
                "interval": config.DEFAULT_INTERVAL,
                "outputsize": output_size,
                "order": "ASC",
            },
        )

        values = payload.get("values")
        if not values:
            raise FetchError(f"{symbol}: API nije vratio nijednu tacku podataka.")

        df = pd.DataFrame(values)
        df["date"] = pd.to_datetime(df["datetime"])
        df["symbol"] = symbol
        df["name"] = instrument["name"]
        df["category"] = instrument["category"]

        if "volume" not in df.columns:
            df["volume"] = pd.NA

        for col in ("open", "high", "low", "close", "volume"):
            df[col] = pd.to_numeric(df[col], errors="coerce")

        if not instrument.get("has_volume", True):
            df["volume"] = pd.NA

        df = df.dropna(subset=["date", "close"])
        return df.sort_values("date").reset_index(drop=True)[SCHEMA]


def fetch_all(
    provider: MarketDataProvider | None = None,
    instruments: list[dict] | None = None,
    output_size: int = config.INCREMENTAL_SIZE,
) -> pd.DataFrame:
    """
    Povuce sve instrumente. Jedan neuspjeh ne rusi run - greska se loguje,
    a ostali instrumenti se nastavljaju povlaciti.
    """
    provider = provider or TwelveDataProvider()
    instruments = instruments if instruments is not None else config.INSTRUMENTS

    frames: list[pd.DataFrame] = []
    failed: list[str] = []

    for instrument in instruments:
        symbol = instrument["symbol"]
        try:
            df = provider.fetch_instrument(instrument, output_size)
            frames.append(df)
            logger.info("OK   %-10s %4d redova  (%s -> %s)",
                        symbol, len(df), df["date"].min().date(), df["date"].max().date())
        except FetchError as exc:
            failed.append(symbol)
            logger.error("FAIL %-10s %s", symbol, exc)

    if failed:
        logger.warning("Neuspjesni instrumenti: %s", ", ".join(failed))

    if not frames:
        return pd.DataFrame(columns=SCHEMA)

    return pd.concat(frames, ignore_index=True)


def fetch_all_backfill(provider: MarketDataProvider | None = None) -> pd.DataFrame:
    """Prvi run - povuce duboku historiju (config.BACKFILL_SIZE tacaka) po instrumentu."""
    return fetch_all(provider, output_size=config.BACKFILL_SIZE)
