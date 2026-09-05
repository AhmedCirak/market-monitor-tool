"""
Lokalna historija podataka u CSV-u.

Namjerno je odvojena od Excel reporta: Excel je izlaz koji se moze
regenerisati kad god, a history.csv je jedini izvor istine o podacima.
"""

from __future__ import annotations

import logging

import pandas as pd

import config
from data_fetcher import SCHEMA

logger = logging.getLogger(__name__)


def load_history() -> pd.DataFrame:
    """Ucita historiju sa diska. Vrati prazan DataFrame ako fajl ne postoji."""
    if not config.HISTORY_FILE.exists():
        logger.info("history.csv ne postoji - krecemo od nule.")
        return pd.DataFrame(columns=SCHEMA)

    df = pd.read_csv(config.HISTORY_FILE, parse_dates=["date"])
    logger.info(
        "Ucitano %d redova historije za %d instrumenata.",
        len(df), df["symbol"].nunique(),
    )
    return df


def save_history(df: pd.DataFrame) -> None:
    df = df.sort_values(["symbol", "date"])
    df.to_csv(config.HISTORY_FILE, index=False)
    logger.info("Snimljeno %d redova u %s", len(df), config.HISTORY_FILE.name)


def merge_history(existing: pd.DataFrame, fresh: pd.DataFrame) -> pd.DataFrame:
    """
    Spoji novo povucene podatke sa postojecom historijom.

    Deduplikacija ide po (symbol, date), pri cemu novi podaci pobjedjuju -
    Stooq ponekad naknadno koriguje zadnji bar, pa zelimo najsvjeziju verziju.
    """
    if existing.empty:
        combined = fresh.copy()
    elif fresh.empty:
        combined = existing.copy()
    else:
        combined = pd.concat([existing, fresh], ignore_index=True)

    if combined.empty:
        return combined

    combined["date"] = pd.to_datetime(combined["date"])
    before = len(combined)
    combined = combined.drop_duplicates(subset=["symbol", "date"], keep="last")
    removed = before - len(combined)

    # Rezanje prestare historije po instrumentu.
    combined = (
        combined.sort_values(["symbol", "date"])
        .groupby("symbol", group_keys=False)
        .tail(config.KEEP_HISTORY_DAYS)
        .reset_index(drop=True)
    )

    logger.info(
        "Spajanje: %d novih redova, %d duplikata odbaceno, ukupno %d.",
        len(fresh), removed, len(combined),
    )
    return combined[SCHEMA]
