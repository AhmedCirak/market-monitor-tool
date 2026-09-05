"""
Glavni orchestrator.

    python main.py                 # povuci podatke, analiziraj, generisi report
    python main.py --no-fetch      # preskoci mrezu, radi nad postojecim history.csv
    python main.py --schedule      # ostani upaljen i pokreni se svaki dan u 22:30
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from datetime import datetime

import analytics
import alerts as alerts_module
import config
import data_fetcher
import excel_writer
import storage


def setup_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%H:%M:%S",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(config.LOG_FILE, encoding="utf-8"),
        ],
    )


def run(fetch: bool = True) -> str | None:
    log = logging.getLogger("main")
    started = time.time()
    log.info("=" * 62)
    log.info("Pokretanje: %s", datetime.now().strftime("%d.%m.%Y %H:%M:%S"))

    # 1. Podaci
    history = storage.load_history()
    if fetch:
        is_first_run = history.empty
        if is_first_run:
            log.info("Prvi run - povlacim duboku historiju (%d tacaka po instrumentu)...",
                     config.BACKFILL_SIZE)
            fresh = data_fetcher.fetch_all_backfill()
        else:
            log.info("Povlacim svjeze podatke sa Twelve Data (%d instrumenata)...",
                     len(config.INSTRUMENTS))
            fresh = data_fetcher.fetch_all()

        if fresh.empty and history.empty:
            log.error("Nema podataka - ni sa mreze ni iz historije. Prekidam.")
            return None
        history = storage.merge_history(history, fresh)
        storage.save_history(history)
    else:
        log.info("--no-fetch: koristim postojecu historiju.")
        if history.empty:
            log.error("history.csv je prazan. Pokreni bar jednom bez --no-fetch.")
            return None

    # 2. Analitika
    log.info("Racunam pokazatelje...")
    enriched = analytics.enrich(history)
    snapshot = analytics.latest_snapshot(enriched)
    correlation = analytics.correlation_matrix(enriched)
    betas = analytics.beta_vs_benchmark(enriched)

    # 3. Alerti
    log.info("Trazim zanimljiva kretanja...")
    alerts_df = alerts_module.detect(snapshot)

    # 4. Report
    log.info("Generisem Excel report...")
    path = excel_writer.build_report(snapshot, enriched, alerts_df, correlation, betas)

    log.info(
        "Gotovo za %.1fs | %d instrumenata | %d redova historije | %d alerta",
        time.time() - started, snapshot["symbol"].nunique(), len(history), len(alerts_df),
    )
    log.info("Report: %s", path)
    return path


def run_scheduled(interval_hours: float = 12.0) -> None:
    """Jednostavan scheduler bez vanjskih zavisnosti - pokrece se svakih N sati."""
    log = logging.getLogger("scheduler")
    log.info("Scheduler aktivan. Run svakih %.0fh.", interval_hours)

    while True:
        try:
            run(fetch=True)
        except Exception:
            log.exception("Run je pukao - nastavljam dalje.")
        log.info("Spavam %.0fh do sljedeceg runa...", interval_hours)
        time.sleep(interval_hours * 3600)


def main() -> None:
    parser = argparse.ArgumentParser(description="Financial Market Monitoring & Reporting Tool")
    parser.add_argument("--no-fetch", action="store_true",
                        help="preskoci mrezu, koristi postojeci history.csv")
    parser.add_argument("--schedule", action="store_true",
                        help="ostani upaljen i pokreni se periodicno")
    parser.add_argument("--interval-hours", type=float, default=12.0,
                        help="razmak izmedju runova sa --schedule (default: 12)")
    parser.add_argument("-v", "--verbose", action="store_true", help="detaljniji log")
    args = parser.parse_args()

    setup_logging(args.verbose)

    if args.schedule:
        run_scheduled(args.interval_hours)
    else:
        run(fetch=not args.no_fetch)


if __name__ == "__main__":
    main()
