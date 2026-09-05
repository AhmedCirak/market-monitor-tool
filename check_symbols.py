"""
Provjeri koji simboli iz config.py rade na tvom Twelve Data planu.

    python check_symbols.py

Povlaci samo po 1 tacku po instrumentu (jeftino), pa javi sta prolazi
a sta ne. Korisno prije prvog punog runa, ili kad dodas novi instrument.
"""

import logging

import config
from data_fetcher import FetchError, TwelveDataProvider

logging.basicConfig(level=logging.ERROR, format="%(message)s")


def main() -> None:
    if not config.TWELVE_DATA_API_KEY:
        print("Nema API kljuca. Napravi .env fajl i upisi TWELVE_DATA_API_KEY.")
        return

    print(f"Provjeravam {len(config.INSTRUMENTS)} simbola...\n")
    provider = TwelveDataProvider()

    ok, failed = [], []
    for instrument in config.INSTRUMENTS:
        symbol = instrument["symbol"]
        try:
            df = provider.fetch_instrument(instrument, output_size=1)
            price = df["close"].iloc[-1]
            date = df["date"].iloc[-1].date()
            print(f"  OK    {symbol:<10} {date}  close={price:,.4f}")
            ok.append(symbol)
        except FetchError as exc:
            print(f"  FAIL  {symbol:<10} {exc}")
            failed.append(symbol)

    print(f"\nRadi: {len(ok)}/{len(config.INSTRUMENTS)}")
    if failed:
        print(f"Ne radi: {', '.join(failed)}")
        print("\nIzbaci ili zamijeni ove simbole u config.py -> INSTRUMENTS.")
    else:
        print("Svi simboli rade. Mozes pokrenuti: python main.py")


if __name__ == "__main__":
    main()
