# Financial Market Monitoring & Reporting Tool

Automatizovan sistem koji sam **prikuplja → analizira → cuva → izvjestava** o stanju
finansijskih trzista. Povlaci OHLCV podatke sa Twelve Data API-ja, racuna tehnicke i
statisticke pokazatelje u Pandasu, detektuje zanimljiva kretanja i generise
formatiran Excel report sa dashboardom, tabelama, korelacionom matricom i grafikonom.

Koristi **besplatan Twelve Data API kljuc** (bez kartice, 800 poziva/dan) —
registracija traje minut na [twelvedata.com](https://twelvedata.com).

---

## Sta prati

15 instrumenata u tri kategorije:

| Kategorija | Instrumenti |
|---|---|
| Stocks / ETF | Apple, Microsoft, NVIDIA, Amazon, JPMorgan, S&P 500 ETF, Nasdaq 100 ETF |
| FX | EUR/USD, GBP/USD, USD/JPY, USD/CHF |
| Commodities / Crypto | Zlato, Srebro, Bitcoin, Ethereum |

Lista se mijenja u `config.py` — dodavanje instrumenta je jedan red.

## Sta racuna

- Prinosi: 1d, 5d, 20d, YTD, 1Y
- Moving averages 20 / 50 / 200 + trend (MA50 vs MA200) + odnos cijene prema MA200
- RSI(14) po Wilderovom izgladjivanju
- MACD (12/26/9) i histogram
- Anualizovana volatilnost (20-dnevna)
- Volume i odnos prema 20-dnevnom prosjeku
- 52-nedjeljni maksimum/minimum i udaljenost cijene od njih
- Max drawdown od vrha
- Beta naspram S&P 500 ETF-a
- Korelaciona matrica dnevnih prinosa

## Alerti

Rule-based detekcija, pragovi u `config.py`:

| Tip | Uslov (default) |
|---|---|
| Promjena cijene | \|dnevna promjena\| ≥ 3% |
| RSI overbought / oversold | RSI ≥ 70 / RSI ≤ 30 |
| Volume anomalija | volume ≥ 2× 20-dnevni prosjek |
| Blizu 52w maksimuma / minimuma | unutar 2% od godisnjeg ekstrema |
| Drawdown | pad od vrha ≥ 20% |
| Test MA200 | cijena unutar 1% od MA200 |

## Excel report

| Sheet | Sadrzaj |
|---|---|
| **Market Overview** | Dashboard sa **pravim Excel formulama** — sazetak trzista, alerti, pregled po kategorijama, top 5 najjacih/najslabijih, alerti po tipu. Pragovi su u plavim celijama koje mozes mijenjati i formule se same preracunaju. |
| **Stocks / FX / Commodities** | Zadnje stanje po instrumentu: OHLC, prinosi, volume, 52w raspon, volatilnost, RSI, trend |
| **Technical Analysis** | Svi pokazatelji na jednom mjestu, sa betom |
| **Alerts** | Sta je flagovano, koja vrijednost i zasto |
| **Correlation** | Korelaciona matrica sa color scale formatiranjem |
| **Price History** | Normalizovano kretanje (prvi dan = 100) + linijski grafikon |
| **Pivot Data** | Long-format podaci (sa kolonama Godina/Mjesec) spremni kao izvor za tvoje pivot tabele |

Svi data sheetovi su prave **Excel Tabele** (`tbl_Stocks`, `tbl_Technical`, …), pa se
automatski prosiruju kad skripta doda nove redove — pivot tabele i formule ostaju tacne.

---

## Pokretanje

```bash
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt

cp .env.example .env              # Windows: copy .env.example .env
# otvori .env i upisi svoj kljuc sa https://twelvedata.com

python main.py
```

Prvi run povlaci duboku historiju (260 tacaka po instrumentu) da odmah rade MA200
i svi ostali indikatori. Svaki sljedeci run povlaci samo par svjezih tacaka.
Report se snima u `reports/market_report.xlsx`.

### Opcije

```bash
python main.py --no-fetch              # radi nad postojecim history.csv, bez mreze
python main.py --schedule              # ostani upaljen, pokreni se svakih 12h
python main.py --schedule --interval-hours 6
python main.py -v                      # detaljniji log
```

### Demo bez interneta / bez kljuca

```bash
python demo_data.py               # generise sinteticku historiju
python main.py --no-fetch         # napravi report nad njom
```

---

## Arhitektura

```
main.py            orchestrator + CLI + scheduling
config.py          instrumenti, pragovi, periodi, putanje
data_fetcher.py     Twelve Data API (iza MarketDataProvider interfejsa)
storage.py          history.csv - upis, citanje, deduplikacija
analytics.py        svi pokazatelji (Pandas)
alerts.py            rule-based detekcija
excel_writer.py       generisanje reporta (openpyxl)
demo_data.py         sinteticki podaci za testiranje bez mreze/kljuca
```

Dvije odluke vrijedne objasnjenja:

**Analitika ide u Pandasu, ne u Excel formulama.** RSI i rolling prozori u Excelu
traze pomocne kolone i tesko se testiraju; u Pandasu su par linija i mijenjaju se
jednim parametrom. Excel formule postoje samo na Overview sheetu, gdje im je vrijednost
u tome sto dashboard ostaje ziv.

**Historija je odvojena od reporta.** `history.csv` je jedini izvor istine; Excel je
izlaz koji se moze obrisati i regenerisati kad god. Deduplikacija ide po
`(symbol, date)` sa prednoscu novijem podatku.

**Izvor podataka je apstrahovan** (`MarketDataProvider`) — `TwelveDataProvider` je
trenutna implementacija, ali storage/analytics/alerts/excel_writer ne znaju niti mari
im odakle podaci dolaze. Zamjena ili dodavanje drugog izvora ne dira ostatak sistema.

---

## Ogranicenja

- Free tier Twelve Data: 8 poziva/minut, 800/dan. Za 15 instrumenata i refresh na 12h
  (2 runa dnevno) to je ~30 poziva dnevno — daleko ispod limita.
- FX parovi obicno nemaju volume kod vecine providera, pa volume alerti tamo ne rade.

## Tehnologije

Python · Pandas · NumPy · Requests · openpyxl · Twelve Data API
