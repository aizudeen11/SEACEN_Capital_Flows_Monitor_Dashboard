# cfm_data — SEACEN Capital Flow Monitor data pipeline

`cfm_data.py` builds every table the Capital Flow Monitor dashboard needs. It reads two Excel workbooks and five public data sources (APIs and downloads). It was compiled from the working notebook (now `testing.ipynb` in this folder), with the exploratory cells removed. Every table it produces was checked against the notebook's output and matches exactly. The one exception is the policy-rate fix described under [Policy rates](#policy-rates).

## Contents

- [Quick start](#quick-start)
- [Uploading to Google Sheets](#uploading-to-google-sheets)
- [Inputs](#inputs)
- [Functions and outputs](#functions-and-outputs)
- [Settings to update](#settings-to-update)
- [Things that can crash it](#things-that-can-crash-it)
- [Things that give wrong numbers without an error](#things-that-give-wrong-numbers-without-an-error)
- [Method notes and open questions](#method-notes-and-open-questions)

## Quick start

**Python environment.** The project's environment is `C:\Users\AhmadAizudeen\.venvs\cfm-final` (Python 3.14, pandas 3.1). Python is not on PATH, so call it by its full path:

```powershell
$py = "$env:USERPROFILE\.venvs\cfm-final\Scripts\python.exe"

# Build everything and print a summary (uses the workbooks in cfm_data/)
& $py cfm_data\cfm_data.py

# Other workbook copies
& $py cfm_data\cfm_data.py --raw "C:\path\raw data for CFM.xlsx" --fsi "C:\path\00 dlx.xlsx"

# Override the automatic FSI cutoff
& $py cfm_data\cfm_data.py --cutoff 2026-09
```

**Packages:** `pandas`, `numpy`, `requests`, `openpyxl` and `arch` (needed for the FSI's GARCH step):

```powershell
& $py -m pip install pandas numpy requests openpyxl arch
```

**From other code** (e.g. `processed_data.py`), import the module and use the functions:

```python
from cfm_data.cfm_data import load_all, gdp_growth, cpi_inflation

tables = load_all()             # dict of every table, see below
yoy, yoy_3mma, gaps = cpi_inflation()
```

A full `load_all()` takes about a minute. Most of the time goes on the downloads and the FSI's rolling-beta loop.

## Uploading to Google Sheets

`data download.py` builds the tables with `cfm_data.py` and writes them to two Google Sheets.

**What the dashboard reads:** `processed_data.py` downloads the *published* copies of both Google Sheets (main and FSI) and caches them for a day. So after an upload, every chart updates on the next app start once its cached copy has expired (see [Upload problems](#upload-problems)). The local `test (1).xlsx` is used only if the main sheet can't be downloaded.

### One-time setup

1. **Packages:** everything both scripts need is listed in `requirements.txt`:
   ```powershell
   & $py -m pip install -r cfm_data\requirements.txt
   ```
2. **Service-account key:** put the Google service-account key file in `cfm_data/` as `credential.json`. To use another name or location, set `GOOGLE_CREDENTIALS_FILE` in `.env`.
3. **Sheet access:** share both Google Sheets with the service account's email address (the `client_email` in the key file), with **Editor** access.
4. **Settings:** check `cfm_data/.env`. It holds the key-file path and the sheet IDs. `.env.example` shows the format, for anyone setting this up fresh.

`.env` and any `credential*.json` file are listed in the project's `.gitignore`, so they're never committed. Only `.env.example` (no real values) is meant to be shared.

### Running it

```powershell
cd cfm_data
& $py "data download.py" --dry-run                  # build everything, upload nothing, show what would go where
& $py "data download.py"                            # build and upload everything
& $py "data download.py" --only fx cds fsi          # upload only these worksheets
& $py "data download.py" --xlsx backup.xlsx         # also save every table to a local Excel file
& $py "data download.py" --cutoff 2026-09           # override the automatic FSI cutoff
```

### What goes where

| Sheet (`.env`) | Worksheets |
|---|---|
| `SHEET_ID_MAIN` | `vix_data`, `policy_rate1`, `fx`, `cds`, `liquidity`, `gdp_growth`, `stock_price_index`, `sovereign_bond_yields`, `capital_flows`, `inflation_forecast`, `gdp_growth_forecast`, `cpi_yoy`, `cpi_yoy_3mma`, `today` (upload log, see below) |
| `SHEET_ID_FSI` | `fsi`, `yield_data_calc`, `stock_data_calc2`, `empi`, `financial_sector_beta`, `garch`, `fsi_missing` |

The worksheet names in `SHEET_ID_MAIN` are the ones `processed_data.py` reads. **Don't rename them** without changing the app too. The app builds its two policy-rate subsets from `policy_rate1`, so only the full table is uploaded. `SHEET_ID_SECTION2` is kept in `.env` from the old script but isn't uploaded yet.

### How uploading behaves

- **Build first:** every requested table is built before anything is uploaded, so a data error stops the run before Google Sheets is touched.
- **Existing worksheets are overwritten in place, then trimmed.** The new table is written over the old one, every cell included, so nothing old survives inside it. Then any leftover rows and columns are cut off.
  - A worksheet's link, its position and any formulas pointing at it keep working. (The old script deleted and re-created each worksheet, which broke those.)
  - Nothing is cleared first, so a failed request leaves that worksheet's **old data in place**, never an empty worksheet.
  - A worksheet that doesn't exist yet is created.
- **Other worksheets:** worksheets the script doesn't manage (e.g. `economics_event`) are never touched.
- **How values are written:**
  - Dates are written as `YYYY-MM-DD`, so Sheets stores them as real dates.
  - Wide tables get their row labels as the first column, and month headers as `YYYY-MM`.
  - Missing values become empty cells.
- **Pacing:** Google allows 60 requests per minute per user. The script counts its requests and pauses ("pausing …s to stay under Google's 50-per-minute limit") once 50 have gone out in the last minute, to leave headroom. A full upload therefore takes a couple of minutes. The limit is `REQUESTS_PER_MINUTE`.
- **Retries:** if Google still answers "too many requests" (HTTP 429) or with a server error, the script waits 65 seconds (long enough for Google's one-minute window to reset) and retries, up to 3 times.
- **If a run stops partway:** worksheets not yet updated keep their previous data. The script lists the ones it finished and prints the exact `--only …` command to upload the rest.
- **Upload log (`today` worksheet):** every run rewrites it with one row per table: `Table`, `Spreadsheet` (main or FSI), `Rows`, `Columns`, `Covers to` (latest period) and `Last updated`.
  - With `--only`, the rows of tables not uploaded keep their previous values. The uploaded rows get a new `Last updated`, so the sidebar's refresh time moves even if only one table changed.
  - It's written as plain text, so Sheets doesn't turn "Aug 2026" into a date.
  - The dashboard reads it for the sidebar's **Data last refreshed** time and for the sizes in its **Data pipeline** tab.
  - The `fsi` row's `Covers to` is the FSI cutoff that was used.

### Upload problems

| Message | Cause | Fix |
|---|---|---|
| `SHEET_ID_MAIN is not set` | `.env` is missing or the value is empty | Fill it in `cfm_data/.env` |
| `Service-account key file not found` | No `credential.json` in `cfm_data/` | Copy the key file there, or point `GOOGLE_CREDENTIALS_FILE` at it |
| `APIError: [403] The caller does not have permission` | The sheet isn't shared with the service account | Share it with the key's `client_email` as Editor |
| `APIError: [404] Requested entity was not found` | Wrong sheet ID | Copy the ID from the sheet's URL |
| `APIError: [429]` after all retries | Something else is using the same Google account's quota at the same time (pacing normally prevents this) | Wait a minute, then run the `--only …` command the script printed |
| Dashboard still shows old data | The app downloads the *published* sheets and caches them for a day in the project's `.cache/` folder; it also reads them only when it starts | Wait a day, or delete the files in `.cache/` and restart the app |
| Dashboard shows "unknown" refresh date | The main sheet couldn't be downloaded, so the app fell back to `test (1).xlsx` | Check the internet connection and that the main sheet is still published to the web |

## Inputs

### Workbooks

Both workbooks sit in `cfm_data/`, next to `cfm_data.py`. The default paths are found relative to that file, so no user-specific path is hardcoded and the folder can be moved as a whole.

```
cfm_data/
├── cfm_data.py
├── data download.py        # uploads the tables to Google Sheets
├── .env                    # private settings (git-ignored)
├── .env.example            # template for .env
├── credential.json         # Google service-account key (git-ignored, add yourself)
├── README.md
├── testing.ipynb           # the working notebook the code came from (reference only, git-ignored: its file paths are outdated)
├── raw data for CFM.xlsx   # RAW_PATH
└── 00 dlx.xlsx             # FSI_PATH
```

| Constant | File | Sheets used |
|---|---|---|
| `RAW_PATH` | `raw data for CFM.xlsx` | `GDP`, `Policy Rate`, `FX`, `CDS`, `Yields YTD`, `stock indices`, `Inflation forecast`, `GDP growth`, `Inflation`, `VN CPI` |
| `FSI_PATH` | `00 dlx.xlsx` | `00 dlx` (first sheet), `Yields2`, `Stock Price Indexes`, `FX`, `ORA` |

The `Metadata` sheet in `raw data for CFM.xlsx` isn't read. Every sheet is read **by name**, so the sheet order doesn't matter.

### Online sources

| Source | Used by | What |
|---|---|---|
| IMF WEO SDMX API (`api.imf.org`, dataflow `WEO/9.0.0`) | `gdp_growth`, `ppp_weighted_forecast` | Annual PPP GDP (`PPPGDP`), used as the weights for the regional averages |
| CBOE (`cdn.cboe.com`) | `vix_history` | Daily VIX history |
| BIS API (`stats.bis.org`) | `policy_rate`, `liquidity` | China policy rate; global liquidity indicator |
| IIF (`iif.com`) | `capital_flows` | Monthly EM portfolio flows workbook |

Everything is downloaded fresh on every run; nothing is cached.

## Functions and outputs

`load_all(raw_path, fsi_path, fsi_cutoff)` runs everything and returns a dict with these keys:

Sizes and latest periods are as of the latest upload (<!--live:*:updated-->28 Oct 2026<!--/live-->). In the dashboard's **Data pipeline** tab they're filled in from the upload log, so they stay current; in this file they show the values when it was last edited.

| Key | Function | Rows × columns | Covers to | Contents |
|---|---|---|---|---|
| `gdp_growth` | `gdp_growth()` | <!--live:gdp_growth:shape-->95 × 3<!--/live--> | <!--live:gdp_growth:covers-->3Q2026<!--/live--> | Quarter, Region, Value: quarterly real GDP growth (%) for Asia, China, India, ASEAN-5, Asian Advanced Economies |
| `vix` | `vix_history()` | <!--live:vix_data:shape-->442 × 2<!--/live--> | <!--live:vix_data:covers-->Oct 2026<!--/live--> | Date, CLOSE: monthly average VIX, **latest month first** |
| `policy_rate_all` / `_1` / `_2` | `policy_rate()` | <!--live:policy_rate1:shape-->811 × 3<!--/live--> | <!--live:policy_rate1:covers-->Oct 2026<!--/live--> | Region, Date, Value: policy rates from Jan 2022. `_1` and `_2` are the two region subsets (the size shown is the full table) |
| `fx` | `forex_exchange()` | <!--live:fx:shape-->77 × 3<!--/live--> | <!--live:fx:covers-->2026<!--/live--> | Year, Region, Value: % change of the exchange rate within each year (+ = local currency appreciated) |
| `cds` | `cds()` | <!--live:cds:shape-->1809 × 3<!--/live--> | <!--live:cds:covers-->Sep 2026<!--/live--> | Date, Region, Value: 5-year sovereign CDS (bps) |
| `liquidity` | `liquidity()` | <!--live:liquidity:shape-->41 × 2<!--/live--> | <!--live:liquidity:covers-->2026-Q1<!--/live--> | TIME_PERIOD, OBS_VALUE: BIS global liquidity (year-on-year %), last 11 years |
| `fsi` | `financial_stress_index()` | <!--live:fsi:shape-->200 × 13<!--/live--> | <!--live:fsi:covers-->Aug 2026<!--/live--> | Financial Stress Index per economy plus Asia, Asian Advanced Economies, ASEAN-4; Jan 2010 to the cutoff |
| `fsi_missing` | `fsi_missing_values()` | <!--live:fsi_missing:shape-->5 × 6<!--/live--> |  | Which FSI inputs were filled up to the cutoff, and with what |
| `bond_yields` | `sovereign_bond_yields()` | <!--live:sovereign_bond_yields:shape-->65 × 3<!--/live--> | <!--live:sovereign_bond_yields:covers-->YTD2026<!--/live--> | Year, Region, Value: change in the 5-year yield (percentage points) within each year |
| `stock_index` | `stock_price_index()` | <!--live:stock_price_index:shape-->77 × 3<!--/live--> | <!--live:stock_price_index:covers-->YTD2026<!--/live--> | Year, Region, Value: % change of the benchmark stock index within each year |
| `capital_flows` | `capital_flows()` | <!--live:capital_flows:shape-->321 × 3<!--/live--> | <!--live:capital_flows:covers-->Sep 2026<!--/live--> | Date, Portfolio Debt, Portfolio Equity: USD billion |
| `inflation_forecast` | `ppp_weighted_forecast("Inflation forecast")` | <!--live:inflation_forecast:shape-->7 × 3<!--/live--> | <!--live:inflation_forecast:covers-->2025-2027<!--/live--> | Groups × (last, this, next year): FocusEconomics inflation forecast |
| `gdp_growth_forecast` | `ppp_weighted_forecast("GDP growth")` | <!--live:gdp_growth_forecast:shape-->7 × 3<!--/live--> | <!--live:gdp_growth_forecast:covers-->2025-2027<!--/live--> | Same for GDP growth |
| `cpi_yoy` | `cpi_inflation()[0]` | <!--live:cpi_yoy:shape-->17 × 47<!--/live--> | <!--live:cpi_yoy:covers-->Sep 2026<!--/live--> | Economies × months: CPI inflation YoY (%) |
| `cpi_yoy_3mma` | `cpi_inflation()[1]` | <!--live:cpi_yoy_3mma:shape-->17 × 47<!--/live--> | <!--live:cpi_yoy_3mma:covers-->Sep 2026<!--/live--> | 3-month moving average of `cpi_yoy` |
| `cpi_gaps` | `cpi_inflation()[2]` | dict | | Months still missing a YoY, per economy |

### How each table is calculated

#### GDP growth (`gdp_growth`)

1. **Growth rates.** The `GDP` sheet holds levels for China, India, Malaysia, Mongolia, the Philippines, Chinese Taipei and Thailand. These become YoY growth (each quarter vs the same quarter a year earlier). For Hong Kong, Indonesia, South Korea, Singapore and Vietnam the sheet already holds growth rates.
2. **Weights.** Annual PPP GDP from the IMF WEO, used as the weight for all four quarters of that year, from 2022 to the current year.
3. **Groups.**
   - **Asia:** PPP-weighted average of all 12 economies.
   - **ASEAN-5:** Indonesia, Malaysia, Philippines, Thailand, Vietnam.
   - **Asian Advanced Economies:** Hong Kong, Korea, Singapore, Chinese Taipei.
   - **China and India:** their own growth.

#### Yearly change tables (`forex_exchange`, `sovereign_bond_yields`, `stock_price_index`)

- **Value:** the change from the first to the last month with data in each year.
  - Stock indices: % change.
  - Bond yields: difference in percentage points.
  - FX: % change of the rate, computed as first ÷ last, so a positive value means the local currency appreciated.
- **YTD label:** stocks, bonds and FX label an unfinished year `YTD2026`. A year counts as unfinished when the sheet has no December row for it yet.
- **Order:** regions are ordered by their value in the latest year, highest first. `Region` is an ordered categorical.
- **Stock indices sheet:** `stock indices` holds two downloads side by side. On the left is the CEIC block, with economies in the `Region` row. On the right is the Haver/MSCI Vietnam block, described in the `.DESC` row. Each block has its own date column, and the two are joined by month.

#### Policy rates

- **China:** the `Policy Rate` sheet has no China rate, so China comes from the BIS API.
- **Chinese Taipei:** `Taiwan (China)` is renamed `Chinese Taipei`, as in every table (see **Chinese Taipei** below).
- **Region subsets:** `POLICY_REGIONS_1` and `POLICY_REGIONS_2` use the names in the data: `European Union`, `South Korea` and `Chinese Taipei`.

  **Changed from the notebook:** the notebook filters with `Euro Area` and `Korea`. Those names don't exist in the data, so the Euro Area and Korea were silently left out of the "Selected Region 1/2" views. The dashboard builds these views itself, in `processed_data.py`. It renames `South Korea` to `Korea` when loading, so Korea is now included, but it still looks for `Euro Area` and needs the same fix for that.

#### Chinese Taipei

Every table names this economy `Chinese Taipei` (the constant `TAIWAN` in `cfm_data.py`). The source workbooks call it `Taiwan (China)`, `Taiwan` or `Taiwan, China`; each reader renames it as it reads the sheet. The column names used to look it up in the workbooks (e.g. in `FX_REGIONS`, `FSI_COUNTRIES` and `FLOW_COUNTRIES`) keep the source spelling.

#### Financial Stress Index (`financial_stress_index`)

For each economy, the FSI adds up five standardised components:

1. **Yield spread:** the 10-year yield minus the US 10-year yield.
2. **Stock market fall:** the negative log change of the stock index.
3. **EMPI (exchange market pressure):** FX depreciation minus the change in foreign reserves (ORA).
4. **Financial-sector beta:** a rolling 13-month beta of the financial-sector index on the market index. Values of 1 or below are set to 0.
5. **Stock volatility:** the GARCH(1,1) conditional variance of stock returns.

**Built series:** some economies have no ready-made financial-sector index, so one is built:
- **Singapore:** Straits Times growth × the financials' share of turnover.
- **Thailand:** banking, finance and insurance index growth, weighted by market cap.
- **Indonesia:** the old `(DC)` index until it stops, then the new one.

**Group averages:** "Asian Advanced Economies" (Hong Kong, Singapore, South Korea, Chinese Taipei) and "ASEAN-4" (Indonesia, Malaysia, Philippines, Thailand) are simple averages.

**Cutoff:** every input is cut at the cutoff month. Any series that stops before it is filled forward with its last value, and `fsi_missing_values()` lists each series filled that way.

- **Automatic by default** (`FSI_CUTOFF = None`, see `fsi_auto_cutoff()`). The cutoff is the latest month for which **every market input** (10-year yields, stock indices, financial-sector indices, FX) has data.
- **Allowed to lag:** reserves (ORA) don't hold the cutoff back; they're carried forward and reported.
- **Stale series:** a market series more than `FSI_STALE_MONTHS` (12) months behind the newest one (e.g. the Thai market caps that stopped in Jul 2023) is treated as stale. It's ignored when choosing the cutoff, then carried forward and reported.
- **Latest upload:** the cutoff was **<!--live:fsi:covers-->Aug 2026<!--/live-->**. In that data the Chinese Taipei and India yields ended in Aug 2026, while the other market inputs reached Sep 2026.
- **Override:** to force a month, set `FSI_CUTOFF = "2026-09"` or pass `--cutoff 2026-09`.

#### Forecasts (`ppp_weighted_forecast`)

- **Coverage:** reads a FocusEconomics sheet for last year, this year and next year.
- **Groups:** each group is a PPP-weighted average, using that year's IMF WEO PPP GDP. Membership is listed in `FORECAST_GROUPS`.
- **Missing forecasts:** an economy with no forecast for a year is left out, and the rest of its group is re-weighted.
- **Reuse:** works for any sheet with the same layout, e.g. `ppp_weighted_forecast("Current account")`.

#### CPI inflation (`cpi_inflation`)

1. **YoY:** each month vs the same month a year earlier, `((x / x 12 months ago) − 1) × 100`. Rows that are already year-on-year are handled differently:
   - Unit `Prev Year=100` (China): YoY = value − 100.
   - Unit `%` (Vietnam): the value is the YoY rate itself.
2. **Vietnam's earlier months:** Vietnam's months before the `Inflation` sheet's data starts (Nov 2025) are filled from the `VN CPI` sheet. Where both sheets have a month, the `Inflation` sheet wins.
3. **3-month average:** the 3-month moving average of the month and the two before it. All three months must have a value.
4. **Gaps:** months still missing a YoY after the first year are returned in `cpi_gaps`. In the Oct 2026 data these were Brunei and Nepal, Sep 2026.

## Settings to update

These constants at the top of each section control the output:

| Constant | Now | When to change |
|---|---|---|
| `FSI_CUTOFF` | `None` (automatic; latest upload: <!--live:fsi:covers-->Aug 2026<!--/live-->) | Only to override the automatic cutoff (or pass `--cutoff`) |
| `FSI_STALE_MONTHS` | 12 | How far behind a market series can be before it no longer holds back the automatic cutoff |
| `GDP_FIRST_YEAR` | 2022 | First year of quarterly GDP growth |
| `FX_FIRST_YEAR`, `STOCK_FIRST_YEAR`, `YIELDS_FIRST_YEAR` | 2020, 2020, 2021 | First year of the yearly-change tables |
| `POLICY_START` | 2022-01-01 | First month of the policy rate tables |
| `LIQUIDITY_YEARS` | 11 | How far back the BIS liquidity series goes |
| `FLOW_COUNTRIES` | 11 economies | Which IIF economies are summed into the portfolio flows |
| `FORECAST_GROUPS`, `ASEAN5_GDP`, `ADVANCED_ASIA_GDP` | — | Group membership |

The forecast years (last, this and next year) and the GDP weights' last year follow today's date automatically.

## Things that can crash it

| Symptom | Cause | Fix |
|---|---|---|
| `PermissionError: [Errno 13] Permission denied: '...xlsx'` | The workbook is open in Excel, or OneDrive is syncing it | Close it in Excel or wait for the sync. Alternatively, copy it and pass the copy with `--raw` / `--fsi` |
| `OSError: [Errno 22] Invalid argument` when reading a workbook | OneDrive "online-only" placeholder: the file isn't downloaded | Right-click the file → **Always keep on this device** |
| `ConnectionResetError`, `ChunkedEncodingError`, `ReadTimeout` | A download was cut off. This happened once with the IIF file during testing | Run again. There's no automatic retry |
| `ValueError: ... did not return an Excel file; the IIF link may have changed` | IIF moved or renamed its portfolio flows file | Find the new link on iif.com and update `IIF_FLOWS_URL` |
| `HTTPError 404` from `api.imf.org` | The IMF released a new WEO dataflow version and retired `9.0.0` | Update the version in `IMF_WEO_URL` |
| `KeyError` with a sheet name (e.g. `'Worksheet named 'GDP' not found'`) | A sheet was renamed | Rename it back, or update the `sheet_name=` in the function |
| `KeyError` with an economy or column name | A column header changed in a download (e.g. CEIC renamed a series) | Update the matching constant: `GDP_LEVEL_COLUMNS` / `GDP_RATE_COLUMNS`, `FX_REGIONS`, `FSI_STOCK_KEYS` / `FSI_FINANCIAL_KEYS`, `FSI_COUNTRIES`, `FORECAST_IMF_CODES` |
| `StopIteration` in `stock_price_index` or `_read_monthly_sheet` | The `.DESC` (Vietnam stock block) or `Region` header row is missing | Restore the header row in the sheet |
| `ValueError: could not convert string to float` in the FSI or forecasts | A row was added to or removed from the top of a sheet, so a text row is read as data | See the layout assumptions below |
| `ModuleNotFoundError: No module named 'arch'` | `arch` isn't installed in the Python being used | `pip install arch` |

### Sheet-layout assumptions (positions, not labels)

Most of the code finds columns by name, but these parts still rely on **positions**. Inserting or deleting rows or columns in these places will crash the code or, worse, read the wrong data:

| Where | Assumption |
|---|---|
| `GDP` sheet | Header on row 2; the first 13 columns are Date + the 12 economies |
| `Policy Rate` sheet | Columns 1, 3, 4 (description, unit, last update) are dropped; column 2 is the region; the rest are months |
| `CDS`, `Yields YTD` | Header on row 2 (Haver layout, dates like `Jan-2010`) |
| `Inflation forecast`, `GDP growth` | Header on row 2; the 3 rows after it are `.T1`, `.TN`, `.LSOURCE`; years written as two digits (`25`) |
| `00 dlx.xlsx` → `00 dlx` sheet | Columns 1, 2, 4 are Date, Indonesia 10Y, Philippines 10Y; 3 metadata rows |
| `00 dlx.xlsx` → `Yields2` | First 11 columns; 2 metadata rows |
| `00 dlx.xlsx` → `Stock Price Indexes`, `FX`, `ORA` | 3 (stocks) or 2 (FX, ORA) metadata rows under the header |
| IIF `Country Flows_nw` | Country names on row 4, Equity/Debt labels on row 5, data from row 6 |

## Things that give wrong numbers without an error

These are worth checking whenever the data is refreshed:

1. **Thailand market caps stopped in Jul 2023.** The three `SET: Market Capitalization` series (Banking, Finance and Securities, Insurance) haven't updated since Jul 2023. Their Jul 2023 values are carried forward to the cutoff, so Thailand's financial-sector index has used fixed sector weights since then. `fsi_missing` lists this every run. Check at the source whether these series were discontinued.
2. **FSI cutoff filling.** Any FSI input that's behind (e.g. China's reserves, ORA) repeats its last value up to the cutoff. Those months' FSI is partly carried-forward data. Check `fsi_missing` before publishing.
3. **FX counts a missing month as 0.** If an exchange rate is missing in the first or last month of a year, the "change" becomes ±100% or the full level. Bond yields and stocks skip missing months instead. FX was kept as in the notebook. Changing it to match the others is a two-line fix.
4. **Vietnam CPI comes from two sources.** Before Nov 2025 it's from `VN CPI`; from Nov 2025 it's from `Inflation`. If both are updated inconsistently, there can be a jump where they meet.
5. **Unit detection in CPI.** A row counts as already-YoY only if its unit is exactly `Prev Year=100` or `%`. If CEIC changes a unit label, the row would get the YoY formula applied to a YoY rate, giving nonsense numbers without an error.
6. **Partial group averages.**
   - Forecasts re-weight when an economy is missing, so a group's figure can be based on fewer economies than usual.
   - The FSI's Asian Advanced Economies and ASEAN-4 simple averages also skip missing economies.
7. **Region names differ between tables.** Matching across tables (e.g. in the dashboard's colour map) needs care:

   | Table | Korea | Hong Kong | ASEAN group |
   |---|---|---|---|
   | GDP growth | Korea | Hong Kong | ASEAN-5 |
   | Forecasts | Korea | Hong Kong | ASEAN5 |
   | FSI | South Korea | Hong Kong SAR (China) | ASEAN-4 |
   | Policy rate | South Korea | — | — |
   | Stock / FX / CPI | South Korea | Hong Kong SAR (China) | — |

   Chinese Taipei is the exception: it's spelled the same in every table (see below). The dashboard doesn't show these differences: as it loads the sheets, it renames `South Korea` to `Korea` and `Hong Kong SAR (China)` to `Hong Kong` (`CHART_NAMES` in `processed_data.py`), so its charts use one spelling.

8. **IIF revises its data.** Portfolio flows for past months change between downloads, sometimes by tens of billions. Compared with the app's Apr 2025 snapshot, 64 months of debt flows differ. Rerunning will change history.
9. **VIX is sorted latest-first**, unlike every other time series.

## Method notes and open questions

These are carried over unchanged from the original notebook code. They're worth a methodology review, but they aren't code bugs:

- **FSI beta:** the covariance divides by *n* (`bias=True`) but the variance by *n − 1* (`ddof=1`), so every beta is scaled down slightly.
- **FSI financial-index base:** Singapore's built index resets to 100 at Dec 2008, but Thailand's never resets.
- **FSI beta threshold:** months where the beta can't be computed (missing data) become 0 rather than staying blank, because NaN fails the `> 1` test.
- **FSI GARCH:** fitted with a constant mean and `lags=1`. A comment in the original notes that an AR mean may be intended (`## need to review maybe mean = 'AR'`).
- **CPI 3-month average:** needs all three months. Excel's `AVERAGE` would average whichever months exist; use `rolling(3, min_periods=1)` to match Excel.
- **GDP growth PPP weights:** each year's annual PPP GDP is applied to all four quarters (no interpolation within the year).
