# SEACEN Capital Flow Monitor Dashboard

A [Shiny for Python](https://shiny.posit.co/py/) dashboard that charts capital-flow, market-stress and macro indicators for SEACEN economies.

**What it covers:**
- **Markets:** VIX, trade and global policy uncertainty, exchange rates, stock indices, bond yields, sovereign CDS.
- **Rates and flows:** policy rates, non-resident portfolio flows, global liquidity.
- **Macro:** GDP growth, headline CPI inflation, GDP growth and inflation forecasts.
- **Stress:** the Financial Stress Index (FSI).

**Links:**
- **Live app:** <https://ahmadaizudeen.shinyapps.io/cfm-final/> (shinyapps.io, app id `14341527`)
- **Data pipeline:** [cfm_data/](cfm_data/) builds every table from the raw workbooks and public APIs, and uploads them to Google Sheets. See [cfm_data/README.md](cfm_data/README.md).

## Contact and disclaimer

For comments and questions, please contact:<br>
**SEACEN Research Support Unit**<br>
Email: [rsu@seacen.org](mailto:rsu@seacen.org)

The SEACEN Capital Flow Monitor Dashboard should not be reported as representing the views of SEACEN member central banks/monetary authorities. The views expressed in this dashboard are those of The SEACEN Centre and do not necessarily represent those of its member central banks/monetary authorities.

## Notes

- The SEACEN Centre recognises "China" as the People's Republic of China, "Hong Kong, China" as the Hong Kong SAR, China, and "Korea" as the Republic of Korea.
- USD or US$ refers to the U.S. dollar.
- IMF and national source data were accessed through the CEIC Database and Haver Analytics. Other sources are named under each chart.
- The data were last refreshed on <!--live:*:updated-->28 Oct 2026<!--/live-->. Each chart shows the latest period available at that time.
- Asian economies include the nineteen economies of the SEACEN member central banks and monetary authorities with available data: Brunei Darussalam; Cambodia; China; Hong Kong, China; India; Indonesia; Korea; Lao PDR; Malaysia; Mongolia; Myanmar; Nepal; Papua New Guinea; Philippines; Singapore; Sri Lanka; Chinese Taipei; Thailand and Vietnam.
- This dashboard also shows subregional groupings, whenever data are available:
  - **Asian Advanced Economies:** Hong Kong, China; Korea; Singapore; and Chinese Taipei.
  - **ASEAN5:** Indonesia, Malaysia, Philippines, Thailand, and Vietnam.
  - **Asian Emerging and Developing Market Economies (EDMEs):** Brunei Darussalam, Cambodia, Lao PDR, Mongolia, Myanmar, Nepal, Papua New Guinea, and Sri Lanka.

## Abbreviations

| Abbreviation | Meaning |
|---|---|
| ASEAN | Association of Southeast Asian Nations |
| ASEAN4 / ASEAN5 | Indonesia, Malaysia, Philippines and Thailand / plus Vietnam |
| BIS | Bank for International Settlements |
| Cboe | Chicago Board Options Exchange |
| CDS | Credit default swap |
| CPI | Consumer price index |
| EDMEs | Emerging and Developing Market Economies |
| FSI | Financial Stress Index |
| GDP | Gross domestic product |
| GEPU | Global Economic Policy Uncertainty index |
| IIF | Institute of International Finance |
| IMF | International Monetary Fund |
| p.a. | Per annum |
| PPP | Purchasing power parity |
| RHS | Right-hand side (axis) |
| SAR | Special Administrative Region |
| SEACEN | South East Asian Central Banks |
| TPU | Trade Policy Uncertainty index |
| USD / US$ | U.S. dollar |
| VIX | Cboe Volatility Index |
| YoY | Year-on-year |
| YTD | Year-to-date |

## Contents

- [Project layout](#project-layout)
- [Quick start](#quick-start)
- [How data flows](#how-data-flows)
- [The dashboard](#the-dashboard)
- [Updating the data](#updating-the-data)
- [Deploying](#deploying)
- [Git and files not in the repository](#git-and-files-not-in-the-repository)
- [Known issues](#known-issues)

## Project layout

```
cfm-final/
├── main.py               # the Shiny app (UI + server)
├── processed_data.py     # loads the tables the app uses (dfs1 = live, dfs = local dev)
├── excel_export.py       # builds the "Download Figures (Excel)" workbook
├── test (1).xlsx         # offline fallback copy of the main Google Sheet   [not in git]
├── requirements.txt      # the app's packages (installed by shinyapps.io on deploy)
├── manifest.json         # shinyapps.io deployment manifest
├── rsconnect-python/     # shinyapps.io deployment record                   [not in git]
├── .cache/               # one-day download cache (both sheets, TPU, GEPU)   [not in git]
├── .gitignore
├── README.md
└── cfm_data/             # data pipeline: build tables, upload to Google Sheets
    ├── cfm_data.py           # builds every table
    ├── data download.py      # uploads the tables to Google Sheets
    ├── README.md             # pipeline documentation
    ├── requirements.txt      # the pipeline's packages (not needed by the app)
    ├── .env                  # sheet IDs and key-file path                  [not in git]
    ├── .env.example          # template for .env
    ├── credential.json       # Google service-account key, add yourself     [not in git]
    ├── testing.ipynb         # the working notebook the pipeline came from   [not in git]
    ├── raw data for CFM.xlsx # CEIC/Haver/FocusEconomics download            [not in git]
    └── 00 dlx.xlsx           # FSI inputs                                     [not in git]
```

## Quick start

**Python.** Python isn't on PATH on this PC. The project environment is `C:\Users\AhmadAizudeen\.venvs\cfm-final` (Python 3.14, pandas 3.1). It sits outside OneDrive because the project path is too long for some package files.

```powershell
$py = "$env:USERPROFILE\.venvs\cfm-final\Scripts\python.exe"

# App packages (first time only); the data pipeline has its own in cfm_data\requirements.txt
& $py -m pip install -r requirements.txt

# Run the dashboard locally (reloads on file changes, opens a browser at http://127.0.0.1:8000)
& $py -m shiny run --reload --launch-browser main.py
```

The app downloads its data from Google Sheets and other websites, so it needs internet access. If a download fails, it uses the last cached copy; for the main tables it falls back to `test (1).xlsx` if no cached copy exists. There are no tests, linters or build steps.

## How data flows

```
raw data for CFM.xlsx ─┐
00 dlx.xlsx ───────────┤                             ┌─► Google Sheet (main) ─► published link ───────────┐
IMF / BIS / CBOE / IIF ┴─► cfm_data/cfm_data.py ─► data download.py                                      ├─► processed_data.dfs1() ─► main.py
                                                     └─► Google Sheet (FSI) ─► published link ───────────┤
                              TPU (matteoiacoviello.com) / GEPU (policyuncertainty.com) ─────────────────┘
```

`processed_data.dfs1()` is the loader the app uses (`all_df = dfs1()` in [main.py](main.py)). It returns 18 tables, plus the upload time, the upload log and the two display switches:

| `all_df[i]` | Table | Source in `dfs1()` |
|---|---|---|
| 0 | `vix_data` | main Google Sheet |
| 1–3 | `policy_rate1` (all), `policy_rate2`, `policy_rate3` (region subsets filtered from 1) | main Google Sheet |
| 4 | `fx` | main Google Sheet |
| 5 | `cds` | main Google Sheet |
| 6 | `liquidity` | main Google Sheet |
| 7 | `gdp_growth` | main Google Sheet |
| 8 | `fsi` | published FSI Google Sheet (live, cached 1 day) |
| 9 | `stock_price_index` | main Google Sheet |
| 10 | `sovereign_bond_yields` | main Google Sheet |
| 11 | `capital_flows` | main Google Sheet |
| 12 | `tpu` (Trade Policy Uncertainty) | matteoiacoviello.com (live, cached 1 day) |
| 13 | `economics_event` (event markers, maintained by hand) | main Google Sheet |
| 14 | `gepu` (Global Economic Policy Uncertainty) | policyuncertainty.com (live, cached 1 day) |
| 15 | `gdp_growth_forecast` (Region, Year, Value) | main Google Sheet, sheet `gdp_growth_forecast` |
| 16 | `inflation_forecast` (Region, Year, Value) | main Google Sheet, sheet `inflation_forecast` |
| 17 | `cpi_3mma` (Date, Region, Value) | main Google Sheet, sheet `cpi_yoy_3mma` |
| 18 | `last_updated`: time of the last upload (a timestamp, not a table; `None` if unknown) | main Google Sheet, sheet `today` |
| 19 | `upload_log`: one row per table with its size, latest period and upload time (`None` if unavailable) | main Google Sheet, sheet `today` |
| 20 | `show_credit`: whether the navbar shows the "Developed by" credit (`True` if the sheet is missing) | main Google Sheet, sheet `to_show`, column `developed` |
| 21 | `excel_full`: whether the Excel download includes the CEIC/Haver-based charts (`True` if the sheet is missing) | main Google Sheet, sheet `to_show`, column `excel_use` |

The positions matter: `main.py` uses `all_df[i]` by number. **Add new tables at the end** of the tuple in [processed_data.py](processed_data.py), in both `dfs()` and `dfs1()`; never insert in the middle.

- **Loader functions.** `dfs()` is a development loader that reads a local `test.xlsx` and `financial stress index.xlsx`, which aren't in the project. To use it, change `all_df = dfs1()` to `dfs()` in `main.py`.
- **Caching.** Both loaders are memoised for the life of the app process. Downloads are also cached on disk in `.cache/` for a day; if a download fails, the last cached copy is used.
- **Online main sheet.** `dfs1()` downloads the *published* copy of the main Google Sheet (`MAIN_URL` in `processed_data.py`), the sheet `data download.py` uploads to. Like the FSI, it's cached in `.cache/` for a day. If it can't be downloaded and there's no cached copy, the app falls back to the local `test (1).xlsx` (the sidebar date then shows "unknown").
- **Refresh date.** The top of the sidebar shows **Data last refreshed**, read from the sheet's `today` worksheet, which the uploader stamps on every run.
- **Display switches (`to_show` worksheet).** Row 1 holds the column names, row 2 a TRUE or FALSE under each. They're read by `_switches()` in `processed_data.py` when the app starts, like the data.

  | Column | FALSE | TRUE |
  |---|---|---|
  | `developed` | hides the "Developed by …" link in the navbar | shows it |
  | `excel_use` | the Excel download leaves out Policy Rate, 5-Year Sovereign CDS, Headline Inflation, and GDP Growth and Inflation Outlook (CEIC/Haver data; the list is `EXCEL_RESTRICTED` in `main.py`) | the download includes every chart |

  - **Missing values:** a missing worksheet, column or value counts as TRUE.
  - **Old layout:** a worksheet with just one TRUE/FALSE cell in A1, with no column names, still works as the `developed` switch.
- **Names in the charts.** `_chart_names()` in `processed_data.py` applies `CHART_NAMES` as each table loads, so every chart and legend uses one spelling:
  - `South Korea` → **Korea** and `Hong Kong SAR (China)` → **Hong Kong**. The Google Sheets keep the source spellings; only the charts (and the Excel download, which uses the same data) change. Card footers keep their own wording (e.g. "Hong Kong, China").
  - The group of Hong Kong, Korea, Singapore and Chinese Taipei is **Asian Advanced Economies**. Sheets uploaded before that rename call it "Asia Advanced Economies" (GDP, FSI) or "Advanced Asian Economies" (forecasts); those are mapped too, so older uploads still work.
  - Chinese Taipei is already spelled that way in the sheets.
- **Wide sheets.** The three newest sheets are stored wide: `Region` plus one column per year or month, the layout `data download.py` uploads. `_forecast_long()` and `_monthly_long()` turn them into long tables. They accept month and year headers whether saved as text (`2026-09`, `2026`) or as dates and numbers (as a Google Sheets export gives).

## The dashboard

A single page of cards, each with its own Plotly chart and controls:

| Card | Controls | Data |
|---|---|---|
| VIX Liquidity | date range; event markers | `all_df[0]` |
| Trade & Global Policy Uncertainty | date range; TPU only / GEPU only / Both; event markers | `all_df[12]`, `[14]` |
| Policy Rate | date range; All Region / Selected Region 1 / Selected Region 2; event markers | `all_df[1–3]` |
| Exchange Rate Changes | years to show; year to sort by (latest 5 years only, `CHANGE_YEARS_SHOWN` in `main.py`) | `all_df[4]` |
| Benchmark Stock Price Indices | years to show; year to sort by (latest 5 years only) | `all_df[9]` |
| Benchmark Sovereign Bond Yields | years to show; year to sort by (latest 5 years only) | `all_df[10]` |
| Nonresident Portfolio Flows | date range; event markers | `all_df[11]` |
| 5-Year Sovereign Credit Default Swap | date range; event markers | `all_df[5]` |
| GDP Growth | quarters to show (latest 9 only, `GDP_QUARTERS_SHOWN` in `main.py`) | `all_df[7]` |
| GDP Growth and Inflation Outlook | — | `all_df[15]`, `[16]` |
| Headline Inflation | date range (from Jan 2024, `CPI_START` in `main.py`); economies | `all_df[17]` |
| Financial Stress Index | event markers | `all_df[8]` |
| Liquidity | — | `all_df[6]` |

**The two newest cards**
- **GDP Growth and Inflation Outlook.** Two grouped bar charts side by side: GDP growth and inflation forecasts.
  - **Layout:** regions on the x-axis (World, Asia-Pacific, China, India, Asian Advanced Economies, ASEAN5, Asia EDMEs) and one bar per year (last, this and next year).
  - **Colours:** the year colours are set in `FORECAST_YEAR_COLORS`.
- **Headline Inflation.** A line chart of CPI inflation, year-on-year, as a 3-month moving average, for 17 economies.
  - **Right-hand axis:** economies with much higher inflation are plotted on the right axis as dashed lines, labelled "(RHS)". Today that's Laos and Sri Lanka. The list is `CPI_RHS_REGIONS` at the top of `main.py`.
  - **Colours:** each economy's colour comes from `CPI_REGION_COLORS`.

**Tabs.** The navigation bar has three tabs:
- **Dashboard:** the cards above.
- **About:** this README.
- **Data pipeline:** `cfm_data/README.md`.

Both READMEs are rendered by `doc_page()` in `main.py` when the app starts. Values wrapped in `<!--live:…-->` comment markers in a README (table sizes, latest periods, the FSI cutoff, upload dates) are filled in from the upload log, so they stay current after every upload. Links to other project files show as plain text, since the app doesn't serve those files. After editing a README, restart the app to see the change.

**Card notes and sources.** Each card's footer shows a **Note** (what the chart shows, group membership, method) and, on its own line, the **Source**. They're built with `card_notes(note=..., source=...)` in `main.py`.
- **VIX:** source only.
- **GDP Growth and Inflation Outlook:** a separate, labelled note for each chart and one shared source.

**Axis titles.** Each chart names its unit on the value axis, for example "Policy rate (% p.a.)", "Basis points", "USD billion" or "% change, year-on-year", instead of a generic "Value".

**Features across the dashboard**
- **Event markers.**
  - Charts with a **Show event markers** button draw the events from the `economics_event` sheet. A `range` event is a shaded band; a `line` event is a dashed line.
  - The sidebar can add custom markers or clear them.
  - **Event details** lists every marker.
- **Download Figures (Excel)** ([excel_export.py](excel_export.py)) builds a workbook with one sheet per card. Each sheet has the data the card is currently showing, after its filters, plus a native Excel chart you can still edit. Any active event markers are listed under the data. Headline Inflation keeps Laos and Sri Lanka on a dashed right-hand axis, as in the dashboard.
  - **Which charts:** when `excel_use` in the `to_show` worksheet is FALSE, the CEIC/Haver-based charts are left out (see **Display switches** above).
  - **What users see:** above the button, a line says whether all charts are included, with a ? icon whose tooltip explains. The button itself has the same tooltip.
- **Viewing tip.** The top of the sidebar suggests zooming the page out (Ctrl/⌘ and −) on laptop screens, since the layout is designed for a large monitor.
- **Contact and disclaimer.** The bottom of the sidebar shows the Research Support Unit's email and a short disclaimer. The full disclaimer, notes and abbreviations are at the top of this README (the **About** tab).
- **Look:** `shinyswatch` Flatly theme; shared chart styling in `style_fig()`. Colours are kept consistent per region and per year across charts.

## Updating the data

The full data-refresh process lives in [cfm_data/](cfm_data/):

1. **Refresh the source workbooks.** Refresh `cfm_data/raw data for CFM.xlsx` and `cfm_data/00 dlx.xlsx` from CEIC/Haver/FocusEconomics.
2. **FSI cutoff (automatic).** The FSI runs to the latest month for which every market input (yields, stock indices, FX) has data. Reserves (ORA) may lag and are carried forward. The run prints the month it chose. To force a different month, pass `--cutoff 2026-09` or set `FSI_CUTOFF` in `cfm_data/cfm_data.py`.
3. **Build and upload.**
   ```powershell
   cd cfm_data
   & $py "data download.py" --dry-run    # check every table builds
   & $py "data download.py"              # upload to both Google Sheets
   ```
   - **Setup:** this needs `cfm_data/.env` and a Google service-account key; see [cfm_data/README.md](cfm_data/README.md#uploading-to-google-sheets).
   - **Timing:** the upload is paced to stay under Google's request limit, so it takes a couple of minutes.
   - **The upload log:** it also refreshes the `today` worksheet, which drives the sidebar date and the live values in the **Data pipeline** tab.
   - **If it stops partway:** run the `--only …` command it prints.
4. **Wait for the app to pick it up.** The app downloads the published main and FSI sheets and keeps them for a day.
   - **Timing:** a new upload shows on the next app start after the cached copy is a day old. To see it immediately, delete the files in `.cache/` and restart the app.
   - **On shinyapps.io:** the data and the live README values are read once, when the app process starts. An idle app is stopped and restarts on the next visit, so it usually picks up new data by itself; an app that never goes idle keeps its old data until it's restarted from the shinyapps.io dashboard.
   - **Check:** the sidebar's **Data last refreshed** shows the upload time.
   - **Required sheets:** the main sheet must contain every sheet the app reads: `vix_data`, `policy_rate1`, `fx`, `cds`, `liquidity`, `gdp_growth`, `stock_price_index`, `sovereign_bond_yields`, `capital_flows`, `economics_event`, `gdp_growth_forecast`, `inflation_forecast` and `cpi_yoy_3mma`. A missing sheet stops the app at startup. `today` (upload log) and `to_show` (display switches) are optional.
   - **Offline copy (optional):** download the main sheet as `.xlsx` (File → Download → Microsoft Excel) and save it over `test (1).xlsx`, the app's fallback when Google can't be reached.
5. **Check locally, then redeploy** (below).

## Deploying

The app is deployed to shinyapps.io with [rsconnect-python](https://docs.posit.co/rsconnect-python/):

```powershell
& $py -m pip install rsconnect-python
& "$env:USERPROFILE\.venvs\cfm-final\Scripts\rsconnect.exe" deploy shiny . --entrypoint main:app --title cfm-final `
    --exclude "cfm_data/.env" --exclude "cfm_data/*.xlsx" --exclude "cfm_data/*.json" `
    --exclude "cfm_data/*.ipynb" --exclude "cfm_data/*.py" --exclude "cfm_data/requirements.txt"
```

- **Updating the live app:** the deployment record in `rsconnect-python/cfm-final.json` makes later deploys from this PC update the same app.
- **From another computer:** that folder is git-ignored, so pass `--app-id 14341527` there; otherwise a new app is created.
- **Files the app needs on the server:** `main.py`, `processed_data.py`, `excel_export.py` and `requirements.txt`, plus `test (1).xlsx` as the offline fallback.
  - Deploying uploads files from your PC, not from Git, so `test (1).xlsx` is included even though Git ignores it.
  - **Empty `.cache/` before deploying.** Cached files uploaded with the app look freshly downloaded on the server, so it would serve your local copy for up to a day.
  - From `cfm_data/`, the app needs only `README.md`, shown in the **Data pipeline** tab.
  - **Keep the pipeline's secrets and data off the server.** `deploy shiny .` uploads the whole folder, including `cfm_data/.env`, the raw workbooks and the notebook. The `--exclude` options above leave them out; keep them whenever you deploy.

## Git and files not in the repository

[.gitignore](.gitignore) keeps these out of Git:

| Ignored | Why |
|---|---|
| `.env`, `.env.*` (except `.env.example`), `credential*.json`, `*service*account*.json` | secrets |
| `*.xlsx` (any folder) | data files, including `test (1).xlsx` and both `cfm_data` workbooks (licensed CEIC/Haver/FocusEconomics data) |
| `rsconnect-python/` | shinyapps.io deployment record |
| `__pycache__/`, `*.py[cod]`, `.cache/`, `.ipynb_checkpoints/` | Python, app and Jupyter caches |
| `cfm_data/testing.ipynb` | scratch notebook the pipeline was developed in |
| `.claude/`, `CLAUDE.local.md` | Claude Code local files |
| `.venv/`, `venv/` | virtual environments |

**Setting up a fresh clone needs these files, obtained separately:**
- **The workbooks:** `test (1).xlsx` (the app's offline fallback, optional), plus `cfm_data/raw data for CFM.xlsx` and `cfm_data/00 dlx.xlsx` (for the pipeline).
- **For uploading:** `cfm_data/.env` (copy `.env.example`) and the service-account key.

Git is installed per-user at `%LOCALAPPDATA%\Programs\Git`.

## Known issues

- **The new requirements haven't been deployed yet.** `requirements.txt` pins the app's packages at the versions tested locally: Python 3.14, pandas 3.1, shiny 1.8. The live app was last deployed in March 2025 with Python 3.12 and pandas 2.2. If shinyapps.io doesn't offer Python 3.14 yet, deploy from a Python version it supports. Check that the pinned versions install there.
- **`manifest.json` is from March 2025.** It doesn't list `excel_export.py` or `test (1).xlsx`. It's regenerated on deploy, but check the new one includes them.
- **Euro Area is missing from the policy-rate "Selected Region 1" view.** In `processed_data.py`, that subset list uses `Euro Area`, but the data says `European Union`, so the view silently leaves it out. (`Korea` and `Chinese Taipei` now match.) `cfm_data.py` already uses the correct name.
- **`cfm_data/testing.ipynb` points at old file locations.** Its cells still read `raw data for CFM.xlsx` and `for FSI/00 dlx.xlsx` from the project root, where they no longer are. It's kept locally for reference (git-ignored); `cfm_data.py` is the working version.
- **Region names vary between the uploaded tables** (e.g. Korea / South Korea; ASEAN-5 / ASEAN5 / ASEAN-4). The charts show Korea, Hong Kong and Chinese Taipei consistently (see **Names in the charts** above), but the sheets and the pipeline keep the source names. The details are in [cfm_data/README.md](cfm_data/README.md#things-that-give-wrong-numbers-without-an-error).
- **Data caveats** (stale Thailand market caps in the FSI, IIF revisions, FX treating a missing month as 0, Vietnam CPI from two sources, and others) are listed in [cfm_data/README.md](cfm_data/README.md).
- **Git repository location.** The `.git` folder is in the parent `NLP Project` folder, not in `cfm-final`, so `git status` there also shows unrelated projects.
