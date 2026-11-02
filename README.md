# SEACEN Capital Flow Monitor Dashboard

A [Shiny for Python](https://shiny.posit.co/py/) dashboard that charts capital-flow, market-stress and macro indicators for SEACEN economies.

**What it covers:**
- **Markets:** VIX, trade and global policy uncertainty, exchange rates, stock indices, bond yields, sovereign CDS.
- **Rates and flows:** policy rates, non-resident portfolio flows, global liquidity.
- **Macro:** GDP growth, headline CPI inflation, GDP growth and inflation forecasts.
- **Stress:** the Financial Stress Index (FSI).

**Links:**
- **Live app:** <https://ahmadaizudeen.shinyapps.io/SEACEN-CFM-Dashboard/>
- **Source code:** <https://github.com/aizudeen11/SEACEN_Capital_Flows_Monitor_Dashboard>
- **What's new:** [UPDATE.md](UPDATE.md) (the **Updates** tab in the dashboard) lists every version and its changes.
- **Data pipeline:** [cfm_data/README.md](cfm_data/README.md) (the **Data pipeline** tab) explains how the data is built and uploaded.

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

## Using the dashboard

The **Dashboard** tab is a page of cards. Each card has its own chart, its own controls, and a footer with a note on what the chart shows and its source.

| Card | What it shows | Controls |
|---|---|---|
| Volatility Index (VIX) | Monthly average of the Cboe VIX | date range; event markers |
| Economic and Trade Policy Uncertainty Indices | Global economic policy uncertainty and trade policy uncertainty | date range; TPU only / GEPU only / both; event markers |
| Policy Rates | Central bank policy rates | date range; all economies or one of two selected groups; event markers |
| Exchange Rate Changes | Change in each currency against the U.S. dollar within each year (positive = appreciation) | years to show and year to sort by (latest 5 years) |
| Changes in Benchmark Stock Price Indices | % change of each benchmark index within each year | years to show and year to sort by (latest 5 years) |
| Changes in Sovereign Bond Yields | Change in each 5-year government bond yield within each year (percentage points) | years to show and year to sort by (latest 5 years) |
| Nonresident Portfolio Flows | Monthly portfolio equity and debt flows (USD billion) | date range; event markers |
| 5-Year Sovereign Credit Default Swap | CDS spreads (basis points) | date range; event markers |
| Quarterly GDP Growth | Quarterly real GDP growth for Asia and its subregions | quarters to show (latest 9) |
| GDP Growth and Inflation Outlook | Forecasts for last year, this year and next year | — |
| Monthly Inflation | CPI inflation, year-on-year, 3-month moving average, from January 2024. Laos and Sri Lanka are on the right-hand axis (dashed lines) | date range; economies |
| Financial Stress Indices | Financial stress for Asia and its subregions | event markers |
| Liquidity | BIS global liquidity indicator: growth of banks' international claims worldwide (year-on-year %) | — |

An unfinished year is labelled `YTD` (year-to-date), for example `YTD2026`.

**Sidebar**
- **Data last refreshed:** when the data was last updated, and the dashboard's version number.
- **Viewing tip:** the layout is designed for a large monitor. On a laptop, zoom the page out (Ctrl and − on Windows, ⌘ and − on a Mac; Ctrl/⌘ and 0 resets it).
- **Download figures (Excel):** a workbook with one sheet per card, holding the data each card is showing (after its filters) and an Excel chart you can edit. Any active event markers are listed under the data. A note above the button says whether every chart is included.
- **Economic event markers:** each chart with a **Show event markers** button can show events as dashed lines (single dates) or shaded bands (periods). You can add your own markers (date and label) or clear them; **Event details** lists every marker.
- **Contact and disclaimer.**

**Other tabs:** **About** (this page), **Updates** (version history) and **Data pipeline** (how the data is built).

---

The rest of this page is for people who run, maintain or deploy the dashboard.

## Contents

- [Project layout](#project-layout)
- [Quick start](#quick-start)
- [How data flows](#how-data-flows)
- [Updating the data](#updating-the-data)
- [Deploying](#deploying)
- [Git and files not in the repository](#git-and-files-not-in-the-repository)
- [Known issues](#known-issues)

## Project layout

```
├── main.py               # the Shiny app (UI + server)
├── processed_data.py     # loads the tables the app uses (dfs1 = live, dfs = local dev)
├── excel_export.py       # builds the "Download figures (Excel)" workbook
├── requirements.txt      # the app's packages
├── README.md             # this file (the About tab)
├── UPDATE.md             # version history (the Updates tab)
├── www/                  # images served by the app: the SEACEN symbol (seacen-icon.png), used as favicon and in the navbar
├── test (1).xlsx         # offline fallback copy of the main Google Sheet   [not in git]
├── .cache/               # one-day download cache                            [not in git]
├── rsconnect-python/     # shinyapps.io deployment record                    [not in git]
└── cfm_data/             # data pipeline: build tables, upload to Google Sheets
    ├── cfm_data.py           # builds every table
    ├── data download.py      # uploads the tables to Google Sheets
    ├── README.md             # pipeline documentation (the Data pipeline tab)
    ├── requirements.txt      # the pipeline's packages (not needed by the app)
    ├── .env.example          # template for .env
    ├── .env                  # sheet IDs and key-file path                  [not in git]
    ├── credential.json       # Google service-account key                   [not in git]
    ├── raw data for CFM.xlsx # CEIC/Haver/FocusEconomics download            [not in git]
    └── 00 dlx.xlsx           # FSI inputs                                     [not in git]
```

## Quick start

You need **Python 3.11 or later** (pandas 3.1 requires it). From the project folder:

```bash
python -m venv .venv
# Windows (PowerShell):  .venv\Scripts\Activate.ps1
# macOS / Linux:         source .venv/bin/activate

pip install -r requirements.txt
shiny run --reload --launch-browser main.py     # opens http://127.0.0.1:8000
```

- **Internet access:** the app downloads its data from Google Sheets and other websites when it starts. If a download fails, it uses the last cached copy in `.cache/`; for the main tables it falls back to `test (1).xlsx` if there's no cached copy.
- **Cloud-synced folders:** if `pip install` fails with long-path errors (common inside OneDrive and similar folders), create the virtual environment in a short path outside the synced folder and activate it from there.
- **No build step:** there are no tests, linters or build steps.

## How data flows

```
raw data for CFM.xlsx ─┐
00 dlx.xlsx ───────────┤                             ┌─► Google Sheet (main) ─► published link ───────────┐
IMF / BIS / CBOE / IIF ┴─► cfm_data/cfm_data.py ─► data download.py                                      ├─► processed_data.dfs1() ─► main.py
                                                     └─► Google Sheet (FSI) ─► published link ───────────┤
                              TPU / GEPU (policyuncertainty.com) ────────────────────────────────────────┘
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
| 8 | `fsi` | published FSI Google Sheet |
| 9 | `stock_price_index` | main Google Sheet |
| 10 | `sovereign_bond_yields` | main Google Sheet |
| 11 | `capital_flows` | main Google Sheet |
| 12 | `tpu` (Trade Policy Uncertainty) | policyuncertainty.com (the data file itself is hosted on matteoiacoviello.com; `TPU_URL` in `processed_data.py`) |
| 13 | `economics_event` (event markers, maintained by hand) | main Google Sheet |
| 14 | `gepu` (Global Economic Policy Uncertainty) | policyuncertainty.com |
| 15 | `gdp_growth_forecast` (Region, Year, Value) | main Google Sheet, sheet `gdp_growth_forecast` |
| 16 | `inflation_forecast` (Region, Year, Value) | main Google Sheet, sheet `inflation_forecast` |
| 17 | `cpi_3mma` (Date, Region, Value) | main Google Sheet, sheet `cpi_yoy_3mma` |
| 18 | `last_updated`: time of the last upload (a timestamp, not a table; `None` if unknown) | main Google Sheet, sheet `today` |
| 19 | `upload_log`: one row per table with its size, latest period and upload time (`None` if unavailable) | main Google Sheet, sheet `today` |
| 20 | `show_credit`: whether the navbar shows the "Developed by" credit | main Google Sheet, sheet `to_show`, column `developed` |
| 21 | `excel_full`: whether the Excel download includes the CEIC/Haver-based charts | main Google Sheet, sheet `to_show`, column `excel_use` |

The positions matter: `main.py` uses `all_df[i]` by number. **Add new tables at the end** of the tuple in [processed_data.py](processed_data.py), in both `dfs()` and `dfs1()`; never insert in the middle.

- **Loader functions.** `dfs()` is a development loader that reads a local `test.xlsx` and `financial stress index.xlsx` instead of the Google Sheets. To use it, change `all_df = dfs1()` to `dfs()` in `main.py`.
- **When data is read.** Both loaders run once, when the app starts. Downloads are also cached on disk in `.cache/` for a day; if a download fails, the last cached copy is used.
- **Google Sheets.** `dfs1()` downloads the *published* copies of the main and FSI Google Sheets (`MAIN_URL` and `FSI_URL` in `processed_data.py`), the sheets `data download.py` uploads to. If the main sheet can't be downloaded and there's no cached copy, the app falls back to `test (1).xlsx`, and the sidebar date shows "unknown".
- **Refresh date.** The sidebar's **Data last refreshed** comes from the sheet's `today` worksheet, which the uploader stamps on every run.
- **Version number.** The sidebar's version comes from the first `## <version>` heading in [UPDATE.md](UPDATE.md).
- **Display switches (`to_show` worksheet).** Row 1 holds the column names, row 2 a TRUE or FALSE under each. They're read by `_switches()` in `processed_data.py` when the app starts, like the data.

  | Column | FALSE | TRUE |
  |---|---|---|
  | `developed` | hides the "Developed by …" link in the navbar | shows it |
  | `excel_use` | the Excel download leaves out Policy Rate, 5-Year Sovereign CDS, Monthly Inflation, and GDP Growth and Inflation Outlook (CEIC/Haver data; the list is `EXCEL_RESTRICTED` in `main.py`) | the download includes every chart |

  - **Missing values:** a missing worksheet, column or value counts as TRUE.
  - **Single-cell layout:** a worksheet with just one TRUE/FALSE cell in A1, with no column names, works as the `developed` switch.
- **Names in the charts.** `_chart_names()` in `processed_data.py` applies `CHART_NAMES` as each table loads, so every chart and legend uses one spelling:
  - `South Korea` → **Korea** and `Hong Kong SAR (China)` → **Hong Kong**. The Google Sheets keep the source spellings; only the charts (and the Excel download, which uses the same data) change. Card footers keep their own wording (e.g. "Hong Kong, China").
  - `Asia Advanced Economies` and `Advanced Asian Economies` → **Asian Advanced Economies**.
- **Wide sheets.** The forecast and CPI sheets are stored wide: `Region` plus one column per year or month. `_forecast_long()` and `_monthly_long()` turn them into long tables. They accept month and year headers whether saved as text (`2026-09`, `2026`) or as dates and numbers.
- **README tabs.** `doc_page()` in `main.py` renders `README.md`, `UPDATE.md` and `cfm_data/README.md` as tabs when the app starts.
  - **Live values:** values wrapped in `<!--live:…-->` comment markers (table sizes, latest periods, the FSI cutoff, upload dates) are filled in from the upload log.
  - **Links:** links to other project files show as plain text, since the app doesn't serve those files.
  - **Editing:** after editing one of these files, restart the app to see the change.

## Updating the data

The full data-refresh process is in [cfm_data/README.md](cfm_data/README.md). In short, with the virtual environment active:

1. **Refresh the source workbooks.** Refresh `cfm_data/raw data for CFM.xlsx` and `cfm_data/00 dlx.xlsx` from CEIC/Haver/FocusEconomics.
2. **FSI cutoff (automatic).** The FSI runs to the latest month for which every market input (yields, stock indices, FX) has data. Reserves (ORA) may lag and are carried forward. The run prints the month it chose. To force a different month, pass `--cutoff YYYY-MM` or set `FSI_CUTOFF` in `cfm_data/cfm_data.py`.
3. **Build and upload.**
   ```bash
   cd cfm_data
   python "data download.py" --dry-run    # check every table builds
   python "data download.py"              # upload to both Google Sheets
   ```
   - **Setup:** this needs `cfm_data/.env` and a Google service-account key; see [cfm_data/README.md](cfm_data/README.md#uploading-to-google-sheets).
   - **Timing:** the upload is paced to stay under Google's request limit, so it takes a couple of minutes.
   - **The upload log:** it also refreshes the `today` worksheet, which drives the sidebar date and the live values in the **Data pipeline** tab.
   - **If it stops partway:** run the `--only …` command it prints.
4. **Let the app pick it up.** The app reads the published sheets when it starts and keeps a copy for a day.
   - **Locally:** delete the files in `.cache/` and restart the app.
   - **On shinyapps.io:** click **Restart** on the app's page. An idle app also stops by itself and restarts on the next visit. If a restart still shows the old data, the server's cached copy is less than a day old: redeploy, or wait for it to expire.
   - **Published copy:** Google can take a few minutes to update a published sheet after an upload or edit. In the sheet, **File → Share → Publish to web** should have **Automatically republish when changes are made** ticked.
   - **Check:** the sidebar's **Data last refreshed** shows the upload time.
   - **Required sheets:** the main sheet must contain every sheet the app reads: `vix_data`, `policy_rate1`, `fx`, `cds`, `liquidity`, `gdp_growth`, `stock_price_index`, `sovereign_bond_yields`, `capital_flows`, `economics_event`, `gdp_growth_forecast`, `inflation_forecast` and `cpi_yoy_3mma`. A missing sheet stops the app at startup. `today` (upload log) and `to_show` (display switches) are optional.
   - **Offline copy (optional):** download the main sheet as `.xlsx` (File → Download → Microsoft Excel) and save it as `test (1).xlsx`, the app's fallback when Google can't be reached.

## Deploying

The app is published on [shinyapps.io](https://www.shinyapps.io) with [rsconnect-python](https://docs.posit.co/rsconnect-python/). Deploying uploads files **from your computer**, not from GitHub. Run these with the virtual environment active, from the project folder.

1. **Install the deploy tool:**
   ```bash
   pip install rsconnect-python
   ```
   If your computer blocks the `rsconnect` program (e.g. "Access is denied"), run it through Python instead: replace `rsconnect` with `python -m rsconnect.main` in the commands below.
2. **Connect your shinyapps.io account** (once per computer). On shinyapps.io go to **Account → Tokens**, click **Show** on a token (or **Add Token**), then **Show secret**, and copy the token and secret into:
   ```bash
   rsconnect add --account <ACCOUNT> --name shinyapps --token <TOKEN> --secret <SECRET>
   ```
   The token is a password: don't commit it or paste it into files in the project.
3. **Check locally first.** Run the app (see [Quick start](#quick-start)) and look over every card, the Excel download and the tabs.
4. **Empty `.cache/`.** Cached files uploaded with the app look freshly downloaded on the server, so it would serve your local copy for up to a day.
5. **Deploy:**
   ```bash
   rsconnect deploy shiny . --name shinyapps --entrypoint main:app --title SEACEN-CFM-Dashboard \
       --override-python-version 3.12.2 \
       --exclude "cfm_data/.env" --exclude "cfm_data/*.xlsx" --exclude "cfm_data/*.json" \
       --exclude "cfm_data/*.ipynb" --exclude "cfm_data/*.py" --exclude "cfm_data/requirements.txt" \
       --exclude ".cache/*" --exclude ".claude/*" --exclude ".venv/*"
   ```
   (In PowerShell, end each continued line with a backtick `` ` `` instead of `\`.)
   - **Which app is updated:** after the first deploy, `rsconnect-python/` remembers the app, so later deploys from the same computer update it. To create a separate app, add `--new` and choose another `--title`. To update an existing app from another computer, add `--app-id <id>` (shown on the app's page in shinyapps.io).
   - **Python version:** shinyapps.io doesn't offer every Python version. `--override-python-version` asks for one it does offer, whatever version you run locally. Any version from 3.11 works with `requirements.txt`.
   - **Excludes:** the `--exclude` options keep the pipeline's secrets, raw workbooks and code, and local caches, off the server. Keep them on every deploy.
   - **Timing:** the upload takes a few minutes while shinyapps.io installs the packages. It prints the app's URL when it's done.
6. **Check the live app.** Open the URL and check the sidebar's **Data last refreshed** date and version, a few charts and the Excel download. If it doesn't start, the error is under **Applications → (app) → Logs** on shinyapps.io.
7. **Instance settings (optional).** Under **Applications → (app) → Settings**:
   - **Instance idle timeout:** how long the app stays up without visitors. A shorter timeout means it restarts, and so reloads the Google Sheets, more often.
   - **Startup timeout:** raise it if the logs show the app being stopped while it downloads the sheets at startup.

**What the server needs:** `main.py`, `processed_data.py`, `excel_export.py`, `requirements.txt`, `README.md`, `UPDATE.md` and `cfm_data/README.md` (the three document tabs), the `www/` folder (favicon and navbar logo), plus `test (1).xlsx` as the offline fallback if you have it.

**After a data upload** no redeploy is needed: click **Restart** on the app's page in shinyapps.io.

**Releasing a new version:** add an entry at the top of [UPDATE.md](UPDATE.md), then deploy. The sidebar picks up the new version number.

## Git and files not in the repository

The code is on GitHub at <https://github.com/aizudeen11/SEACEN_Capital_Flows_Monitor_Dashboard>. To get a copy:

```bash
git clone https://github.com/aizudeen11/SEACEN_Capital_Flows_Monitor_Dashboard.git
```

[.gitignore](.gitignore) keeps these out of Git:

| Ignored | Why |
|---|---|
| `.env`, `.env.*` (except `.env.example`), `credential*.json`, `*service*account*.json` | secrets |
| `*.xlsx` (any folder) | data files, including `test (1).xlsx` and both `cfm_data` workbooks (licensed CEIC/Haver/FocusEconomics data) |
| `rsconnect-python/`, `manifest.json` | shinyapps.io deployment record and manifest (the manifest is regenerated on every deploy) |
| `__pycache__/`, `*.py[cod]`, `.cache/`, `.ipynb_checkpoints/` | Python, app and Jupyter caches |
| `cfm_data/testing.ipynb` | local scratch notebook |
| `.claude/`, `CLAUDE.local.md` | Claude Code local files |
| `.venv/`, `venv/` | virtual environments |

**A fresh clone needs these files, obtained separately:**
- **To run the app:** nothing. It reads the published Google Sheets. `test (1).xlsx` is an optional offline fallback.
- **To build and upload the data:** `cfm_data/raw data for CFM.xlsx` and `cfm_data/00 dlx.xlsx` (licensed data), `cfm_data/.env` (copy `.env.example`) and a Google service-account key with Editor access to both sheets.

## Known issues

- **Euro Area is missing from the policy-rate "Selected Region 1" view.** In `processed_data.py`, that subset list uses `Euro Area`, but the data says `European Union`, so the view leaves it out. `cfm_data.py` already uses the correct name.
- **Region names vary between the uploaded tables** (e.g. Korea / South Korea; ASEAN-5 / ASEAN5 / ASEAN-4). The charts show one spelling (see **Names in the charts** above), but the sheets and the pipeline keep the source names. Details are in [cfm_data/README.md](cfm_data/README.md#things-that-give-wrong-numbers-without-an-error).
- **Data caveats** (stale Thailand market caps in the FSI, IIF revisions, FX treating a missing month as 0, Vietnam CPI from two sources, and others) are listed in [cfm_data/README.md](cfm_data/README.md).
