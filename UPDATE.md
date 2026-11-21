# Updates

What has changed in the SEACEN Capital Flow Monitor Dashboard, newest version first.

**Version numbers** follow `MAJOR.MINOR.PATCH`:
- **MAJOR:** a redesign, or a change in how the data is produced or read.
- **MINOR:** new cards, charts or features.
- **PATCH:** fixes and small wording or layout changes.

The version shown in the dashboard's sidebar is read from the first entry below. To release a new version, add an entry at the top, in the same format (`## <version> — <date>`).

## 2.2.0 — 21 November 2026

### New
- **Capital Flows, quarterly IIP:** an economy that hasn't reported the latest quarter yet now keeps its latest position, instead of showing a gap and dropping out of the regional totals. For example, if Cambodia has data for 1Q2026 but not 2Q2026, its 2Q2026 position repeats 1Q2026. Chinese Taipei's latest annual position fills the quarters after it in the same way. The BoP flows are not filled.
- **Data pipeline, `iip_missing`:** a new worksheet in the Section 2 sheet lists every quarterly IIP item that was filled: the economy, the item, its last quarter with data, the value used, the quarters filled and how many. It works like `fsi_missing` for the Financial Stress Indices.
- **Capital Flows, central-bank data:** blanks in the IMF quarterly BoP and IIP data are now filled from the central banks' own data, where they have it. IMF values are never replaced. Where a central bank records BoP assets with the opposite sign to the IMF (e.g. Malaysia, Vietnam), its values are reversed to match. For example, Sri Lanka's IIP for 2Q2020 to 4Q2021, missing from the IMF data, now comes from its central bank. A new `cb_filled` worksheet lists every item filled. This replaces two unused functions that overwrote a fixed list of 2025 quarters.

### Changed
- **New address:** the dashboard is now published at <https://aizudeenjapri.shinyapps.io/seacen-cfm-dashboard/>, on a paid shinyapps.io account. The previous address, <https://ahmadaizudeen.shinyapps.io/seacen-cfm-dashboard/>, still works for now but no longer receives updates.
- **About tab (Deploying):** how to publish under another shinyapps.io account, and the instance settings of paid plans.
- **FSI cutoff:** the automatic cutoff of the Financial Stress Indices now follows the date of the data run, not the data: last month after the 7th of the month, otherwise the month before last (e.g. a run on 18 November covers up to October). Inputs without data for that month are carried forward and listed, as before. A month can still be set by hand with `--cutoff YYYY-MM`.
- **About tab (Updating the data) and Data pipeline tab:** after a data upload, wait about 5 minutes for Google to republish the sheets, then restart the app on shinyapps.io (**Applications → (the app) → Restart**) so the live dashboard shows the new data.
- **Capital Flows, regional groups:** as in the SEACEN Capital Flows Monitor, the groups now cover 16 economies. Asian EDMEs are Cambodia, Lao PDR, Mongolia, Nepal, and Sri Lanka, and the Asia-Pacific total (and the current account) leaves out Brunei Darussalam, Myanmar and Papua New Guinea. These three are still shown in the By economy view. The chart notes say so.

### Fixed
- **About tab (Deploying), Windows:** the deploy command's wildcard excludes (e.g. `cfm_data/*.xlsx`) didn't work on Windows: rsconnect expands them itself, so only the first matching file was excluded and the others were uploaded with the app. The About tab now gives a Windows form of the command that keeps them all off the server, a way to list what will be uploaded, and excludes `cfm_data/__pycache__`, `cfm_data/rsconnect-python` and `reference/` too.
- **Capital Flows, Korea in the regional totals:** Korea was left out of Asian Advanced Economies and the Asia-Pacific total, in every BoP and IIP table (half-yearly and annual). It is now included, so these totals are higher than before (e.g. Asia-Pacific IIP assets by about US$3 trillion).
- **Capital Flows, latest half-year of BoP:** the half-yearly BoP tables always dropped the latest quarter, so when the data ended at a 2Q or 4Q the latest half-year (e.g. 1H2026) was missing. Quarters are now paired by date. An economy that has reported only one of the two quarters is left out of that half-year, rather than counted with a one-quarter sum.
- **Quarterly GDP Growth, regional groups:** when a member economy had no data yet for a quarter (e.g. China and Vietnam in 2026), the Asia and ASEAN-5 averages counted it as 0% growth and fell sharply. A group is now left blank for that quarter.

## 2.1.1 — 4 November 2026

### Fixed
- **Capital Flows charts:** when one series had gaps (for example Mongolia before 2022 in FDI Assets by economy), the bars after it in the stack restarted from zero and were drawn over the others, so the bars fell short of the Total line. Gaps are now skipped, so the bars always add up to the total.

## 2.1.0 — 3 November 2026

### New
- **Section 2 data (BoP and IIP).** The data pipeline now also builds the Section 2 balance of payments and international investment position tables (quarterly, half-yearly and annual, by economy and by regional group) and uploads them to their own Google Sheet, with its own upload log.
  - It reads eight compiled workbooks (IMF and central-bank BoP and IIP, quarterly and annual), one sheet per economy.
  - Only the 19 SEACEN economies are kept; the non-SEACEN economies in the IMF workbooks are left out.
  - Chinese Taipei is read from the central-bank workbooks, so it no longer needs a separate file. Its IIP now runs to 2025.
- **Capital Flows tab.** A new tab for the Section 2 data: BoP flows and IIP positions of the SEACEN economies (US$ billion), annual, half-yearly or quarterly.
  - Three views: by regional group, by category (direct investment, portfolio equity and debt, loans, …) for one group, or one item by economy.
  - A stacked bar chart with the total as a line, and a data table with a CSV download.
  - Its own "data last refreshed" date, from the Section 2 sheet.

### Changed
- **Tab name:** the **Dashboard** tab is now **Macro-Financial**. Its sidebar is headed **SEACEN Macro-Financial Dashboard**, without the "About this dashboard" title.
- **Excel download:** covers the Macro-Financial tab; its note now says so.

## 2.0.0 — 2 November 2026

### New
- **GDP Growth and Inflation Outlook card.** GDP growth and inflation forecasts (FocusEconomics) for last year, this year and next year, for World, Asia-Pacific, China, India, Asian Advanced Economies, ASEAN5 and Asian EDMEs. Regional figures are PPP-weighted averages.
- **Monthly Inflation card.** Monthly CPI inflation, year-on-year, as a 3-month moving average, from January 2024. Economies with much higher inflation (Laos and Sri Lanka) are on a dashed right-hand axis.
- **Download figures (Excel).** One sheet per card, with the data the card is showing (after its filters), a native Excel chart that can still be edited, and any active event markers.
- **Tabs.** **About** (what the dashboard is and how to use and maintain it), **Updates** (this page) and **Data pipeline** (how the data is built).
- **Sidebar.**
  - The date the data was last refreshed, and this version number.
  - A viewing tip for laptop screens (zoom the page out).
  - Contact details and a short disclaimer.
- **SEACEN branding.** The SEACEN logo appears in the navigation bar and as the browser-tab icon.
- **Disclaimer, notes and abbreviations**, in the About tab.
- **Display switches.** Two TRUE/FALSE settings in the Google Sheet (`to_show` worksheet):
  - `developed` shows or hides the "Developed by" link.
  - `excel_use` decides whether the Excel download includes the charts built from CEIC and Haver Analytics data (Policy Rate, 5-Year Sovereign CDS, Monthly Inflation, and GDP Growth and Inflation Outlook). A note and tooltip above the download button say which applies.

### Changed
- **Name:** now the **SEACEN Capital Flow Monitor Dashboard**, and published at a new address, <https://ahmadaizudeen.shinyapps.io/seacen-cfm-dashboard/> (previously `cfm-final`).
- **Where the data comes from:**
  - The dashboard now reads published Google Sheets, kept for a day, instead of a workbook bundled with the app. A local copy is used only if the sheets can't be reached.
  - A new data pipeline (`cfm_data/`) builds every table and uploads it to the Google Sheets. It replaces the working notebook.
- **Financial Stress Index:**
  - The last month is now chosen automatically: the latest month for which every market input has data.
  - Inputs that end earlier are carried forward, and the pipeline reports each one.
- **Charts:**
  - Every chart names its unit on the value axis.
  - Chart titles and card headers were reworded.
  - Legends were moved below most charts.
  - Each card's footer gives a note and, on its own line, the source. The Liquidity card now has one too.
- **Year filters:** Exchange Rate, Stock Price Index and Sovereign Bond Yield changes offer the latest 5 years. An unfinished year is labelled `YTD` (for example `YTD2026`) on all three, where before only stocks and bonds had it.
- **Quarterly GDP Growth:** shows the latest 9 quarters.
- **Navigation bar:** stays visible while scrolling.
- **Names in the charts:** Taiwan is shown as **Chinese Taipei**, South Korea as **Korea**, and Hong Kong SAR (China) as **Hong Kong**. The Hong Kong / Korea / Singapore / Chinese Taipei group is **Asian Advanced Economies** everywhere. Card footers keep their formal names (for example "Hong Kong, China").
- **Uploading to Google Sheets:**
  - Worksheets are now overwritten in place instead of deleted and re-created, so their links keep working and a failed upload never leaves a worksheet empty.
  - Uploads are paced to stay within Google's request limit.
  - Every upload is logged in a `today` worksheet.

### Fixed
- **Policy-rate views:** "Selected Region 1/2" left out Korea and Chinese Taipei, because the region lists didn't match the names in the data.
- **Stocks and inflation:** Vietnam is now included in the stock index changes. Vietnam's earlier CPI months are filled from a second source.
- **Startup crash:** the dashboard no longer stops at startup when Google Sheets formats the upload log's row and column counts as dates.

## 1.0.0 — March 2025

- **First release on shinyapps.io**, as `cfm-final`.
- **Cards:** VIX, policy uncertainty, policy rates, exchange rate, stock index and bond yield changes, non-resident portfolio flows, sovereign CDS, GDP growth, Financial Stress Index and global liquidity.
- **Data:** read from a workbook bundled with the app.
