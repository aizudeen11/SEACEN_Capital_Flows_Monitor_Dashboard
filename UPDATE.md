# Updates

What has changed in the SEACEN Capital Flow Monitor Dashboard, newest version first.

**Version numbers** follow `MAJOR.MINOR.PATCH`:
- **MAJOR:** a redesign, or a change in how the data is produced or read.
- **MINOR:** new cards, charts or features.
- **PATCH:** fixes and small wording or layout changes.

The version shown in the dashboard's sidebar is read from the first entry below. To release a new version, add an entry at the top, in the same format (`## <version> — <date>`).

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
- **Name:** now the **SEACEN Capital Flow Monitor Dashboard**, and published at a new address, <https://ahmadaizudeen.shinyapps.io/SEACEN-CFM-Dashboard/> (previously `cfm-final`).
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
