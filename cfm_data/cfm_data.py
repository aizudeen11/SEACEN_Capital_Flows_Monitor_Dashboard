"""
Data pipeline for the Capital Flow Monitor dashboard.

Builds every table the dashboard needs from the raw workbooks and public APIs.
See README.md in this folder for what each function returns, the inputs it
expects and the known ways it can break.

Run it directly to build everything and print a summary:

    python cfm_data.py
    python cfm_data.py --raw "path/to/raw data for CFM.xlsx" --fsi "path/to/00 dlx.xlsx"
    python cfm_data.py --cutoff 2026-08     # override the automatic FSI cutoff
"""

import argparse
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import requests
from arch import arch_model

# The workbooks live in the same folder as this file.
DATA_DIR = Path(__file__).resolve().parent
RAW_PATH = DATA_DIR / "raw data for CFM.xlsx"
FSI_PATH = DATA_DIR / "00 dlx.xlsx"

IMF_WEO_URL = "https://api.imf.org/external/sdmx/3.0/data/dataflow/IMF.RES/WEO/9.0.0/"
VIX_URL = "https://cdn.cboe.com/api/global/us_indices/daily_prices/VIX_History.csv"
BIS_CHINA_POLICY_URL = "https://stats.bis.org/api/v2/data/dataflow/BIS/WS_CBPOL/1.0/M.CN?format=csv"
BIS_LIQUIDITY_URL = (
    "https://stats.bis.org/api/v2/data/dataflow/BIS/WS_GLI/1.0/Q.TO1.5J.A.B.I.A.771"
    "?startPeriod={start}&endPeriod={end}&format=csv"
)
IIF_FLOWS_URL = "https://www.iif.com/Portals/0/Files/Databases/monthly_em_portfolio_flows_database.xlsx"
REQUEST_TIMEOUT = 60


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def sdmx_to_dataframe(payload: dict) -> pd.DataFrame:
    """
    Flattens an SDMX-JSON response into one row per observation. Series keys
    (e.g. "0:0:0") are positions into the series dimensions (country,
    indicator, frequency) and observation keys are positions into the
    time-period list.
    """
    data = payload["data"]
    structure = data["structures"][0]
    series_dims = structure["dimensions"]["series"]
    periods = [v["value"] for v in structure["dimensions"]["observation"][0]["values"]]

    rows = []
    for series_key, series in data["dataSets"][0]["series"].items():
        positions = [int(i) for i in series_key.split(":")]
        labels = {dim["id"]: dim["values"][pos]["id"] for dim, pos in zip(series_dims, positions)}
        for obs_key, obs in series["observations"].items():
            rows.append({**labels, "TIME_PERIOD": periods[int(obs_key)], "OBS_VALUE": obs[0]})

    df = pd.DataFrame(rows)
    df["OBS_VALUE"] = pd.to_numeric(df["OBS_VALUE"])
    return df


def fetch_ppp_gdp(codes: list[str], first_year: int, last_year: int) -> pd.DataFrame:
    """
    Annual PPP GDP from the IMF WEO, with years ("2025", ...) as rows and the
    IMF country codes as columns.
    """
    # The API only returns a year whose whole span is inside TIME_PERIOD, so
    # the range runs from the last day of the year before to Dec 31.
    url = (
        f"{IMF_WEO_URL}{'+'.join(codes)}.PPPGDP.A"
        f"?c%5BTIME_PERIOD%5D=ge:{first_year - 1}-12-31+le:{last_year}-12-31"
        "&attributes=all&detail=full&includeHistory=true&limit=100"
    )
    response = requests.get(url, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    ppp = sdmx_to_dataframe(response.json())
    return ppp.pivot(index="TIME_PERIOD", columns="COUNTRY", values="OBS_VALUE")


def _label_ytd(df: pd.DataFrame) -> pd.DataFrame:
    """Relabels the Year of a year without December rows yet as "YTD<year>"."""
    last_month = df.groupby("Year")["Date"].max()
    ytd = {y: f"YTD{y}" for y, d in last_month.items() if d.month < 12}
    df["Year"] = df["Year"].replace(ytd)
    return df


def _yearly_change(df: pd.DataFrame, first_year: int, pct: bool) -> pd.DataFrame:
    """
    Long (Date, Region, Value) monthly data -> (Year, Region, Value), where
    Value is the change from the first to the last month with data in each
    year: a % change if `pct`, otherwise a difference. A year without
    December rows yet is labelled "YTD<year>". Regions are ordered by their
    value in the latest year, highest first.
    """
    df["Year"] = df["Date"].dt.year.astype(str)
    df = _label_ytd(df[df["Year"].astype(int).between(first_year, pd.Timestamp.today().year)])

    # Missing months are skipped rather than counted as 0, which would turn
    # the change into the whole level (e.g. a series that starts mid-year).
    df = df.dropna(subset=["Value"]).sort_values(["Region", "Year", "Date"])
    by_year = df.groupby(["Region", "Year"])["Value"].agg(
        first=lambda s: s.iloc[0], last=lambda s: s.iloc[-1], months="size"
    )
    if pct:
        change = ((by_year["last"] / by_year["first"]) - 1) * 100
    else:
        change = by_year["last"] - by_year["first"]
    change = change.where(by_year["months"] > 1)  # a single month has no change

    df = change.rename("Value").reset_index()
    latest_year = df["Year"].max()
    region_sort = list(df[df["Year"] == latest_year].sort_values(by="Value", ascending=False)["Region"])
    df["Region"] = pd.Categorical(df["Region"], categories=region_sort, ordered=True)
    return df.sort_values(by=["Region", "Year"], ascending=False)


# Taiwan's name in every table this module builds. The source workbooks call it
# "Taiwan (China)", "Taiwan" or "Taiwan, China"; each reader renames it to this.
TAIWAN = "Chinese Taipei"


# ---------------------------------------------------------------------------
# GDP growth (actual, quarterly)
# ---------------------------------------------------------------------------

GDP_FIRST_YEAR = 2022
# Rows at the top of the GDP sheet that describe each column rather than hold data.
GDP_METADATA_ROWS = ["Unit", "Source", "Last Update Time"]
GDP_IMF_NAMES = {
    "CHN": "China", "HKG": "Hong Kong", "IDN": "Indonesia", "IND": "India",
    "KOR": "Korea", "MYS": "Malaysia", "MNG": "Mongolia", "PHL": "Philippines",
    "SGP": "Singapore", "THA": "Thailand", "TWN": TAIWAN, "VNM": "Vietnam",
}
# GDP sheet columns holding levels (turned into YoY growth) vs growth rates already.
GDP_LEVEL_COLUMNS = ["China", "India", "Malaysia", "Mongolia", "Philippines", "Taiwan (China)", "Thailand"]
GDP_RATE_COLUMNS = ["Hong Kong SAR (China)", "Indonesia", "South Korea", "Singapore", "Vietnam"]
ASEAN5_GDP = ["Indonesia", "Malaysia", "Philippines", "Thailand", "Vietnam"]
ADVANCED_ASIA_GDP = ["Hong Kong", "Korea", "Singapore", TAIWAN]


def gdp_growth(path: Path = RAW_PATH) -> pd.DataFrame:
    """
    Quarterly real GDP growth (Quarter, Region, Value) for Asia, China, India,
    ASEAN-5 and Asian Advanced Economies. The groups are PPP-weighted
    averages, using each year's IMF WEO PPP GDP for all four of its quarters.
    """
    annual = fetch_ppp_gdp(list(GDP_IMF_NAMES), GDP_FIRST_YEAR, date.today().year).T
    annual.index = annual.index.map(GDP_IMF_NAMES)
    annual = annual.sort_index()
    weights = pd.concat(
        {f"{q}Q{year}": annual[year] for year in annual.columns for q in range(1, 5)}, axis=1
    )

    df = pd.read_excel(path, sheet_name="GDP", header=1)
    df = df[~df["Region"].isin(GDP_METADATA_ROWS)]
    df = (
        df[df.columns[:13]]
        .rename(columns={"Region": "Date"})
        .set_index("Date")
        .apply(pd.to_numeric, errors="coerce")
    )
    levels = df[GDP_LEVEL_COLUMNS].apply(lambda x: ((x / x.shift(4)) - 1) * 100)
    growth = pd.concat([df[GDP_RATE_COLUMNS], levels], axis=1).reset_index()
    dates = pd.to_datetime(growth["Date"]) + pd.offsets.MonthEnd(0)
    growth["Date"] = [f"{(d.month - 1) // 3 + 1}Q{d.year}" for d in dates]

    # Economies as rows, quarters as columns, named to match the PPP weights.
    growth = growth.set_index("Date").T.rename(
        index={"Hong Kong SAR (China)": "Hong Kong", "South Korea": "Korea", "Taiwan (China)": TAIWAN}
    )

    # Keep only the quarters that have both weights and growth data.
    quarters = [q for q in weights.columns if q in growth.columns]
    weights = weights[quarters]
    growth = growth.reindex(weights.index)[quarters]

    def weighted(members: list[str]) -> pd.Series:
        w = weights.loc[members]
        return (w / w.sum() * growth.loc[members]).sum()

    result = pd.concat(
        {
            "Asia": weighted(list(weights.index)),
            "China": growth.loc["China"],
            "India": growth.loc["India"],
            "ASEAN-5": weighted(ASEAN5_GDP),
            "Asian Advanced Economies": weighted(ADVANCED_ASIA_GDP),
        },
        axis=1,
    )
    result = result.rename_axis("Quarter").reset_index()
    return result.melt(id_vars="Quarter", var_name="Region", value_name="Value")


# ---------------------------------------------------------------------------
# Market data
# ---------------------------------------------------------------------------

def vix_history() -> pd.DataFrame:
    """CBOE VIX: monthly average of the daily close as (Date, CLOSE), latest month first."""
    df = pd.read_csv(VIX_URL, storage_options={"User-Agent": "Mozilla/5.0"})
    df["DATE"] = pd.to_datetime(df["DATE"])
    monthly = df.groupby(df["DATE"].dt.to_period("M"))["CLOSE"].mean()
    monthly.index = monthly.index.to_timestamp() + pd.offsets.MonthEnd(0)
    return monthly.rename_axis("Date").reset_index().iloc[::-1].reset_index(drop=True)


POLICY_START = "2022-01-01"
# Spelled as in the "Policy Rate" sheet (after renaming "Taiwan (China)").
POLICY_REGIONS_1 = [
    "United States", "European Union", "Mongolia", "Nepal", "Philippines",
    "South Korea", "Sri Lanka", "Thailand", "China",
]
POLICY_REGIONS_2 = ["Indonesia", "India", "Malaysia", TAIWAN, "Vietnam"]


def policy_rate(path: Path = RAW_PATH) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Policy rates as (Region, Date, Value) from January 2022: (all regions,
    POLICY_REGIONS_1, POLICY_REGIONS_2). The sheet has no China rate, so China
    comes from the BIS API.
    """
    df = pd.read_excel(path, sheet_name="Policy Rate")
    df = df.drop(df.columns[[0, 2, 3]], axis=1)
    df = df.rename(columns={d: d.strftime("%Y-%m") for d in df.columns[1:]})
    df = df.melt(id_vars="Region", var_name="Date", value_name="Value")
    df["Date"] = pd.to_datetime(df["Date"]) + pd.offsets.MonthEnd(0)
    df["Region"] = df["Region"].replace({"Taiwan (China)": TAIWAN})

    china = pd.read_csv(BIS_CHINA_POLICY_URL)
    china = pd.DataFrame({
        "Region": "China",
        "Date": pd.to_datetime(china["TIME_PERIOD"]) + pd.offsets.MonthEnd(0),
        "Value": china["OBS_VALUE"],
    })
    df = pd.concat([df[df["Region"] != "China"], china], ignore_index=True)
    df = df[df["Date"] >= POLICY_START].reset_index(drop=True)

    region1 = df[df["Region"].isin(POLICY_REGIONS_1)].reset_index(drop=True)
    region2 = df[df["Region"].isin(POLICY_REGIONS_2)].reset_index(drop=True)
    return df, region1, region2


FX_FIRST_YEAR = 2020
FX_REGIONS = [
    "China", "Hong Kong SAR (China)", "India", "Indonesia", "Malaysia", "Philippines",
    "Singapore", "South Korea", "Taiwan (China)", "Thailand", "Vietnam",
]


def forex_exchange(path: Path = RAW_PATH) -> pd.DataFrame:
    """
    (Year, Region, Value): % change of each exchange rate from the first to
    the last month of each year (positive = the local currency appreciated
    against the US dollar). A year without December rows yet is labelled
    "YTD<year>". Regions ordered by the latest year, highest first.
    """
    df = pd.read_excel(path, sheet_name="FX", header=1)
    df = df[["Region"] + FX_REGIONS].rename(columns={"Region": "Date", "Taiwan (China)": TAIWAN})
    # Keep only the monthly rows; this drops the "Unit" / "Last Update Time" rows.
    df["Date"] = pd.to_datetime(df["Date"], format="%Y-%m-%d %H:%M:%S", errors="coerce")
    df = df.dropna(subset=["Date"])
    df = df.melt(id_vars="Date", var_name="Region", value_name="Value")
    df["Date"] = df["Date"] + pd.offsets.MonthEnd(0)
    df["Year"] = df["Date"].dt.year.astype(str)
    df = _label_ytd(df[df["Year"].astype(int).between(FX_FIRST_YEAR, pd.Timestamp.today().year)])

    # First and last monthly rate of each year per region (a missing rate counts as 0).
    df = df.sort_values(["Region", "Year", "Date"])
    by_year = df.groupby(["Region", "Year"])["Value"].agg(
        first=lambda s: s.iloc[0], last=lambda s: s.iloc[-1], months="size"
    )
    first, last = by_year["first"].fillna(0), by_year["last"].fillna(0)
    change = ((first / last) - 1) * 100
    change = change.where(~((first == 0) & (last == 0)), 0)
    change = change.where(by_year["months"] > 1)
    change = change.replace([np.inf, -np.inf], np.nan)

    df = change.rename("Value").reset_index()
    latest_year = df["Year"].max()
    region_sort = list(df[df["Year"] == latest_year].sort_values(by="Value", ascending=False)["Region"])
    df["Region"] = pd.Categorical(df["Region"], categories=region_sort, ordered=True)
    return df.sort_values(by=["Region", "Year"], ascending=False)


def cds(path: Path = RAW_PATH) -> pd.DataFrame:
    """5-year sovereign CDS (bps) as (Date, Region, Value); months with no quote stay NaN."""
    df = pd.read_excel(path, sheet_name="CDS", header=1)
    # Header cells look like "China: 5 Year USD Credit Default Swap ..."; keep the country.
    df.columns = ["Date"] + [c.split(":")[0] for c in df.columns[1:]]
    # Keep only the monthly rows ("Jan-2010"); this drops the .T1 / .TN / .SOURCE rows.
    df["Date"] = pd.to_datetime(df["Date"], format="%b-%Y", errors="coerce")
    df = df.dropna(subset=["Date"])
    df = df.melt(id_vars="Date", var_name="Region", value_name="Value")
    df["Date"] = df["Date"] + pd.offsets.MonthEnd(0)
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce")
    return df


LIQUIDITY_YEARS = 11


def liquidity() -> pd.DataFrame:
    """BIS global liquidity indicator (quarterly) for the past LIQUIDITY_YEARS years as (TIME_PERIOD, OBS_VALUE)."""
    today = pd.Timestamp.today().normalize()
    start = today - pd.DateOffset(years=LIQUIDITY_YEARS)
    url = BIS_LIQUIDITY_URL.format(start=start.date(), end=today.date())
    return pd.read_csv(url)[["TIME_PERIOD", "OBS_VALUE"]]


YIELDS_FIRST_YEAR = 2021


def sovereign_bond_yields(path: Path = RAW_PATH) -> pd.DataFrame:
    """
    (Year, Region, Value): change in the 5-year government bond yield
    (percentage points) from the first to the last month with data in each year.
    """
    df = pd.read_excel(path, sheet_name="Yields YTD", header=1)
    # Header cells look like "Hong Kong: 5-Yr Government Bond Yield (%)"; keep the country.
    df.columns = ["Date"] + [c.split(":")[0] for c in df.columns[1:]]
    df = df.rename(columns={"Taiwan": TAIWAN})
    # Keep only the monthly rows ("Jan-2010"); this drops the .T1 / .TN / .SOURCE rows.
    df["Date"] = pd.to_datetime(df["Date"], format="%b-%Y", errors="coerce")
    df = df.dropna(subset=["Date"])
    df = df.melt(id_vars="Date", var_name="Region", value_name="Value")
    df["Date"] = df["Date"] + pd.offsets.MonthEnd(0)
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce")
    return _yearly_change(df, YIELDS_FIRST_YEAR, pct=False)


STOCK_FIRST_YEAR = 2020


def stock_price_index(path: Path = RAW_PATH) -> pd.DataFrame:
    """
    (Year, Region, Value): % change of each benchmark stock index from the
    first to the last month with data in each year.
    """
    raw = pd.read_excel(path, sheet_name="stock indices", header=None)
    # The sheet holds two downloads side by side, each with its own date column:
    # the CEIC block (economies named in the "Region" row) and, to its right,
    # the Haver block for Vietnam (series described in the ".DESC" row).
    haver_col = next(c for c in raw.columns if (raw[c] == ".DESC").any())

    ceic = raw.loc[:, :haver_col - 1]
    region_row = ceic.index[ceic[0] == "Region"][0]
    ceic = ceic.loc[region_row + 1:]
    ceic.columns = ["Date"] + raw.loc[region_row, 1:haver_col - 1].tolist()
    # Keep only the monthly rows; this drops the "Unit" / "Last Update Time" rows.
    ceic["Date"] = pd.to_datetime(ceic["Date"], format="%Y-%m-%d %H:%M:%S", errors="coerce")

    haver = raw.loc[:, haver_col:]
    desc_row = haver.index[haver[haver_col] == ".DESC"][0]
    names = [d.split(":")[0] for d in raw.loc[desc_row, haver_col + 1:]]
    haver = haver.loc[desc_row + 1:]
    haver.columns = ["Date"] + names
    # Haver dates look like "Jan-2010"; this drops the .T1 / .TN / .SOURCE rows.
    haver["Date"] = pd.to_datetime(haver["Date"], format="%b-%Y", errors="coerce")

    df = ceic.dropna(subset=["Date"]).merge(haver.dropna(subset=["Date"]), on="Date", how="outer")
    df = df.rename(columns={"Taiwan (China)": TAIWAN})
    df = df.melt(id_vars="Date", var_name="Region", value_name="Value")
    df["Date"] = df["Date"] + pd.offsets.MonthEnd(0)
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce")
    return _yearly_change(df, STOCK_FIRST_YEAR, pct=True)


# Economies summed into the portfolio-flow totals, spelled as in the IIF sheet.
# Where an economy has no Debt (or Equity) column it drops out of that total.
FLOW_COUNTRIES = [
    "China", "India", "Indonesia", "Korea", "Malaysia", "Philippines",
    "Taiwan, China", "Thailand", "Vietnam", "Sri lanka", "Mongolia",
]


def capital_flows(url: str = IIF_FLOWS_URL) -> pd.DataFrame:
    """
    IIF monthly EM portfolio flows: (Date, Portfolio Debt, Portfolio Equity),
    net non-resident purchases summed over FLOW_COUNTRIES, in USD billion.
    """
    response = requests.get(url, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    if "spreadsheetml" not in response.headers.get("Content-Type", ""):
        raise ValueError(f"{url} did not return an Excel file; the IIF link may have changed.")

    raw = pd.read_excel(BytesIO(response.content), sheet_name="Country Flows_nw", header=None)
    # Row 3 names each economy once above its Equity/Debt columns and row 4
    # says which is which, so pick columns by those labels, not by position.
    country = raw.iloc[3].ffill()
    flow_type = raw.iloc[4].astype(str).str.strip()
    data = raw.iloc[5:]
    data = data[pd.to_datetime(data[0], errors="coerce").notna()]

    def total(kind: str) -> pd.Series:
        cols = [c for c in raw.columns[1:] if country[c] in FLOW_COUNTRIES and flow_type[c] == kind]
        return data[cols].apply(pd.to_numeric, errors="coerce").sum(axis=1) / 1000

    return pd.DataFrame({
        "Date": pd.to_datetime(data[0]).values,
        "Portfolio Debt": total("Debt").values,
        "Portfolio Equity": total("Equity").values,
    })


# ---------------------------------------------------------------------------
# Financial Stress Index
# ---------------------------------------------------------------------------

# Last month the FSI covers. None = automatic (fsi_auto_cutoff): the latest
# month every market input has data. Set a month such as "2026-08" (or pass
# --cutoff) to override. Any input series that stops before the cutoff is
# filled with its latest value; fsi_missing_values() lists what was filled.
FSI_CUTOFF = None
# For the automatic cutoff, a market series this many months behind the newest
# one counts as stale (e.g. the Thai market caps that stopped in Jul 2023) and
# doesn't hold the cutoff back; it is still filled and reported.
FSI_STALE_MONTHS = 12

FSI_STOCK_KEYS = {
    'Select this link and click Refresh/Edit Download to update data and add or remove series': 'Date',
    'Index: Shenzhen Stock Exchange: Composite': 'China',
    'Equity Market Index: Month End: Hang Seng': 'Hong Kong SAR (China)',
    'Equity Market Index: Month End: BSE: Sensitive 30 (Sensex)': 'India',
    'Equity Market Index: Month End: Jakarta Composite': 'Indonesia',
    'Equity Market Index: Month End: FTSE Bursa Malaysia: Composite': 'Malaysia',
    'Equity Market Index: Month End: PSEi': 'Philippines',
    'Equity Market Index: Month End: FTSE Strait Times': 'Singapore',
    'Equity Market Index: Month End: KOSPI': 'South Korea',
    'TWSE: Equity Market Index: TAIEX Capitalization Weighted: Month Avg': TAIWAN,
    'Equity Market Index: Month End: SET': 'Thailand',
}
FSI_FINANCIAL_KEYS = {
    'Select this link and click Refresh/Edit Download to update data and add or remove series': 'Date',
    'Index: Shenzhen Stock Exchange: Financials': 'China',
    'Index: Hang Seng Finance': 'Hong Kong SAR (China)',
    'BSE: Index: Bankex': 'India',
    '(DC)IDX: Index: Finance': 'Indonesia1',
    'IDX: Index: Financials': 'Indonesia2',
    'Bursa Malaysia: Index: Financial Services': 'Malaysia',
    'Index: BankingFinancial Services': 'Philippines',
    'Equity Market Index: Month End: FTSE Strait Times': 'Singapore1',
    'Singapore Exchange Turnover: Value: Mainboard': 'Singapore2',
    'Singapore Exchange Turnover: Value: Mainboard: Financials': 'Singapore3',
    'Index: KOSPI: Financial Institutions': 'South Korea',
    'TWSE: Equity Market Index: Finance and Insurance': TAIWAN,
    'SET: Index: Banking': 'Thailand1',
    'SET: Index: Finance and Securities': 'Thailand2',
    'SET: Index: Insurance': 'Thailand3',
    'SET: Market Capitalization: Banking': 'Thailand4',
    'SET: Market Capitalization: Finance and Securities': 'Thailand5',
    'SET: Market Capitalization: Insurance': 'Thailand6',
}
FSI_COUNTRIES = [
    'China', 'Hong Kong SAR (China)', 'India', 'Indonesia', 'Malaysia', 'Philippines',
    'Singapore', 'South Korea', 'Taiwan (China)', 'Thailand',
]


def _fill_to_cutoff(df, cutoff, sheet, report, labels=None, skip=()):
    """
    Cuts a Date-indexed (month-end) table at `cutoff`, extends it to `cutoff`
    if the sheet ends earlier, and fills each series' trailing gap with its
    latest available value. Gaps before a series' last value are left alone.
    Each filled series is added to `report`.
    """
    labels = labels or {}
    df = df.loc[df.index <= cutoff]
    df = df.reindex(df.index.union(pd.date_range(df.index.max(), cutoff, freq="ME")))
    df.index.name = "Date"
    for col in df.columns:
        if col in skip:
            continue
        last = df[col].last_valid_index()
        if last is None or last >= cutoff:
            continue
        gap = df.index[df.index > last]
        df.loc[gap, col] = df.at[last, col]
        report.append({
            "Sheet": sheet,
            "Series": labels.get(col, col),
            "Last available": last.strftime("%b %Y"),
            "Value used": df.at[last, col],
            "Months filled": ", ".join(d.strftime("%b %Y") for d in gap),
            "Count": len(gap),
        })
    return df


# How each FSI input is labelled in the fill report, which of its columns are
# never filled, and whether it can hold back the automatic cutoff (ORA, the
# reserves data, is allowed to lag and is carried forward instead).
_FSI_INPUTS = {
    "yields1": ("00 dlx (10Y yield)", {}, (), True),
    "yields2": ("Yields2 (10Y yield)", {}, (), True),
    "stocks": ("Stock Price Indexes", {v: k for k, v in FSI_STOCK_KEYS.items()}, (), True),
    # Indonesia1 is the discontinued (DC) index; the switch to Indonesia2
    # relies on it staying empty after it stops, so it is never filled.
    "financials": (
        "Stock Price Indexes (financials)", {v: k for k, v in FSI_FINANCIAL_KEYS.items()}, ("Indonesia1",), True,
    ),
    "fx": ("FX", {}, (), True),
    "ora": ("ORA", {}, (), False),
}


def _read_fsi_inputs(path: Path = FSI_PATH) -> dict[str, pd.DataFrame]:
    """
    Reads every FSI input from `path` as Date-indexed (month-end) float
    tables, keyed as in _FSI_INPUTS, without cutting or filling.
    """
    # Indonesia and Philippines 10-year yields come from the "00 dlx" sheet,
    # with dates like "Sep-10"; the rest come from Yields2.
    yield_data1 = pd.read_excel(path, header=1)
    yield_data1 = yield_data1[[*yield_data1.columns[[0, 1, 3]]]]
    yield_data1 = yield_data1.drop([0, 1, 2])
    yield_data1 = yield_data1.rename(columns={
        '.DESC': 'Date',
        'Indonesia: 10 Year Treasury Bond Mid Yield (% p.a.)': 'Indonesia',
        'Philippines: 10 Year Treasury Bond Mid Yield (% p.a.)': 'Philippines',
    })
    yield_data1["Date"] = pd.to_datetime(yield_data1["Date"], format="%b-%y") + pd.offsets.MonthEnd(0)
    yield_data1 = yield_data1.set_index('Date').astype(float)

    yield_data2 = pd.read_excel(path, sheet_name='Yields2', header=1)
    yield_data2 = yield_data2[yield_data2.columns[:11]]
    yield_data2 = yield_data2.drop([0, 1])
    yield_data2 = yield_data2.rename(columns={'Region': 'Date', 'Taiwan (China)': TAIWAN})
    yield_data2["Date"] = pd.to_datetime(yield_data2["Date"]) + pd.offsets.MonthEnd(0)
    yield_data2 = yield_data2.set_index('Date').astype(float)

    stock_sheet = pd.read_excel(path, sheet_name='Stock Price Indexes')

    stock_data = stock_sheet[FSI_STOCK_KEYS.keys()]
    stock_data = stock_data.drop([0, 1, 2])
    stock_data = stock_data.rename(columns=FSI_STOCK_KEYS)
    stock_data["Date"] = pd.to_datetime(stock_data["Date"]) + pd.offsets.MonthEnd(0)
    stock_data = stock_data.set_index('Date').astype(float)

    fb = stock_sheet[FSI_FINANCIAL_KEYS.keys()]
    fb = fb.drop([0, 1, 2])
    fb = fb.rename(columns=FSI_FINANCIAL_KEYS)
    fb['Date'] = pd.to_datetime(fb['Date']) + pd.offsets.MonthEnd(0)
    fb = fb.set_index('Date').astype(float)

    fx = pd.read_excel(path, sheet_name='FX', header=1)
    fx = fx.drop([0, 1])
    fx = fx[['Region'] + FSI_COUNTRIES]
    fx = fx.rename(columns={'Region': 'Date', 'Taiwan (China)': TAIWAN})
    fx["Date"] = pd.to_datetime(fx["Date"]) + pd.offsets.MonthEnd(0)
    fx = fx.set_index('Date').astype(float)

    ora = pd.read_excel(path, sheet_name='ORA', header=1)
    ora = ora.drop([0, 1])
    ora = ora[['Region'] + FSI_COUNTRIES]
    ora = ora.rename(columns={'Region': 'Date', 'Taiwan (China)': TAIWAN})
    ora["Date"] = pd.to_datetime(ora["Date"]) + pd.offsets.MonthEnd(0)
    ora = ora.set_index('Date').astype(float)

    return {
        "yields1": yield_data1, "yields2": yield_data2, "stocks": stock_data,
        "financials": fb, "fx": fx, "ora": ora,
    }


def fsi_auto_cutoff(path: Path = FSI_PATH, inputs: dict | None = None) -> pd.Timestamp:
    """
    The automatic FSI cutoff: the latest month for which every market input
    (yields, stock and financial-sector indices, FX) has data. ORA may lag and
    is carried forward instead; series more than FSI_STALE_MONTHS behind the
    newest one (stale or discontinued) are ignored.
    """
    inputs = inputs if inputs is not None else _read_fsi_inputs(path)
    lasts = []
    for key, (_, _, skip, sets_cutoff) in _FSI_INPUTS.items():
        if not sets_cutoff:
            continue
        for col, series in inputs[key].items():
            last = series.last_valid_index()
            if col not in skip and last is not None:
                lasts.append(last)
    newest = max(lasts)
    return min(d for d in lasts if d > newest - pd.DateOffset(months=FSI_STALE_MONTHS))


def _load_fsi_inputs(path: Path = FSI_PATH, cutoff: str | None = FSI_CUTOFF):
    """
    Reads every FSI input from `path`, cut and filled up to `cutoff` (None =
    automatic). Returns (yields, stock indices, financial-sector indices, FX,
    ORA, report, cutoff used).
    """
    inputs = _read_fsi_inputs(path)
    if cutoff is None:
        cutoff = fsi_auto_cutoff(path, inputs)
    cutoff = pd.Timestamp(cutoff) + pd.offsets.MonthEnd(0)

    report = []
    filled = {}
    for key, (sheet, labels, skip, _) in _FSI_INPUTS.items():
        filled[key] = _fill_to_cutoff(inputs[key], cutoff, sheet, report, labels=labels, skip=skip)

    report = pd.DataFrame(report, columns=["Sheet", "Series", "Last available", "Value used", "Months filled", "Count"])
    yield_data = filled["yields2"].join(filled["yields1"])
    return yield_data, filled["stocks"], filled["financials"], filled["fx"], filled["ora"], report, cutoff


def fsi_missing_values(path: Path = FSI_PATH, cutoff: str | None = FSI_CUTOFF) -> pd.DataFrame:
    """
    Every FSI input series with no data up to `cutoff` (None = automatic):
    the last month it was available and the months filled with that value.
    Empty = nothing filled.
    """
    return _load_fsi_inputs(path, cutoff)[5]


def _chain_index(dates: pd.Series, growth: pd.Series, base_date=None) -> pd.Series:
    """
    Builds an index from monthly % growth, starting at 100 and resetting to
    100 at `base_date` if given.
    """
    index = [100.0]
    for i in range(1, len(growth)):
        if base_date is not None and dates.iloc[i] == base_date:
            index.append(100.0)
        else:
            index.append(index[-1] * ((growth.iloc[i] / 100) + 1))
    return pd.Series(index, index=growth.index)


def financial_stress_index(
    path: Path = FSI_PATH, cutoff: str | None = FSI_CUTOFF,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Financial Stress Index, monthly from January 2010 up to `cutoff` (None =
    automatic, see fsi_auto_cutoff). Returns (yield spread, stock returns,
    EMPI, financial-sector beta, GARCH volatility, FSI), each standardised.
    The FSI is their sum per economy, plus "Asian Advanced Economies" and
    "ASEAN-4" simple averages.
    """
    yield_data, stock_data, fb, fx, ora, _, _ = _load_fsi_inputs(path, cutoff)

    # Yield spread over the US 10-year.
    yield_spread = yield_data.sub(yield_data['United States'], axis=0).drop('United States', axis=1)
    yield_spread = yield_spread.loc[yield_spread.index >= "2010-01-31"]
    yield_spread['Asia'] = yield_spread.mean(axis=1)
    yield_spread = (yield_spread - yield_spread.mean()) / yield_spread.std()

    # Stock market returns (log change, sign flipped so falls raise stress).
    stock_data = stock_data.loc[stock_data.index >= "2008-12-01"]
    stock_returns = np.log(stock_data)
    stock_returns = (stock_returns.shift(1) - stock_returns) * 100
    stock_returns['Asia'] = stock_returns.mean(axis=1)
    stock_std = (stock_returns - stock_returns.mean()) / stock_returns.std()
    stock_std = stock_std.loc[stock_std.index >= "2009-12-01"]

    fb = fb.reset_index()

    # Singapore has no financial-sector index, so build one from the Straits
    # Times growth weighted by the financials' share of mainboard turnover.
    sg_weight = fb['Singapore3'] / fb['Singapore2']
    sg_growth = ((fb['Singapore1'] / fb['Singapore1'].shift(1)) - 1) * 100 * sg_weight
    sg_index = _chain_index(fb['Date'], sg_growth, base_date=pd.Timestamp('2008-12-31'))

    # Thailand: banking / finance / insurance index growth weighted by market cap.
    th_levels = fb[['Thailand1', 'Thailand2', 'Thailand3']].to_numpy()
    th_caps = fb[['Thailand4', 'Thailand5', 'Thailand6']].to_numpy()
    th_sector_growth = (th_levels[1:] / th_levels[:-1] - 1) * 100
    th_weights = th_caps / th_caps.sum(axis=1, keepdims=True)
    # A missing sector or market cap makes that month's growth NaN.
    th_growth = pd.Series(
        np.r_[np.nan, (th_sector_growth * th_weights[1:]).sum(axis=1)], index=fb.index
    )
    th_index = _chain_index(fb['Date'], th_growth)

    # Indonesia: the old (DC) finance index until it stops, then the new one.
    switch = fb['Indonesia1'].isna().idxmax()
    indonesia = pd.concat([fb.loc[:switch - 1, 'Indonesia1'], fb.loc[switch:, 'Indonesia2']])

    financials = pd.DataFrame({
        'China': fb['China'], 'Hong Kong SAR (China)': fb['Hong Kong SAR (China)'], 'India': fb['India'],
        'Indonesia': indonesia, 'Malaysia': fb['Malaysia'], 'Philippines': fb['Philippines'],
        'Singapore': sg_index, 'South Korea': fb['South Korea'], TAIWAN: fb[TAIWAN], 'Thailand': th_index,
    }).set_axis(fb['Date']).astype(float)
    fin_returns = np.log(financials)
    fin_returns = (fin_returns.shift(1) - fin_returns) * 100
    fin_returns['Asia'] = fin_returns.mean(axis=1)

    # Rolling 13-month beta of each financial-sector index on its market index.
    dates = stock_returns.index
    covariance = pd.DataFrame(index=dates, columns=stock_returns.columns, dtype=float)
    variance = pd.DataFrame(index=dates, columns=stock_returns.columns, dtype=float)
    for col in stock_returns.columns:
        for i in range(1, len(dates) - 12):
            window = slice(dates[i], dates[i + 12])
            market = stock_returns.loc[window, col].values
            covariance.at[dates[i + 12], col] = np.cov(market, fin_returns.loc[window, col].values, bias=True)[0, 1]
            variance.at[dates[i + 12], col] = np.var(market, ddof=1)
    beta = covariance / variance
    beta = (beta - beta.mean()) / beta.std()
    beta = beta.map(lambda x: x if x > 1 else 0)

    # Exchange market pressure: FX depreciation minus reserve change.
    fx = fx.loc[fx.index >= "2009-12-31"]
    fx = ((fx / fx.shift(1)) - 1) * 100
    fx['Asia'] = fx.mean(axis=1)
    fx = (fx - fx.mean()) / fx.std()

    ora = ora.loc[ora.index >= "2009-12-31"] / 1000
    ora = ((ora.shift(1) / ora) - 1) * 100
    ora['Asia'] = ora.mean(axis=1)
    ora = (ora - ora.mean()) / ora.std()

    empi = fx - ora
    yield_spread = yield_spread[stock_std.columns]

    # Stock market volatility: GARCH(1,1) conditional variance of the returns.
    garch = pd.DataFrame(index=stock_std.index)
    for col in stock_std.columns:
        res = arch_model(stock_std[col].tolist(), lags=1, vol='GARCH', p=1, q=1).fit(disp='off')
        garch[col] = res.conditional_volatility ** 2
    garch = (garch - garch.mean()) / garch.std()

    stock_std = stock_std.loc[stock_std.index >= "2010-01-01"]
    empi = empi.loc[empi.index >= "2010-01-01"]
    beta = beta.loc[beta.index >= "2010-01-01"]
    garch = garch.loc[garch.index >= "2010-01-01"]
    yield_spread = yield_spread.fillna(0)

    fsi = yield_spread + stock_std + empi + beta + garch
    fsi['Asian Advanced Economies'] = fsi[['Hong Kong SAR (China)', 'Singapore', 'South Korea', TAIWAN]].mean(axis=1)
    fsi['ASEAN-4'] = fsi[['Indonesia', 'Malaysia', 'Philippines', 'Thailand']].mean(axis=1)

    return yield_spread, stock_std, empi, beta, garch, fsi


# ---------------------------------------------------------------------------
# Forecasts (FocusEconomics) and CPI inflation
# ---------------------------------------------------------------------------

# Forecast sheet name -> IMF WEO country code, used to fetch the PPP weights.
FORECAST_IMF_CODES = {
    "Brunei": "BRN", "Cambodia": "KHM", "China": "CHN", "Hong Kong": "HKG",
    "India": "IND", "Indonesia": "IDN", "Korea": "KOR", "Laos": "LAO",
    "Malaysia": "MYS", "Mongolia": "MNG", "Myanmar": "MMR", "Philippines": "PHL",
    "Singapore": "SGP", "Sri Lanka": "LKA", TAIWAN: "TWN", "Thailand": "THA",
    "Vietnam": "VNM",
}
FORECAST_GROUPS = {
    "Asia-Pacific": list(FORECAST_IMF_CODES),
    "Asian Advanced Economies": ["Hong Kong", "Korea", "Singapore", TAIWAN],
    "ASEAN5": ["Indonesia", "Malaysia", "Philippines", "Thailand", "Vietnam"],
    "Asia EDMEs": ["Brunei", "Cambodia", "Laos", "Mongolia", "Myanmar", "Sri Lanka"],
}
FORECAST_ROWS = ['World', 'Asia-Pacific', 'China', 'India', 'Asian Advanced Economies', 'ASEAN5', 'Asia EDMEs']


def ppp_weighted_forecast(sheet_name: str, path: Path = RAW_PATH) -> pd.DataFrame:
    """
    A FocusEconomics forecast sheet ("Inflation forecast", "GDP growth") for
    last year, this year and next year: FORECAST_ROWS as rows, years as
    columns. Each group is the PPP-weighted average of its members using
    that year's IMF WEO PPP GDP; an economy with no forecast in a year is
    left out and the rest re-weighted.
    """
    years = [str(y) for y in range(date.today().year - 1, date.today().year + 2)]

    df = pd.read_excel(path, sheet_name=sheet_name, header=1)
    df = df.drop([0, 1, 2], axis=0)  # .T1 / .TN / .LSOURCE rows
    df = df.rename(columns={col: col.split(':')[0] for col in df.columns.tolist()[1:]})
    df = df.rename(columns={'.DESC': 'Year', 'Taiwan': TAIWAN})
    df['Year'] = df['Year'].map(lambda x: f"20{x}")
    df = df[df['Year'].isin(years)].set_index('Year').astype(float)

    ppp = fetch_ppp_gdp(list(FORECAST_IMF_CODES.values()), int(years[0]), int(years[-1]))
    ppp = ppp.rename(columns={v: k for k, v in FORECAST_IMF_CODES.items()}).reindex(df.index)

    for group, members in FORECAST_GROUPS.items():
        x = df[members]
        w = ppp[members].where(x.notna())
        df[group] = (x * w).sum(axis=1) / w.sum(axis=1)

    return df[FORECAST_ROWS].T


def _read_monthly_sheet(sheet_name: str, path: Path = RAW_PATH) -> tuple[pd.DataFrame, pd.Series]:
    """
    Reads a sheet with economies as rows and months as columns (a "Region"
    header row, optional "Unit" column). Returns (values with month-end
    dates as columns, each economy's Unit or an empty Series).
    """
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    region_col = next(c for c in raw.columns if (raw[c] == "Region").any())
    header_row = raw.index[raw[region_col] == "Region"][0]
    header = raw.loc[header_row]
    # Month columns are the ones whose header cell is a date.
    month_cols = [c for c in raw.columns if isinstance(header[c], datetime)]
    unit_col = next((c for c in raw.columns if header[c] == "Unit"), None)

    body = raw.loc[header_row + 1:].dropna(subset=[region_col])
    values = body[month_cols].apply(pd.to_numeric, errors="coerce")
    values.index = body[region_col].values
    values.columns = [pd.Timestamp(header[c]) + pd.offsets.MonthEnd(0) for c in month_cols]
    units = body[unit_col].set_axis(body[region_col].values) if unit_col is not None else pd.Series(dtype=str)
    return values, units


def cpi_inflation(path: Path = RAW_PATH) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, list[str]]]:
    """
    Monthly CPI inflation from the "Inflation" sheet, economies as rows and
    month-end dates as columns. Returns (YoY %, 3-month moving average of
    YoY, months still missing a YoY after the first year per economy).
    """
    cpi, units = _read_monthly_sheet("Inflation", path)
    cpi, units = cpi.rename(index={"Taiwan (China)": TAIWAN}), units.rename(index={"Taiwan (China)": TAIWAN})
    units = units.astype(str).str.strip()

    # Index levels -> YoY, i.e. Excel's =((BP7/BD7)-1)*100 twelve months apart.
    yoy = ((cpi / cpi.shift(12, axis=1)) - 1) * 100
    # Rows that are already year-on-year: "Prev Year=100" (China: 102.1 means
    # +2.1%) and "%" (Vietnam: the YoY rate itself).
    prev_year_100 = units.index[units == "Prev Year=100"]
    yoy.loc[prev_year_100] = cpi.loc[prev_year_100] - 100
    yoy_rate = units.index[units == "%"]
    yoy.loc[yoy_rate] = cpi.loc[yoy_rate]

    # Vietnam's YoY in the "Inflation" sheet only starts in Nov 2025; earlier
    # months come from the "VN CPI" sheet. Months already covered are kept.
    vn_yoy, _ = _read_monthly_sheet("VN CPI", path)
    vn_yoy = vn_yoy.reindex(columns=yoy.columns)
    yoy.loc[vn_yoy.index] = yoy.loc[vn_yoy.index].fillna(vn_yoy)

    # Three-month moving average, i.e. Excel's =AVERAGE(BX28:BZ28). All three
    # months must have a value.
    yoy_3mma = yoy.T.rolling(3).mean().T

    gaps = {
        region: [d.strftime("%b %Y") for d in row.index[row.isna()] if d >= yoy.columns[12]]
        for region, row in yoy.iterrows()
    }
    return yoy, yoy_3mma, {k: v for k, v in gaps.items() if v}


# ---------------------------------------------------------------------------
# Everything at once
# ---------------------------------------------------------------------------

def load_all(raw_path: Path = RAW_PATH, fsi_path: Path = FSI_PATH, fsi_cutoff: str | None = FSI_CUTOFF) -> dict:
    """Builds every table. Keys are the table names used in README.md."""
    policy_all, policy_1, policy_2 = policy_rate(raw_path)
    _, _, _, _, _, fsi = financial_stress_index(fsi_path, fsi_cutoff)
    cpi_yoy, cpi_yoy_3mma, cpi_gaps = cpi_inflation(raw_path)
    return {
        "gdp_growth": gdp_growth(raw_path),
        "vix": vix_history(),
        "policy_rate_all": policy_all,
        "policy_rate_1": policy_1,
        "policy_rate_2": policy_2,
        "fx": forex_exchange(raw_path),
        "cds": cds(raw_path),
        "liquidity": liquidity(),
        "fsi": fsi,
        "fsi_missing": fsi_missing_values(fsi_path, fsi_cutoff),
        "bond_yields": sovereign_bond_yields(raw_path),
        "stock_index": stock_price_index(raw_path),
        "capital_flows": capital_flows(),
        "inflation_forecast": ppp_weighted_forecast("Inflation forecast", raw_path),
        "gdp_growth_forecast": ppp_weighted_forecast("GDP growth", raw_path),
        "cpi_yoy": cpi_yoy,
        "cpi_yoy_3mma": cpi_yoy_3mma,
        "cpi_gaps": cpi_gaps,
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Build every Capital Flow Monitor table and print a summary.")
    parser.add_argument("--raw", type=Path, default=RAW_PATH, help="path to 'raw data for CFM.xlsx'")
    parser.add_argument("--fsi", type=Path, default=FSI_PATH, help="path to '00 dlx.xlsx'")
    parser.add_argument(
        "--cutoff", default=FSI_CUTOFF, help="last month of the FSI, e.g. 2026-08 (default: automatic)"
    )
    args = parser.parse_args()

    tables = load_all(args.raw, args.fsi, args.cutoff)
    for name, table in tables.items():
        if isinstance(table, pd.DataFrame):
            print(f"{name:22} {table.shape}")
    print(f"\nFSI cutoff: {tables['fsi'].index.max():%b %Y}" + (" (automatic)" if args.cutoff is None else ""))
    print("FSI inputs filled up to the cutoff:")
    print(tables["fsi_missing"][["Sheet", "Series", "Last available", "Count"]].to_string(index=False)
          if len(tables["fsi_missing"]) else "  none")
    print("\nCPI months still missing a YoY:", tables["cpi_gaps"] or "none")
