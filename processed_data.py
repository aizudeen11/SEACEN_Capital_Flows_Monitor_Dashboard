import pandas as pd
import requests
from io import BytesIO
# import concurrent.futures as cf  # MultiThreaded code processing
import time
import hashlib
from concurrent.futures import ThreadPoolExecutor
from cache_pandas import timed_lru_cache
# from data import Data

from pathlib import Path

path = Path(__file__).parent
path1 = path / "test.xlsx"
path2 = path / "financial stress index.xlsx"
path3 = path / "test (1).xlsx"
print(path2)

TPU_URL = "https://www.matteoiacoviello.com/tpu_files/tpu_web_latest.xlsx"
GEPU_URL = "https://www.policyuncertainty.com/media/Global_Policy_Uncertainty_Data.xlsx"
FSI_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSRMFoOOi8AXat4kH62KTNp56FbRc7XN9sSioYbfzOzPU0TDYWqkNqxhu8rWtnUy6erG9WVp9noknFp/pub?output=xlsx"
# Published ("File > Share > Publish to web") copy of the main dashboard sheet,
# the one cfm_data/"data download.py" uploads to as SHEET_ID_MAIN.
MAIN_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS5DDYL_MTv8b1AsgYvUEDuxiIFq81-CR9Yccc4sL1pvv1wD2GyHNabQQiixJsRpwaxG59Qyg-Asr11/pub?output=xlsx"

CACHE_DIR = path / ".cache"
CACHE_MAX_AGE_SECONDS = 24 * 60 * 60


def _fetch_cached(url):
    """Downloads `url`, reusing an on-disk copy younger than a day. If the
    download fails, falls back to a stale copy rather than breaking startup."""
    cache_file = CACHE_DIR / (hashlib.sha1(url.encode()).hexdigest()[:16] + ".xlsx")
    if cache_file.exists() and time.time() - cache_file.stat().st_mtime < CACHE_MAX_AGE_SECONDS:
        return cache_file.read_bytes()
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
    except requests.RequestException:
        if cache_file.exists():
            return cache_file.read_bytes()
        raise
    try:
        CACHE_DIR.mkdir(exist_ok=True)
        cache_file.write_bytes(response.content)
    except OSError:
        pass
    return response.content


def _prefetch(urls):
    """Warms the cache for all `urls` at once so the sequential loaders below
    read from disk instead of waiting on each download in turn. A failed
    download is ignored here; the loader that needs it fetches again and
    handles the failure itself."""

    def _try(url):
        try:
            _fetch_cached(url)
        except requests.RequestException:
            pass

    with ThreadPoolExecutor(max_workers=len(urls)) as pool:
        list(pool.map(_try, urls))


def _main_workbook():
    """The main dashboard workbook: the published Google Sheet (cached for a
    day), or the local snapshot "test (1).xlsx" if the sheet can't be
    downloaded and there is no cached copy."""
    try:
        return pd.ExcelFile(BytesIO(_fetch_cached(MAIN_URL)))
    except requests.RequestException:
        if not path3.exists():
            raise
        print(f"Main Google Sheet unavailable; using local snapshot {path3.name}")
        return pd.ExcelFile(path3)


def _upload_log(workbook):
    """The "today" sheet that cfm_data's uploader writes: one row per table
    (Table, Spreadsheet, Rows, Columns, Covers to, Last updated). Older
    uploads wrote just a "Last updated" column. None if the sheet is missing
    or empty."""
    try:
        today = pd.read_excel(workbook, sheet_name="today")
    except ValueError:  # no such sheet
        return None
    return None if today.empty else today


# Names as the charts show them. The sheets keep their source spellings
# ("South Korea", "Hong Kong SAR (China)"), and those uploaded before
# cfm_data renamed the Hong Kong / Korea / Singapore / Chinese Taipei group
# still use its earlier names.
CHART_NAMES = {
    "South Korea": "Korea",
    "Hong Kong SAR (China)": "Hong Kong",
    "Asia Advanced Economies": "Asian Advanced Economies",
    "Advanced Asian Economies": "Asian Advanced Economies",
}


def _chart_names(df):
    """Applies CHART_NAMES to column names and to a Region column if there
    is one."""
    df = df.rename(columns=CHART_NAMES)
    if "Region" in df.columns:
        df["Region"] = df["Region"].replace(CHART_NAMES)
    return df


def _switches(workbook):
    """The TRUE/FALSE switches in the "to_show" sheet: a header row
    ("developed", "excel_use") over one row of values. Returns
    (show the "Developed by" credit, include the CEIC/Haver-based sheets in
    the Excel download). A switch is True unless its cell says FALSE, so a
    missing sheet, column or value leaves things as they were. An older
    sheet with just one cell (A1) holds only the credit switch."""
    try:
        sheet = pd.read_excel(workbook, sheet_name="to_show", header=None)
    except ValueError:  # no such sheet
        return True, True
    if sheet.empty:
        return True, True
    header = [str(h).strip().lower() for h in sheet.iloc[0]]
    if "developed" not in header and "excel_use" not in header:
        values = {"developed": sheet.iat[0, 0]}
    else:
        values = dict(zip(header, sheet.iloc[1])) if len(sheet) > 1 else {}

    def on(name):
        return str(values.get(name, "")).strip().upper() != "FALSE"

    return on("developed"), on("excel_use")


def _last_updated(upload_log):
    """Time of the latest upload in the upload log, or None."""
    if upload_log is None or "Last updated" not in upload_log.columns:
        return None
    value = pd.to_datetime(upload_log["Last updated"], errors="coerce").max()
    return None if pd.isna(value) else value


def _load_tpu_monthly():
    """Trade Policy Uncertainty index (Caldara, Iacoviello, Molligo, Prestipino
    and Raffo), monthly series, aligned to month-end to match the VIX sheet's
    date convention. Public dataset, not part of the SEACEN workbooks, so both
    dfs() and dfs1() fetch it the same way."""
    tpu = pd.read_excel(BytesIO(_fetch_cached(TPU_URL)), sheet_name="TPU_MONTHLY")
    tpu = tpu[["DATE", "TPU"]].rename(columns={"DATE": "Date", "TPU": "Value"})
    tpu["Date"] = pd.to_datetime(tpu["Date"]) + pd.offsets.MonthEnd(0)
    return tpu


def _load_gepu_monthly():
    """Global Economic Policy Uncertainty index (Davis, 2016), current-price
    weighted version (GEPU_current, as opposed to the PPP-weighted variant),
    monthly, aligned to month-end. The sheet ends with a blank row and a
    citation row in the Year column, so both are dropped."""
    gepu = pd.read_excel(BytesIO(_fetch_cached(GEPU_URL)), sheet_name="Sheet1")
    gepu = gepu[pd.to_numeric(gepu["Year"], errors="coerce").notna()]
    gepu = gepu.dropna(subset=["Year", "Month"])
    year = gepu["Year"].astype(int)
    month = gepu["Month"].astype(int)
    gepu = gepu[["GEPU_current"]].rename(columns={"GEPU_current": "Value"})
    gepu["Date"] = pd.to_datetime(dict(year=year, month=month, day=1)) + pd.offsets.MonthEnd(0)
    return gepu[["Date", "Value"]]


def _forecast_long(df):
    """Forecast sheet (Region + one column per year, as uploaded by
    cfm_data) -> (Region, Year, Value), keeping the sheet's region order.
    Year headers may come back as text or numbers depending on how the
    workbook was saved."""
    df = df.rename(columns={c: str(c).split(".")[0] for c in df.columns[1:]})
    return df.melt(id_vars="Region", var_name="Year", value_name="Value")


def _monthly_long(df):
    """Monthly sheet (Region + one column per month, as uploaded by
    cfm_data) -> (Date, Region, Value) with Date at month-end. Month headers
    may be "YYYY-MM" text or dates (Google Sheets turns them into dates)."""
    df = df.melt(id_vars="Region", var_name="Date", value_name="Value")
    df["Date"] = pd.to_datetime(df["Date"].astype(str).str[:7], format="%Y-%m") + pd.offsets.MonthEnd(0)
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce")
    return df[["Date", "Region", "Value"]]


# vix, policy_rate1, policy_rate2, policy_rate3, fx, cds, liquidity, gdp_growth
@timed_lru_cache(seconds=None, maxsize=None)
def dfs():
    t1 = time.perf_counter()
    _prefetch([TPU_URL, GEPU_URL])
    vix_data = pd.read_excel(path1, sheet_name="vix_data")
    policy_rate1 = _chart_names(pd.read_excel(path1, sheet_name="policy_rate1"))
    fx = _chart_names(pd.read_excel(path1, sheet_name="fx"))
    fx['Year'] = fx['Year'].astype(str)
    cds = _chart_names(pd.read_excel(path1, sheet_name="cds"))
    liquidity = pd.read_excel(path1, sheet_name="liquidity")
    gdp_growth = _chart_names(pd.read_excel(path1, sheet_name="gdp_growth"))
    fsi = _chart_names(pd.read_excel(path2, sheet_name="fsi"))
    stock_price_index = _chart_names(pd.read_excel(path1, sheet_name="stock_price_index"))
    stock_price_index['Year'] = stock_price_index['Year'].astype(str)
    sovereign_bond_yields = _chart_names(pd.read_excel(path1, sheet_name="sovereign_bond_yields"))
    sovereign_bond_yields['Year'] = sovereign_bond_yields['Year'].astype(str)
    capital_flows = pd.read_excel(path1, sheet_name="capital_flows")
    tpu = _load_tpu_monthly()
    gepu = _load_gepu_monthly()
    economics_event = pd.read_excel(path1, sheet_name="economics_event")
    gdp_growth_forecast = _chart_names(_forecast_long(pd.read_excel(path1, sheet_name="gdp_growth_forecast")))
    inflation_forecast = _chart_names(_forecast_long(pd.read_excel(path1, sheet_name="inflation_forecast")))
    cpi_3mma = _chart_names(_monthly_long(pd.read_excel(path1, sheet_name="cpi_yoy_3mma")))
    upload_log = _upload_log(path1)
    last_updated = _last_updated(upload_log)
    show_credit, excel_full = _switches(path1)
    t2 = time.perf_counter()
    policy_rate2 = policy_rate1[
        policy_rate1["Region"].isin(
            [
                "United States",
                "Euro Area",
                "Mongolia",
                "Nepal",
                "Philippines",
                "Korea",
                "Sri Lanka",
                "Thailand",
                "China",
            ]
        )
    ]
    policy_rate3 = policy_rate1[
        policy_rate1["Region"].isin(
            ["Indonesia", "India", "Malaysia", "Chinese Taipei", "Vietnam"]
        )
    ]
    print(f"Read excel method:{t2 - t1} seconds")

    return (
        vix_data,
        policy_rate1,
        policy_rate2,
        policy_rate3,
        fx,
        cds,
        liquidity,
        gdp_growth,
        fsi,
        stock_price_index,
        sovereign_bond_yields,
        capital_flows,
        tpu,
        economics_event,
        gepu,
        gdp_growth_forecast,
        inflation_forecast,
        cpi_3mma,
        last_updated,
        upload_log,
        show_credit,
        excel_full,
    )

@timed_lru_cache(seconds=None, maxsize=None)
def dfs1():
    t1 = time.perf_counter()
    # The main dashboard tables come from the published main Google Sheet
    # (cached for a day, falling back to the local snapshot "test (1).xlsx"
    # if it can't be downloaded); FSI, TPU and GEPU are fetched live too.
    _prefetch([MAIN_URL, FSI_URL, TPU_URL, GEPU_URL])
    xl3 = _main_workbook()
    vix_data = pd.read_excel(xl3, sheet_name="vix_data")
    policy_rate1 = _chart_names(pd.read_excel(xl3, sheet_name="policy_rate1"))
    fx = _chart_names(pd.read_excel(xl3, sheet_name="fx"))
    fx['Year'] = fx['Year'].astype(str)
    cds = _chart_names(pd.read_excel(xl3, sheet_name="cds"))
    liquidity = pd.read_excel(xl3, sheet_name="liquidity")
    gdp_growth = _chart_names(pd.read_excel(xl3, sheet_name="gdp_growth"))
    fsi = _chart_names(pd.read_excel(BytesIO(_fetch_cached(FSI_URL)), sheet_name="fsi"))
    stock_price_index = _chart_names(pd.read_excel(xl3, sheet_name="stock_price_index"))
    stock_price_index['Year'] = stock_price_index['Year'].astype(str)
    sovereign_bond_yields = _chart_names(pd.read_excel(xl3, sheet_name="sovereign_bond_yields"))
    sovereign_bond_yields['Year'] = sovereign_bond_yields['Year'].astype(str)
    capital_flows = pd.read_excel(xl3, sheet_name="capital_flows")
    tpu = _load_tpu_monthly()
    gepu = _load_gepu_monthly()
    economics_event = pd.read_excel(xl3, sheet_name="economics_event")
    gdp_growth_forecast = _chart_names(_forecast_long(pd.read_excel(xl3, sheet_name="gdp_growth_forecast")))
    inflation_forecast = _chart_names(_forecast_long(pd.read_excel(xl3, sheet_name="inflation_forecast")))
    cpi_3mma = _chart_names(_monthly_long(pd.read_excel(xl3, sheet_name="cpi_yoy_3mma")))
    upload_log = _upload_log(xl3)
    last_updated = _last_updated(upload_log)
    show_credit, excel_full = _switches(xl3)
    t2 = time.perf_counter()
    policy_rate2 = policy_rate1[
        policy_rate1["Region"].isin(
            [
                "United States",
                "Euro Area",
                "Mongolia",
                "Nepal",
                "Philippines",
                "Korea",
                "Sri Lanka",
                "Thailand",
                "China",
            ]
        )
    ]
    policy_rate3 = policy_rate1[
        policy_rate1["Region"].isin(
            ["Indonesia", "India", "Malaysia", "Chinese Taipei", "Vietnam"]
        )
    ]
    print(f"Google Sheets + live fetch method:{t2 - t1} seconds")

    return (
        vix_data,
        policy_rate1,
        policy_rate2,
        policy_rate3,
        fx,
        cds,
        liquidity,
        gdp_growth,
        fsi,
        stock_price_index,
        sovereign_bond_yields,
        capital_flows,
        tpu,
        economics_event,
        gepu,
        gdp_growth_forecast,
        inflation_forecast,
        cpi_3mma,
        last_updated,
        upload_log,
        show_credit,
        excel_full,
    )

# @timed_lru_cache(seconds=None, maxsize=None)
# def dataF():
#     df = Data()
#     vix_data = df.vix_history()
#     policy_rate1, policy_rate2, policy_rate3 = df.policy_rate()
#     fx = df.forex_exchange()
#     cds = df.cds()
#     liquidity = df.liquidity()
#     gdp_growth = df.gdp_growth()
#     return (
#         vix_data,
#         policy_rate1,
#         policy_rate2,
#         policy_rate3,
#         fx,
#         cds,
#         liquidity,
#         gdp_growth,
#     )