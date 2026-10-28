"""
Uploads the Capital Flow Monitor tables (built by cfm_data.py) to Google Sheets.

Settings and secrets (sheet IDs, the service-account key file) come from the
.env file in this folder; see .env.example and README.md.

    python "data download.py"                       # build and upload everything
    python "data download.py" --dry-run             # build only, show what would be uploaded
    python "data download.py" --only fx cds         # upload just these worksheets
    python "data download.py" --xlsx backup.xlsx    # also save every table to a local workbook
    python "data download.py" --cutoff 2026-09      # override the automatic FSI cutoff

Every upload also refreshes the "today" worksheet of the main sheet: one row
per table with its size, the latest period it covers and when it was last
uploaded. The dashboard shows the upload time in its sidebar and the sizes in
its "Data pipeline" tab.
"""

import argparse
import os
import time
from collections import deque
from functools import cache
from pathlib import Path

import gspread
import pandas as pd
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe

import cfm_data as cfm

HERE = Path(__file__).resolve().parent
load_dotenv(HERE / ".env")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
# Google allows 60 requests per minute per user (reads and writes are counted
# separately). Stay under it with some headroom by pausing once this many
# requests have gone out in the last minute.
REQUESTS_PER_MINUTE = 50
# If Google still answers "too many requests" (429) or a server error, wait
# long enough for its one-minute window to reset, then try again.
RETRIES = 4
RETRY_WAIT_SECONDS = 65
# The worksheet in the main sheet that records what was uploaded and when.
TODAY_SHEET = "today"
TODAY_COLUMNS = ["Table", "Spreadsheet", "Rows", "Columns", "Covers to", "Last updated"]


def _env(name: str, required: bool = True) -> str | None:
    value = os.getenv(name, "").strip()
    if required and not value:
        raise SystemExit(f"{name} is not set. Add it to {HERE / '.env'} (see .env.example).")
    return value or None


# ---------------------------------------------------------------------------
# Which table goes to which worksheet
# ---------------------------------------------------------------------------

def build_worksheets(raw_path: Path, fsi_path: Path, cutoff: str) -> dict[str, dict]:
    """
    Worksheet layout of each Google Sheet: {sheet env var: {worksheet: builder}}.
    Builders run only when their worksheet is uploaded. The main sheet's
    worksheet names are the ones processed_data.py reads, so keep them.
    """

    @cache
    def fsi():
        return cfm.financial_stress_index(fsi_path, cutoff)

    @cache
    def policy():
        return cfm.policy_rate(raw_path)

    @cache
    def cpi():
        return cfm.cpi_inflation(raw_path)

    return {
        "SHEET_ID_MAIN": {
            "vix_data": cfm.vix_history,
            "policy_rate1": lambda: policy()[0],
            "fx": lambda: cfm.forex_exchange(raw_path),
            "cds": lambda: cfm.cds(raw_path),
            "liquidity": cfm.liquidity,
            "gdp_growth": lambda: cfm.gdp_growth(raw_path),
            "stock_price_index": lambda: cfm.stock_price_index(raw_path),
            "sovereign_bond_yields": lambda: cfm.sovereign_bond_yields(raw_path),
            "capital_flows": cfm.capital_flows,
            "inflation_forecast": lambda: cfm.ppp_weighted_forecast("Inflation forecast", raw_path),
            "gdp_growth_forecast": lambda: cfm.ppp_weighted_forecast("GDP growth", raw_path),
            "cpi_yoy": lambda: cpi()[0],
            "cpi_yoy_3mma": lambda: cpi()[1],
        },
        "SHEET_ID_FSI": {
            "fsi": lambda: fsi()[5],
            "yield_data_calc": lambda: fsi()[0],
            "stock_data_calc2": lambda: fsi()[1],
            "empi": lambda: fsi()[2],
            "financial_sector_beta": lambda: fsi()[3],
            "garch": lambda: fsi()[4],
            "fsi_missing": lambda: cfm.fsi_missing_values(fsi_path, cutoff),
        },
    }


def covers_to(df: pd.DataFrame) -> str:
    """The latest period a table covers, as text ("Aug 2026", "3Q2026",
    "2025-2027", ...), or "" for tables that aren't time series."""
    if isinstance(df.index, pd.DatetimeIndex):
        return f"{df.index.max():%b %Y}"
    if "Date" in df.columns:
        return f"{pd.to_datetime(df['Date']).max():%b %Y}"
    month_cols = [c for c in df.columns if isinstance(c, pd.Timestamp)]
    if month_cols:
        return f"{max(month_cols):%b %Y}"
    if "Quarter" in df.columns:
        return max(df["Quarter"].astype(str), key=lambda q: (int(q[2:]), int(q[0])))
    if "TIME_PERIOD" in df.columns:
        return str(df["TIME_PERIOD"].max())
    if "Year" in df.columns:
        return max(df["Year"].astype(str), key=lambda y: (int(y[-4:]), y))
    year_cols = [str(c) for c in df.columns if str(c).isdigit()]
    if year_cols:
        return f"{min(year_cols)}-{max(year_cols)}"
    return ""


def today_frame(meta: dict[str, dict], previous: pd.DataFrame | None = None) -> pd.DataFrame:
    """The "today" worksheet: `meta` rows for the tables just built, plus the
    previous rows of any table not rebuilt this time (e.g. with --only)."""
    rows = pd.DataFrame([{"Table": name, **m} for name, m in meta.items()], columns=TODAY_COLUMNS)
    if previous is not None and set(TODAY_COLUMNS) <= set(previous.columns):
        kept = previous.loc[~previous["Table"].isin(rows["Table"]), TODAY_COLUMNS]
        rows = pd.concat([rows, kept], ignore_index=True)
    # Counts go in as text too: as numbers, Sheets can display them in a date
    # format left on the worksheet (442 -> 1901-03-17).
    rows[["Rows", "Columns"]] = rows[["Rows", "Columns"]].astype(str)
    return rows


def to_sheet_frame(df: pd.DataFrame) -> pd.DataFrame:
    """
    Makes a table safe to write: a meaningful index becomes a column, date
    column headers become "YYYY-MM", date values "YYYY-MM-DD" (Sheets reads
    them as dates) and categories plain text. Missing values are left for
    gspread_dataframe to write as blank cells.
    """
    # An unnamed whole-number index is just row numbers (possibly shuffled by
    # sorting) and is dropped; anything else (dates, regions) becomes a column.
    if df.index.name is None and pd.api.types.is_integer_dtype(df.index):
        df = df.reset_index(drop=True)
    else:
        df = df.reset_index(names=df.index.name or "Region")
    df.columns = [c.strftime("%Y-%m") if isinstance(c, pd.Timestamp) else str(c) for c in df.columns]
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%Y-%m-%d")
        elif isinstance(df[col].dtype, pd.CategoricalDtype):
            df[col] = df[col].astype(str)
    return df


# ---------------------------------------------------------------------------
# Google Sheets
# ---------------------------------------------------------------------------

class _RateLimiter:
    """Keeps the number of API requests in any 60-second window at or below
    `per_minute`, sleeping when the next request would go over."""

    def __init__(self, per_minute: int):
        self.per_minute = per_minute
        self.sent = deque()  # monotonic times of recent requests

    def wait(self, cost: int = 1):
        while True:
            now = time.monotonic()
            while self.sent and now - self.sent[0] >= 60:
                self.sent.popleft()
            if len(self.sent) + cost <= self.per_minute:
                break
            pause = 60 - (now - self.sent[0]) + 0.5
            print(f"    pausing {pause:.0f}s to stay under Google's {self.per_minute}-per-minute limit")
            time.sleep(pause)
        self.sent.extend([time.monotonic()] * cost)


_limiter = _RateLimiter(REQUESTS_PER_MINUTE)


def _with_retry(action, what: str, cost: int = 1):
    """Runs one API action (`cost` = how many requests it makes), paced by the
    rate limiter, retrying rate-limit and server errors."""
    for attempt in range(1, RETRIES + 1):
        _limiter.wait(cost)
        try:
            return action()
        except gspread.exceptions.APIError as e:
            status = e.response.status_code
            if status not in (429, 500, 502, 503) or attempt == RETRIES:
                raise
            print(f"    Google API {status} on {what}; retrying in {RETRY_WAIT_SECONDS}s ({attempt}/{RETRIES - 1})")
            time.sleep(RETRY_WAIT_SECONDS)


def _write_frame(ws, df: pd.DataFrame, what: str, **options) -> None:
    """Overwrites a worksheet with `df`, then trims it to fit. Writing first
    (rather than clearing first) means a failed request leaves the old data
    in place instead of an empty worksheet; every cell of the new table is
    written, blanks included, so no old values survive inside it."""
    rows, cols = len(df) + 1, max(len(df.columns), 1)
    # Up to 2 requests: grow the sheet if it's too small, then write the cells.
    _with_retry(lambda: set_with_dataframe(ws, df, resize=False, **options), f"writing {what}", cost=2)
    if (ws.row_count, ws.col_count) != (rows, cols):
        _with_retry(lambda: ws.resize(rows=rows, cols=cols), f"trimming {what}")


def connect() -> gspread.Client:
    key_file = Path(_env("GOOGLE_CREDENTIALS_FILE"))
    if not key_file.is_absolute():
        key_file = HERE / key_file
    if not key_file.exists():
        raise SystemExit(f"Service-account key file not found: {key_file}")
    creds = Credentials.from_service_account_file(key_file, scopes=SCOPES)
    return gspread.authorize(creds)


def write_today(client: gspread.Client, sheet_id: str, meta: dict[str, dict]) -> pd.DataFrame:
    """Refreshes the "today" worksheet of the main sheet with `meta`, keeping
    the earlier rows of tables not uploaded this run. Written as plain text
    so Sheets doesn't turn "Aug 2026" or the timestamps into dates."""
    spreadsheet = _with_retry(lambda: client.open_by_key(sheet_id), "opening the spreadsheet")
    try:
        ws = spreadsheet.worksheet(TODAY_SHEET)
        previous = pd.DataFrame(_with_retry(ws.get_all_records, "reading today"))
    except gspread.exceptions.WorksheetNotFound:
        ws, previous = None, None
    today = today_frame(meta, previous)
    if ws is None:
        ws = _with_retry(
            lambda: spreadsheet.add_worksheet(title=TODAY_SHEET, rows=len(today) + 1, cols=len(TODAY_COLUMNS)),
            "adding today",
        )
    _write_frame(ws, today, TODAY_SHEET, string_escaping="full")
    print(f"     {TODAY_SHEET:24} {len(today)} rows (upload log)")
    return today


def upload(client: gspread.Client, sheet_id: str, tables: dict[str, pd.DataFrame], done: list[str]) -> None:
    """
    Writes each table to the worksheet of the same name, appending each
    finished name to `done`. An existing worksheet is overwritten in place
    (so its link and any references to it keep working); a missing one is
    created. Worksheets not in `tables` (e.g. economics_event) are left
    untouched.
    """
    spreadsheet = _with_retry(lambda: client.open_by_key(sheet_id), "opening the spreadsheet")
    print(f"  -> {spreadsheet.title}")
    existing = {ws.title: ws for ws in _with_retry(spreadsheet.worksheets, "listing worksheets")}

    for name, df in tables.items():
        rows, cols = len(df) + 1, max(len(df.columns), 1)
        ws = existing.get(name)
        if ws is None:
            ws = _with_retry(lambda: spreadsheet.add_worksheet(title=name, rows=rows, cols=cols), f"adding {name}")
        _write_frame(ws, df, name)
        done.append(name)
        print(f"     {name:24} {df.shape[0]} rows x {df.shape[1]} cols")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Upload the Capital Flow Monitor tables to Google Sheets.")
    parser.add_argument("--only", nargs="+", metavar="WORKSHEET", help="upload only these worksheets")
    parser.add_argument("--dry-run", action="store_true", help="build the tables but don't upload")
    parser.add_argument("--xlsx", type=Path, help="also save every built table to this Excel file")
    parser.add_argument("--raw", type=Path, default=cfm.RAW_PATH, help="path to 'raw data for CFM.xlsx'")
    parser.add_argument("--fsi", type=Path, default=cfm.FSI_PATH, help="path to '00 dlx.xlsx'")
    parser.add_argument(
        "--cutoff", default=cfm.FSI_CUTOFF, help="last month of the FSI, e.g. 2026-08 (default: automatic)"
    )
    args = parser.parse_args()

    layout = build_worksheets(args.raw, args.fsi, args.cutoff)
    all_names = [name for worksheets in layout.values() for name in worksheets]
    unknown = set(args.only or []) - set(all_names)
    if unknown:
        raise SystemExit(f"Unknown worksheet(s): {', '.join(sorted(unknown))}. Choose from: {', '.join(all_names)}")

    # Build every requested table first, so a data error stops the run before
    # anything in Google Sheets is touched.
    built = {}
    meta = {}  # per table: size and latest period, for the "today" worksheet
    now = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")
    for env_var, worksheets in layout.items():
        for name, builder in worksheets.items():
            if args.only and name not in args.only:
                continue
            print(f"Building {name} ...")
            df = builder()
            meta[name] = {
                "Spreadsheet": "main" if env_var == "SHEET_ID_MAIN" else "FSI",
                "Rows": df.shape[0], "Columns": df.shape[1], "Covers to": covers_to(df), "Last updated": now,
            }
            built.setdefault(env_var, {})[name] = to_sheet_frame(df)
    if "fsi" in meta:
        print(f"FSI cutoff: {meta['fsi']['Covers to']}" + (" (automatic)" if args.cutoff is None else ""))

    if args.xlsx:
        with pd.ExcelWriter(args.xlsx) as writer:
            for tables in built.values():
                for name, df in tables.items():
                    df.to_excel(writer, sheet_name=name[:31], index=False)
            today_frame(meta).to_excel(writer, sheet_name=TODAY_SHEET, index=False)
        print(f"Saved {sum(len(t) for t in built.values())} tables to {args.xlsx}")

    if args.dry_run:
        for env_var, tables in built.items():
            print(f"\n{env_var} would receive:")
            for name, df in tables.items():
                print(f"     {name:24} {df.shape[0]} rows x {df.shape[1]} cols")
        print(f"\nSHEET_ID_MAIN would also get the '{TODAY_SHEET}' upload log:")
        print(today_frame(meta).to_string(index=False))
        return

    client = connect()
    done = []
    try:
        for env_var, tables in built.items():
            print(f"\nUploading to {env_var}")
            upload(client, _env(env_var), tables, done)
        print(f"\nUpdating the upload log in SHEET_ID_MAIN")
        write_today(client, _env("SHEET_ID_MAIN"), meta)
    except Exception:
        # Say exactly what's left, so the run can be finished with --only.
        remaining = [name for tables in built.values() for name in tables if name not in done]
        print("\nUpload stopped partway. Worksheets not yet updated keep their previous data.")
        if done:
            print(f"Updated: {', '.join(done)}")
        if remaining:
            print(f"To finish, run:\n  python \"data download.py\" --only {' '.join(remaining)}")
        else:
            print("All tables were written; only the upload log failed. Rerun any --only to refresh it.")
        raise
    print("\nDone.")


if __name__ == "__main__":
    main()
