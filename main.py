# shiny run --reload --launch-browser cfm-nlp/main.py
# shiny run --reload --launch-browser main.py
# SHIFT + END to go to end of line
import itertools
import re
from pathlib import Path

from markdown_it import MarkdownIt
from mdit_py_plugins.anchors import anchors_plugin
from shiny import App, ui, render, reactive
import pandas as pd
from processed_data import dfs, dfs1
from excel_export import build_capital_flow_workbook

import shinyswatch
from faicons import icon_svg
from shinywidgets import render_plotly, output_widget
import plotly.express as px
from plotly.subplots import make_subplots
from cache_pandas import timed_lru_cache


all_df = dfs1()  # publish use this
# all_df = dfs()

TODAY = pd.Timestamp.today().date()
TEN_YEARS_AGO = (pd.Timestamp.today() - pd.DateOffset(years=10)).date()

# ---------------------------------------------------------------------------
# Design tokens
#
# Validated categorical palette (fixed hue order, CVD-safe on adjacent pairs)
# plus the chart chrome/status colors it ships with. See the project's
# dataviz skill reference for how this set was derived/validated.
# ---------------------------------------------------------------------------
CATEGORICAL_PALETTE = [
    "#2a78d6",  # blue
    "#eb6834",  # orange
    "#1baf7a",  # aqua
    "#eda100",  # yellow
    "#e87ba4",  # magenta
    "#008300",  # green
    "#4a3aa7",  # violet
    "#e34948",  # red
]
GRID_COLOR = "#e1e0d9"
AXIS_COLOR = "#c3c2b7"
MUTED_TEXT = "#898781"
PRIMARY_INK = "#0b0b0b"
GOOD_COLOR = "#0ca30c"
CRITICAL_COLOR = "#d03b3b"

px.defaults.color_discrete_sequence = CATEGORICAL_PALETTE
px.defaults.template = "plotly_white"


def _entity_color_map(*dataframes, column="Region", palette=CATEGORICAL_PALETTE):
    """Assign a fixed palette slot per distinct entity so the same region/year
    keeps the same color across every chart it appears in. `palette` must have
    at least as many colors as the largest simultaneous series count across
    the charts sharing this map, or two series in the same chart will collide."""
    entities = sorted(set().union(*(set(df[column].unique()) for df in dataframes)))
    cycle = itertools.cycle(palette)
    return {entity: next(cycle) for entity in entities}


# Policy Rate (up to 14 regions at once), CDS (9) and GDP Growth (5) share a
# region-name space with up to 14 simultaneous series in one chart — more than
# the validated 8-hue categorical set can keep collision-free, so this map
# draws from Plotly's larger qualitative set instead (still one fixed color
# per region, just without the CVD-validation guarantee of the 8-hue set).
REGION_COLORS = _entity_color_map(
    all_df[1], all_df[5], all_df[7], palette=px.colors.qualitative.Dark24
)
FSI_REGION_COLORS = dict(
    zip(["Asia", "China", "India", "ASEAN-4", "Asian Advanced Economies"], CATEGORICAL_PALETTE)
)
# Shared across the three year-over-year "changes" charts (FX, Stock Price
# Index, Sovereign Bond Yields) so the same year keeps the same color in all
# three.
YEAR_COLORS = dict(
    zip(
        sorted(set(all_df[4]["Year"].unique()) | set(all_df[9]["Year"].unique()) | set(all_df[10]["Year"].unique())),
        CATEGORICAL_PALETTE,
    )
)
# The year filters of those three charts offer only the latest
# CHANGE_YEARS_SHOWN years, newest first. Years look like "2025" or "YTD2026".
CHANGE_YEARS_SHOWN = 5


def _latest_years(df):
    return sorted(df["Year"].unique(), key=lambda y: int(y[-4:]), reverse=True)[:CHANGE_YEARS_SHOWN]


FX_YEARS, STOCK_YEARS, BOND_YEARS = (_latest_years(all_df[i]) for i in (4, 9, 10))
FLOW_COLORS = {"Portfolio Equity": CATEGORICAL_PALETTE[0], "Portfolio Debt": CATEGORICAL_PALETTE[1]}
# Forecast years (last, this, next) in the reference chart's yellow / blue /
# green order, shared by the GDP growth and inflation forecast charts.
FORECAST_YEAR_COLORS = dict(
    zip(
        sorted(set(all_df[15]["Year"]) | set(all_df[16]["Year"])),
        [CATEGORICAL_PALETTE[3], CATEGORICAL_PALETTE[0], CATEGORICAL_PALETTE[2]],
    )
)
# Up to 17 economies in one inflation chart, so (like REGION_COLORS) this
# draws from Plotly's larger qualitative set.
CPI_REGION_COLORS = _entity_color_map(all_df[17], palette=px.colors.qualitative.Dark24)
# The Headline Inflation card shows CPI from this month on (earlier months
# only exist for the few economies already reported as year-on-year).
CPI_START = pd.Timestamp("2024-01-01")
CPI_DATA = all_df[17][all_df[17]["Date"] >= CPI_START]
# When the dashboard tables were last uploaded to the Google Sheet (its
# "today" sheet); None if unknown.
DATA_LAST_UPDATED = all_df[18]
# The full upload log ("today" sheet: one row per table with its size, the
# latest period it covers and its upload time); None if unavailable.
UPLOAD_LOG = all_df[19]
# Whether the navbar shows the "Developed by" credit: the TRUE/FALSE cell of
# the sheet's "to_show" sheet (shown if that sheet is missing).
SHOW_CREDIT = all_df[20]
# Whether the Excel download includes the charts built from CEIC/Haver data
# (EXCEL_RESTRICTED): the "excel_use" TRUE/FALSE cell of the "to_show" sheet
# (included if that sheet or cell is missing).
EXCEL_FULL = all_df[21]
EXCEL_RESTRICTED = [
    "Policy Rate", "5-Year Sovereign CDS", "Headline Inflation", "GDP Growth and Inflation Outlook",
]
if EXCEL_FULL:
    EXCEL_NOTE = "The Excel download includes every chart on the dashboard."
else:
    EXCEL_NOTE = (
        f"The Excel download leaves out {', '.join(EXCEL_RESTRICTED[:-1])} and {EXCEL_RESTRICTED[-1]}, "
        "as these use data from CEIC and Haver Analytics. Every other chart is included."
    )
# The GDP Growth card shows only the latest GDP_QUARTERS_SHOWN quarters.
# Quarters are labelled like "3Q2026", so sort them by (year, quarter).
GDP_QUARTERS_SHOWN = 9
GDP_QUARTERS = sorted(all_df[7]["Quarter"].unique(), key=lambda q: (int(q[2:]), int(q[0])))[-GDP_QUARTERS_SHOWN:]
GDP_DATA = all_df[7][all_df[7]["Quarter"].isin(GDP_QUARTERS)]
# Economies whose inflation runs far above the rest; plotted on the right-hand
# axis so the others stay readable on the left.
CPI_RHS_REGIONS = ["Laos", "Sri Lanka"]

# Default event markers, shared by every chart with a "Show event markers"
# button, sourced from the "economics_event" sheet (columns: type
# ["range"|"line"], start, end, label). "range" events render as a shaded
# band (like the reference chart's tariff-increase box); "line" events
# render as a single dashed marker.
DEFAULT_EVENTS = all_df[13].to_dict("records")


def style_fig(fig, hovermode="closest", legend="top", bottom_space=70):
    """Apply shared chart chrome so every figure in the dashboard reads as one
    system: transparent background (blends into the card), muted gridlines,
    and consistent typography.

    `legend` picks where a multi-series legend goes: "top" stacks it above the
    plot (room reserved in the top margin so it never collides with the title;
    fine up to ~6 entries on one row), "right" uses a vertical list beside the
    plot instead (for charts with many simultaneous series, where a top legend
    would wrap across several rows), "bottom" centres a horizontal legend
    under the chart (anchored to the figure's bottom edge, so long x-axis
    labels never collide with it), and "none" leaves Plotly's default
    (single-series charts get no legend box at all).

    For a "bottom" legend, `bottom_space` (px) is the room below the plot for
    the x-axis tick labels plus the legend; charts with an x-axis title need
    more (about 100).
    """
    layout = dict(
        font=dict(family="Segoe UI, system-ui, -apple-system, sans-serif", color=PRIMARY_INK, size=13),
        title=dict(font=dict(size=15, color=PRIMARY_INK)),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        hovermode=hovermode,
        hoverlabel=dict(bgcolor="white", font=dict(color=PRIMARY_INK)),
    )
    if legend == "top":
        layout["margin"] = dict(l=10, r=10, t=90, b=10)
        layout["legend"] = dict(orientation="h", yanchor="top", y=1.32, xanchor="left", x=0, title=None)
    elif legend == "bottom":
        layout["margin"] = dict(l=10, r=10, t=48, b=bottom_space)
        layout["legend"] = dict(
            orientation="h", yref="container", yanchor="bottom", y=0, xanchor="center", x=0.5, title=None
        )
    elif legend == "right":
        layout["margin"] = dict(l=10, r=10, t=48, b=10)
        layout["legend"] = dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02, title=None)
    else:
        layout["margin"] = dict(l=10, r=10, t=48, b=10)
    fig.update_layout(**layout)
    fig.update_xaxes(showgrid=False, linecolor=AXIS_COLOR, tickfont=dict(color=MUTED_TEXT))
    fig.update_yaxes(showgrid=True, gridcolor=GRID_COLOR, zerolinecolor=AXIS_COLOR, tickfont=dict(color=MUTED_TEXT))
    return fig


def _make_yoy_diff(df):
    """Build a cached year-over-year diff accessor for a (Year, Region, Value)
    sheet, shared by the FX, Stock Price Index, and Sovereign Bond Yields
    "changes" cards."""

    @timed_lru_cache(seconds=None, maxsize=None)
    def _diff():
        d = df.copy()
        years = sorted(d["Year"].unique())[-2:]
        d = d[d["Year"].isin(years)]
        d["diff"] = d.groupby("Region")["Value"].diff(-1)
        d.dropna(inplace=True)
        d = d.iloc[::-1]
        d["diff"] = d["diff"].round(2)
        return d, years

    return _diff


ex_rate_diff = _make_yoy_diff(all_df[4])
stock_price_diff = _make_yoy_diff(all_df[9])
bond_yield_diff = _make_yoy_diff(all_df[10])


def diff_table_html(diff_fn):
    diff_df, years = diff_fn()
    html = f"""
    <table class='table table-sm table-hover'>
        <thead class='table-light'>
            <tr>
                <th>Region</th>
                <th>Diff: ({years[1]} - {years[0]})</th>
            </tr>
        </thead>
        <tbody>
    """
    for _, row in diff_df.iterrows():
        is_up = row["diff"] > 0
        color = GOOD_COLOR if is_up else CRITICAL_COLOR
        arrow = "&#9650;" if is_up else "&#9660;"
        html += (
            f"<tr><td>{row['Region']}</td>"
            f"<td style='color:{color};font-weight:600;'>{arrow} {row['diff']}</td></tr>"
        )
    html += "</tbody></table>"
    return ui.HTML(html)


def card_header(icon_name, text):
    return ui.card_header(icon_svg(icon_name), text, class_="d-flex align-items-center gap-2 fw-semibold")


def card_notes(note=None, source=None):
    """A card footer with the note (if any) and, on its own line below, the
    source (if any)."""
    parts = []
    if note is not None:
        parts.append(ui.div(note))
    if source is not None:
        parts.append(ui.div(source, class_="mt-1" if parts else None))
    return ui.card_footer(*parts, class_="text-muted small")


# The project's README files, shown as tabs next to the dashboard. They are
# read once at startup, so an edited README shows after the app restarts.
APP_DIR = Path(__file__).resolve().parent
DOC_TABS = [
    ("About", "book-open", APP_DIR / "README.md"),
    ("Data pipeline", "database", APP_DIR / "cfm_data" / "README.md"),
]
# GitHub-style Markdown (tables etc.) with heading ids, so the READMEs'
# "Contents" links (#section-name) jump within the page.
_MARKDOWN = MarkdownIt("gfm-like").use(anchors_plugin, max_level=3)
DOC_CSS = """
.cfm-doc { max-width: 1100px; }
.cfm-doc h1 { font-size: 1.8rem; margin-bottom: 1rem; }
.cfm-doc h2 { font-size: 1.4rem; margin-top: 2rem; padding-bottom: .3rem; border-bottom: 1px solid #dee2e6; }
.cfm-doc h3 { font-size: 1.15rem; margin-top: 1.5rem; }
.cfm-doc table { margin-bottom: 1rem; border-collapse: collapse; }
.cfm-doc th, .cfm-doc td { border: 1px solid #dee2e6; padding: .35rem .6rem; vertical-align: top; }
.cfm-doc th { background: #f6f8fa; }
.cfm-doc pre { background: #f6f8fa; padding: .75rem 1rem; border-radius: 6px; overflow-x: auto; }
.cfm-doc code { color: #b4235a; }
.cfm-doc pre code { color: inherit; }
"""


def _as_count(value):
    """A row/column count from the upload log. If Sheets formatted the cell as
    a date, the count comes back as a date (442 -> 1901-03-17), so convert it
    back to the day number it encodes. (Excel counts a non-existent
    29 Feb 1900, so numbers below 60 sit one day later than the rest.)"""
    if hasattr(value, "to_pydatetime"):
        stamp = pd.Timestamp(value)
        days = (stamp - pd.Timestamp("1899-12-30")).days
        return days - 1 if stamp < pd.Timestamp("1900-03-01") else days
    return int(float(value))


def _live_value_or_none(table, field):
    """_live_value, but an unreadable log entry falls back to the README's own
    text rather than stopping the app."""
    try:
        return _live_value(table, field)
    except (TypeError, ValueError, KeyError) as e:
        print(f"Upload log entry {table}:{field} unreadable ({e}); using the README's value")
        return None


def _live_value(table, field):
    """A value from the upload log for a README "live" marker, as text, or
    None if the log doesn't have it. `table` "*" means the whole upload."""
    if table == "*":
        return DATA_LAST_UPDATED.strftime("%d %b %Y") if field == "updated" and DATA_LAST_UPDATED is not None else None
    if UPLOAD_LOG is None or "Table" not in UPLOAD_LOG.columns:
        return None
    rows = UPLOAD_LOG[UPLOAD_LOG["Table"] == table]
    if rows.empty:
        return None
    row = rows.iloc[0]
    if field == "shape":
        return f"{_as_count(row['Rows'])} × {_as_count(row['Columns'])}"
    value = row["Covers to"] if field == "covers" else row["Last updated"]
    if pd.isna(value) or value == "":
        return None
    if field == "updated":
        return pd.to_datetime(value).strftime("%d %b %Y")
    # Sheets may have stored a month as a date.
    return value.strftime("%b %Y") if hasattr(value, "strftime") else str(value)


def doc_page(path):
    """A README rendered as a page. Links to files in the project become
    plain text (the app doesn't serve those files); web, email and
    same-page section links keep working. Values marked
    <!--live:TABLE:FIELD-->default<!--/live--> (FIELD: shape, covers or
    updated) are replaced with the latest ones from the upload log; the
    default text is kept when the log doesn't have them."""
    if not path.exists():
        return ui.p(f"{path.name} was not found.", class_="text-muted p-4")
    text = path.read_text(encoding="utf-8")
    text = re.sub(
        r"<!--live:([\w*]+):(\w+)-->(.*?)<!--/live-->",
        lambda m: _live_value_or_none(m[1], m[2]) or m[3],
        text,
    )
    text = re.sub(r"\[([^\]]+)\]\((?!https?://|#|mailto:)[^)]*\)", r"\1", text)
    return ui.div(ui.HTML(_MARKDOWN.render(text)), class_="cfm-doc container-fluid py-3 px-4")


sidebar = ui.sidebar(
    ui.div(
        ui.h5("SEACEN Capital Flow Monitor", class_="mb-1 fw-bold"),
        ui.p(
            icon_svg("clock-rotate-left"),
            " Data last refreshed: ",
            ui.span(
                DATA_LAST_UPDATED.strftime("%d %b %Y, %H:%M") if DATA_LAST_UPDATED is not None else "unknown",
                class_="fw-semibold",
            ),
            class_="small mb-2",
        ),
        ui.p(
            "Cross-market indicators for SEACEN economies: volatility, policy rates, "
            "exchange rates, stock and bond markets, sovereign credit risk, portfolio "
            "flows, growth, and financial stress.",
            class_="text-muted small mb-3",
        ),
        ui.div(
            icon_svg("magnifying-glass-minus"),
            ui.strong(" Viewing tip: "),
            "the dashboard is laid out for a large monitor. On a laptop screen, if charts look "
            "cramped, zoom the page out (Ctrl and − on Windows, ⌘ and − on a Mac; Ctrl and 0 "
            "resets it).",
            class_="alert alert-info small py-2 px-2 mb-0",
        ),
    ),
    ui.hr(),
    # What the download contains (EXCEL_NOTE) sits above the button, with the
    # same text as a tooltip on the button.
    ui.p(
        icon_svg("circle-check" if EXCEL_FULL else "circle-exclamation"),
        " All charts included" if EXCEL_FULL else " Some charts not included",
        ui.tooltip(
            ui.span(icon_svg("circle-question"), class_="ms-1", style="cursor: help;"),
            EXCEL_NOTE,
        ),
        class_="small fw-semibold mb-1 " + ("text-success" if EXCEL_FULL else "text-warning"),
    ),
    ui.tooltip(
        ui.download_button(
            "download_excel",
            ui.span(icon_svg("file-excel"), " Download figures (Excel)"),
            class_="btn-sm btn-success w-100",
        ),
        EXCEL_NOTE,
        placement="right",
    ),
    ui.p(
        "Exports each included chart's current data, with native (editable) Excel "
        "charts, matching whatever filters are selected below - including "
        "any active event markers.",
        class_="text-muted small mt-1",
    ),
    ui.hr(),
    ui.p(
        icon_svg("circle-info"),
        " Each panel has its own date, region, or year controls - use them to "
        "narrow the chart beside them.",
        class_="text-muted small",
    ),
    ui.hr(),
    ui.h6("Economic Event Markers", class_="fw-bold mb-2"),
    ui.p(
        "Each chart has its own “Show event markers” button - click it "
        "to toggle markers on just that chart.",
        class_="text-muted small",
    ),
    ui.input_date(
        "custom_event_date",
        "Add event marker - date",
        value=all_df[0]["Date"].max().date(),
    ),
    ui.input_text(
        "custom_event_label",
        "Event label",
        placeholder="e.g. Middle East conflict escalation",
    ),
    ui.input_action_button(
        "add_event_btn", "Add marker", class_="btn-sm btn-outline-primary w-100 mb-2"
    ),
    ui.input_action_button(
        "clear_events_btn", "Clear custom markers", class_="btn-sm btn-outline-secondary w-100 mb-3"
    ),
    ui.accordion(
        ui.accordion_panel("Event details", ui.output_ui("event_details"), value="event_details"),
        id="event_details_accordion",
        open=False,
    ),
    ui.hr(),
    ui.h6("Contact", class_="fw-bold mb-2"),
    ui.p(
        "For comments and questions, please contact the SEACEN Research Support Unit: ",
        ui.tags.a("rsu@seacen.org", href="mailto:rsu@seacen.org"),
        class_="small",
    ),
    ui.p(
        ui.strong("Disclaimer: "),
        "The views expressed in this dashboard are those of The SEACEN Centre and do not "
        "necessarily represent those of its member central banks/monetary authorities. "
        "See the About tab for notes and abbreviations.",
        class_="text-muted small",
    ),
    title="About this dashboard",
    width=300,
)

dashboard_page = ui.layout_sidebar(
    sidebar,
    ui.layout_columns(
        ui.card(
            card_header("gauge-high", "Volatility Index (VIX)"),
            ui.layout_columns(
                ui.input_date_range(
                    "date_range",
                    "Select Date Range:",
                    start=all_df[0]["Date"].min(),
                    end=TODAY,
                    max=TODAY,
                ),
                ui.input_action_button(
                    "toggle_markers_vix",
                    ui.span(icon_svg("flag"), "Show event markers"),
                    class_="btn-sm btn-outline-secondary mb-2",
                ),
                output_widget("hist"),
                col_widths={"sm": (12, 12, 12)},
                row_heights=["auto", "auto", 1],
            ),
            card_notes(source="Source: Chicago Board Options Exchange (Cboe)."),
            full_screen=True,
        ),
        ui.card(
            card_header("globe", "Economic and Trade Policy Uncertainty Indices"),
            ui.layout_columns(
                ui.input_date_range(
                    "date_range_pu",
                    "Select Date Range:",
                    start=TEN_YEARS_AGO,
                    end=TODAY,
                    max=TODAY,
                ),
                ui.input_radio_buttons(
                    "pu_mode",
                    "Show",
                    choices=["TPU only", "GEPU only", "Both"],
                    selected="Both",
                    inline=True,
                ),
                ui.input_action_button(
                    "toggle_markers_pu",
                    ui.span(icon_svg("flag"), "Hide event markers"),
                    class_="btn-sm btn-outline-secondary mb-2",
                ),
                output_widget("hist10"),
                col_widths={"sm": (8, 4, 12, 12)},
                row_heights=["auto", "auto", 1],
            ),
            card_notes(
                note="Note: Economic policy uncertainty index refers to global uncertainty. Trade policy "
                "uncertainty computed following Caldara, Iacoviello, Molligo, Prestipino and Raffo (online). "
                "Higher values mean higher uncertainty.",
                source=ui.TagList(
                    "Source: Economic Policy Uncertainty (",
                    ui.tags.a(
                        "https://www.policyuncertainty.com/trade_cimpr.html",
                        href="https://www.policyuncertainty.com/trade_cimpr.html",
                        target="_blank",
                        rel="noopener noreferrer",
                    ),
                    ").",
                ),
            ),
            full_screen=True,
        ),
        ui.card(
            card_header("landmark", "Policy Rates in"),
            ui.layout_columns(
                ui.input_date_range(
                    "date_range2",
                    "Select Date Range:",
                    start=all_df[1]["Date"].min(),
                    end=all_df[1]["Date"].max(),
                ),
                ui.input_selectize(
                    "var",
                    "Policy Rate",
                    choices=[
                        "All region",
                        "Selected region 1",
                        "Selected region 2",
                    ],
                ),
                ui.input_action_button(
                    "toggle_markers_policy",
                    ui.span(icon_svg("flag"), "Show event markers"),
                    class_="btn-sm btn-outline-secondary mb-2",
                ),
                output_widget("hist1"),
                col_widths={"sm": (6, 6, 12, 12)},
                row_heights=["auto", "auto", 1],
            ),
            card_notes(
                note="Note: The policy rate for the United States refers to the effective Fed Funds rate. "
                "Data for China pertains to the one-year loan prime rate sourced from the Bank for "
                "International Settlements (BIS) Data Portal. The policy rate for the Euro Area is the "
                "main refinancing fixed rate of the European Central Bank.",
                source="Source: Data taken from national sources accessed through CEIC Database and BIS Data Portal.",
            ),
            full_screen=True,
        ),
        ui.card(
            card_header("money-bill-trend-up", "Exchange Rate Changes"),
            ui.layout_columns(
                ui.input_checkbox_group(
                    "ex_rate",
                    "Select years",
                    choices=FX_YEARS,
                    selected=FX_YEARS[:2],
                    inline=True,
                ),
                ui.input_radio_buttons(
                    "ex_rate2",
                    "Select Order by Year",
                    choices=FX_YEARS,
                    selected=FX_YEARS[0],
                    inline=True,
                ),
                output_widget("hist2"),
                ui.tags.div(ui.output_ui("styled_table")),
                col_widths={"sm": (6, 6, 8, 4)},
                row_heights=["auto", 1],
            ),
            card_notes(
                note="Note: Year-to-date values are computed as the monthly difference between the first and "
                "last data points within a year. Positive changes refer to an appreciation of the local "
                "currency versus the U.S. dollar, and negative changes refer to depreciation.",
                source="Source: SEACEN staff calculations using data accessed through the CEIC Database.",
            ),
            full_screen=True,
        ),
        ui.card(
            card_header("arrow-trend-up", "Changes in Benchmark Stock Price Indices"),
            ui.layout_columns(
                ui.input_checkbox_group(
                    "stock_year",
                    "Select years",
                    choices=STOCK_YEARS,
                    selected=STOCK_YEARS[:2],
                    inline=True,
                ),
                ui.input_radio_buttons(
                    "stock_year2",
                    "Select Order by Year",
                    choices=STOCK_YEARS,
                    selected=STOCK_YEARS[0],
                    inline=True,
                ),
                output_widget("hist7"),
                ui.tags.div(ui.output_ui("styled_table_stock")),
                col_widths={"sm": (6, 6, 8, 4)},
                row_heights=["auto", 1],
            ),
            card_notes(
                note="Note: Year-to-date values are computed as the monthly difference between the first and "
                "last data points within a year.",
                source="Source: SEACEN staff calculations using data accessed through the CEIC Database.",
            ),
            full_screen=True,
        ),
        ui.card(
            card_header("percent", "Changes in Sovereign Bond Yields"),
            ui.layout_columns(
                ui.input_checkbox_group(
                    "bond_year",
                    "Select years",
                    choices=BOND_YEARS,
                    selected=BOND_YEARS[:2],
                    inline=True,
                ),
                ui.input_radio_buttons(
                    "bond_year2",
                    "Select Order by Year",
                    choices=BOND_YEARS,
                    selected=BOND_YEARS[0],
                    inline=True,
                ),
                output_widget("hist8"),
                ui.tags.div(ui.output_ui("styled_table_bond")),
                col_widths={"sm": (6, 6, 8, 4)},
                row_heights=["auto", 1],
            ),
            card_notes(
                note="Note: Year-to-date values are computed as the monthly difference between the first and "
                "last data points within a year.",
                source="Source: SEACEN staff calculations using data accessed from Haver Analytics.",
            ),
            full_screen=True,
        ),
        ui.card(
            card_header("money-bill-transfer", "Nonresident Portfolio Flows"),
            ui.layout_columns(
                ui.input_date_range(
                    "date_range4",
                    "Select Date Range:",
                    start=TEN_YEARS_AGO,
                    end=all_df[11]["Date"].max(),
                ),
                ui.input_action_button(
                    "toggle_markers_portfolio",
                    ui.span(icon_svg("flag"), "Show event markers"),
                    class_="btn-sm btn-outline-secondary mb-2",
                ),
                output_widget("hist9"),
                col_widths={"sm": (12, 12, 12)},
                row_heights=["auto", "auto", 1],
            ),
            ui.card_footer(
                "Notes: The sample for nonresident portfolio equity flows includes China, India, Indonesia, Korea, Malaysia, Mongolia, the Philippines, Sri Lanka, Chinese Taipei, Thailand, and Vietnam. The sample for nonresident portfolio debt flows includes China, India, Indonesia, Korea, Malaysia, Mongolia, the Philippines, and Thailand.",
                class_="text-muted small",
            ),
            full_screen=True,
        ),
        ui.card(
            card_header("shield-halved", "5-Year Sovereign Credit Default Swap"),
            ui.layout_columns(
                ui.input_date_range(
                    "date_range3",
                    "Select Date Range:",
                    start=all_df[5]["Date"].min(),
                    end=all_df[5]["Date"].max(),
                ),
                ui.input_action_button(
                    "toggle_markers_cds",
                    ui.span(icon_svg("flag"), "Show event markers"),
                    class_="btn-sm btn-outline-secondary mb-2",
                ),
                output_widget("hist3"),
                col_widths={"sm": (12, 12, 12)},
                row_heights=["auto", "auto", 1],
            ),
            card_notes(
                note="Note: 5-Year USD Credit Default Swap par mid-rate in basis points.",
                source="Source: CMA Datavision and Haver Analytics, downloaded from Haver.",
            ),
            full_screen=True,
        ),
        ui.card(
            card_header("chart-column", "Quarterly GDP Growth"),
            ui.layout_columns(
                ui.input_checkbox_group(
                    "gdp_quarterly",
                    "GDP Quarterly",
                    choices=GDP_QUARTERS,
                    selected=GDP_QUARTERS,
                    inline=True,
                ),
                output_widget("hist5"),
                col_widths={"sm": (12, 12)},
                row_heights=["auto", 1],
            ),
            card_notes(
                note="Note: Regional growth rates are weighted averages of individual growth rates, using GDP "
                "in PPP as weights. Asia Economies include China, India, Mongolia, ASEAN-5 (Indonesia, "
                "Malaysia, Philippines, Thailand, and Vietnam), and Asian Advanced Economies (Hong Kong, "
                "China; Korea; Singapore; and Chinese Taipei).",
                source="Source: SEACEN staff calculations using data from national sources accessed through "
                "the CEIC Database.",
            ),
            full_screen=True,
        ),
        ui.card(
            card_header("binoculars", "GDP Growth and Inflation Outlook"),
            ui.layout_columns(
                output_widget("forecast_gdp"),
                output_widget("forecast_inflation"),
                col_widths={"sm": (12, 12), "lg": (6, 6)},
            ),
            card_notes(
                # One note per chart, labelled; one shared source.
                note=ui.TagList(
                    ui.div(
                        ui.strong("GDP Growth Outlook. "),
                        "Note: Values for regional growth rates were weighted averages of individual growth "
                        "rates, using GDP in PPP current international dollars ($) as weights. Asia-Pacific "
                        "economies include China, India, Asian Advanced Economies (Hong Kong, China; Korea; "
                        "Singapore; and Chinese Taipei), ASEAN5 (Indonesia, Malaysia, Philippines, Thailand, "
                        "and Vietnam), and Asian EDMEs (Brunei Darussalam, Cambodia, Lao PDR, Mongolia, "
                        "Myanmar, and Sri Lanka).",
                    ),
                    ui.div(
                        ui.strong("Inflation Outlook. "),
                        "Note: Values for regional inflation were weighted averages of individual inflation, "
                        "using GDP in PPP current international dollars ($) as weights. Asia-Pacific economies "
                        "include China, India, Asian Advanced Economies (Hong Kong, China; Korea; Singapore; "
                        "and Chinese Taipei), ASEAN5 (Indonesia, Malaysia, Philippines, Thailand, and Vietnam), "
                        "and Asian EDMEs (Brunei Darussalam, Cambodia, Lao PDR, Mongolia, Myanmar, and "
                        "Sri Lanka).",
                        class_="mt-1",
                    ),
                ),
                source="Source: SEACEN staff calculations using data from Focus Economics consensus forecasts "
                "accessed through Haver Analytics.",
            ),
            full_screen=True,
        ),
        ui.card(
            card_header("cart-shopping", "Monthly Inflation"),
            ui.layout_columns(
                ui.input_date_range(
                    "date_range_cpi",
                    "Select Date Range:",
                    start=CPI_START,
                    end=CPI_DATA["Date"].max(),
                    min=CPI_START,
                    max=CPI_DATA["Date"].max(),
                ),
                ui.input_selectize(
                    "cpi_regions",
                    "Economies:",
                    choices=list(CPI_DATA["Region"].unique()),
                    selected=list(CPI_DATA["Region"].unique()),
                    multiple=True,
                ),
                output_widget("cpi_chart"),
                col_widths={"sm": (12, 12, 12), "lg": (4, 8, 12)},
                row_heights=["auto", 1],
            ),
            card_notes(
                note="Note: Year-on-year change in the consumer price index, 3-month moving average. "
                f"{' and '.join(CPI_RHS_REGIONS)} (dashed lines) are on the right-hand axis.",
                source="Source: SEACEN staff calculations using data from national sources accessed through "
                "the CEIC Database.",
            ),
            full_screen=True,
        ),
        ui.card(
            card_header("heart-pulse", "Financial Stress Indices"),
            ui.layout_columns(
                ui.input_action_button(
                    "toggle_markers_fsi",
                    ui.span(icon_svg("flag"), "Show event markers"),
                    class_="btn-sm btn-outline-secondary mb-2",
                ),
                output_widget("hist6"),
                col_widths={"sm": (12, 12)},
                row_heights=["auto", 1],
            ),
            ui.card_footer(
                ui.div(
                    "Note: Asian Advanced Economies include Hong Kong, China; Korea; Singapore; "
                    "and Chinese Taipei. ASEAN4 includes Indonesia, Malaysia, Philippines, and "
                    "Thailand. Financial stress indices for Asian Advanced Economies and ASEAN4 are "
                    "computed as a simple average of individual country financial stress "
                    "indices. Individual country financial stress indices are calculated following the "
                    "methodology of Park and Mercado (2014) but use financial sector beta instead of "
                    "banking sector beta. Financial market stress indices are computed starting in "
                    "2010."
                ),
                ui.div(
                    "Source: SEACEN staff calculations using CEIC Database and Haver Analytics data.",
                    class_="mt-1",
                ),
                class_="text-muted small",
            ),
            full_screen=True,
        ),
        ui.card(
            card_header("droplet", "Liquidity"),
            ui.layout_columns(
                output_widget("hist4"),
                col_widths={"sm": (12)},
                row_heights=["auto", 1],
            ),
            full_screen=True,
        ),
        fill=False,
        col_widths={"sm": (6, 6, 6, 6, 6, 6, 6, 6, 6, 12, 12, 12, 12)},
    ),
)

app_ui = ui.page_navbar(
    ui.nav_panel("Dashboard", dashboard_page, value="dashboard", icon=icon_svg("chart-line")),
    *[ui.nav_panel(name, doc_page(path), icon=icon_svg(icon)) for name, icon, path in DOC_TABS],
    ui.nav_spacer(),
    *(
        [
            ui.nav_control(
                ui.tags.a(
                    "Developed by Ahmad Aizudeen of The SEACEN Centre",
                    href="https://www.linkedin.com/in/ahmad-aizudeen-japri-5461ba1a9/",
                    target="_blank",
                    rel="noopener noreferrer",
                    class_="nav-link small text-decoration-underline",
                )
            )
        ]
        if SHOW_CREDIT
        else []
    ),
    title=ui.tags.span(icon_svg("chart-line"), " SEACEN Capital Flow Monitor Dashboard"),
    header=ui.tags.style(DOC_CSS),
    theme=shinyswatch.theme.flatly(),
    # Keep the tab bar visible while scrolling down the dashboard.
    navbar_options=ui.navbar_options(position="sticky-top"),
    fillable=["dashboard"],
)


def server(input, output, session):
    @output
    @render.ui
    def styled_table():
        return diff_table_html(ex_rate_diff)

    @output
    @render.ui
    def styled_table_stock():
        return diff_table_html(stock_price_diff)

    @output
    @render.ui
    def styled_table_bond():
        return diff_table_html(bond_yield_diff)

    def parent_filtered_df(df, input_date_range):
        start_date, end_date = input_date_range
        start_date = pd.Timestamp(start_date)
        end_date = pd.Timestamp(end_date)

        filt_df = df[(df["Date"] >= start_date) & (df["Date"] <= end_date)]

        return filt_df

    filltered_vix = reactive.Calc(
        lambda: parent_filtered_df(all_df[0], input.date_range())
    )
    filltered_tpu = reactive.Calc(
        lambda: parent_filtered_df(all_df[12], input.date_range_pu())
    )
    filltered_gepu = reactive.Calc(
        lambda: parent_filtered_df(all_df[14], input.date_range_pu())
    )

    custom_events = reactive.Value([])

    @reactive.effect
    @reactive.event(input.add_event_btn)
    def _add_custom_event():
        event_date = input.custom_event_date()
        if not event_date:
            return
        label = input.custom_event_label().strip() or "Custom event"
        custom_events.set(custom_events() + [{"date": event_date, "label": label}])
        ui.update_text("custom_event_label", value="")

    @reactive.effect
    @reactive.event(input.clear_events_btn)
    def _clear_custom_events():
        custom_events.set([])

    @render.ui
    def event_details():
        def fmt(d):
            return pd.Timestamp(d).strftime("%d %b %Y")

        rows = []
        for ev in DEFAULT_EVENTS:
            start = pd.Timestamp(ev["start"])
            when = f"{fmt(start)} - {fmt(ev['end'])}" if ev["type"] == "range" else fmt(start)
            rows.append((start, when, ev["label"], False))
        for ev in custom_events():
            rows.append((pd.Timestamp(ev["date"]), fmt(ev["date"]), ev["label"], True))
        rows.sort(key=lambda r: r[0], reverse=True)

        items = [
            ui.tags.li(
                ui.tags.div(when, class_="fw-semibold"),
                ui.tags.div(label + (" (custom)" if is_custom else "")),
                class_="mb-2",
            )
            for _, when, label, is_custom in rows
        ]
        return ui.tags.ul(*items, class_="list-unstyled small mb-0")

    # Each chart has its own "Show event markers" button, toggled independently
    # so clicking one only refreshes that chart, not every chart at once.
    def _marker_toggle(button_id, default=False):
        state = reactive.Value(default)

        @reactive.effect
        @reactive.event(getattr(input, button_id))
        def _toggle():
            state.set(not state())
            new_label = ("Hide" if state() else "Show") + " event markers"
            ui.update_action_button(button_id, label=ui.span(icon_svg("flag"), new_label))

        return state

    show_markers_vix = _marker_toggle("toggle_markers_vix")
    show_markers_pu = _marker_toggle("toggle_markers_pu", default=True)
    show_markers_policy = _marker_toggle("toggle_markers_policy")
    show_markers_portfolio = _marker_toggle("toggle_markers_portfolio")
    show_markers_cds = _marker_toggle("toggle_markers_cds")
    show_markers_fsi = _marker_toggle("toggle_markers_fsi")
    filltered_pr1 = reactive.Calc(
        lambda: parent_filtered_df(all_df[1], input.date_range2())
    )
    filltered_pr2 = reactive.Calc(
        lambda: parent_filtered_df(all_df[2], input.date_range2())
    )
    filltered_pr3 = reactive.Calc(
        lambda: parent_filtered_df(all_df[3], input.date_range2())
    )
    filltered_cds = reactive.Calc(
        lambda: parent_filtered_df(all_df[5], input.date_range3())
    )
    filltered_cf = reactive.Calc(
        lambda: parent_filtered_df(all_df[11], input.date_range4())
    )

    def changes_bar_chart(df, selected_years, sort_year, title, value_label):
        df = df.copy()
        region_sort = list(
            df[df["Year"] == sort_year].sort_values(by="Value", ascending=False)["Region"]
        )
        df["Region"] = pd.Categorical(df["Region"], categories=region_sort, ordered=True)
        df = df.sort_values(by=["Region", "Year"], ascending=False)
        df = df[df["Year"].isin(selected_years)]
        fig = px.histogram(
            df,
            x="Value",
            y="Region",
            color="Year",
            color_discrete_map=YEAR_COLORS,
            barmode="group",
            title=title,
        )
        fig.update_traces(hovertemplate="%{y}: %{x:.2f}")
        fig.update_xaxes(title_text=value_label)
        fig.update_yaxes(title_text=None)
        return style_fig(fig, legend="bottom", bottom_space=100)

    def _add_event_annotations(fig, x_min, x_max, show):
        if not show:
            return
        # Stagger label heights (in px, via yshift) so nearby markers — e.g. the
        # 2022/2023/2025 cluster — don't print on top of one another.
        line_slot = 0
        for ev in DEFAULT_EVENTS:
            if ev["type"] == "range":
                start, end = pd.Timestamp(ev["start"]), pd.Timestamp(ev["end"])
                if end < x_min or start > x_max:
                    continue
                fig.add_vrect(
                    x0=max(start, x_min),
                    x1=min(end, x_max),
                    fillcolor=MUTED_TEXT,
                    opacity=0.15,
                    line_width=0,
                    annotation_text=ev["label"],
                    annotation_position="top left",
                    annotation_font_size=10,
                    annotation_font_color=MUTED_TEXT,
                )
            else:
                d = pd.Timestamp(ev["start"])
                if d < x_min or d > x_max:
                    continue
                fig.add_vline(
                    x=d,
                    line_dash="dash",
                    line_color=CRITICAL_COLOR,
                    line_width=1,
                    annotation_text=ev["label"],
                    # "top left" anchors the label so it extends left of the
                    # marker rather than being centered on it — otherwise a
                    # marker near the right edge spills its label into a
                    # right-side legend (Policy Rate, CDS, etc. use legend="right").
                    annotation_position="top left",
                    annotation_font_size=10,
                    annotation_font_color=CRITICAL_COLOR,
                    annotation_yshift=-(line_slot % 3) * 14,
                )
                line_slot += 1
        for i, ev in enumerate(custom_events()):
            d = pd.Timestamp(ev["date"])
            if d < x_min or d > x_max:
                continue
            fig.add_vline(
                x=d,
                line_dash="dot",
                line_color=PRIMARY_INK,
                line_width=1,
                annotation_text=ev["label"],
                annotation_position="bottom",
                annotation_font_size=10,
                annotation_font_color=PRIMARY_INK,
                annotation_yshift=(i % 3) * 14,
            )

    @render_plotly
    def hist():
        vix = filltered_vix()
        start_date, end_date = input.date_range()
        x_min, x_max = pd.Timestamp(start_date), pd.Timestamp(end_date)

        fig = px.line(vix, x="Date", y="CLOSE")
        fig.update_traces(line_color=CATEGORICAL_PALETTE[0])
        fig.update_yaxes(title_text="VIX Close")

        _add_event_annotations(fig, x_min, x_max, show_markers_vix())
        fig.update_xaxes(range=[x_min, x_max])
        return style_fig(fig, hovermode="x unified", legend="none")

    @render_plotly
    def hist10():
        tpu = filltered_tpu()
        gepu = filltered_gepu()
        mode = input.pu_mode()
        start_date, end_date = input.date_range_pu()
        x_min, x_max = pd.Timestamp(start_date), pd.Timestamp(end_date)

        if mode == "GEPU only":
            fig = px.line(gepu, x="Date", y="Value", title="Uncertainty Indices")
            fig.update_traces(line_color=CATEGORICAL_PALETTE[2])
            fig.update_yaxes(title_text="GEPU Index")
            legend_mode = "none"
        elif mode == "Both":
            tpu_s = tpu.assign(Series="TPU")[["Date", "Series", "Value"]]
            gepu_s = gepu.assign(Series="GEPU")[["Date", "Series", "Value"]]
            combo = pd.concat([tpu_s, gepu_s], ignore_index=True)
            fig = px.line(
                combo,
                x="Date",
                y="Value",
                color="Series",
                color_discrete_map={"TPU": CATEGORICAL_PALETTE[1], "GEPU": CATEGORICAL_PALETTE[2]},
                title="Uncertainty Indices",
            )
            fig.update_yaxes(title_text="Index value")
            legend_mode = "top"
        else:
            fig = px.line(tpu, x="Date", y="Value", title="Uncertainty Indices")
            fig.update_traces(line_color=CATEGORICAL_PALETTE[1])
            fig.update_yaxes(title_text="TPU Index")
            legend_mode = "none"

        _add_event_annotations(fig, x_min, x_max, show_markers_pu())
        fig.update_xaxes(range=[x_min, x_max])
        return style_fig(fig, hovermode="x unified", legend=legend_mode)

    @render_plotly
    def hist1():
        policy_dict = {
            "All region": filltered_pr1(),
            "Selected region 1": filltered_pr2(),
            "Selected region 2": filltered_pr3(),
        }
        policy_rate = policy_dict.get(input.var(), filltered_pr1())
        fig = px.line(
            policy_rate,
            x="Date",
            y="Value",
            color="Region",
            color_discrete_map=REGION_COLORS,
            title="Selected Economies (% per annum)",
        )
        fig.update_yaxes(title_text="Policy rate (% p.a.)")
        start_date, end_date = input.date_range2()
        x_min, x_max = pd.Timestamp(start_date), pd.Timestamp(end_date)
        _add_event_annotations(fig, x_min, x_max, show_markers_policy())
        fig.update_xaxes(range=[x_min, x_max])
        return style_fig(fig, hovermode="x unified", legend="right")

    @render_plotly
    def hist2():
        return changes_bar_chart(
            all_df[4], input.ex_rate(), input.ex_rate2(), "Selected Asian Economies (% change, year-on-year)",
            "% change against the US dollar (+ = appreciation)",
        )

    @render_plotly
    def hist7():
        return changes_bar_chart(
            all_df[9], input.stock_year(), input.stock_year2(), "Selected Asian Economies (% change, year-on-year)",
            "% change in the index",
        )

    @render_plotly
    def hist8():
        return changes_bar_chart(
            all_df[10], input.bond_year(), input.bond_year2(), "Selected Asian Economies (% change, year-on-year)",
            "Change in the 5-year yield (percentage points)",
        )

    @render_plotly
    def hist9():
        cf_wide = filltered_cf()
        cf = cf_wide.melt(id_vars=["Date"], var_name="Flow", value_name="Value")
        fig = px.bar(
            cf,
            x="Date",
            y="Value",
            color="Flow",
            color_discrete_map=FLOW_COLORS,
            barmode="relative",
            title="Selected Asian Economies (US$ billion)",
        )
        total = cf_wide["Portfolio Debt"] + cf_wide["Portfolio Equity"]
        fig.add_scatter(
            x=cf_wide["Date"],
            y=total,
            mode="lines",
            name="Net Portfolio Flows",
            line=dict(color=PRIMARY_INK, width=2),
        )
        fig.update_yaxes(title_text="USD billion")
        start_date, end_date = input.date_range4()
        x_min, x_max = pd.Timestamp(start_date), pd.Timestamp(end_date)
        _add_event_annotations(fig, x_min, x_max, show_markers_portfolio())
        fig.update_xaxes(range=[x_min, x_max], title_text=None)
        return style_fig(fig, legend="bottom")

    @render_plotly
    def hist3():
        cds = filltered_cds()
        fig = px.line(
            cds,
            x="Date",
            y="Value",
            color="Region",
            color_discrete_map=REGION_COLORS,
            title="Selected Asian Economies (basis points)",
        )
        fig.update_yaxes(title_text="Basis points")
        start_date, end_date = input.date_range3()
        x_min, x_max = pd.Timestamp(start_date), pd.Timestamp(end_date)
        _add_event_annotations(fig, x_min, x_max, show_markers_cds())
        fig.update_xaxes(range=[x_min, x_max])
        return style_fig(fig, hovermode="x unified", legend="right")

    @render_plotly
    def hist4():
        liquidity = all_df[6]
        fig = px.histogram(
            liquidity,
            x="TIME_PERIOD",
            y="OBS_VALUE",
            title="Liquidity",
        )
        fig.update_traces(hovertemplate="%{x}: %{y:.1f}%")
        fig.update_xaxes(title_text=None)
        fig.update_yaxes(title_text="% change, year-on-year")
        return style_fig(fig, legend="none")

    @render_plotly
    def hist5():
        gdp_growth = GDP_DATA[GDP_DATA["Quarter"].isin(input.gdp_quarterly())]
        fig = px.histogram(
            gdp_growth,
            x="Quarter",
            y="Value",
            color="Region",
            color_discrete_map=REGION_COLORS,
            category_orders={"Quarter": GDP_QUARTERS},
            barmode="group",
            title="Selected Asian Economies (% change, year-on-year)",
        )
        fig.update_traces(hovertemplate="%{x}: %{y:.1f}%")
        fig.update_xaxes(title_text=None)
        fig.update_yaxes(title_text="% change, year-on-year")
        return style_fig(fig, legend="bottom")

    def forecast_bar_chart(df, title):
        """Grouped bars: regions on the x-axis in the sheet's order, one bar
        per forecast year."""
        fig = px.bar(
            df,
            x="Region",
            y="Value",
            color="Year",
            color_discrete_map=FORECAST_YEAR_COLORS,
            barmode="group",
            title=title,
        )
        fig.update_xaxes(title_text=None)
        fig.update_yaxes(title_text="% change, year-on-year")
        return style_fig(fig, legend="bottom")

    @render_plotly
    def forecast_gdp():
        return forecast_bar_chart(all_df[15], "GDP Growth (% change, year-on-year)")

    @render_plotly
    def forecast_inflation():
        return forecast_bar_chart(all_df[16], "Inflation (% change, year-on-year)")

    filltered_cpi = reactive.Calc(
        lambda: parent_filtered_df(CPI_DATA, input.date_range_cpi())
    )

    @render_plotly
    def cpi_chart():
        cpi = filltered_cpi()
        cpi = cpi[cpi["Region"].isin(input.cpi_regions())]
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        for region, series in cpi.groupby("Region", sort=False):
            rhs = region in CPI_RHS_REGIONS
            fig.add_scatter(
                x=series["Date"],
                y=series["Value"],
                mode="lines",
                name=f"{region} (RHS)" if rhs else region,
                line=dict(color=CPI_REGION_COLORS[region], width=2, dash="dash" if rhs else "solid"),
                hovertemplate="%{y:.1f}%",
                secondary_y=rhs,
            )
        fig.update_layout(title="Selected Asian Economies (% change, year-on-year)")
        fig.update_yaxes(title_text="% YoY (LHS)", secondary_y=False)
        fig.update_yaxes(title_text=f"% YoY, {' / '.join(CPI_RHS_REGIONS)} (RHS)", secondary_y=True)
        start_date, end_date = input.date_range_cpi()
        fig.update_xaxes(range=[pd.Timestamp(start_date), pd.Timestamp(end_date)])
        fig = style_fig(fig, hovermode="x unified", legend="right")
        # One set of gridlines (from the left axis) keeps the chart readable.
        fig.update_yaxes(showgrid=False, secondary_y=True)
        return fig

    @render_plotly
    def hist6():
        fsi = all_df[8][['Date', 'China', 'India', 'ASEAN-4', 'Asian Advanced Economies', 'Asia']]
        fsi = fsi.melt(id_vars=["Date"], var_name="Region", value_name="Value")
        fig = px.line(
            fsi,
            x="Date",
            y="Value",
            color="Region",
            color_discrete_map=FSI_REGION_COLORS,
            title="Selected Asian Economies",
        )
        fig.update_yaxes(title_text="Index (sum of standardised components)")
        x_min, x_max = fsi["Date"].min(), fsi["Date"].max()
        _add_event_annotations(fig, x_min, x_max, show_markers_fsi())
        fig.update_xaxes(range=[x_min, x_max])
        return style_fig(fig, hovermode="x unified", legend="right")

    POLICY_LABELS = {
        "All region": "All Region",
        "Selected region 1": "Selected Region 1",
        "Selected region 2": "Selected Region 2",
    }

    def _active_events(x_min, x_max):
        """Mirrors `_add_event_annotations`'s filtering, but collects the
        markers instead of drawing them, for the Excel export."""
        events = []
        for ev in DEFAULT_EVENTS:
            if ev["type"] == "range":
                start, end = pd.Timestamp(ev["start"]), pd.Timestamp(ev["end"])
                if end < x_min or start > x_max:
                    continue
                events.append({"type": "range", "label": ev["label"], "start": start, "end": end})
            else:
                d = pd.Timestamp(ev["start"])
                if d < x_min or d > x_max:
                    continue
                events.append({"type": "line", "label": ev["label"], "start": d})
        for ev in custom_events():
            d = pd.Timestamp(ev["date"])
            if d < x_min or d > x_max:
                continue
            events.append({"type": "custom", "label": ev["label"], "start": d})
        return events

    @render.download_button(
        filename=lambda: f"capital_flow_monitor_{pd.Timestamp.today().date()}.xlsx"
    )
    def download_excel():
        policy_dict = {
            "All region": filltered_pr1(),
            "Selected region 1": filltered_pr2(),
            "Selected region 2": filltered_pr3(),
        }

        vix_start, vix_end = input.date_range()
        pu_start, pu_end = input.date_range_pu()
        policy_start, policy_end = input.date_range2()
        portfolio_start, portfolio_end = input.date_range4()
        cds_start, cds_end = input.date_range3()
        fsi_df = all_df[8][["Date", "China", "India", "ASEAN-4", "Asian Advanced Economies", "Asia"]]

        buf = build_capital_flow_workbook(
            vix_df=filltered_vix(),
            tpu_df=filltered_tpu(),
            gepu_df=filltered_gepu(),
            pu_mode=input.pu_mode(),
            # CEIC/Haver-based sheets (EXCEL_RESTRICTED) only when allowed.
            policy_df=policy_dict.get(input.var(), filltered_pr1()) if EXCEL_FULL else None,
            policy_label=POLICY_LABELS.get(input.var(), "All Region"),
            fx_df=all_df[4],
            fx_years=input.ex_rate(),
            fx_sort_year=input.ex_rate2(),
            stock_df=all_df[9],
            stock_years=input.stock_year(),
            stock_sort_year=input.stock_year2(),
            bond_df=all_df[10],
            bond_years=input.bond_year(),
            bond_sort_year=input.bond_year2(),
            portfolio_df=filltered_cf(),
            cds_df=filltered_cds() if EXCEL_FULL else None,
            liquidity_df=all_df[6],
            gdp_df=GDP_DATA,
            gdp_quarters=[q for q in GDP_QUARTERS if q in input.gdp_quarterly()],
            fsi_df=fsi_df,
            gdp_forecast_df=all_df[15] if EXCEL_FULL else None,
            inflation_forecast_df=all_df[16] if EXCEL_FULL else None,
            cpi_df=filltered_cpi()[lambda d: d["Region"].isin(input.cpi_regions())] if EXCEL_FULL else None,
            cpi_rhs_regions=CPI_RHS_REGIONS,
            vix_events=_active_events(pd.Timestamp(vix_start), pd.Timestamp(vix_end))
            if show_markers_vix() else None,
            pu_events=_active_events(pd.Timestamp(pu_start), pd.Timestamp(pu_end))
            if show_markers_pu() else None,
            policy_events=_active_events(pd.Timestamp(policy_start), pd.Timestamp(policy_end))
            if show_markers_policy() else None,
            portfolio_events=_active_events(pd.Timestamp(portfolio_start), pd.Timestamp(portfolio_end))
            if show_markers_portfolio() else None,
            cds_events=_active_events(pd.Timestamp(cds_start), pd.Timestamp(cds_end))
            if show_markers_cds() else None,
            fsi_events=_active_events(fsi_df["Date"].min(), fsi_df["Date"].max())
            if show_markers_fsi() else None,
        )
        yield buf.getvalue()


app = App(app_ui, server)
