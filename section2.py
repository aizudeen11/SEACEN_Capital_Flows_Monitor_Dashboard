"""The Capital Flows tab: balance of payments (BoP) and international
investment position (IIP) of the SEACEN economies, from the Section 2 sheet
(processed_data.dfs2). A Shiny module, added to the dashboard in main.py:
capital_flows_ui("cf", header, footer) in the page and
capital_flows_server("cf") in the server."""
from shiny import module, ui, render, reactive
import pandas as pd
from faicons import icon_svg
from shinywidgets import render_plotly, output_widget
import plotly.express as px

from processed_data import dfs2, SECTION2_LABELS
from design import CATEGORICAL_PALETTE, MUTED_TEXT, PRIMARY_INK, _entity_color_map, style_fig, card_header, card_notes

SECTION2 = dfs2()

DATASETS = ["Balance of Payments", "International Investment Position"]
FREQUENCIES = ["Annual", "Half-yearly", "Quarterly"]
SEGMENTS = ["By region", "By category", "By economy"]
# The Group2 value of each dataset and frequency in the Section 2 tables
# (the half-yearly tables are built from the quarterly data).
GROUP2 = {
    ("Balance of Payments", "Annual"): "BoP Annual",
    ("Balance of Payments", "Half-yearly"): "BoP Quarterly",
    ("Balance of Payments", "Quarterly"): "BoP Quarterly",
    ("International Investment Position", "Annual"): "IIP Annual",
    ("International Investment Position", "Half-yearly"): "IIP Quarterly",
    ("International Investment Position", "Quarterly"): "IIP Quarterly",
}
# The aggregate of all SEACEN economies: the sum of the other regional
# groups, so the "By region" chart stacks the groups and shows it as the
# Total line instead (except for the current account, only kept for it).
ALL_ECONOMIES = "Asia-Pacific"
CA_BALANCE = "Current Account Balance"


def _group_label(group, dataset):
    """BoP assets and liabilities are resident and non-resident capital
    flows; IIP ones are positions."""
    if dataset == "Balance of Payments":
        return {
            "Assets": "Assets (resident capital flows)",
            "Liabilities": "Liabilities (non-resident capital flows)",
            "Current Account": "Current account",
        }.get(group, group)
    return group


def _periods(df):
    return [c for c in df.columns if c not in SECTION2_LABELS]


def _item_group(item):
    if item.endswith("Assets"):
        return "Assets"
    if item.endswith("Liabilities"):
        return "Liabilities"
    return "Current account"


if SECTION2 is not None:
    HALF_TYPE, QUARTER_TYPE, HALF_REGION, ANNUAL_TYPE, ANNUAL_REGION, _, SECTION2_UPDATED = SECTION2
    BY_ECONOMY = {"Annual": ANNUAL_TYPE, "Half-yearly": HALF_TYPE, "Quarterly": QUARTER_TYPE}
    BY_REGION = {"Annual": ANNUAL_REGION, "Half-yearly": HALF_REGION}
    YEARS = sorted(
        {p[-4:] for df in (*BY_ECONOMY.values(), *BY_REGION.values()) for p in _periods(df)}, reverse=True
    )
    REGION_GROUPS = list(dict.fromkeys(HALF_REGION["Region"]))
    ECONOMIES = sorted(set(HALF_TYPE["Region"]) | set(ANNUAL_TYPE["Region"]))
    # Fixed colors, so a regional group, category or economy looks the same in
    # every view. Up to 9 categories and 19 economies at once - more than the
    # 8-hue palette - so those two draw from Plotly's larger set (as main.py's
    # REGION_COLORS does).
    REGION_GROUP_COLORS = dict(zip(REGION_GROUPS, CATEGORICAL_PALETTE))
    _categories = pd.concat([HALF_REGION, ANNUAL_REGION])
    CATEGORY_COLORS = _entity_color_map(
        _categories[~_categories["Type"].str.endswith(" Total")], column="Type", palette=px.colors.qualitative.Dark24
    )
    ECONOMY_COLORS = _entity_color_map(HALF_TYPE, ANNUAL_TYPE, palette=px.colors.qualitative.Dark24)
else:
    YEARS, REGION_GROUPS, ECONOMIES, SECTION2_UPDATED = [], [], [], None


def _empty_fig(message):
    fig = px.bar()
    fig.add_annotation(text=message, showarrow=False, font=dict(color=MUTED_TEXT, size=14))
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    return style_fig(fig, legend="none")


NO_DATA = "No data for the selected options."


def view(dataset, frequency, segment, years, group=None, regions=(), region=None, item=None, economies=()):
    """The rows behind the chart and table for the given filters, with their
    label columns and the selected years' periods, and the chart title.
    Rows are None if there is nothing to show."""
    group2 = GROUP2[(dataset, frequency)]
    if segment == SEGMENTS[2]:
        df = BY_ECONOMY[frequency]
        df = df[(df["Group2"] == group2) & (df["Type"] == item) & df["Region"].isin(economies)]
        labels = [c for c in ("Region", "Type", "Last Update Time") if c in df.columns]
        title = f"{item} by economy (US$ billion)"
    else:
        if frequency not in BY_REGION or not group:
            return None, None
        df = BY_REGION[frequency]
        df = df[(df["Group2"] == group2) & (df["Group"] == group)]
        label = _group_label(group, dataset)
        if segment == SEGMENTS[0]:
            df = df[df["Region"].isin(regions)]
            title = f"{label} by regional group (US$ billion)"
        else:
            df = df[df["Region"] == region]
            title = f"{label}, {region} (US$ billion)"
        labels = ["Region", "Type"]
    # Periods without any data (e.g. a half-year not yet reported) are left
    # out. Economies without data stay, as blank rows in the table; the chart
    # leaves them out (see figure).
    periods = [p for p in _periods(df) if p[-4:] in years and df[p].notna().any()]
    if df.empty or not periods:
        return None, title
    return df[labels + periods], title


def figure(df, segment, title):
    """Stacked bars (one series per regional group, category or economy) with
    their total as a line."""
    if df is None:
        return _empty_fig(NO_DATA)
    periods = _periods(df)
    # Totals (the "... Total" rows and the current account balance) are the
    # line, not bars.
    parts = df[~df["Type"].str.endswith(" Total") & (df["Type"] != CA_BALANCE)]
    if segment == SEGMENTS[0]:
        parts = parts.groupby("Region", sort=False)[periods].sum(min_count=1).reset_index()
    # Series without any data get no bars, and stay out of the legend.
    parts = parts[parts[periods].notna().any(axis=1)]
    if parts.empty:
        return _empty_fig(NO_DATA)
    series, colors = {
        SEGMENTS[0]: ("Region", REGION_GROUP_COLORS),
        SEGMENTS[1]: ("Type", CATEGORY_COLORS),
        SEGMENTS[2]: ("Region", ECONOMY_COLORS),
    }[segment]
    # Blank values are dropped: Plotly restarts a stack from zero at a blank,
    # so the series after it would be drawn over the ones before it.
    long = parts.melt(id_vars=[series], value_vars=periods, var_name="Period", value_name="Value").dropna(
        subset=["Value"]
    )
    fig = px.bar(
        long,
        x="Period",
        y="Value",
        color=series,
        color_discrete_map=colors,
        category_orders={"Period": periods, series: list(parts[series])},
        barmode="relative",
        title=title,
    )
    fig.update_traces(hovertemplate="%{y:,.1f}")
    fig.add_scatter(
        x=periods,
        y=parts[periods].sum(min_count=1).values,
        mode="lines+markers",
        name="Total",
        line=dict(color=PRIMARY_INK, width=2),
        hovertemplate="%{y:,.1f}",
    )
    fig.update_xaxes(title_text=None, type="category")
    fig.update_yaxes(title_text="US$ billion")
    many = long[series].nunique() > 6
    return style_fig(fig, hovermode="x unified", legend="right" if many else "bottom")


def table_data(df):
    """The chart's rows for the table and the CSV download."""
    if df is None:
        return pd.DataFrame({"Message": [NO_DATA]})
    df = df.copy()
    if "Last Update Time" in df.columns:
        df["Last Update Time"] = pd.to_datetime(df["Last Update Time"], errors="coerce").dt.strftime("%d %b %Y")
    return df


@module.ui
def capital_flows_ui(header, footer):
    """The tab's page. `header` and `footer` go at the top and bottom of its
    sidebar (main.py passes the same ones as the Macro-Financial tab's)."""
    if SECTION2 is None:
        return ui.div(
            ui.p(
                icon_svg("circle-exclamation"),
                " The Capital Flows data (Section 2 Google Sheet) could not be loaded. Please try again later.",
                class_="text-muted",
            ),
            class_="p-4",
        )
    sidebar = ui.sidebar(
        header,
        ui.hr(),
        ui.h6("Filters", class_="fw-bold mb-2"),
        ui.input_radio_buttons("dataset", "Dataset", choices=DATASETS, selected=DATASETS[0]),
        ui.input_radio_buttons("frequency", "Frequency", choices=FREQUENCIES, selected="Half-yearly", inline=True),
        ui.input_radio_buttons("segment", "View", choices=SEGMENTS, selected=SEGMENTS[0]),
        ui.input_checkbox_group("years", "Years", choices=YEARS, selected=YEARS, inline=True),
        ui.panel_conditional(
            f"input.segment !== '{SEGMENTS[2]}'",
            # Choices follow the dataset (see _update_groups).
            ui.input_select("group", "Account", choices=[]),
        ),
        ui.panel_conditional(
            f"input.segment === '{SEGMENTS[0]}'",
            ui.input_checkbox_group("regions", "Regional groups", choices=[]),
        ),
        ui.panel_conditional(
            f"input.segment === '{SEGMENTS[1]}'",
            ui.input_select("region", "Regional group", choices=REGION_GROUPS, selected=ALL_ECONOMIES),
        ),
        ui.panel_conditional(
            f"input.segment === '{SEGMENTS[2]}'",
            ui.input_select("item", "Item", choices=[]),
            ui.input_selectize("economies", "Economies", choices=ECONOMIES, selected=ECONOMIES, multiple=True),
        ),
        ui.p(
            icon_svg("circle-info"),
            " Quarterly data is shown by economy only.",
            class_="text-muted small mt-2",
        ),
        ui.hr(),
        footer,
        width=300,
    )
    return ui.layout_sidebar(
        sidebar,
        ui.card(
            card_header("chart-column", ui.output_text("heading", inline=True)),
            output_widget("chart"),
            card_notes(
                note=ui.TagList(
                    ui.div(
                        "Note: US$ billion. Balance of payments: assets are resident capital flows (net "
                        "acquisition of foreign financial assets) and liabilities are non-resident capital "
                        "flows (net incurrence of liabilities); half-yearly values are the sum of the two "
                        "quarters. International investment position: end-of-period positions (half-yearly: "
                        "end-June and end-December). Bars are stacked; the line is their total."
                    ),
                    ui.div(
                        "Regional groups: Asian Advanced Economies (Hong Kong, China; Korea; Singapore; and "
                        "Chinese Taipei), ASEAN5 (Indonesia, Malaysia, Philippines, Thailand, and Vietnam), "
                        "Asian EDMEs (Brunei Darussalam, Cambodia, Lao PDR, Mongolia, Myanmar, Nepal, Papua "
                        "New Guinea, and Sri Lanka), China and India. Asia-Pacific is the total of all 19 "
                        "SEACEN economies. Singapore's quarterly BoP components are estimated from its totals "
                        "using fixed shares.",
                        class_="mt-1",
                    ),
                ),
                source="Source: SEACEN staff calculations using IMF balance of payments and international "
                "investment position statistics and national sources.",
            ),
            full_screen=True,
            min_height="560px",
        ),
        ui.card(
            ui.card_header(
                ui.span(icon_svg("table"), " Data", class_="fw-semibold"),
                ui.download_link("download_table", ui.span(icon_svg("file-csv"), " Download CSV"), class_="small"),
                class_="d-flex align-items-center justify-content-between",
            ),
            ui.output_data_frame("table"),
            card_notes(note="Note: US$ billion. Regional rows include each group's total."),
            full_screen=True,
            min_height="420px",
        ),
        fillable=False,
    )


@module.server
def capital_flows_server(input, output, session):
    if SECTION2 is None:
        return

    def group2():
        return GROUP2[(input.dataset(), input.frequency())]

    @reactive.effect
    def _update_segments():
        """Quarterly data is only by economy."""
        choices = SEGMENTS[2:] if input.frequency() == "Quarterly" else SEGMENTS
        with reactive.isolate():
            current = input.segment()
        ui.update_radio_buttons("segment", choices=choices, selected=current if current in choices else choices[0])

    @reactive.calc
    def region_rows():
        """The regional table's rows for the dataset (all accounts)."""
        df = BY_REGION.get(input.frequency(), BY_REGION["Half-yearly"])
        return df[df["Group2"] == group2()]

    @reactive.calc
    def economy_rows():
        df = BY_ECONOMY[input.frequency()]
        return df[df["Group2"] == group2()]

    @reactive.effect
    def _update_groups():
        rows = region_rows()
        groups = list(dict.fromkeys(rows.loc[rows[_periods(rows)].notna().any(axis=1), "Group"]))
        with reactive.isolate():
            current = input.group()
        ui.update_select(
            "group",
            choices={g: _group_label(g, input.dataset()) for g in groups},
            selected=current if current in groups else groups[0],
        )

    @reactive.effect
    def _update_regions():
        """The groups to compare: all but the Asia-Pacific total, unless it is
        the only one (the current account)."""
        rows = region_rows()
        available = list(dict.fromkeys(rows.loc[rows["Group"] == input.group(), "Region"]))
        if available != [ALL_ECONOMIES]:
            available = [r for r in available if r != ALL_ECONOMIES]
        with reactive.isolate():
            current = [r for r in input.regions() if r in available]
        ui.update_checkbox_group("regions", choices=available, selected=current or available)

    @reactive.effect
    def _update_items():
        """Items with data for the dataset, grouped as assets, liabilities and
        current account."""
        df = economy_rows()
        has_data = df[_periods(df)].notna().any(axis=1)
        items = list(dict.fromkeys(df.loc[has_data, "Type"]))
        grouped = {
            group: {item: item for item in items if _item_group(item) == group}
            for group in ("Assets", "Liabilities", "Current account")
        }
        grouped = {group: choices for group, choices in grouped.items() if choices}
        with reactive.isolate():
            current = input.item()
        default = "FDI Assets" if "FDI Assets" in items else items[0]
        ui.update_select("item", choices=grouped, selected=current if current in items else default)

    @reactive.calc
    def current_view():
        return view(
            input.dataset(),
            input.frequency(),
            input.segment(),
            input.years(),
            group=input.group(),
            regions=input.regions(),
            region=input.region(),
            item=input.item(),
            economies=input.economies(),
        )

    @render.text
    def heading():
        return f"{input.dataset()}, {input.frequency().lower()} data"

    @render_plotly
    def chart():
        df, title = current_view()
        return figure(df, input.segment(), title)

    @render.data_frame
    def table():
        df = table_data(current_view()[0])
        shown = df.copy()
        if "Message" not in df.columns:
            for p in _periods(df):
                shown[p] = df[p].map(lambda x: "" if pd.isna(x) else f"{x:,.2f}")
        return render.DataGrid(shown, width="100%")

    @render.download_link(filename=lambda: f"capital_flows_{pd.Timestamp.today().date()}.csv")
    def download_table():
        yield table_data(current_view()[0]).to_csv(index=False)
