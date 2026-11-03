"""Shared look of the dashboard: colors, chart styling and card parts, used by
the Macro-Financial tab (main.py) and the Capital Flows tab (section2.py)."""
import itertools

from shiny import ui
from faicons import icon_svg
import plotly.express as px

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
