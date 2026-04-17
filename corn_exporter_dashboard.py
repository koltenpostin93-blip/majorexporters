"""
JSA Global Export Dashboard — John Stewart & Associates
Monthly export shipments in thousands of metric tons (TMT).

Commodities:
  🌽 Corn     — US, Brazil, Argentina, Ukraine, Total Non-US, Major Exporters
  🫘 Soybeans — US, Brazil, Argentina, Total Non-US, Major Exporters, China Imports

Marketing Year conventions:
  Oct–Sep : US, Ukraine, Total Non-US, Major Exporters, China Imports
  Mar–Feb : Argentina, Brazil

Data source: Corn Exporter Dashboard Data.xlsx
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import base64
import os
import shutil
import tempfile

# ─────────────────────────────────────────────────────────────────────────────
# PATHS
# ─────────────────────────────────────────────────────────────────────────────
_HERE = os.path.dirname(os.path.abspath(__file__))
DASHBOARD_DIR = (
    r"C:\Users\KoltenPostin\John Stewart and Associates"
    r"\JSA - Documents\Research Analyst\Dashboards"
    r"\Corn by Major Exports"
)

def _find(filename: str) -> str:
    local = os.path.join(_HERE, filename)
    if os.path.exists(local):
        return local
    return os.path.join(DASHBOARD_DIR, filename)

EXCEL_PATH      = _find("Corn Exporter Dashboard Data.xlsx")
LOGO_FULL_PATH  = _find("jsa_logo_full.png")
LOGO_WHITE_PATH = _find("jsa_logo_white.png")

# ─────────────────────────────────────────────────────────────────────────────
# BRAND COLORS
# ─────────────────────────────────────────────────────────────────────────────
JSA_GREEN = "#4a6741"
JSA_CYAN  = "#0693e3"
JSA_DARK  = "#1e2124"
JSA_CHAR  = "#32373c"
JSA_MID   = "#2a2f35"

# ─────────────────────────────────────────────────────────────────────────────
# MONTH LISTS
# ─────────────────────────────────────────────────────────────────────────────
OCT_SEP_MONTHS = ["Oct","Nov","Dec","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep"]
MAR_FEB_MONTHS = ["Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb"]
APR_MAR_MONTHS = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]
ALL_MONTHS     = set(OCT_SEP_MONTHS)

# ─────────────────────────────────────────────────────────────────────────────
# COMMODITY CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────
COMMODITY_CONFIG = {
    "corn": {
        "sheet":          "Corn",
        "emoji":          "🌽",
        "label":          "Corn",
        "col_names":      ["MarketYear","Date","Month","US","Brazil","Argentina",
                           "Ukraine","TotalNonUS","MajorExporter"],
        "numeric_cols":   ["US","Brazil","Argentina","Ukraine","TotalNonUS","MajorExporter"],
        "fields": {
            "US":            "United States",
            "Brazil":        "Brazil",
            "Argentina":     "Argentina",
            "Ukraine":       "Ukraine",
            "TotalNonUS":    "Total Non-US",
            "MajorExporter": "Major Exporters",
        },
        "mar_feb_fields":   {"Brazil","Argentina"},
        "arbr_months":      MAR_FEB_MONTHS,
        "arbr_last_month":  "Feb",
        "arbr_label":       "Mar–Feb",
        "arbr_prev_months": frozenset({"Jan","Feb"}),
        "import_fields":  set(),
        "non_us_comps":   ["Brazil","Argentina","Ukraine"],
        "major_comps":    ["US","Brazil","Argentina","Ukraine"],
        "tile_order":     ["US","Brazil","Argentina","Ukraine","TotalNonUS","MajorExporter"],
        "tile_accents": {
            "US":            "#f9a825",
            "Brazil":        "#2e7d32",
            "Argentina":     "#29b6f6",
            "Ukraine":       "#fdd835",
            "TotalNonUS":    "#7e57c2",
            "MajorExporter": "#ef5350",
        },
        "country_colors": {
            "US":            "#f9a825",
            "Brazil":        "#1b7e35",
            "Argentina":     "#4fc3f7",
            "Ukraine":       "#fdd835",
            "TotalNonUS":    "#7e57c2",
            "MajorExporter": "#ef5350",
        },
        "bu_lbs": 56.0,   # lbs per bushel of corn
    },
    "soybeans": {
        "sheet":          "Soybeans",
        "emoji":          "🫘",
        "label":          "Soybeans",
        "col_names":      ["MarketYear","Date","Month","US","Brazil","Argentina",
                           "TotalNonUS","MajorExporter","ChinaImports"],
        "numeric_cols":   ["US","Brazil","Argentina","TotalNonUS","MajorExporter","ChinaImports"],
        "fields": {
            "US":            "United States",
            "Brazil":        "Brazil",
            "Argentina":     "Argentina",
            "TotalNonUS":    "Total Non-US",
            "MajorExporter": "Major Exporters",
            "ChinaImports":  "China Imports",
        },
        "mar_feb_fields":   {"Brazil","Argentina"},
        "arbr_months":      APR_MAR_MONTHS,
        "arbr_last_month":  "Mar",
        "arbr_label":       "Apr–Mar",
        "arbr_prev_months": frozenset({"Jan","Feb","Mar"}),
        "import_fields":  {"ChinaImports"},
        "non_us_comps":   ["Brazil","Argentina"],
        "major_comps":    ["US","Brazil","Argentina"],
        "tile_order":     ["US","Brazil","Argentina","TotalNonUS","MajorExporter","ChinaImports"],
        "tile_accents": {
            "US":            "#f9a825",
            "Brazil":        "#2e7d32",
            "Argentina":     "#29b6f6",
            "TotalNonUS":    "#7e57c2",
            "MajorExporter": "#ef5350",
            "ChinaImports":  "#e53935",
        },
        "country_colors": {
            "US":            "#f9a825",
            "Brazil":        "#1b7e35",
            "Argentina":     "#4fc3f7",
            "TotalNonUS":    "#7e57c2",
            "MajorExporter": "#ef5350",
            "ChinaImports":  "#e53935",
        },
        "bu_lbs": 60.0,   # lbs per bushel of soybeans
    },
}


def _bu_conv_factor(cfg: dict) -> float:
    """TMT → Million Bushels conversion factor for this commodity."""
    return 1_000 * (2_204.62 / cfg["bu_lbs"]) / 1_000_000


def _apply_unit(pivot: dict, factor: float) -> dict:
    return {
        m: {y: (v * factor if v is not None else None) for y, v in yr.items()}
        for m, yr in pivot.items()
    }


# ─────────────────────────────────────────────────────────────────────────────
# FAVICON
# ─────────────────────────────────────────────────────────────────────────────
def _make_favicon():
    try:
        from PIL import Image
        img  = Image.open(LOGO_FULL_PATH).convert("RGBA")
        _, h = img.size
        icon = img.crop((0, 0, h, h))
        data = np.array(icon, dtype=np.uint8)
        mask = data[:, :, 3] > 10
        data[mask, 0], data[mask, 1], data[mask, 2] = 74, 103, 65
        return Image.fromarray(data, "RGBA")
    except Exception:
        return "📊"


st.set_page_config(
    page_title="JSA Export Dashboard",
    page_icon=_make_favicon(),
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────────────────────
# LOGO UTILITIES
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def _load_logo_b64(path: str) -> str | None:
    try:
        with open(path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode()
        return f"data:image/png;base64,{encoded}"
    except FileNotFoundError:
        return None


def _add_chart_watermark(fig: go.Figure, logo_b64: str | None) -> go.Figure:
    if logo_b64 is None:
        return fig
    fig.add_layout_image(dict(
        source=logo_b64, xref="paper", yref="paper",
        x=0.5, y=0.5, sizex=0.30, sizey=0.30,
        xanchor="center", yanchor="middle",
        opacity=0.10, layer="above",
    ))
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(commodity: str) -> pd.DataFrame:
    cfg = COMMODITY_CONFIG[commodity]

    # Copy to a temp file first so a locked/open Excel workbook on Windows
    # doesn't cause a PermissionError or a stale read.
    tmp_path = None
    try:
        tmp_path = tempfile.mktemp(suffix=".xlsx")
        shutil.copy2(EXCEL_PATH, tmp_path)
        read_path = tmp_path
    except Exception:
        read_path = EXCEL_PATH   # fall back to reading directly

    try:
        df = pd.read_excel(read_path, sheet_name=cfg["sheet"], header=0)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
    n   = len(cfg["col_names"])
    df  = df.iloc[:, :n].copy()
    df.columns = cfg["col_names"]
    df["MarketYear"] = df["MarketYear"].astype(str).str.strip()
    df["Month"]      = df["Month"].astype(str).str.strip()
    df["Date"]       = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Month"].isin(ALL_MONTHS)].copy()
    for col in cfg["numeric_cols"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return _enforce_aggregate_completeness(df, cfg)


def _enforce_aggregate_completeness(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """
    Recompute TotalNonUS and MajorExporter from components.
    Null out any row where a required component is missing.
    """
    df = df.copy()
    non_us_cols = cfg["non_us_comps"]
    major_cols  = cfg["major_comps"]
    all_non_us  = df[non_us_cols].notna().all(axis=1)
    all_major   = df[major_cols].notna().all(axis=1)
    df["TotalNonUS"]    = np.where(all_non_us, df[non_us_cols].sum(axis=1), np.nan)
    df["MajorExporter"] = np.where(all_major,  df[major_cols].sum(axis=1),  np.nan)
    return df


# ─────────────────────────────────────────────────────────────────────────────
# PIVOT BUILDERS
# ─────────────────────────────────────────────────────────────────────────────
def build_pivot(df: pd.DataFrame, field: str) -> tuple[dict, list[str]]:
    months = OCT_SEP_MONTHS
    pivot  = {m: {} for m in months}
    for _, row in df.iterrows():
        m, y = row["Month"], row["MarketYear"]
        if m not in pivot:
            continue
        val = row[field]
        pivot[m][y] = None if pd.isna(val) else float(val)
    all_years = sorted({y for md in pivot.values() for y in md})
    return pivot, all_years


def build_arbr_pivot(df: pd.DataFrame, field: str,
                     months_list=None,
                     prev_months=None) -> tuple[dict, list[str]]:
    """
    Build a marketing-year pivot for AR/BR fields.
    months_list : ordered list of months for this MY (e.g. MAR_FEB_MONTHS or APR_MAR_MONTHS)
    prev_months : months that belong to the *previous* marketing year label
                  (Jan+Feb for Mar-Feb; Jan+Feb+Mar for Apr-Mar)
    """
    if months_list is None:
        months_list = MAR_FEB_MONTHS
    if prev_months is None:
        prev_months = frozenset({"Jan", "Feb"})
    pivot = {m: {} for m in months_list}
    for _, row in df.iterrows():
        month = row["Month"]
        if month not in pivot:
            continue
        date = row["Date"]
        if pd.isna(date):
            continue
        val   = row[field]
        year  = date.year
        start = year - 1 if month in prev_months else year
        label = f"{start}/{str(start + 1)[-2:]}"
        if label not in pivot[month]:
            pivot[month][label] = None if pd.isna(val) else float(val)
    all_years = sorted({y for md in pivot.values() for y in md})
    return pivot, all_years


def build_cumulative_pivot(monthly_pivot: dict, years: list[str],
                            months: list[str]) -> dict:
    cum = {m: {} for m in months}
    for year in years:
        running = 0.0
        for month in months:
            val = monthly_pivot[month].get(year)
            if val is not None:
                running += val
                cum[month][year] = running
            else:
                cum[month][year] = None
    return cum


# ─────────────────────────────────────────────────────────────────────────────
# STATISTICS
# ─────────────────────────────────────────────────────────────────────────────
def get_complete_years(pivot: dict, last_month: str) -> list[str]:
    return sorted(y for y, v in pivot.get(last_month, {}).items() if v is not None)


def olympic_avg(values: list) -> float | None:
    valid = [v for v in values if v is not None and not np.isnan(v)]
    if len(valid) < 3:
        return None
    s = sorted(valid)
    return sum(s[1:-1]) / len(s[1:-1])


def compute_stats(data_pivot, all_years, complete_years, cy, ly,
                  months, is_cumulative=False) -> dict:
    oly_years  = sorted(complete_years)[-6:]
    hist_years = sorted(complete_years)
    stats: dict = {}

    for month in months:
        hist_vals  = [data_pivot[month].get(y) for y in hist_years]
        clean_hist = [v for v in hist_vals if v is not None]
        oly_vals   = [data_pivot[month].get(y) for y in oly_years]
        cy_val     = data_pivot[month].get(cy)
        ly_val     = data_pivot[month].get(ly) if ly else None
        oly = olympic_avg(oly_vals)
        stats[month] = dict(
            oly_avg    = oly,
            min        = min(clean_hist) if clean_hist else None,
            max        = max(clean_hist) if clean_hist else None,
            pct_vs_ly  = _pct(cy_val, ly_val),
            pct_vs_oly = _pct(cy_val, oly),
        )

    totals: dict[str, float | None] = {}
    for year in all_years:
        yr_vals  = [data_pivot[m].get(year) for m in months]
        non_none = [v for v in yr_vals if v is not None]
        totals[year] = (non_none[-1] if non_none else None) if is_cumulative \
                       else (sum(non_none) if non_none else None)

    hist_totals = [totals.get(y) for y in hist_years]
    clean_ht    = [v for v in hist_totals if v is not None]
    oly_t       = olympic_avg([totals.get(y) for y in oly_years])
    stats["TOTAL"] = dict(
        oly_avg    = oly_t,
        min        = min(clean_ht) if clean_ht else None,
        max        = max(clean_ht) if clean_ht else None,
        pct_vs_ly  = _pct(totals.get(cy), totals.get(ly) if ly else None),
        pct_vs_oly = _pct(totals.get(cy), oly_t),
    )
    stats["_totals"] = totals
    return stats


def _pct(a, b) -> float | None:
    if a is None or b is None or b == 0:
        return None
    return (a - b) / abs(b) * 100


# ─────────────────────────────────────────────────────────────────────────────
# FORMATTING
# ─────────────────────────────────────────────────────────────────────────────
def fmt_num(val, decimals: int = 0) -> str:
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return "—"
    return f"{val:,.{decimals}f}"


def fmt_pct(val) -> str:
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return "—"
    sign = "+" if val >= 0 else ""
    return f"{sign}{val:.1f}%"


def pct_color(val) -> str:
    if val is None:
        return "#555555"
    return "#1b7e35" if val > 0 else "#c0392b" if val < 0 else "#555555"


# ─────────────────────────────────────────────────────────────────────────────
# TABLE RENDERER
# ─────────────────────────────────────────────────────────────────────────────
_TABLE_CSS = """
<style>
.corn-wrap {
    overflow-x: auto; max-height: 580px; overflow-y: auto;
    border-radius: 6px; border: 1px solid #484f56;
    font-family: Arial, sans-serif; font-size: 12.5px;
    line-height: 1.35; margin-bottom: 8px;
}
.corn-tbl { border-collapse: collapse; width: max-content; min-width: 100%; }
.corn-tbl th {
    background: #1e2124; color: #ffffff; font-weight: 600;
    text-align: center; padding: 7px 10px; white-space: nowrap;
    position: sticky; top: 0; z-index: 3;
    border-right: 1px solid #484f56; border-bottom: 2px solid #0d0f11;
}
.corn-tbl th.stat-hdr  { background: #2a2f35; color: #ffffff; z-index: 5; }
.corn-tbl th.cy-hdr    { background: #0555a0; color: #ffffff; }
.corn-tbl td {
    padding: 5px 10px; text-align: right; white-space: nowrap;
    border-bottom: 1px solid #484f56; border-right: 1px solid #484f56;
    background: #32373c; color: #ffffff;
}
.corn-tbl tr:hover td  { filter: brightness(1.2); }
.corn-tbl td.m-cell    { text-align: left; font-weight: 700;
                          background: #1e2124 !important; color: #ffffff !important; }
.corn-tbl td.s-cell    { background: #f4f6f8 !important; color: #000000 !important; }
.corn-tbl td.p-cell    { background: #f4f6f8 !important; font-weight: 700; }
.corn-tbl td.left-divider { border-left: 2px solid #0693e3 !important; }
.corn-tbl th.left-divider { border-left: 2px solid #0693e3 !important; }
.corn-tbl td.cy-cell   { background: #0693e3 !important; font-weight: 600; color: #ffffff !important; }
.corn-tbl tr.total-row td {
    background: #232729 !important; font-weight: 700;
    border-top: 2px solid #0693e3; color: #ffffff;
}
.corn-tbl tr.total-row td.cy-cell { background: #0555a0 !important; color: #ffffff !important; }
.corn-tbl tr.total-row td.s-cell  { background: #2a2f35 !important; color: #ffffff !important; }
.corn-tbl tr.total-row td.p-cell  { background: #2a2f35 !important; }
</style>
"""

def render_table_html(data_pivot, stats, all_years, cy, ly, months,
                      decimals: int = 0) -> str:
    fn = lambda v: fmt_num(v, decimals)
    W  = dict(month=65, stat=94, pct=112, year=90)

    def sticky_left(w):
        return f"position:sticky;left:0;min-width:{w}px;z-index:2;"

    R = [0]
    R.append(R[-1] + W["pct"])
    R.append(R[-1] + W["pct"])
    R.append(R[-1] + W["stat"])
    R.append(R[-1] + W["stat"])
    R.append(R[-1] + W["stat"])

    def sticky_right(r_idx, w):
        return f"position:sticky;right:{R[r_idx]}px;min-width:{w}px;z-index:5;"

    hdr = (f'<th class="stat-hdr" '
           f'style="{sticky_left(W["month"])};text-align:left;z-index:10;">Month</th>')
    for year in all_years:
        if year == cy:
            continue
        hdr += f'<th style="min-width:{W["year"]}px;">{year}</th>'

    for r_idx, w, cls, lbl in [
        (5, W["year"], "cy-hdr left-divider", cy),
        (4, W["stat"], "stat-hdr",            "6-Yr<br>Oly Avg"),
        (3, W["stat"], "stat-hdr",            "Min"),
        (2, W["stat"], "stat-hdr",            "Max"),
        (1, W["pct"],  "stat-hdr",            "% Chg<br>CY vs LY"),
        (0, W["pct"],  "stat-hdr",            "% Chg CY<br>vs Oly Avg"),
    ]:
        hdr += f'<th class="{cls}" style="{sticky_right(r_idx, w)}">{lbl}</th>'

    def build_row(label, s, year_data, is_total=False):
        valid = [(y, year_data[y]) for y in all_years
                 if y != cy and year_data.get(y) is not None]
        srt  = sorted(valid, key=lambda x: x[1])
        n    = len(srt)
        bot2 = {y for y, _ in srt[:2]}  if n >= 2 else set()
        top2 = {y for y, _ in srt[-2:]} if n >= 2 else set()
        bot2 -= top2

        row = f'<td class="m-cell" style="{sticky_left(W["month"])}">{label}</td>'
        for year in all_years:
            if year == cy:
                continue
            val = year_data.get(year)
            if year in top2:
                row += f'<td style="background:#2e7d32;color:#fff;font-weight:600;">{fn(val)}</td>'
            elif year in bot2:
                row += f'<td style="background:#c62828;color:#fff;font-weight:600;">{fn(val)}</td>'
            else:
                row += f'<td>{fn(val)}</td>'

        pc_ly  = s["pct_vs_ly"]
        pc_oly = s["pct_vs_oly"]
        cy_val = year_data.get(cy)
        for r_idx, w, cls, val_str, xtra in [
            (5, W["year"], "cy-cell left-divider", fn(cy_val),       ""),
            (4, W["stat"], "s-cell",               fn(s["oly_avg"]), ""),
            (3, W["stat"], "s-cell",               fn(s["min"]),     ""),
            (2, W["stat"], "s-cell",               fn(s["max"]),     ""),
            (1, W["pct"],  "p-cell", fmt_pct(pc_ly),  f"color:{pct_color(pc_ly)};"),
            (0, W["pct"],  "p-cell", fmt_pct(pc_oly), f"color:{pct_color(pc_oly)};"),
        ]:
            row += (f'<td class="{cls}" style="{sticky_right(r_idx, w)}{xtra}">'
                    f'{val_str}</td>')

        tag = 'class="total-row"' if is_total else ""
        return f"<tr {tag}>{row}</tr>"

    rows_html = ""
    for month in months:
        rows_html += build_row(month, stats[month],
                               {y: data_pivot[month].get(y) for y in all_years})
    rows_html += build_row("TOTAL", stats["TOTAL"], stats["_totals"], is_total=True)

    return (_TABLE_CSS
            + '<div class="corn-wrap"><table class="corn-tbl">'
            + f'<thead><tr>{hdr}</tr></thead>'
            + f'<tbody>{rows_html}</tbody>'
            + '</table></div>')


# ─────────────────────────────────────────────────────────────────────────────
# CHARTS
# ─────────────────────────────────────────────────────────────────────────────
def _base_layout(title_text, x_title, y_title) -> dict:
    axis = dict(
        gridcolor="#484f56", linecolor="#484f56",
        tickcolor="#aaaaaa", tickfont=dict(color="#cccccc"),
        title_font=dict(color="#ffffff"), zerolinecolor="#484f56",
    )
    return dict(
        title=dict(text=title_text, font=dict(color="#ffffff", size=15, family="Arial")),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#32373c",
        font=dict(family="Arial", color="#ffffff", size=12),
        legend=dict(bgcolor="rgba(30,33,36,0.92)", bordercolor="#484f56",
                    borderwidth=1, font=dict(size=11)),
        margin=dict(l=70, r=20, t=55, b=60),
        hovermode="x unified",
        xaxis=dict(**axis, title=x_title),
        yaxis=dict(**axis, title=y_title, tickformat=","),
    )


def _year_style(year, cy, ly, all_years) -> tuple:
    if year == cy:
        return "#0693e3", 4.5, 1.0
    if year == ly:
        return "#ffffff", 3.0, 0.92
    other = [y for y in all_years if y != cy and y != ly]
    idx   = other.index(year) if year in other else 0
    n     = max(len(other) - 1, 1)
    t     = idx / n
    return (f"rgb({int(100+80*t)},{int(120+80*t)},{int(155+65*t)})",
            round(1.0 + 0.4 * t, 1), round(0.18 + 0.45 * t, 2))


def make_seasonal_chart(data_pivot, all_years, cy, complete_years,
                        field_label, is_cumulative, months,
                        logo_b64=None, unit_short="TMT") -> go.Figure:
    lbl = "Cumulative " if is_cumulative else ""
    fig = go.Figure()
    ly  = complete_years[-1] if complete_years else None

    hist_8 = sorted(complete_years)[-8:]
    if len(hist_8) >= 2:
        band_upper, band_lower = [], []
        for m in months:
            vals  = [data_pivot[m].get(y) for y in hist_8]
            clean = [v for v in vals if v is not None]
            if len(clean) >= 2:
                mu, sig = float(np.mean(clean)), float(np.std(clean, ddof=1))
                band_upper.append(mu + sig)
                band_lower.append(max(0.0, mu - sig))
            else:
                band_upper.append(None)
                band_lower.append(None)
        fig.add_trace(go.Scatter(
            x=months + months[::-1], y=band_upper + band_lower[::-1],
            fill="toself", fillcolor="rgba(6,147,227,0.10)",
            line=dict(color="rgba(0,0,0,0)"),
            name=f"±1 Std Dev ({len(hist_8)} yr)",
            hoverinfo="skip", showlegend=True,
        ))

    oly_6 = sorted(complete_years)[-6:]
    if len(oly_6) >= 3:
        oly_vals = [olympic_avg([data_pivot[m].get(y) for y in oly_6]) for m in months]
        fig.add_trace(go.Scatter(
            x=months, y=oly_vals, mode="lines+markers",
            name="6-Yr Olympic Avg",
            line=dict(color="#0693e3", width=2.5, dash="dash"),
            marker=dict(symbol="diamond", size=6, color="#0693e3"),
        ))

    draw_order = [y for y in all_years if y != cy and y != ly]
    if ly:
        draw_order.append(ly)
    draw_order.append(cy)

    for year in draw_order:
        vals  = [data_pivot[m].get(year) for m in months]
        color, width, opacity = _year_style(year, cy, ly, all_years)
        is_key = year in (cy, ly)
        fig.add_trace(go.Scatter(
            x=months, y=vals, mode="lines+markers", name=year,
            line=dict(color=color, width=width),
            marker=dict(size=5 if is_key else 3, color=color),
            opacity=opacity, connectgaps=False,
        ))

    fig.update_layout(**_base_layout(
        f"Seasonal {lbl}Shipments — {field_label} ({unit_short})",
        "Month", f"{lbl}Volume ({unit_short})",
    ))
    _add_chart_watermark(fig, logo_b64)
    return fig


_BAR_COLORS = [
    "#1565c0","#2e7d32","#6a1b9a","#ad1457",
    "#00695c","#e65100","#37474f","#4527a0",
    "#558b2f","#00838f","#4e342e","#283593",
]


def make_column_chart(data_pivot, stats, selected_years, cy,
                      field_label, is_cumulative, months,
                      logo_b64=None, unit_short="TMT") -> go.Figure:
    lbl = "Cumulative " if is_cumulative else ""
    fig = go.Figure()

    mins = [stats[m]["min"] for m in months]
    maxs = [stats[m]["max"] for m in months]
    fig.add_trace(go.Scatter(
        x=months + months[::-1], y=maxs + mins[::-1],
        fill="toself", fillcolor="rgba(6,147,227,0.08)",
        line=dict(color="rgba(0,0,0,0)"),
        name="Hist. Min–Max Range", hoverinfo="skip",
    ))
    olys = [stats[m]["oly_avg"] for m in months]
    fig.add_trace(go.Scatter(
        x=months, y=olys, mode="lines+markers", name="6-Yr Olympic Avg",
        line=dict(color="#0693e3", width=2.5, dash="dot"),
        marker=dict(symbol="diamond", size=7, color="#0693e3"),
    ))

    color_idx = 0
    for year in selected_years:
        vals  = [data_pivot[m].get(year) for m in months]
        color = "#f9a825" if year == cy else _BAR_COLORS[color_idx % len(_BAR_COLORS)]
        if year != cy:
            color_idx += 1
        fig.add_trace(go.Bar(x=months, y=vals, name=year,
                             marker_color=color, opacity=0.85))

    fig.update_layout(
        barmode="group",
        **_base_layout(
            f"{lbl}Monthly Volume Comparison — {field_label} ({unit_short})",
            "Month", f"{lbl}Volume ({unit_short})",
        ),
    )
    _add_chart_watermark(fig, logo_b64)
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# STAT TILES
# ─────────────────────────────────────────────────────────────────────────────
def _compute_tile_stats(df, use_bushels, unit_factor, cfg,
                        arbr_oct_sep: bool = False) -> list:
    tiles = []
    for field in cfg["tile_order"]:
        # If the user toggled Oct-Sep for AR/BR, treat those fields as Oct-Sep too
        mar_feb     = field in cfg["mar_feb_fields"] and not arbr_oct_sep
        months_list = cfg["arbr_months"] if mar_feb else OCT_SEP_MONTHS
        last_month  = cfg["arbr_last_month"] if mar_feb else "Sep"
        my_label    = cfg["arbr_label"] if mar_feb else "Oct–Sep"

        pivot, all_years = (
            build_arbr_pivot(df, field,
                             months_list=cfg["arbr_months"],
                             prev_months=cfg["arbr_prev_months"])
            if mar_feb else build_pivot(df, field)
        )
        if not all_years:
            tiles.append(None)
            continue

        cy = all_years[-1]
        ly = all_years[-2] if len(all_years) >= 2 else None
        complete_years = [y for y in get_complete_years(pivot, last_month) if y != cy]
        oly_years      = sorted(complete_years)[-6:]

        if use_bushels:
            pivot = _apply_unit(pivot, unit_factor)

        latest_month = latest_val = None
        for m in reversed(months_list):
            v = pivot[m].get(cy)
            if v is not None:
                latest_month, latest_val = m, v
                break

        if latest_month is None:
            tiles.append(None)
            continue

        ly_m    = pivot[latest_month].get(ly) if ly else None
        oly_m   = olympic_avg([pivot[latest_month].get(y) for y in oly_years])
        cum_piv = build_cumulative_pivot(pivot, all_years, months_list)
        cy_cum  = cum_piv[latest_month].get(cy)
        ly_cum  = cum_piv[latest_month].get(ly) if ly else None
        oly_cum = olympic_avg([cum_piv[latest_month].get(y) for y in oly_years])

        tiles.append(dict(
            field        = field,
            label        = cfg["fields"][field],
            my_label     = my_label,
            cy           = cy,
            latest_month = latest_month,
            accent       = cfg["tile_accents"].get(field, JSA_CYAN),
            is_import    = field in cfg["import_fields"],
            monthly_val  = latest_val,
            pct_ly_m     = _pct(latest_val, ly_m),
            pct_oly_m    = _pct(latest_val, oly_m),
            cum_val      = cy_cum,
            pct_ly_c     = _pct(cy_cum, ly_cum),
            pct_oly_c    = _pct(cy_cum, oly_cum),
        ))
    return tiles


def _pct_chip(val, label) -> str:
    if val is None:
        return f'<span style="color:#3d4a58;">{label}&thinsp;—</span>'
    color = "#4caf50" if val >= 0 else "#ef5350"
    sign  = "+" if val >= 0 else ""
    return f'<span style="color:{color};font-weight:700;">{label}&thinsp;{sign}{val:.1f}%</span>'


def _render_tile_grid(tiles, unit_short, unit_decimals) -> str:
    """
    Render stat tiles as unified per-country cards (Monthly + MYTD seamlessly
    stacked inside one card), all cards flush against each other and centred.
    """
    cards_html = ""
    total      = sum(1 for t in tiles if t is not None)
    idx        = 0   # position among non-None tiles

    for t in tiles:
        if t is None:
            continue

        accent   = t.get("accent", JSA_CYAN)
        imp_note = (' <span style="font-size:9px;color:#8a9aaa;font-weight:400;">'
                    '(imports)</span>'
                    if t.get("is_import") else "")

        # rounded corners only on the outer edges of the first/last card
        br_tl = "6px" if idx == 0 else "0"
        br_bl = "6px" if idx == 0 else "0"
        br_tr = "6px" if idx == total - 1 else "0"
        br_br = "6px" if idx == total - 1 else "0"
        border_radius = f"{br_tl} {br_tr} {br_br} {br_bl}"

        # right border only on non-last cards (avoids double border)
        right_border = "" if idx == total - 1 else "border-right:1px solid #2e353d;"

        m_vol = fmt_num(t["monthly_val"], unit_decimals) if t["monthly_val"] is not None else "—"
        c_vol = fmt_num(t["cum_val"],     unit_decimals) if t["cum_val"]     is not None else "—"

        card = (
            # ── outer card ──────────────────────────────────────────────────
            f'<div style="flex:1;min-width:140px;background:#21262c;'
            f'border-top:3px solid {accent};'
            f'border-bottom:1px solid #2e353d;border-left:1px solid #2e353d;'
            f'{right_border}'
            f'border-radius:{border_radius};overflow:hidden;">'

            # ── country header ───────────────────────────────────────────────
            f'<div style="text-align:center;font-size:10.5px;font-weight:700;'
            f'color:#ffffff;text-transform:uppercase;letter-spacing:0.7px;'
            f'padding:6px 8px 5px;background:#1a1e22;'
            f'border-bottom:1px solid #2e353d;">'
            f'{t["label"]}{imp_note}</div>'

            # ── monthly section ──────────────────────────────────────────────
            f'<div style="padding:9px 10px 8px;text-align:center;">'
            f'<div style="font-size:8.5px;font-weight:700;color:#5a6878;'
            f'text-transform:uppercase;letter-spacing:0.5px;margin-bottom:3px;">Monthly</div>'
            f'<div style="color:{JSA_CYAN};font-size:19px;font-weight:700;'
            f'line-height:1.1;font-family:Arial;">{m_vol}</div>'
            f'<div style="font-size:9px;color:#3d4a58;margin:2px 0 5px;">'
            f'{unit_short}&nbsp;·&nbsp;'
            f'<span style="color:#5a6878;">{t["latest_month"]}</span>'
            f'&nbsp;<span style="color:#3d4a58;">{t["cy"]}</span></div>'
            f'<div style="font-size:9.5px;display:flex;gap:6px;justify-content:center;">'
            f'{_pct_chip(t["pct_ly_m"],"LY")}&nbsp;{_pct_chip(t["pct_oly_m"],"Avg")}'
            f'</div></div>'

            # ── divider ──────────────────────────────────────────────────────
            f'<div style="border-top:1px solid #2e353d;margin:0;"></div>'

            # ── MYTD section ─────────────────────────────────────────────────
            f'<div style="padding:9px 10px 8px;text-align:center;">'
            f'<div style="font-size:8.5px;font-weight:700;color:{accent};'
            f'text-transform:uppercase;letter-spacing:0.5px;margin-bottom:3px;">MYTD</div>'
            f'<div style="color:{JSA_CYAN};font-size:19px;font-weight:700;'
            f'line-height:1.1;font-family:Arial;">{c_vol}</div>'
            f'<div style="font-size:9px;color:#3d4a58;margin:2px 0 5px;">'
            f'{unit_short}&nbsp;·&nbsp;'
            f'<span style="color:#5a6878;">thru&nbsp;{t["latest_month"]}</span>'
            f'&nbsp;<span style="color:#3d4a58;">{t["cy"]}&nbsp;({t["my_label"]})</span></div>'
            f'<div style="font-size:9.5px;display:flex;gap:6px;justify-content:center;">'
            f'{_pct_chip(t["pct_ly_c"],"LY")}&nbsp;{_pct_chip(t["pct_oly_c"],"Avg")}'
            f'</div></div>'

            f'</div>'  # close card
        )
        cards_html += card
        idx += 1

    return (
        f'<div style="font-family:Arial;margin-bottom:14px;'
        f'display:flex;justify-content:center;">'
        f'<div style="display:flex;flex-wrap:nowrap;width:fit-content;">'
        f'{cards_html}'
        f'</div></div>'
    )


# ─────────────────────────────────────────────────────────────────────────────
# COMMODITY TAB RENDERER
# ─────────────────────────────────────────────────────────────────────────────
def _run_commodity_tab(commodity: str, use_bushels: bool,
                       unit_short: str, unit_decimals: int, unit_long: str,
                       logo_white_b64: str | None):
    """Render the full dashboard for one commodity inside its top-level tab."""
    cfg = COMMODITY_CONFIG[commodity]
    pfx = commodity   # namespace all widget keys

    # Per-commodity unit factor (different bu_lbs for corn vs soybeans)
    unit_factor = _bu_conv_factor(cfg) if use_bushels else 1.0

    # ── AR/BR marketing-year convention toggle ────────────────────────────
    arbr_oct_sep: bool = st.toggle(
        f"🗓️  Argentina & Brazil: use Oct–Sep MY",
        value=False,
        key=f"{pfx}_arbr_oct_sep",
        help=(
            f"By default Argentina & Brazil use a **{cfg['arbr_label']}** "
            f"marketing year for {cfg['label']}.\n\n"
            "Enable this to align them with the **Oct–Sep** convention used by "
            "the US and aggregates — useful for apples-to-apples comparisons."
        ),
    )

    # ── Load data ─────────────────────────────────────────────────────────
    try:
        df = load_data(commodity)
    except FileNotFoundError:
        st.error(f"Excel file not found at:\n`{EXCEL_PATH}`")
        st.stop()
    except Exception as exc:
        st.error(f"Error loading {cfg['label']} data: {exc}")
        st.stop()

    FIELDS         = cfg["fields"]
    MAR_FEB_FIELDS = cfg["mar_feb_fields"]

    # ── Stat tiles ────────────────────────────────────────────────────────
    tile_stats = _compute_tile_stats(df, use_bushels, unit_factor, cfg,
                                     arbr_oct_sep=arbr_oct_sep)
    st.markdown(_render_tile_grid(tile_stats, unit_short, unit_decimals),
                unsafe_allow_html=True)
    st.markdown(
        '<div style="border-top:1px solid #2e353d;margin:4px 0 16px;"></div>',
        unsafe_allow_html=True,
    )

    # ── Country / Category filter ─────────────────────────────────────────
    st.markdown("#### Select Country or Category")
    field_key = f"{pfx}_field"
    if field_key not in st.session_state:
        st.session_state[field_key] = "US"

    filter_cols = st.columns(len(FIELDS))
    for i, (fk, fl) in enumerate(FIELDS.items()):
        with filter_cols[i]:
            btn_type = "primary" if st.session_state[field_key] == fk else "secondary"
            if st.button(fl, key=f"{pfx}_btn_{fk}", type=btn_type,
                         use_container_width=True):
                st.session_state[field_key] = fk
                st.rerun()

    field       = st.session_state[field_key]
    field_label = FIELDS[field]
    # Honour the AR/BR Oct-Sep toggle when determining which MY convention to use
    mar_feb    = field in MAR_FEB_FIELDS and not arbr_oct_sep
    months     = cfg["arbr_months"] if mar_feb else OCT_SEP_MONTHS
    last_month = cfg["arbr_last_month"] if mar_feb else "Sep"
    my_label   = cfg["arbr_label"] if mar_feb else "Oct–Sep"

    # ── Build pivots ──────────────────────────────────────────────────────
    if mar_feb:
        monthly_pivot, all_years = build_arbr_pivot(
            df, field,
            months_list=cfg["arbr_months"],
            prev_months=cfg["arbr_prev_months"],
        )
    else:
        monthly_pivot, all_years = build_pivot(df, field)

    if not all_years:
        st.warning("No valid marketing-year data found.")
        return

    cy = all_years[-1]
    ly = all_years[-2] if len(all_years) >= 2 else None

    complete_years = get_complete_years(monthly_pivot, last_month)
    complete_years = [y for y in complete_years if y != cy]
    oly_years  = sorted(complete_years)[-6:]
    oly_label  = " → ".join(oly_years) if oly_years else "N/A"

    if use_bushels:
        monthly_pivot = _apply_unit(monthly_pivot, unit_factor)

    cum_pivot     = build_cumulative_pivot(monthly_pivot, all_years, months)
    monthly_stats = compute_stats(monthly_pivot, all_years, complete_years,
                                  cy, ly, months, is_cumulative=False)
    cum_stats     = compute_stats(cum_pivot, all_years, complete_years,
                                  cy, ly, months, is_cumulative=True)

    # ── Info strip ────────────────────────────────────────────────────────
    is_import_field = field in cfg["import_fields"]
    flow_type = "Import" if is_import_field else "Export"
    st.markdown(f"""
    <div style="background:{JSA_MID};padding:9px 18px;border-radius:6px;
                margin:10px 0 6px 0;font-family:Arial;font-size:13px;
                display:flex;gap:28px;flex-wrap:wrap;color:#d0d8e0;
                border-left:3px solid {JSA_GREEN};">
        <span>📊 <b style="color:#fff;">Showing:</b> {field_label}
          {"&nbsp;<span style='color:#e57373;font-size:11px;'>(import data)</span>"
           if is_import_field else ""}</span>
        <span>📐 <b style="color:#fff;">Units:</b>
              <span style="color:{JSA_CYAN};font-weight:700;">{unit_long}</span></span>
        <span>📅 <b style="color:#fff;">Marketing Year:</b> {my_label}</span>
        <span>📅 <b style="color:#fff;">Current Year (CY):</b>
              <span style="color:{JSA_CYAN};font-weight:700;">{cy}</span></span>
        <span>📅 <b style="color:#fff;">Last Year (LY):</b> {ly or "N/A"}</span>
        <span>📈 <b style="color:#fff;">Olympic Avg (prior yrs):</b> {oly_label}</span>
    </div>
    """, unsafe_allow_html=True)

    # ── Legend ────────────────────────────────────────────────────────────
    st.markdown(f"""
    <div style="font-family:Arial;font-size:12px;color:#aab4c0;
                margin:6px 0 14px 0;display:flex;gap:20px;flex-wrap:wrap;
                padding:7px 14px;background:#252a2f;border-radius:5px;">
        <span><span style="background:{JSA_CYAN};padding:2px 8px;border-radius:3px;
              color:#fff;font-weight:600;">CY</span>&nbsp;Current Year ({cy})</span>
        <span><span style="background:#2e7d32;padding:2px 8px;border-radius:3px;
              color:#fff;">■</span>&nbsp;2 Highest (prior yrs)</span>
        <span><span style="background:#c62828;padding:2px 8px;border-radius:3px;
              color:#fff;">■</span>&nbsp;2 Lowest (prior yrs)</span>
        <span><span style="background:#f4f6f8;padding:2px 8px;border-radius:3px;
              color:#000;">■</span>&nbsp;Stat Columns (prior yrs only)</span>
        <span style="color:#4caf50;font-weight:600;">+x.x%</span>&nbsp;Above reference&nbsp;
        <span style="color:#ef5350;font-weight:600;">-x.x%</span>&nbsp;Below reference
    </div>
    """, unsafe_allow_html=True)

    # ── Monthly / Cumulative tabs ─────────────────────────────────────────
    tab1, tab2 = st.tabs(["📊  Monthly Shipments", "📈  Cumulative Shipments"])

    with tab1:
        st.markdown(
            f"**Monthly — {field_label}** &nbsp;({unit_short}) &nbsp;|&nbsp; "
            f"{my_label} marketing year &nbsp;|&nbsp; "
            f"Stats reflect prior marketing years only.",
            unsafe_allow_html=True,
        )
        st.markdown(
            render_table_html(monthly_pivot, monthly_stats, all_years,
                              cy, ly, months, decimals=unit_decimals),
            unsafe_allow_html=True,
        )
        st.plotly_chart(
            make_seasonal_chart(monthly_pivot, all_years, cy, complete_years,
                                field_label, False, months,
                                logo_white_b64, unit_short=unit_short),
            use_container_width=True,
        )

    with tab2:
        st.markdown(
            f"**Cumulative — {field_label}** &nbsp;({unit_short}) &nbsp;|&nbsp; "
            f"{my_label} marketing year &nbsp;|&nbsp; "
            f"Stats reflect prior marketing years only.",
            unsafe_allow_html=True,
        )
        st.markdown(
            render_table_html(cum_pivot, cum_stats, all_years,
                              cy, ly, months, decimals=unit_decimals),
            unsafe_allow_html=True,
        )
        st.plotly_chart(
            make_seasonal_chart(cum_pivot, all_years, cy, complete_years,
                                field_label, True, months,
                                logo_white_b64, unit_short=unit_short),
            use_container_width=True,
        )

    # ── Volume Comparison Column Chart ────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📊 Volume Comparison by Month")

    COUNTRY_COLORS = cfg["country_colors"]
    ctrl_ctry, ctrl_yr, ctrl_type = st.columns([2, 3, 1])

    with ctrl_ctry:
        cmp_countries = st.multiselect(
            "Countries / categories",
            options=list(FIELDS.keys()),
            default=[field],
            max_selections=2,
            format_func=lambda k: FIELDS[k],
            key=f"{pfx}_cmp_countries",
        )
    with ctrl_type:
        cmp_mode    = st.radio("Data type", ["Monthly", "Cumulative"],
                               key=f"{pfx}_cmp_mode", horizontal=False)
    use_cum_cmp = cmp_mode == "Cumulative"
    lbl_cum     = "Cumulative " if use_cum_cmp else ""

    multi_country = len(set(cmp_countries)) > 1

    # TWO-COUNTRY MODE
    if multi_country:
        with ctrl_yr:
            st.caption(
                "ℹ️ Two-country comparison uses **Oct–Sep** marketing year "
                "for both to keep things consistent."
            )
            _, oct_sep_years = build_pivot(df, "US")
            cmp_single_year  = st.selectbox(
                "Marketing year", options=oct_sep_years[::-1],
                key=f"{pfx}_cmp_single_year",
            )

        fig = go.Figure()
        for fk in cmp_countries:
            p, _ = build_pivot(df, fk)
            if use_bushels:
                p = _apply_unit(p, unit_factor)
            if use_cum_cmp:
                p = build_cumulative_pivot(p, oct_sep_years, OCT_SEP_MONTHS)
            vals = [p[m].get(cmp_single_year) for m in OCT_SEP_MONTHS]
            fig.add_trace(go.Bar(
                x=OCT_SEP_MONTHS, y=vals, name=FIELDS[fk],
                marker_color=COUNTRY_COLORS.get(fk, "#aaaaaa"), opacity=0.88,
            ))
        fig.update_layout(
            barmode="group",
            **_base_layout(
                f"{lbl_cum}Country Comparison — {cmp_single_year}  (Oct–Sep MY)",
                "Month", f"{lbl_cum}Volume ({unit_short})",
            ),
        )
        _add_chart_watermark(fig, logo_white_b64)
        st.plotly_chart(fig, use_container_width=True)

    # SINGLE-COUNTRY MODE
    else:
        cmp_field = cmp_countries[0] if cmp_countries else field

        if cmp_field == field:
            c_pivot_m, c_pivot_c = monthly_pivot, cum_pivot
            c_all_years, c_cy, c_months = all_years, cy, months
            c_stats_m, c_stats_c = monthly_stats, cum_stats
        else:
            c_mar_feb = cmp_field in MAR_FEB_FIELDS and not arbr_oct_sep
            c_months  = cfg["arbr_months"] if c_mar_feb else OCT_SEP_MONTHS
            c_last_m  = cfg["arbr_last_month"] if c_mar_feb else "Sep"
            c_pivot_m, c_all_years = (
                build_arbr_pivot(df, cmp_field,
                                 months_list=cfg["arbr_months"],
                                 prev_months=cfg["arbr_prev_months"])
                if c_mar_feb else build_pivot(df, cmp_field)
            )
            c_cy      = c_all_years[-1]
            c_ly      = c_all_years[-2] if len(c_all_years) >= 2 else None
            c_complete= [y for y in get_complete_years(c_pivot_m, c_last_m) if y != c_cy]
            if use_bushels:
                c_pivot_m = _apply_unit(c_pivot_m, unit_factor)
            c_pivot_c = build_cumulative_pivot(c_pivot_m, c_all_years, c_months)
            c_stats_m = compute_stats(c_pivot_m, c_all_years, c_complete,
                                      c_cy, c_ly, c_months, is_cumulative=False)
            c_stats_c = compute_stats(c_pivot_c, c_all_years, c_complete,
                                      c_cy, c_ly, c_months, is_cumulative=True)

        with ctrl_yr:
            default_yrs = [y for y in [c_cy,
                           c_all_years[-2] if len(c_all_years) > 1 else None]
                           if y is not None]
            cmp_years = st.multiselect(
                "Marketing years to compare",
                options=c_all_years, default=default_yrs,
                key=f"{pfx}_cmp_years",
            )

        st.caption(
            f"Showing: **{FIELDS[cmp_field]}** &nbsp;|&nbsp; "
            "Dashed line = 6-yr Olympic Avg (prior years) &nbsp;|&nbsp; "
            "Faint band = historical Min–Max range (prior years)."
        )

        if cmp_years:
            c_pivot = c_pivot_c if use_cum_cmp else c_pivot_m
            c_stats = c_stats_c if use_cum_cmp else c_stats_m
            st.plotly_chart(
                make_column_chart(c_pivot, c_stats, cmp_years, c_cy,
                                  FIELDS[cmp_field], use_cum_cmp, c_months,
                                  logo_white_b64, unit_short=unit_short),
                use_container_width=True,
            )
        else:
            st.info("Select at least one marketing year above to display the chart.")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────────────────────────────────────
def main():

    logo_white_b64 = _load_logo_b64(LOGO_WHITE_PATH)
    logo_full_b64  = _load_logo_b64(LOGO_FULL_PATH)

    # ── Header ───────────────────────────────────────────────────────────
    logo_img_tag = (
        f'<img src="{logo_full_b64}" '
        f'style="height:52px;width:auto;display:block;" alt="JSA Logo">'
        if logo_full_b64 else
        '<span style="font-size:22px;font-weight:700;color:#fff;'
        'font-family:Georgia,serif;">JSA</span>'
    )
    st.markdown(f"""
    <div style="background:#1e2124;padding:16px 28px;border-radius:10px;
                margin-bottom:18px;border-bottom:3px solid {JSA_GREEN};
                display:flex;align-items:center;gap:24px;">
        <div style="flex-shrink:0;">{logo_img_tag}</div>
        <div style="border-left:1px solid #484f56;padding-left:22px;flex:1;">
            <h1 style="color:#fff;margin:0;font-size:24px;font-family:Arial;
                       letter-spacing:0.3px;">
                Global Agricultural Export Dashboard
            </h1>
            <p style="color:#aab4c0;margin:5px 0 0 0;font-size:12.5px;font-family:Arial;">
                Monthly shipments in thousands of metric tons (TMT)
                &nbsp;&nbsp;•&nbsp;&nbsp;
                US / Ukraine / Aggregates: Oct–Sep MY
                &nbsp;&nbsp;•&nbsp;&nbsp;
                Corn AR/BR: Mar–Feb MY &nbsp;•&nbsp; Soybeans AR/BR: Apr–Mar MY
            </p>
        </div>
        <div style="flex-shrink:0;text-align:right;">
            <span style="display:inline-block;background:{JSA_GREEN};color:#fff;
                         font-family:Arial;font-size:11px;font-weight:600;
                         padding:4px 12px;border-radius:9999px;letter-spacing:0.5px;">
                RESEARCH ANALYTICS
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Refresh + Unit toggle ─────────────────────────────────────────────
    col_btn, col_toggle, col_note = st.columns([1, 1.6, 4.4])
    with col_btn:
        if st.button("🔄 Refresh Data", use_container_width=True):
            st.cache_data.clear()
            st.toast("Data cache cleared — reloading…", icon="🔄")
            st.rerun()
    with col_toggle:
        use_bushels = st.toggle(
            "📐 Million Bushels (Mbu)",
            value=False,
            key="unit_toggle",
            help=(
                "Convert volumes to Million Bushels (Mbu).\n\n"
                "Corn: 56 lbs/bu  →  1 TMT ≈ 0.03937 Mbu\n"
                "Soybeans: 60 lbs/bu  →  1 TMT ≈ 0.03674 Mbu"
            ),
        )
    with col_note:
        st.caption(
            "Reads live from the Excel file. After updating, "
            "click **Refresh Data** — tables, stats, and charts update automatically."
        )

    unit_short    = "Mbu"  if use_bushels else "TMT"
    unit_long     = "Million Bushels (Mbu)" if use_bushels else "Thousand Metric Tons (TMT)"
    unit_decimals = 1      if use_bushels else 0

    # ── Top-level commodity tabs ──────────────────────────────────────────
    corn_tab, soy_tab = st.tabs(["🌽  Corn", "🫘  Soybeans"])

    with corn_tab:
        _run_commodity_tab("corn", use_bushels, unit_short,
                           unit_decimals, unit_long, logo_white_b64)

    with soy_tab:
        _run_commodity_tab("soybeans", use_bushels, unit_short,
                           unit_decimals, unit_long, logo_white_b64)

    # ── Footer ────────────────────────────────────────────────────────────
    footer_logo = (
        f'<img src="{logo_white_b64}" style="height:36px;width:auto;opacity:0.85;" alt="JSA">'
        if logo_white_b64 else
        '<span style="font-family:Georgia,serif;font-size:18px;font-weight:700;'
        'color:#fff;">JSA</span>'
    )
    st.markdown(f"""
    <div style="background:{JSA_DARK};border-top:3px solid {JSA_GREEN};
                padding:22px 28px;border-radius:8px;margin-top:24px;
                font-family:Arial;display:flex;align-items:center;
                justify-content:space-between;flex-wrap:wrap;gap:16px;">
        <div style="display:flex;align-items:center;gap:18px;">
            {footer_logo}
            <div style="border-left:1px solid #484f56;padding-left:18px;">
                <div style="color:#fff;font-size:13px;font-weight:600;">
                    John Stewart and Associates</div>
                <div style="color:#7a8a9a;font-size:11px;margin-top:3px;">
                    Commodity Research &amp; Analytics</div>
            </div>
        </div>
        <div style="text-align:center;color:#5a6a7a;font-size:11px;line-height:1.6;">
            📁 Source: Corn Exporter Dashboard Data.xlsx<br>
            Stats reflect prior completed marketing years only — current year excluded.
        </div>
        <div style="text-align:right;color:#5a6a7a;font-size:11px;line-height:1.6;">
            <a href="https://www.jpsi.com" target="_blank"
               style="color:{JSA_CYAN};text-decoration:none;font-weight:600;">
               jpsi.com</a><br>info@jpsi.com
        </div>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
