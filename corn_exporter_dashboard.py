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
SEP_AUG_MONTHS = ["Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug"]
JUL_JUN_MONTHS = ["Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May","Jun"]
DEC_NOV_MONTHS = ["Dec","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov"]
JUN_MAY_MONTHS = ["Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May"]
AUG_JUL_MONTHS = ["Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May","Jun","Jul"]
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
        "us_local_months":      SEP_AUG_MONTHS,
        "us_local_last_month":  "Aug",
        "us_local_label":       "Sep–Aug",
        "us_local_prev_months": frozenset({"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug"}),
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
        "us_local_months":      SEP_AUG_MONTHS,
        "us_local_last_month":  "Aug",
        "us_local_label":       "Sep–Aug",
        "us_local_prev_months": frozenset({"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug"}),
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
    "soybeanmeal": {
        "sheet":          "Meal",
        "emoji":          "🌾",
        "label":          "Soybean Meal",
        # 6 raw columns only — TotalNonUS & MajorExporter are computed
        "col_names":      ["MarketYear","Date","Month","US","Brazil","Argentina"],
        "numeric_cols":   ["US","Brazil","Argentina"],
        "fields": {
            "US":            "United States",
            "Brazil":        "Brazil",
            "Argentina":     "Argentina",
            "TotalNonUS":    "Total Non-US",
            "MajorExporter": "Major Exporters",
        },
        "mar_feb_fields":   {"Brazil","Argentina"},
        "arbr_months":      APR_MAR_MONTHS,
        "arbr_last_month":  "Mar",
        "arbr_label":       "Apr–Mar",
        "arbr_prev_months": frozenset({"Jan","Feb","Mar"}),
        "us_local_months":      SEP_AUG_MONTHS,
        "us_local_last_month":  "Aug",
        "us_local_label":       "Sep–Aug",
        "us_local_prev_months": frozenset({"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug"}),
        "import_fields":    set(),
        "non_us_comps":     ["Brazil","Argentina"],
        "major_comps":      ["US","Brazil","Argentina"],
        "tile_order":       ["US","Brazil","Argentina","TotalNonUS","MajorExporter"],
        "tile_accents": {
            "US":            "#f9a825",
            "Brazil":        "#43a047",
            "Argentina":     "#29b6f6",
            "TotalNonUS":    "#8d6e63",
            "MajorExporter": "#ef6c00",
        },
        "country_colors": {
            "US":            "#f9a825",
            "Brazil":        "#43a047",
            "Argentina":     "#4fc3f7",
            "TotalNonUS":    "#8d6e63",
            "MajorExporter": "#ef6c00",
        },
        "bu_lbs":  None,   # meal is not measured in bushels
    },
    "wheat": {
        "sheet":        "Wheat",
        "emoji":        "🌾",
        "label":        "Wheat",
        # Adjust col_names to match your Excel column order exactly.
        # Remove India / China / Brazil if those columns are not yet present.
        # No MarketYear column needed — wheat pivots are built from the Date column.
        "col_names":    ["Date","Month","US","Canada","EU","Russia","Ukraine",
                         "India","China","Argentina","Australia","Brazil",
                         "TotalNonUS","MajorExporter"],
        "numeric_cols": ["US","Canada","EU","Russia","Ukraine",
                         "India","China","Argentina","Australia","Brazil",
                         "TotalNonUS","MajorExporter"],
        "fields": {
            "US":            "United States",
            "Canada":        "Canada",
            "EU":            "EU",
            "Russia":        "Russia",
            "Ukraine":       "Ukraine",
            "India":         "India",
            "China":         "China",
            "Argentina":     "Argentina",
            "Australia":     "Australia",
            "Brazil":        "Brazil",
            "TotalNonUS":    "Total Non-US",
            "MajorExporter": "Major Exporters",
        },
        # Per-field MY convention — all wheat uses year_offset=0 in build_arbr_pivot.
        # Aggregates (TotalNonUS, MajorExporter) are aligned to Jul-Jun (dominant NH).
        "field_my": {
            "US":            dict(months=JUN_MAY_MONTHS,
                                  prev=frozenset({"Jan","Feb","Mar","Apr","May"}),
                                  last="May", label="Jun–May", hemisphere="North"),
            "Canada":        dict(months=AUG_JUL_MONTHS,
                                  prev=frozenset({"Jan","Feb","Mar","Apr","May","Jun","Jul"}),
                                  last="Jul", label="Aug–Jul", hemisphere="North"),
            "EU":            dict(months=JUL_JUN_MONTHS,
                                  prev=frozenset({"Jan","Feb","Mar","Apr","May","Jun"}),
                                  last="Jun", label="Jul–Jun", hemisphere="North"),
            "Russia":        dict(months=JUL_JUN_MONTHS,
                                  prev=frozenset({"Jan","Feb","Mar","Apr","May","Jun"}),
                                  last="Jun", label="Jul–Jun", hemisphere="North"),
            "Ukraine":       dict(months=JUL_JUN_MONTHS,
                                  prev=frozenset({"Jan","Feb","Mar","Apr","May","Jun"}),
                                  last="Jun", label="Jul–Jun", hemisphere="North"),
            "India":         dict(months=APR_MAR_MONTHS,
                                  prev=frozenset({"Jan","Feb","Mar"}),
                                  last="Mar", label="Apr–Mar", hemisphere="North"),
            "China":         dict(months=JUN_MAY_MONTHS,
                                  prev=frozenset({"Jan","Feb","Mar","Apr","May"}),
                                  last="May", label="Jun–May", hemisphere="North"),
            "Argentina":     dict(months=DEC_NOV_MONTHS,
                                  prev=frozenset({"Jan","Feb","Mar","Apr","May","Jun",
                                                  "Jul","Aug","Sep","Oct","Nov"}),
                                  last="Nov", label="Dec–Nov", hemisphere="South"),
            "Australia":     dict(months=DEC_NOV_MONTHS,
                                  prev=frozenset({"Jan","Feb","Mar","Apr","May","Jun",
                                                  "Jul","Aug","Sep","Oct","Nov"}),
                                  last="Nov", label="Dec–Nov", hemisphere="South"),
            "Brazil":        dict(months=OCT_SEP_MONTHS,
                                  prev=frozenset({"Jan","Feb","Mar","Apr","May","Jun",
                                                  "Jul","Aug","Sep"}),
                                  last="Sep", label="Oct–Sep", hemisphere="South"),
            "TotalNonUS":    dict(months=JUL_JUN_MONTHS,
                                  prev=frozenset({"Jan","Feb","Mar","Apr","May","Jun"}),
                                  last="Jun", label="Jul–Jun", hemisphere="North"),
            "MajorExporter": dict(months=JUL_JUN_MONTHS,
                                  prev=frozenset({"Jan","Feb","Mar","Apr","May","Jun"}),
                                  last="Jun", label="Jul–Jun", hemisphere="North"),
        },
        # NH comparison window (Jul–Jun)
        "nh_months": JUL_JUN_MONTHS,
        "nh_prev":   frozenset({"Jan","Feb","Mar","Apr","May","Jun"}),
        "nh_last":   "Jun",
        "nh_label":  "Jul–Jun",
        # SH comparison window (Dec–Nov)
        "sh_months": DEC_NOV_MONTHS,
        "sh_prev":   frozenset({"Jan","Feb","Mar","Apr","May","Jun",
                                "Jul","Aug","Sep","Oct","Nov"}),
        "sh_last":   "Nov",
        "sh_label":  "Dec–Nov",
        "nh_fields": {"US","Canada","EU","Russia","Ukraine","India","China",
                      "TotalNonUS","MajorExporter"},
        "sh_fields": {"Argentina","Australia","Brazil"},
        "mar_feb_fields": set(),
        "import_fields":  set(),
        "non_us_comps":   ["Canada","EU","Russia","Ukraine","India","China",
                           "Argentina","Australia","Brazil"],
        "major_comps":    ["US","Canada","EU","Russia","Ukraine","India","China",
                           "Argentina","Australia","Brazil"],
        "tile_order":     ["US","Canada","EU","Russia","Ukraine","China",
                           "Argentina","Australia","Brazil","TotalNonUS","MajorExporter"],
        "tile_accents": {
            "US":            "#f9a825",
            "Canada":        "#ef5350",
            "EU":            "#1565c0",
            "Russia":        "#78909c",
            "Ukraine":       "#fdd835",
            "India":         "#ff6d00",
            "China":         "#b71c1c",
            "Argentina":     "#29b6f6",
            "Australia":     "#00897b",
            "Brazil":        "#43a047",
            "TotalNonUS":    "#7e57c2",
            "MajorExporter": "#ef6c00",
        },
        "country_colors": {
            "US":            "#f9a825",
            "Canada":        "#ef5350",
            "EU":            "#1565c0",
            "Russia":        "#78909c",
            "Ukraine":       "#fdd835",
            "India":         "#ff6d00",
            "China":         "#b71c1c",
            "Argentina":     "#4fc3f7",
            "Australia":     "#4db6ac",
            "Brazil":        "#43a047",
            "TotalNonUS":    "#7e57c2",
            "MajorExporter": "#ef6c00",
        },
        "bu_lbs": 60.0,   # lbs per bushel of wheat
    },
}


def _bu_conv_factor(cfg: dict) -> float:
    """TMT → Million Bushels conversion factor for this commodity.
    Returns 1.0 for commodities with no bushel equivalent (e.g. soybean meal)."""
    if not cfg.get("bu_lbs"):
        return 1.0
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
    # For wheat, map whatever header names are in the Excel to our internal
    # field names, then keep only recognised columns.
    # For other commodities use the exact positional slice as before.
    if commodity == "wheat":
        _WHEAT_HEADER_MAP = {
            # Date / Month
            "Date": "Date", "Month": "Month",
            "MarketYear": "MarketYear", "Market Year": "MarketYear",
            # United States
            "US": "US", "U.S.": "US", "United States": "US", "USA": "US",
            # Canada
            "Canada": "Canada",
            # EU
            "EU": "EU", "E.U.": "EU", "European Union": "EU", "Euro Union": "EU",
            # Russia
            "Russia": "Russia", "Russian Federation": "Russia",
            # Ukraine
            "Ukraine": "Ukraine",
            # India
            "India": "India",
            # China
            "China": "China", "China (Mainland)": "China",
            # Argentina
            "Argentina": "Argentina",
            # Australia
            "Australia": "Australia",
            # Brazil
            "Brazil": "Brazil",
            # Aggregates
            "TotalNonUS": "TotalNonUS", "Total Non-US": "TotalNonUS",
            "Total Non US": "TotalNonUS", "Non-US Total": "TotalNonUS",
            "MajorExporter": "MajorExporter", "Major Exporters": "MajorExporter",
            "Major Exporter": "MajorExporter",
        }
        df = df.rename(columns=_WHEAT_HEADER_MAP)
        keep = [c for c in cfg["col_names"] if c in df.columns]
        df   = df[keep].copy()
    else:
        n   = len(cfg["col_names"])
        df  = df.iloc[:, :n].copy()
        df.columns = cfg["col_names"]

    if "MarketYear" in df.columns:
        df["MarketYear"] = df["MarketYear"].astype(str).str.strip()
    df["Month"] = df["Month"].astype(str).str.strip()
    # Normalise full month names → 3-letter abbreviations (e.g. "January" → "Jan")
    _FULL_TO_ABB = {
        "January":"Jan","February":"Feb","March":"Mar","April":"Apr",
        "May":"May","June":"Jun","July":"Jul","August":"Aug",
        "September":"Sep","October":"Oct","November":"Nov","December":"Dec",
    }
    df["Month"] = df["Month"].replace(_FULL_TO_ABB)
    df["Date"]       = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Month"].isin(ALL_MONTHS)].copy()
    for col in cfg["numeric_cols"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return _enforce_aggregate_completeness(df, cfg)


def _enforce_aggregate_completeness(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """
    Recompute TotalNonUS and MajorExporter from components.
    Only uses component columns that actually exist in the dataframe
    (handles wheat where some country columns may not be present yet).
    """
    df = df.copy()
    non_us_cols = [c for c in cfg["non_us_comps"] if c in df.columns]
    major_cols  = [c for c in cfg["major_comps"]  if c in df.columns]
    if non_us_cols:
        all_non_us = df[non_us_cols].notna().all(axis=1)
        df["TotalNonUS"] = np.where(all_non_us, df[non_us_cols].sum(axis=1), np.nan)
    if major_cols:
        all_major = df[major_cols].notna().all(axis=1)
        df["MajorExporter"] = np.where(all_major, df[major_cols].sum(axis=1), np.nan)
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
                     prev_months=None,
                     year_offset: int = -1) -> tuple[dict, list[str]]:
    """
    Build a marketing-year pivot keyed by derived year label.

    months_list  : ordered month list for this MY convention
    prev_months  : months that belong to the *previous* MY label
    year_offset  : controls the year-label formula:
                     -1  Southern Hemisphere planting convention (AR/BR default)
                         non-prev → year-1,  prev → year-2
                      0  Standard convention (US Sep-Aug)
                         non-prev → year,    prev → year-1
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
        start = (year + year_offset - 1) if month in prev_months \
                else (year + year_offset)
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

    # Find the latest month CY has reported, then compare CY / LY / Oly Avg
    # all at that same reference month — always apples-to-apples mid-year.
    cy_reported = [m for m in months if data_pivot[m].get(cy) is not None]
    latest_m    = cy_reported[-1] if cy_reported else None
    m_idx       = months.index(latest_m) if latest_m is not None else -1
    months_thru = months[: m_idx + 1]

    if latest_m is not None:
        if is_cumulative:
            # For cumulative pivot, value at latest_m IS the running total
            cy_ref  = data_pivot[latest_m].get(cy)
            ly_ref  = data_pivot[latest_m].get(ly) if ly else None
            oly_ref = olympic_avg([data_pivot[latest_m].get(y) for y in oly_years])
        else:
            # For monthly pivot, sum only the months CY has reported
            def _sum_thru(year):
                nv = [v for m in months_thru
                      if (v := data_pivot[m].get(year)) is not None]
                return sum(nv) if nv else None
            cy_ref  = _sum_thru(cy)
            ly_ref  = _sum_thru(ly) if ly else None
            oly_ref = olympic_avg([_sum_thru(y) for y in oly_years])
    else:
        cy_ref = ly_ref = oly_ref = None

    stats["TOTAL"] = dict(
        oly_avg    = oly_t,
        min        = min(clean_ht) if clean_ht else None,
        max        = max(clean_ht) if clean_ht else None,
        pct_vs_ly  = _pct(cy_ref, ly_ref),
        pct_vs_oly = _pct(cy_ref, oly_ref),
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
    dec = 1 if unit_short == "Mbu" else 0          # decimal places for volumes
    vol_fmt = f",.{dec}f"                           # Plotly format string
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
            legendrank=900,
        ))

    oly_6 = sorted(complete_years)[-6:]
    if len(oly_6) >= 3:
        oly_vals = [olympic_avg([data_pivot[m].get(y) for y in oly_6]) for m in months]
        oly_hover = [
            f"{v:,.{dec}f} {unit_short}" if v is not None else "—"
            for v in oly_vals
        ]
        fig.add_trace(go.Scatter(
            x=months, y=oly_vals, mode="lines+markers",
            name="6-Yr Olympic Avg",
            line=dict(color="#0693e3", width=2.5, dash="dash"),
            marker=dict(symbol="diamond", size=6, color="#0693e3"),
            legendrank=800,
            customdata=oly_hover,
            hovertemplate="%{x}: %{customdata}<extra>6-Yr Olympic Avg</extra>",
        ))

    # Build year draw order (oldest first for rendering so newer lines sit on top)
    draw_order = [y for y in all_years if y != cy and y != ly]
    if ly:
        draw_order.append(ly)
    draw_order.append(cy)

    # Legend rank: CY=1 (top), LY=2, then historical newest→oldest (3, 4, 5…)
    hist_years_newest_first = [y for y in reversed(all_years) if y != cy and y != ly]
    legend_rank = {cy: 1}
    if ly:
        legend_rank[ly] = 2
    for rank, yr in enumerate(hist_years_newest_first, start=3):
        legend_rank[yr] = rank

    for year in draw_order:
        vals  = [data_pivot[m].get(year) for m in months]
        color, width, opacity = _year_style(year, cy, ly, all_years)
        is_key = year in (cy, ly)
        yr_hover = [
            f"{v:,.{dec}f} {unit_short}" if v is not None else "—"
            for v in vals
        ]
        fig.add_trace(go.Scatter(
            x=months, y=vals, mode="lines+markers", name=year,
            line=dict(color=color, width=width),
            marker=dict(size=5 if is_key else 3, color=color),
            opacity=opacity, connectgaps=False,
            legendrank=legend_rank.get(year, 500),
            customdata=yr_hover,
            hovertemplate="%{x}: %{customdata}<extra>%{fullData.name}</extra>",
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
    dec = 1 if unit_short == "Mbu" else 0
    vol_fmt = f",.{dec}f"
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
    oly_hover = [
        f"{v:,.{dec}f} {unit_short}" if v is not None else "—"
        for v in olys
    ]
    fig.add_trace(go.Scatter(
        x=months, y=olys, mode="lines+markers", name="6-Yr Olympic Avg",
        line=dict(color="#0693e3", width=2.5, dash="dot"),
        marker=dict(symbol="diamond", size=7, color="#0693e3"),
        customdata=oly_hover,
        hovertemplate="%{x}: %{customdata}<extra>6-Yr Olympic Avg</extra>",
    ))

    color_idx = 0
    for year in selected_years:
        vals  = [data_pivot[m].get(year) for m in months]
        color = "#f9a825" if year == cy else _BAR_COLORS[color_idx % len(_BAR_COLORS)]
        if year != cy:
            color_idx += 1
        bar_hover = [
            f"{v:,.{dec}f} {unit_short}" if v is not None else "—"
            for v in vals
        ]
        fig.add_trace(go.Bar(
            x=months, y=vals, name=year,
            marker_color=color, opacity=0.85,
            customdata=bar_hover,
            hovertemplate="%{x}: %{customdata}<extra>%{fullData.name}</extra>",
        ))

    fig.update_layout(
        barmode="group",
        **_base_layout(
            f"{lbl}Monthly Volume Comparison — {field_label} ({unit_short})",
            "Month", f"{lbl}Volume ({unit_short})",
        ),
    )
    _add_chart_watermark(fig, logo_b64)
    return fig


def make_snapshot_chart(snap_data, commodity_label, selected_year_label,
                        unit_short, logo_b64=None) -> go.Figure:
    """
    Single grouped horizontal bar chart.
    For each country:
      • Top bar  — % vs Olympic Avg  (solid, full opacity)
      • Bottom bar — % vs Last Year  (same colour family, 55% opacity)
    Countries sorted highest → lowest by % vs Avg.
    Green = above reference, Red = below.
    """
    valid = [d for d in snap_data
             if d.get("pct_avg") is not None or d.get("pct_ly") is not None]
    if not valid:
        return go.Figure()

    # Sort ascending → highest ends up at the top of the horizontal chart
    sorted_data = sorted(
        valid,
        key=lambda d: (d["pct_avg"] if d.get("pct_avg") is not None else -9999),
    )

    labels   = [d["label"]       for d in sorted_data]
    pct_avgs = [d.get("pct_avg") for d in sorted_data]
    pct_lys  = [d.get("pct_ly")  for d in sorted_data]

    def _colors(vals):
        return ["#4caf50" if (v is not None and v >= 0) else "#ef5350" for v in vals]

    n     = len(labels)
    fig_h = max(400, n * 58 + 160)

    all_vals = [v for v in pct_avgs + pct_lys if v is not None]
    if all_vals:
        mx      = max(abs(v) for v in all_vals)
        pad     = mx * 0.30 + 8
        x_range = [-(mx + pad), mx + pad]
    else:
        x_range = [-30, 30]

    fig = go.Figure()

    # Pre-format hover strings in Python (avoids d3 format flag issues)
    hover_avg = [f"{v:+.1f}%" if v is not None else "N/A" for v in pct_avgs]
    hover_ly  = [f"{v:+.1f}%" if v is not None else "N/A" for v in pct_lys]

    # ── Trace 1: % vs Olympic Avg (solid — drawn first so it appears on TOP) ──
    fig.add_trace(go.Bar(
        x=pct_avgs, y=labels,
        orientation="h",
        name="% vs Olympic Avg",
        marker_color=_colors(pct_avgs),
        marker_line_width=0,
        opacity=1.0,
        text=[f"Avg {v:+.1f}%" if v is not None else "Avg —" for v in pct_avgs],
        textposition="outside",
        cliponaxis=False,
        customdata=hover_avg,
        hovertemplate="%{y}: %{customdata}<extra>% vs Olympic Avg</extra>",
    ))

    # ── Trace 2: % vs Last Year (lighter — appears BELOW in grouped chart) ────
    fig.add_trace(go.Bar(
        x=pct_lys, y=labels,
        orientation="h",
        name="% vs Last Year",
        marker_color=_colors(pct_lys),
        marker_line_width=0,
        opacity=0.55,
        text=[f"LY  {v:+.1f}%" if v is not None else "LY  —" for v in pct_lys],
        textposition="outside",
        cliponaxis=False,
        customdata=hover_ly,
        hovertemplate="%{y}: %{customdata}<extra>% vs Last Year</extra>",
    ))

    # Single zero-reference line across the whole chart
    fig.add_vline(x=0, line_color="#8a9aaa", line_width=1.2)

    fig.update_layout(
        barmode="group",
        height=fig_h,
        title=dict(
            text=(f"{commodity_label} — Cumulative Shipment Snapshot"
                  f"  ·  {selected_year_label}"),
            font=dict(size=14, color="#ffffff", family="Arial"),
            x=0.01,
        ),
        paper_bgcolor="#181c20",
        plot_bgcolor="#1d2227",
        font=dict(family="Arial", color="#d0d8e0", size=11),
        legend=dict(
            orientation="h",
            yanchor="bottom", y=1.02,
            xanchor="left", x=0,
            font=dict(color="#aab4c0", size=11),
            bgcolor="rgba(0,0,0,0)",
            itemsizing="constant",
        ),
        margin=dict(l=10, r=120, t=80, b=20),
        bargap=0.28,
        bargroupgap=0.06,
    )
    fig.update_xaxes(
        ticksuffix="%",
        gridcolor="#2e353d", gridwidth=0.5,
        zeroline=False,
        tickfont=dict(size=10),
        range=x_range,
    )
    fig.update_yaxes(
        gridcolor="#2e353d", gridwidth=0.5,
        tickfont=dict(size=11),
    )

    # Annotation note at bottom explaining opacity difference
    fig.add_annotation(
        text="■ Solid = % vs Olympic Avg     ■ Light = % vs Last Year"
             "     🟢 Above reference     🔴 Below reference",
        xref="paper", yref="paper",
        x=0.0, y=-0.04,
        xanchor="left", yanchor="top",
        showarrow=False,
        font=dict(size=10, color="#5a6878", family="Arial"),
    )

    _add_chart_watermark(fig, logo_b64)
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# STAT TILES
# ─────────────────────────────────────────────────────────────────────────────
def _compute_tile_stats(df, use_bushels, unit_factor, cfg,
                        arbr_local_my: bool = True,
                        us_local_my: bool = True) -> list:
    tiles = []
    for field in cfg["tile_order"]:
        # Determine marketing year convention for this field
        if field == "US" and us_local_my:
            months_list = cfg["us_local_months"]
            last_month  = cfg["us_local_last_month"]
            my_label    = f"Local ({cfg['us_local_label']})"
            pivot, all_years = build_arbr_pivot(
                df, field,
                months_list=cfg["us_local_months"],
                prev_months=cfg["us_local_prev_months"],
                year_offset=0,
            )
        elif field in cfg["mar_feb_fields"] and arbr_local_my:
            months_list = cfg["arbr_months"]
            last_month  = cfg["arbr_last_month"]
            my_label    = f"Local ({cfg['arbr_label']})"
            pivot, all_years = build_arbr_pivot(
                df, field,
                months_list=cfg["arbr_months"],
                prev_months=cfg["arbr_prev_months"],
            )
        else:
            months_list = OCT_SEP_MONTHS
            last_month  = "Sep"
            my_label    = "USDA (Oct–Sep)"
            pivot, all_years = build_pivot(df, field)
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
# SNAPSHOT DATA HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _compute_snapshot_data(df, cfg, use_bushels, unit_factor,
                            selected_year=None, use_local_my=False) -> list:
    """
    Compute per-country snapshot stats.
    use_local_my=False (default) → USDA Oct–Sep for all countries.
    use_local_my=True            → US uses Sep–Aug, AR/BR use Mar–Feb or Apr–Mar.
    selected_year=None  → CY MYTD (latest reported month per country).
    selected_year=str   → full MY total for that prior year.
    Returns list of {field, label, pct_avg, pct_ly, cum_val, year}.
    """
    result = []
    for field in cfg["tile_order"]:
        try:
            if use_local_my and field == "US":
                pivot, all_years = build_arbr_pivot(
                    df, field,
                    months_list=cfg["us_local_months"],
                    prev_months=cfg["us_local_prev_months"],
                    year_offset=0,
                )
                months_list = cfg["us_local_months"]
                last_month  = cfg["us_local_last_month"]
            elif use_local_my and field in cfg["mar_feb_fields"]:
                pivot, all_years = build_arbr_pivot(
                    df, field,
                    months_list=cfg["arbr_months"],
                    prev_months=cfg["arbr_prev_months"],
                )
                months_list = cfg["arbr_months"]
                last_month  = cfg["arbr_last_month"]
            else:
                pivot, all_years = build_pivot(df, field)
                months_list = OCT_SEP_MONTHS
                last_month  = "Sep"
        except Exception:
            continue
        if not all_years:
            continue
        if use_bushels:
            pivot = _apply_unit(pivot, unit_factor)
        cum_piv  = build_cumulative_pivot(pivot, all_years, months_list)
        cy       = all_years[-1]
        ly       = all_years[-2] if len(all_years) >= 2 else None
        complete = [y for y in get_complete_years(pivot, last_month) if y != cy]

        if selected_year is None:
            # Current year: find latest reported cumulative month
            year     = cy
            ref_ly   = ly
            latest_m = None
            for m in reversed(months_list):
                if cum_piv[m].get(cy) is not None:
                    latest_m = m
                    break
            if latest_m is None:
                continue
            cum_val = cum_piv[latest_m].get(cy)
            ly_val  = cum_piv[latest_m].get(ly) if ly else None
            oly_yrs = sorted(complete)[-6:]
            oly_val = olympic_avg([cum_piv[latest_m].get(y) for y in oly_yrs])
        else:
            year = selected_year
            if year not in all_years:
                continue
            yr_idx = all_years.index(year)
            ref_ly = all_years[yr_idx - 1] if yr_idx > 0 else None
            # Full MY total at last month of this MY; fallback to latest available
            cum_val = cum_piv[last_month].get(year)
            if cum_val is None:
                for m in reversed(months_list):
                    v = cum_piv[m].get(year)
                    if v is not None:
                        cum_val = v
                        break
            if cum_val is None:
                continue
            ly_val  = cum_piv[last_month].get(ref_ly) if ref_ly else None
            prior   = [y for y in complete if y != year]
            oly_yrs = sorted(prior)[-6:]
            oly_val = olympic_avg([cum_piv[last_month].get(y) for y in oly_yrs])

        result.append(dict(
            field   = field,
            label   = cfg["fields"][field],
            year    = year,
            cum_val = cum_val,
            pct_avg = _pct(cum_val, oly_val),
            pct_ly  = _pct(cum_val, ly_val),
        ))
    return result


def _compute_wheat_snapshot_data(df, cfg, use_bushels, unit_factor,
                                  nh_compare, sh_compare,
                                  selected_year=None) -> list:
    """
    Per-country wheat snapshot using each country's own MY convention.
    selected_year=None  → newest available data per country (MYTD).
    selected_year=str   → full MY total for that prior year.
    """
    result = []
    for field in cfg["tile_order"]:
        if field not in df.columns:
            continue
        months, prev_months, last_month, _ = _get_wheat_field_my(
            field, cfg, nh_compare, sh_compare
        )
        try:
            pivot, all_years = build_arbr_pivot(
                df, field,
                months_list=months, prev_months=prev_months, year_offset=0,
            )
        except Exception:
            continue
        if not all_years:
            continue
        if use_bushels:
            pivot = _apply_unit(pivot, unit_factor)
        cum_piv  = build_cumulative_pivot(pivot, all_years, months)
        complete = get_complete_years(pivot, last_month)

        if selected_year is None:
            # Newest year with any data (countries report at different times)
            year     = None
            latest_m = None
            for yr in reversed(all_years):
                for m in reversed(months):
                    if cum_piv[m].get(yr) is not None:
                        year, latest_m = yr, m
                        break
                if year:
                    break
            if year is None:
                continue
            yr_idx = all_years.index(year)
            ref_ly = all_years[yr_idx - 1] if yr_idx > 0 else None
            complete_for = [y for y in complete if y != year]
            cum_val = cum_piv[latest_m].get(year)
            ly_val  = cum_piv[latest_m].get(ref_ly) if ref_ly else None
            oly_yrs = sorted(complete_for)[-6:]
            oly_val = olympic_avg([cum_piv[latest_m].get(y) for y in oly_yrs])
        else:
            year = selected_year
            if year not in all_years:
                continue
            yr_idx = all_years.index(year)
            ref_ly = all_years[yr_idx - 1] if yr_idx > 0 else None
            cum_val = cum_piv[last_month].get(year)
            if cum_val is None:
                for m in reversed(months):
                    v = cum_piv[m].get(year)
                    if v is not None:
                        cum_val = v
                        break
            if cum_val is None:
                continue
            ly_val = cum_piv[last_month].get(ref_ly) if ref_ly else None
            complete_for = [y for y in complete if y != year]
            oly_yrs = sorted(complete_for)[-6:]
            oly_val = olympic_avg([cum_piv[last_month].get(y) for y in oly_yrs])

        result.append(dict(
            field   = field,
            label   = cfg["fields"][field],
            year    = year,
            cum_val = cum_val,
            pct_avg = _pct(cum_val, oly_val),
            pct_ly  = _pct(cum_val, ly_val),
        ))
    return result


# ─────────────────────────────────────────────────────────────────────────────
# COMMODITY TAB RENDERER
# ─────────────────────────────────────────────────────────────────────────────
def _run_commodity_tab(commodity: str, use_bushels: bool,
                       unit_short: str, unit_decimals: int, unit_long: str,
                       logo_white_b64: str | None):
    """Render the full dashboard for one commodity inside its top-level tab."""
    cfg = COMMODITY_CONFIG[commodity]
    pfx = commodity   # namespace all widget keys

    # Per-commodity unit factor — meal has no bushel equivalent so always TMT
    supports_bushels = bool(cfg.get("bu_lbs"))
    if not supports_bushels:
        use_bushels   = False
        unit_short    = "TMT"
        unit_long     = "Thousand Metric Tons (TMT)"
        unit_decimals = 0
    unit_factor = _bu_conv_factor(cfg) if use_bushels else 1.0

    # ── Marketing-year convention toggles ─────────────────────────────────
    tog_col1, tog_col2 = st.columns(2)
    with tog_col1:
        arbr_local_my: bool = st.toggle(
            f"🌎  AR/BR: Local MY ({cfg['arbr_label']})",
            value=True,
            key=f"{pfx}_arbr_local_my",
            help=(
                f"**ON** (default) — Argentina & Brazil use their native "
                f"**{cfg['arbr_label']}** marketing year for {cfg['label']}.\n\n"
                "**OFF** — Aligns them with the USDA **Oct–Sep** convention "
                "for apples-to-apples cross-country comparisons."
            ),
        )
    with tog_col2:
        us_local_my: bool = st.toggle(
            "🇺🇸  US: Local MY (Sep–Aug)",
            value=True,
            key=f"{pfx}_us_local_my",
            help=(
                "**ON** (default) — United States uses its native **Sep–Aug** "
                "marketing year.\n\n"
                "**OFF** — Aligns the US with the USDA **Oct–Sep** convention "
                "for apples-to-apples cross-country comparisons."
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
                                     arbr_local_my=arbr_local_my,
                                     us_local_my=us_local_my)
    st.markdown(_render_tile_grid(tile_stats, unit_short, unit_decimals),
                unsafe_allow_html=True)
    st.markdown(
        '<div style="border-top:1px solid #2e353d;margin:4px 0 16px;"></div>',
        unsafe_allow_html=True,
    )

    # ── Snapshot Chart ────────────────────────────────────────────────────
    st.markdown("### 📸 Export Snapshot — All Countries")

    # Build year list from USDA MY on first available tile field
    _snap_ref_field = cfg["tile_order"][0]
    try:
        _, _snap_all_years = build_pivot(df, _snap_ref_field)
    except Exception:
        _snap_all_years = []

    _snap_cy     = _snap_all_years[-1] if _snap_all_years else None
    _snap_prior  = list(reversed(_snap_all_years[:-1])) if len(_snap_all_years) > 1 else []
    _snap_opts   = [f"Current Year YTD ({_snap_cy})"] + _snap_prior if _snap_cy else []

    if _snap_opts:
        _snap_ctrl1, _snap_ctrl2 = st.columns([3, 1])
        with _snap_ctrl1:
            _snap_sel = st.selectbox(
                "Marketing Year",
                options=_snap_opts,
                index=0,
                key=f"{pfx}_snap_year",
                help=(
                    "**Current Year YTD** — cumulative shipments through the most "
                    "recent reported month for each country.\n\n"
                    "**Prior years** — full marketing year total."
                ),
            )
        with _snap_ctrl2:
            _snap_local_my = st.toggle(
                "Local MY",
                value=False,
                key=f"{pfx}_snap_local_my",
                help=(
                    "**OFF** (default) — USDA Oct–Sep for all countries.\n\n"
                    f"**ON** — US uses Sep–Aug; "
                    f"AR/BR use {cfg['arbr_label']} (local harvest MY)."
                ),
            )

        _is_cy_snap = _snap_sel.startswith("Current Year")
        _sel_yr     = None if _is_cy_snap else _snap_sel
        _my_tag     = f"Local MY" if _snap_local_my else "USDA Oct–Sep"
        _snap_label = (f"CY {_snap_cy} (YTD, {_my_tag})"
                       if _is_cy_snap else f"Full MY {_snap_sel} ({_my_tag})")

        _snap_data = _compute_snapshot_data(df, cfg, use_bushels, unit_factor,
                                            selected_year=_sel_yr,
                                            use_local_my=_snap_local_my)
        if _snap_data:
            st.plotly_chart(
                make_snapshot_chart(_snap_data, cfg["label"], _snap_label,
                                    unit_short, logo_white_b64),
                use_container_width=True,
            )
        else:
            st.info("No snapshot data available for the selected year.")

    st.markdown("---")

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

    # Determine marketing year convention for the selected field
    if field == "US" and us_local_my:
        months     = cfg["us_local_months"]
        last_month = cfg["us_local_last_month"]
        my_label   = f"Local ({cfg['us_local_label']})"
    elif field in MAR_FEB_FIELDS and arbr_local_my:
        months     = cfg["arbr_months"]
        last_month = cfg["arbr_last_month"]
        my_label   = f"Local ({cfg['arbr_label']})"
    else:
        months     = OCT_SEP_MONTHS
        last_month = "Sep"
        my_label   = "USDA (Oct–Sep)"

    # ── Build pivots ──────────────────────────────────────────────────────
    if field == "US" and us_local_my:
        monthly_pivot, all_years = build_arbr_pivot(
            df, field,
            months_list=cfg["us_local_months"],
            prev_months=cfg["us_local_prev_months"],
            year_offset=0,
        )
    elif field in MAR_FEB_FIELDS and arbr_local_my:
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
                "ℹ️ Two-country comparison uses **USDA (Oct–Sep)** marketing year "
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
            if cmp_field == "US" and us_local_my:
                c_months  = cfg["us_local_months"]
                c_last_m  = cfg["us_local_last_month"]
                c_pivot_m, c_all_years = build_arbr_pivot(
                    df, cmp_field,
                    months_list=cfg["us_local_months"],
                    prev_months=cfg["us_local_prev_months"],
                    year_offset=0,
                )
            elif cmp_field in MAR_FEB_FIELDS and arbr_local_my:
                c_months  = cfg["arbr_months"]
                c_last_m  = cfg["arbr_last_month"]
                c_pivot_m, c_all_years = build_arbr_pivot(
                    df, cmp_field,
                    months_list=cfg["arbr_months"],
                    prev_months=cfg["arbr_prev_months"],
                )
            else:
                c_months  = OCT_SEP_MONTHS
                c_last_m  = "Sep"
                c_pivot_m, c_all_years = build_pivot(df, cmp_field)
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
                options=c_all_years[::-1], default=default_yrs,
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
# WHEAT TAB
# ─────────────────────────────────────────────────────────────────────────────
def _get_wheat_field_my(field: str, cfg: dict,
                        nh_compare: bool, sh_compare: bool) -> tuple:
    """Return (months, prev_months, last_month, my_label) for a wheat field."""
    fmy = cfg["field_my"][field]
    hem = fmy["hemisphere"]
    if hem == "North" and nh_compare:
        return (cfg["nh_months"], cfg["nh_prev"], cfg["nh_last"],
                f"NH Compare ({cfg['nh_label']})")
    if hem == "South" and sh_compare:
        return (cfg["sh_months"], cfg["sh_prev"], cfg["sh_last"],
                f"SH Compare ({cfg['sh_label']})")
    return fmy["months"], fmy["prev"], fmy["last"], fmy["label"]


def _compute_wheat_tile_stats(df, use_bushels, unit_factor, cfg,
                              nh_compare, sh_compare) -> list:
    tiles = []
    for field in cfg["tile_order"]:
        if field not in df.columns:
            tiles.append(None)
            continue
        months_list, prev_months, last_month, my_label = _get_wheat_field_my(
            field, cfg, nh_compare, sh_compare
        )
        pivot, all_years = build_arbr_pivot(
            df, field,
            months_list=months_list,
            prev_months=prev_months,
            year_offset=0,
        )
        if not all_years:
            tiles.append(None)
            continue

        if use_bushels:
            pivot = _apply_unit(pivot, unit_factor)

        # Search newest → oldest year to find the most recent data for this
        # country. Countries report at different times so the latest year label
        # may have no data yet for some exporters.
        active_cy    = None
        latest_month = None
        latest_val   = None
        for yr in reversed(all_years):
            for m in reversed(months_list):
                v = pivot[m].get(yr)
                if v is not None:
                    active_cy, latest_month, latest_val = yr, m, v
                    break
            if active_cy:
                break

        if active_cy is None:
            tiles.append(None)
            continue

        cy = active_cy
        cy_idx = all_years.index(cy)
        ly     = all_years[cy_idx - 1] if cy_idx > 0 else None

        complete_years = [y for y in get_complete_years(pivot, last_month) if y != cy]
        oly_years      = sorted(complete_years)[-6:]

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
            is_import    = False,
            monthly_val  = latest_val,
            pct_ly_m     = _pct(latest_val, ly_m),
            pct_oly_m    = _pct(latest_val, oly_m),
            cum_val      = cy_cum,
            pct_ly_c     = _pct(cy_cum, ly_cum),
            pct_oly_c    = _pct(cy_cum, oly_cum),
        ))
    return tiles


def _run_wheat_tab(use_bushels: bool, unit_short: str,
                   unit_decimals: int, unit_long: str,
                   logo_white_b64: str | None):
    """Render the full Wheat dashboard tab."""
    cfg         = COMMODITY_CONFIG["wheat"]
    pfx         = "wheat"
    unit_factor = _bu_conv_factor(cfg) if use_bushels else 1.0

    # ── Hemisphere comparison toggles ────────────────────────────────────
    tog1, tog2 = st.columns(2)
    with tog1:
        nh_compare: bool = st.toggle(
            "🌍  NH: Jul–Jun Comparison MY",
            value=False,
            key="wheat_nh_compare",
            help=(
                "**OFF** (default) — each NH country shows its own Local MY:\n"
                "US: Jun–May · Canada: Aug–Jul · EU/Russia/Ukraine: Jul–Jun · "
                "India: Apr–Mar · China: Jun–May\n\n"
                "**ON** — aligns all NH countries to **Jul–Jun** for "
                "apples-to-apples cross-country comparisons."
            ),
        )
    with tog2:
        sh_compare: bool = st.toggle(
            "🌏  SH: Dec–Nov Comparison MY",
            value=False,
            key="wheat_sh_compare",
            help=(
                "**OFF** (default) — each SH country shows its own Local MY:\n"
                "Argentina: Dec–Nov · Australia: Dec–Nov · Brazil: Oct–Sep\n\n"
                "**ON** — aligns all SH countries to **Dec–Nov** for "
                "apples-to-apples cross-country comparisons."
            ),
        )

    # ── Load data ────────────────────────────────────────────────────────
    try:
        df = load_data("wheat")
    except FileNotFoundError:
        st.error(f"Excel file not found at:\n`{EXCEL_PATH}`")
        st.stop()
    except Exception as exc:
        st.error(f"Error loading Wheat data: {exc}")
        st.stop()

    # Only expose fields that actually exist in the loaded dataframe
    FIELDS = {k: v for k, v in cfg["fields"].items() if k in df.columns}

    # Diagnostic: if no country columns matched, show actual Excel headers
    if not FIELDS:
        st.error(
            f"**No country columns matched.** "
            f"Excel headers found: `{list(df.columns)}`\n\n"
            f"Expected one of: `{list(cfg['fields'].keys())}`"
        )
        st.stop()

    # ── Stat tiles ───────────────────────────────────────────────────────
    tile_stats = _compute_wheat_tile_stats(
        df, use_bushels, unit_factor, cfg, nh_compare, sh_compare
    )
    st.markdown(_render_tile_grid(tile_stats, unit_short, unit_decimals),
                unsafe_allow_html=True)
    st.markdown(
        '<div style="border-top:1px solid #2e353d;margin:4px 0 16px;"></div>',
        unsafe_allow_html=True,
    )

    # ── Wheat Snapshot Chart ──────────────────────────────────────────────
    st.markdown("### 📸 Export Snapshot — All Countries")

    # Build year list from the first available wheat field
    _wsnap_ref = next((f for f in cfg["tile_order"] if f in df.columns), None)
    _wsnap_all_years = []
    if _wsnap_ref:
        try:
            _wm, _wp, _, _ = _get_wheat_field_my(_wsnap_ref, cfg, nh_compare, sh_compare)
            _, _wsnap_all_years = build_arbr_pivot(
                df, _wsnap_ref, months_list=_wm, prev_months=_wp, year_offset=0
            )
        except Exception:
            _wsnap_all_years = []

    _wsnap_cy    = _wsnap_all_years[-1] if _wsnap_all_years else None
    _wsnap_prior = list(reversed(_wsnap_all_years[:-1])) if len(_wsnap_all_years) > 1 else []
    _wsnap_opts  = [f"Current Year YTD ({_wsnap_cy})"] + _wsnap_prior if _wsnap_cy else []

    if _wsnap_opts:
        _wsnap_sel = st.selectbox(
            "Marketing Year",
            options=_wsnap_opts,
            index=0,
            key="wheat_snap_year",
            help=(
                "**Current Year YTD** — cumulative shipments through the most recent "
                "reported month (each country uses its own marketing year).\n\n"
                "**Prior years** — full marketing year total per country."
            ),
        )
        _wis_cy     = _wsnap_sel.startswith("Current Year")
        _wsel_yr    = None if _wis_cy else _wsnap_sel
        _wsnap_lbl  = (f"CY {_wsnap_cy} (YTD, per-country MY)"
                       if _wis_cy else f"Full MY {_wsnap_sel} (per-country MY)")

        _wsnap_data = _compute_wheat_snapshot_data(
            df, cfg, use_bushels, unit_factor,
            nh_compare, sh_compare, selected_year=_wsel_yr,
        )
        if _wsnap_data:
            st.plotly_chart(
                make_snapshot_chart(_wsnap_data, "Wheat", _wsnap_lbl,
                                    unit_short, logo_white_b64),
                use_container_width=True,
            )
        else:
            st.info("No snapshot data available for the selected year.")

    st.markdown("---")

    # ── Country / Category filter ────────────────────────────────────────
    st.markdown("#### Select Country or Category")
    field_key = f"{pfx}_field"
    if field_key not in st.session_state or st.session_state[field_key] not in FIELDS:
        st.session_state[field_key] = next(iter(FIELDS))

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

    months, prev_months, last_month, my_label = _get_wheat_field_my(
        field, cfg, nh_compare, sh_compare
    )

    # ── Build pivots ─────────────────────────────────────────────────────
    monthly_pivot, all_years = build_arbr_pivot(
        df, field,
        months_list=months,
        prev_months=prev_months,
        year_offset=0,
    )
    if not all_years:
        st.warning("No valid marketing-year data found.")
        return

    cy = all_years[-1]
    ly = all_years[-2] if len(all_years) >= 2 else None

    complete_years = [y for y in get_complete_years(monthly_pivot, last_month) if y != cy]
    oly_years      = sorted(complete_years)[-6:]
    oly_label      = " → ".join(oly_years) if oly_years else "N/A"

    if use_bushels:
        monthly_pivot = _apply_unit(monthly_pivot, unit_factor)

    cum_pivot     = build_cumulative_pivot(monthly_pivot, all_years, months)
    monthly_stats = compute_stats(monthly_pivot, all_years, complete_years,
                                  cy, ly, months, is_cumulative=False)
    cum_stats     = compute_stats(cum_pivot, all_years, complete_years,
                                  cy, ly, months, is_cumulative=True)

    # ── Info strip ───────────────────────────────────────────────────────
    st.markdown(f"""
    <div style="background:{JSA_MID};padding:9px 18px;border-radius:6px;
                margin:10px 0 6px 0;font-family:Arial;font-size:13px;
                display:flex;gap:28px;flex-wrap:wrap;color:#d0d8e0;
                border-left:3px solid {JSA_GREEN};">
        <span>📊 <b style="color:#fff;">Showing:</b> {field_label}</span>
        <span>📐 <b style="color:#fff;">Units:</b>
              <span style="color:{JSA_CYAN};font-weight:700;">{unit_long}</span></span>
        <span>📅 <b style="color:#fff;">Marketing Year:</b> {my_label}</span>
        <span>📅 <b style="color:#fff;">Current Year (CY):</b>
              <span style="color:{JSA_CYAN};font-weight:700;">{cy}</span></span>
        <span>📅 <b style="color:#fff;">Last Year (LY):</b> {ly or "N/A"}</span>
        <span>📈 <b style="color:#fff;">Olympic Avg (prior yrs):</b> {oly_label}</span>
    </div>
    """, unsafe_allow_html=True)

    # ── Legend ───────────────────────────────────────────────────────────
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

    # ── Volume Comparison ────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📊 Volume Comparison by Month")

    COUNTRY_COLORS = cfg["country_colors"]
    ctrl_ctry, ctrl_yr, ctrl_type = st.columns([2, 3, 1])

    with ctrl_ctry:
        cmp_countries = st.multiselect(
            "Countries / categories",
            options=[k for k in FIELDS],
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

    if multi_country:
        # Two-country comparison: each uses its own (or comparison-group) MY
        fa, fb = cmp_countries[0], cmp_countries[1]
        ma, pma, _, la = _get_wheat_field_my(fa, cfg, nh_compare, sh_compare)
        mb, pmb, _, _  = _get_wheat_field_my(fb, cfg, nh_compare, sh_compare)
        pa, years_a    = build_arbr_pivot(df, fa, months_list=ma, prev_months=pma, year_offset=0)
        if use_bushels:
            pa = _apply_unit(pa, unit_factor)
        with ctrl_yr:
            st.caption(f"ℹ️ Using **{la}** MY for year selection.")
            cmp_yr = st.selectbox("Marketing year", options=years_a[::-1],
                                  key=f"{pfx}_cmp_single_year")
        fig = go.Figure()
        for fk, mfk, pmfk in [(fa, ma, pma), (fb, mb, pmb)]:
            p, yrs = build_arbr_pivot(df, fk, months_list=mfk, prev_months=pmfk, year_offset=0)
            if use_bushels:
                p = _apply_unit(p, unit_factor)
            if use_cum_cmp:
                p = build_cumulative_pivot(p, yrs, mfk)
            vals = [p[m].get(cmp_yr) if m in p else None for m in ma]
            fig.add_trace(go.Bar(x=ma, y=vals, name=FIELDS[fk],
                                 marker_color=COUNTRY_COLORS.get(fk, "#aaa"), opacity=0.88))
        fig.update_layout(barmode="group",
                          **_base_layout(f"{lbl_cum}Country Comparison — {cmp_yr}",
                                         "Month", f"{lbl_cum}Volume ({unit_short})"))
        _add_chart_watermark(fig, logo_white_b64)
        st.plotly_chart(fig, use_container_width=True)

    else:
        cmp_field = cmp_countries[0] if cmp_countries else field

        if cmp_field == field:
            c_pivot_m, c_pivot_c = monthly_pivot, cum_pivot
            c_all_years, c_cy, c_months = all_years, cy, months
            c_stats_m, c_stats_c = monthly_stats, cum_stats
        else:
            c_months, c_pm, c_lm, _ = _get_wheat_field_my(
                cmp_field, cfg, nh_compare, sh_compare
            )
            c_pivot_m, c_all_years = build_arbr_pivot(
                df, cmp_field,
                months_list=c_months, prev_months=c_pm, year_offset=0,
            )
            c_cy      = c_all_years[-1]
            c_ly      = c_all_years[-2] if len(c_all_years) >= 2 else None
            c_complete= [y for y in get_complete_years(c_pivot_m, c_lm) if y != c_cy]
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
                options=c_all_years[::-1], default=default_yrs,
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
# MARKETING YEARS REFERENCE TAB
# ─────────────────────────────────────────────────────────────────────────────
def _render_my_reference_tab():
    """Render the Marketing Years Reference tab — one simple table per commodity."""

    def _ref_table(title, rows):
        """rows: list of (Country, Hemisphere, USDA MY, Local MY)"""
        hdr = "".join(
            f'<th style="padding:9px 16px;text-align:left;color:#aab4c0;'
            f'font-size:11px;text-transform:uppercase;letter-spacing:0.5px;'
            f'border-bottom:2px solid #0693e3;white-space:nowrap;">{h}</th>'
            for h in ["Country", "Hemisphere", "USDA MY", "Local MY"]
        )
        tbl_rows = ""
        for i, (country, hemi, usda_my, local_my) in enumerate(rows):
            bg = "#21262c" if i % 2 == 0 else "#1d2227"
            same = local_my == usda_my
            local_cell = (
                local_my if same else
                f'<b style="color:#f9a825;">{local_my}</b>'
            )
            tbl_rows += (
                f'<tr style="background:{bg};">'
                f'<td style="padding:8px 16px;color:#e0e8f0;font-size:12.5px;">{country}</td>'
                f'<td style="padding:8px 16px;color:#8a9aaa;font-size:12.5px;">{hemi}</td>'
                f'<td style="padding:8px 16px;color:#d0d8e0;font-size:12.5px;">{usda_my}</td>'
                f'<td style="padding:8px 16px;font-size:12.5px;">{local_cell}</td>'
                f'</tr>'
            )
        return f"""
        <div style="font-family:Arial;margin-bottom:24px;">
          <div style="font-size:13px;font-weight:700;color:#fff;
                      margin-bottom:8px;">{title}</div>
          <div style="border-radius:6px;border:1px solid #2e353d;overflow:hidden;">
            <table style="border-collapse:collapse;width:100%;">
              <thead><tr style="background:#151a1f;">{hdr}</tr></thead>
              <tbody>{tbl_rows}</tbody>
            </table>
          </div>
        </div>
        """

    st.markdown(
        '<p style="font-family:Arial;font-size:12.5px;color:#8a9aaa;margin-bottom:20px;">'
        'Local MY (highlighted in amber) differs from the USDA standard. '
        'Toggle <em>Local MY</em> off on any commodity tab to align all countries '
        'to <b style="color:#0693e3;">USDA Oct–Sep</b> for cross-country comparisons.'
        '</p>',
        unsafe_allow_html=True,
    )

    wheat_html = _ref_table("🌾 Wheat Marketing Year Structure", [
        ("United States", "North", "Jun–May", "Jun–May"),
        ("Canada",        "North", "Aug–Jul", "Aug–Jul"),
        ("EU",            "North", "Jul–Jun", "Jul–Jun"),
        ("Russia",        "North", "Jul–Jun", "Jul–Jun"),
        ("Ukraine",       "North", "Jul–Jun", "Jul–Jun"),
        ("India",         "North", "Apr–Mar", "Apr–Mar"),
        ("China",         "North", "Jun–May", "Jun–May"),
        ("Argentina",     "South", "Dec–Nov", "Dec–Nov"),
        ("Australia",     "South", "Dec–Nov", "Dec–Nov"),
        ("Brazil",        "South", "Oct–Sep", "Oct–Sep"),
        ("Total Non-US",  "—",     "Jul–Jun",  "Jul–Jun"),
        ("Major Exporters","—",    "Jul–Jun",  "Jul–Jun"),
    ])

    corn_html = _ref_table("🌽 Corn Marketing Year Structure", [
        ("United States",         "North", "Oct–Sep", "Sep–Aug"),
        ("Brazil",                "South", "Oct–Sep", "Mar–Feb"),
        ("Argentina",             "South", "Oct–Sep", "Mar–Feb"),
        ("Ukraine",               "North", "Oct–Sep", "Oct–Sep"),
        ("Total Non-US",          "—",     "Oct–Sep", "Oct–Sep"),
        ("Major Exporters",       "—",     "Oct–Sep", "Oct–Sep"),
    ])

    soy_html = _ref_table("🫘 Soybeans Marketing Year Structure", [
        ("United States",                              "North", "Oct–Sep", "Sep–Aug"),
        ("Brazil",                                     "South", "Oct–Sep", "Apr–Mar"),
        ("Argentina",                                  "South", "Oct–Sep", "Apr–Mar"),
        ("Total Non-US",                               "—",     "Oct–Sep", "Oct–Sep"),
        ("Major Exporters",                            "—",     "Oct–Sep", "Oct–Sep"),
        ("China Imports",                              "North", "Oct–Sep", "Oct–Sep"),
    ])

    meal_html = _ref_table("🌾 Soybean Meal Marketing Year Structure", [
        ("United States",   "North", "Oct–Sep", "Sep–Aug"),
        ("Brazil",          "South", "Oct–Sep", "Apr–Mar"),
        ("Argentina",       "South", "Oct–Sep", "Apr–Mar"),
        ("Total Non-US",    "—",     "Oct–Sep", "Oct–Sep"),
        ("Major Exporters", "—",     "Oct–Sep", "Oct–Sep"),
    ])

    st.markdown(wheat_html + corn_html + soy_html + meal_html, unsafe_allow_html=True)

    # ── USDA PSD link ─────────────────────────────────────────────────────
    st.markdown(f"""
    <div style="background:#1a1e22;border:1px solid #2e353d;
                border-left:4px solid {JSA_CYAN};border-radius:6px;
                padding:12px 18px;font-family:Arial;margin-top:4px;">
      <span style="color:#fff;font-weight:700;font-size:12.5px;">
        🌐 USDA PSD Database
      </span>
      <span style="color:#8a9aaa;font-size:12px;margin:0 12px;">
        — Authoritative source for marketing year definitions &amp; export data.
      </span>
      <a href="https://apps.fas.usda.gov/psdonline/app/index.html#/app/downloads"
         target="_blank"
         style="color:{JSA_CYAN};font-weight:600;font-size:12px;text-decoration:none;">
        Open USDA PSD App ↗
      </a>
    </div>
    """, unsafe_allow_html=True)


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
                Default: Local MY per country (US: Sep–Aug · Corn AR/BR: Mar–Feb · Soy &amp; Meal AR/BR: Apr–Mar)
                &nbsp;&nbsp;•&nbsp;&nbsp;
                Toggle off Local MY to align with USDA Oct–Sep for cross-country comparisons
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
    corn_tab, soy_tab, meal_tab, wheat_tab, ref_tab = st.tabs(
        ["🌽  Corn", "🫘  Soybeans", "🌾  Soybean Meal", "🌿  Wheat", "📅  Marketing Years"]
    )

    with corn_tab:
        _run_commodity_tab("corn", use_bushels, unit_short,
                           unit_decimals, unit_long, logo_white_b64)

    with soy_tab:
        _run_commodity_tab("soybeans", use_bushels, unit_short,
                           unit_decimals, unit_long, logo_white_b64)

    with meal_tab:
        _run_commodity_tab("soybeanmeal", use_bushels, unit_short,
                           unit_decimals, unit_long, logo_white_b64)

    with wheat_tab:
        try:
            _run_wheat_tab(use_bushels, unit_short, unit_decimals,
                           unit_long, logo_white_b64)
        except Exception as _e:
            st.error(f"Wheat tab error: {_e}")

    with ref_tab:
        try:
            _render_my_reference_tab()
        except Exception as _e:
            st.error(f"Reference tab error: {_e}")

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
