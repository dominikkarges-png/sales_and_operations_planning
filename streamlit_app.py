"""
S&OP Prognose – Streamlit App (Lemken) – Cloud-Version
Hierarchie (Filter): Primary Product Group → Main Product Group → Product Group → Item → Sales Area
Forecast-Editor: Product Group + Item + Sales Area
Excel-Format: Wide (Jan 21 – Dez 26) – deutsche Monatsabkürzungen

Start:
    pip install streamlit pandas plotly openpyxl
    streamlit run streamlit_app.py
"""

import base64
import re
from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ============================================================
# Konfiguration
# ============================================================
BASE_DIR    = Path(__file__).parent
LOGO_FILE   = BASE_DIR / "logo.png"
LEMKEN_FILE = BASE_DIR / "LemkenGmbH_Logo.png"

PRIMARY       = "#000055"
BLUE_MEDIUM   = "#6688BB"
BLUE_LIGHT    = "#BBCCEE"
BLUE_LIGHTEST = "#E4E4F6"
GRAY_MEDIUM   = "#999999"
GRAY_LIGHT    = "#D2D2D2"

ACCENT_ORANGE     = "#FFC000"
ACCENT_GREEN      = "#92D050"
ACCENT_DARK_GREEN = "#00B050"
ACCENT_RED        = "#C31924"

WP_PALETTE = [PRIMARY, BLUE_MEDIUM, BLUE_LIGHT, GRAY_MEDIUM, GRAY_LIGHT, BLUE_LIGHTEST]
LOGO_MAXH = 80

STATUS_COLOR = {
    "ENTWURF":      GRAY_MEDIUM,
    "ZUR_PRUEFUNG": ACCENT_ORANGE,
    "FREIGEGEBEN":  ACCENT_DARK_GREEN,
    "ABGELEHNT":    ACCENT_RED,
}
STATUS_LABEL = {
    "ENTWURF":      "Entwurf",
    "ZUR_PRUEFUNG": "Zur Prüfung",
    "FREIGEGEBEN":  "Freigegeben",
    "ABGELEHNT":    "Abgelehnt",
}

HIERARCHY_COLS = ["Primary Product Group", "Main Product Group", "Product Group", "Item", "Sales Area"]
KEY_COLS = ["Product Group", "Item", "Sales Area"]

MONTH_MAP = {
    "Jan": 1, "Feb": 2, "Mrz": 3, "Mär": 3, "Apr": 4, "Mai": 5, "Jun": 6,
    "Jul": 7, "Aug": 8, "Sep": 9, "Okt": 10, "Nov": 11, "Dez": 12,
    "Mar": 3, "May": 5, "Oct": 10, "Dec": 12,
}

LOOKBACK_MONTHS = 3

st.set_page_config(page_title="S&OP Prognose", page_icon="📈", layout="wide")

# ============================================================
# Session State – DB-Ersatz (forecast + history als DataFrames)
# ============================================================
def _init_session_db():
    if "db_forecast" not in st.session_state:
        st.session_state["db_forecast"] = pd.DataFrame(columns=[
            "product_group", "item", "sales_area", "monat", "menge",
            "status", "ersteller", "kommentar", "created_at", "updated_at"
        ])
    if "db_history" not in st.session_state:
        st.session_state["db_history"] = pd.DataFrame(columns=[
            "id", "product_group", "item", "sales_area", "monat",
            "menge_alt", "menge_neu", "differenz", "status",
            "user", "kommentar", "aktion", "timestamp"
        ])

_init_session_db()

DB_KEY_COLS = ["product_group", "item", "sales_area"]
COL_TO_DB = dict(zip(KEY_COLS, DB_KEY_COLS))
DB_TO_COL = dict(zip(DB_KEY_COLS, KEY_COLS))


def get_current_values() -> dict:
    df_db = st.session_state["db_forecast"]
    if df_db.empty:
        return {}
    result = {}
    for _, row in df_db.iterrows():
        key = (row["product_group"], row["item"], row["sales_area"], row["monat"])
        result[key] = row["menge"]
    return result


def save_forecast_with_diff(df_long: pd.DataFrame, original_lookup: dict,
                             user: str, status: str, kommentar: str = "") -> int:
    now = datetime.now().isoformat(timespec="seconds")
    aktion = "EINREICHUNG" if status == "ZUR_PRUEFUNG" else "ENTWURF_GESPEICHERT"
    db_vorher = get_current_values()
    geaendert = 0

    df_fc = st.session_state["db_forecast"].copy()
    history_rows = []

    for _, row in df_long.iterrows():
        pg   = row[KEY_COLS[0]]
        item = row[KEY_COLS[1]]
        sa   = row[KEY_COLS[2]]
        monat = row["Monat"]
        menge_neu = int(row["Menge"])

        db_key = (pg, item, sa, monat)
        orig_key = (pg, item, sa, monat)

        if db_key in db_vorher:
            menge_alt = db_vorher[db_key]
        elif orig_key in original_lookup:
            menge_alt = int(original_lookup[orig_key])
        else:
            menge_alt = None

        # Upsert in df_fc
        mask = (
            (df_fc["product_group"] == pg) &
            (df_fc["item"] == item) &
            (df_fc["sales_area"] == sa) &
            (df_fc["monat"] == monat)
        )
        if mask.any():
            df_fc.loc[mask, ["menge", "status", "ersteller", "kommentar", "updated_at"]] = [
                menge_neu, status, user, kommentar, now
            ]
        else:
            new_row = {
                "product_group": pg, "item": item, "sales_area": sa,
                "monat": monat, "menge": menge_neu, "status": status,
                "ersteller": user, "kommentar": kommentar,
                "created_at": now, "updated_at": now
            }
            df_fc = pd.concat([df_fc, pd.DataFrame([new_row])], ignore_index=True)

        if menge_alt is not None and menge_alt != menge_neu:
            history_rows.append({
                "id": len(st.session_state["db_history"]) + len(history_rows) + 1,
                "product_group": pg, "item": item, "sales_area": sa,
                "monat": monat, "menge_alt": menge_alt, "menge_neu": menge_neu,
                "differenz": menge_neu - menge_alt,
                "status": status, "user": user, "kommentar": kommentar,
                "aktion": aktion, "timestamp": now
            })
            geaendert += 1

    st.session_state["db_forecast"] = df_fc
    if history_rows:
        st.session_state["db_history"] = pd.concat(
            [st.session_state["db_history"], pd.DataFrame(history_rows)],
            ignore_index=True
        )
    return geaendert


def update_status(status_neu: str, user: str, kommentar: str = "") -> int:
    now = datetime.now().isoformat(timespec="seconds")
    df_fc = st.session_state["db_forecast"].copy()
    mask = df_fc["status"] == "ZUR_PRUEFUNG"
    affected = mask.sum()
    df_fc.loc[mask, ["status", "updated_at"]] = [status_neu, now]
    st.session_state["db_forecast"] = df_fc

    aktion = "FREIGABE" if status_neu == "FREIGEGEBEN" else "ABLEHNUNG"
    history_row = {
        "id": len(st.session_state["db_history"]) + 1,
        "product_group": None, "item": None, "sales_area": None,
        "monat": None, "menge_alt": None, "menge_neu": None, "differenz": None,
        "status": status_neu, "user": user, "kommentar": kommentar,
        "aktion": aktion, "timestamp": now
    }
    st.session_state["db_history"] = pd.concat(
        [st.session_state["db_history"], pd.DataFrame([history_row])],
        ignore_index=True
    )
    return int(affected)


def load_forecast() -> pd.DataFrame:
    return st.session_state["db_forecast"].copy()


def load_change_history(limit: int = 200) -> pd.DataFrame:
    df_h = st.session_state["db_history"].copy()
    if df_h.empty:
        return df_h
    df_h = df_h[
        df_h["menge_alt"].notna() &
        df_h["menge_neu"].notna() &
        (df_h["menge_alt"] != df_h["menge_neu"])
    ].sort_values("id", ascending=False).head(limit)
    return df_h


def get_current_status() -> str:
    df_db = load_forecast()
    if df_db.empty:
        return "KEINE_DATEN"
    return df_db["status"].mode().iloc[0]


def build_styled_excel(pivot_df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pivot_df.to_excel(writer, index=False, sheet_name="Prognose")
        ws = writer.sheets["Prognose"]

        n_rows = ws.max_row
        n_cols = ws.max_column

        header_fill = PatternFill("solid", start_color="000055")
        header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        body_font   = Font(name="Arial", color="000055", size=10)
        thin = Side(border_style="thin", color="BFBFBF")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        center = Alignment(horizontal="center", vertical="center")
        right  = Alignment(horizontal="right",  vertical="center")
        left   = Alignment(horizontal="left",   vertical="center")

        for col_idx in range(1, n_cols + 1):
            c = ws.cell(row=1, column=col_idx)
            c.fill = header_fill
            c.font = header_font
            c.alignment = center
            c.border = border

        n_key_cols = len(KEY_COLS)
        for row_idx in range(2, n_rows + 1):
            for col_idx in range(1, n_cols + 1):
                c = ws.cell(row=row_idx, column=col_idx)
                c.font = body_font
                c.border = border
                if col_idx <= n_key_cols:
                    c.alignment = left
                else:
                    c.alignment = right
                    c.number_format = "#,##0"

        widths = [22, 18, 14]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for col_idx in range(n_key_cols + 1, n_cols + 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = 11

        ws.freeze_panes = f"{get_column_letter(n_key_cols + 1)}2"
        ws.auto_filter.ref = ws.dimensions

    return buffer.getvalue()


# ============================================================
# CSS (W&P-Design + Sticky Header)
# ============================================================
st.markdown(
    f"""
    <style>
        html, body, [class*="css"], [data-testid="stAppViewContainer"],
        [data-testid="stMarkdownContainer"], [data-testid="stMetricLabel"],
        [data-testid="stMetricValue"], [data-testid="stMetricDelta"],
        [data-testid="stSidebar"], [data-testid="stSidebar"] *,
        .stDataFrame, .stDataEditor, h1, h2, h3, h4, h5, h6,
        p, span, label, div, li, th, td, a {{
            color: {PRIMARY} !important;
        }}
        [data-baseweb="tag"] {{
            background-color: {PRIMARY} !important;
            color: #FFFFFF !important;
        }}
        [data-baseweb="tag"] * {{ color: #FFFFFF !important; }}

        .stButton > button, .stDownloadButton > button {{
            color: {PRIMARY} !important;
            background-color: #FFFFFF !important;
            border: 2px solid {PRIMARY} !important;
            font-weight: 600;
        }}
        .stButton > button *, .stDownloadButton > button * {{ color: {PRIMARY} !important; }}
        .stButton > button:hover, .stDownloadButton > button:hover {{
            background-color: {PRIMARY} !important;
            color: #FFFFFF !important;
        }}
        .stButton > button:hover *, .stDownloadButton > button:hover * {{ color: #FFFFFF !important; }}

        .stDataFrame thead tr th, .stDataEditor thead tr th {{
            background-color: {PRIMARY} !important;
            color: #FFFFFF !important;
            font-weight: 600;
        }}
        .stDataFrame thead tr th *, .stDataEditor thead tr th * {{ color: #FFFFFF !important; }}

        [data-testid="stHeader"] {{
            background: transparent;
            height: 0;
        }}

        .block-container {{
            padding-top: 0.5rem !important;
        }}

        .wp-sticky-header {{
            position: sticky;
            top: 0;
            z-index: 999;
            background-color: #FFFFFF;
            padding: 8px 0 0 0;
            margin: 0 0 12px 0;
            border-bottom: 6px solid {PRIMARY};
            box-shadow: 0 2px 6px rgba(0,0,0,0.08);
        }}

        .wp-sticky-header .wp-header-row {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            padding: 0 8px;
        }}

        .wp-sticky-header .wp-logo-left,
        .wp-sticky-header .wp-logo-right {{
            display: flex;
            align-items: center;
            flex: 0 0 auto;
        }}

        .wp-sticky-header .wp-logo-left img,
        .wp-sticky-header .wp-logo-right img {{
            max-height: {LOGO_MAXH}px;
            width: auto;
            image-rendering: -webkit-optimize-contrast;
            image-rendering: crisp-edges;
            display: block;
        }}

        .wp-sticky-header .wp-title {{
            flex: 1 1 auto;
            text-align: center;
        }}

        .wp-sticky-header .wp-title h1 {{
            color: {PRIMARY} !important;
            margin: 4px 0;
            font-family: Arial, sans-serif;
            font-weight: 700;
            font-size: 2.4rem;
            line-height: 1.1;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)


def encode_logo(path: Path) -> str:
    if not path.exists():
        return ""
    return f"data:image/png;base64,{base64.b64encode(path.read_bytes()).decode()}"


# ============================================================
# Passwort-Schutz
# ============================================================
PASSWORD = "wieselhuberNLP2026!"

if 'authenticated' not in st.session_state:
    col1, col2 = st.columns([1.4, 2.6])
    with col1:
        if LOGO_FILE.exists():
            st.image(LOGO_FILE, width=220)
    with col2:
        st.title("Anmeldung erforderlich")

    password = st.text_input("Passwort eingeben", type="password")
    if st.button("Anmelden"):
        if password == PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Falsches Passwort")
    st.stop()

# ============================================================
# Sticky Header
# ============================================================
logo_wp_src     = encode_logo(LOGO_FILE)
logo_lemken_src = encode_logo(LEMKEN_FILE)

st.markdown(
    f"""
    <div class="wp-sticky-header">
        <div class="wp-header-row">
            <div class="wp-logo-left">
                {f'<img src="{logo_wp_src}" alt="W&P" />' if logo_wp_src else ''}
            </div>
            <div class="wp-title">
                <h1>S&amp;OP Prognose</h1>
            </div>
            <div class="wp-logo-right">
                {f'<img src="{logo_lemken_src}" alt="Lemken" />' if logo_lemken_src else ''}
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# Daten laden + Wide-zu-Long Transformation
# ============================================================
def parse_quantity_columns(df: pd.DataFrame) -> list:
    pattern = re.compile(r"^([A-Za-zÄÖÜäöü]{3})\s+(\d{2,4})$")
    matches = []
    for col in df.columns:
        col_clean = str(col).replace("\xa0", " ").strip()
        col_clean = re.sub(r"\s+", " ", col_clean)

        if "total" in col_clean.lower() or "summe" in col_clean.lower():
            continue

        m = pattern.match(col_clean)
        if not m:
            continue

        mon_str, yr_str = m.group(1), m.group(2)
        mon_key = mon_str[0].upper() + mon_str[1:].lower()
        mon_num = MONTH_MAP.get(mon_key)
        if mon_num is None:
            continue

        year = int(yr_str)
        if year < 100:
            year = 2000 + year

        try:
            datum = pd.Timestamp(year=year, month=mon_num, day=1)
        except ValueError:
            continue

        matches.append((col, datum))
    return matches


def transform_wide_to_long(df_wide: pd.DataFrame) -> pd.DataFrame:
    missing = [c for c in HIERARCHY_COLS if c not in df_wide.columns]
    if missing:
        raise ValueError(
            f"Fehlende Hierarchie-Spalten: {missing}\n\n"
            f"Vorhandene Spalten: {list(df_wide.columns)}"
        )

    qty_cols = parse_quantity_columns(df_wide)
    if not qty_cols:
        non_hier_cols = [c for c in df_wide.columns if c not in HIERARCHY_COLS]
        sample = non_hier_cols[:5]
        sample_repr = [f"{repr(str(c))} (Typ: {type(c).__name__})" for c in sample]
        raise ValueError(
            f"Keine Monatsspalten erkannt. "
            f"Erste Datenspalten in Excel:\n  - " + "\n  - ".join(sample_repr) +
            f"\n\nErwartet z.B.: 'Jan 21', 'Mrz 22'."
        )

    qty_col_names = [c for c, _ in qty_cols]
    col_to_date = {c: d for c, d in qty_cols}

    for c in HIERARCHY_COLS:
        df_wide[c] = df_wide[c].fillna("(unbekannt)").astype(str)

    df_long = df_wide.melt(
        id_vars=HIERARCHY_COLS,
        value_vars=qty_col_names,
        var_name="QtyCol",
        value_name="Absatzmenge",
    )
    df_long["Datum"] = df_long["QtyCol"].map(col_to_date)
    df_long["Absatzmenge"] = pd.to_numeric(df_long["Absatzmenge"], errors="coerce").fillna(0).astype(int)
    df_long["Jahr"] = df_long["Datum"].dt.year
    df_long["Monat"] = df_long["Datum"].dt.month
    df_long = df_long.drop(columns=["QtyCol"])

    df_long = (df_long.groupby(HIERARCHY_COLS + ["Datum", "Jahr", "Monat"], as_index=False)
                      ["Absatzmenge"].sum())
    return df_long


def load_data_from_upload(file_bytes: bytes) -> pd.DataFrame:
    df_wide = pd.read_excel(BytesIO(file_bytes))
    return transform_wide_to_long(df_wide)


# ============================================================
# Data Source Management (Upload only – kein Disk-Fallback)
# ============================================================
if "uploaded_bytes" not in st.session_state:
    st.session_state["uploaded_bytes"] = None
if "cache_buster" not in st.session_state:
    st.session_state["cache_buster"] = 0

# ============================================================
# Sidebar
# ============================================================
st.sidebar.header("Datenquelle")

uploaded = st.sidebar.file_uploader(
    "Excel hochladen (max. 200MB)",
    type=["xlsx"],
    help="Prognose_Absatz.xlsx im Wide-Format (Jan 21 – Dez 26)"
)
if uploaded is not None:
    new_bytes = uploaded.getvalue()
    if new_bytes != st.session_state["uploaded_bytes"]:
        st.session_state["uploaded_bytes"] = new_bytes
        st.session_state["cache_buster"] += 1
        # Reset forecast DB bei neuer Datei
        st.session_state["db_forecast"] = pd.DataFrame(columns=[
            "product_group", "item", "sales_area", "monat", "menge",
            "status", "ersteller", "kommentar", "created_at", "updated_at"
        ])
        st.session_state["db_history"] = pd.DataFrame(columns=[
            "id", "product_group", "item", "sales_area", "monat",
            "menge_alt", "menge_neu", "differenz", "status",
            "user", "kommentar", "aktion", "timestamp"
        ])
        for k in list(st.session_state.keys()):
            if k.startswith("forecast_edit"):
                del st.session_state[k]
        st.sidebar.success("Excel geladen.")
        st.rerun()

if not st.session_state["uploaded_bytes"]:
    st.info("👆 Bitte Excel-Datei in der Sidebar hochladen, um zu starten.")
    st.stop()

try:
    df = load_data_from_upload(st.session_state["uploaded_bytes"])
except Exception as e:
    st.error(f"Datei konnte nicht verarbeitet werden: {e}")
    st.stop()

df_orig = df.copy()
df_orig["MonatStr"] = df_orig["Datum"].dt.strftime("%Y-%m")
ORIGINAL_LOOKUP = (df_orig.groupby(KEY_COLS + ["MonatStr"])["Absatzmenge"]
                   .sum().to_dict())

# ============================================================
# Editor-Datumsspalten: 3 Monate rückwirkend + alle Zukunftsmonate
# ============================================================
heute = pd.Timestamp(datetime.now().date())
heute_monatsanfang = pd.Timestamp(year=heute.year, month=heute.month, day=1)
cutoff_monat = heute_monatsanfang - pd.DateOffset(months=LOOKBACK_MONTHS - 1)
EDITOR_DATES = sorted([d for d in df["Datum"].unique() if pd.Timestamp(d) >= cutoff_monat])
EDITOR_MONAT_STRS = [pd.Timestamp(d).strftime("%Y-%m") for d in EDITOR_DATES]

st.sidebar.markdown("---")
st.sidebar.header("Benutzer")
rolle = st.sidebar.radio("Rolle", ["Planer", "Reviewer"], index=0)
user = st.sidebar.text_input("Name", value="", placeholder="Vor- und Nachname")

st.sidebar.markdown("---")
st.sidebar.header("Status")
current_status = get_current_status()
status_col = STATUS_COLOR.get(current_status, GRAY_MEDIUM)
status_lbl = STATUS_LABEL.get(current_status, "Keine Daten")
st.sidebar.markdown(
    f"""
    <div style="padding:8px 12px; background-color:{status_col};
                color:white; border-radius:4px; font-weight:600; text-align:center;">
        {status_lbl}
    </div>
    """,
    unsafe_allow_html=True,
)

st.sidebar.markdown("---")
st.sidebar.header("Filter")

primary_opts = sorted(df["Primary Product Group"].unique())
sel_primary = st.sidebar.multiselect("Primary Product Group", primary_opts, default=primary_opts)

main_opts = sorted(df.loc[df["Primary Product Group"].isin(sel_primary), "Main Product Group"].unique())
sel_main = st.sidebar.multiselect("Main Product Group", main_opts, default=main_opts)

pg_opts = sorted(df.loc[
    df["Primary Product Group"].isin(sel_primary)
    & df["Main Product Group"].isin(sel_main),
    "Product Group"
].unique())
sel_pg = st.sidebar.multiselect("Product Group", pg_opts, default=pg_opts)

item_opts = sorted(df.loc[
    df["Primary Product Group"].isin(sel_primary)
    & df["Main Product Group"].isin(sel_main)
    & df["Product Group"].isin(sel_pg),
    "Item"
].unique())
sel_item = st.sidebar.multiselect("Item", item_opts, default=item_opts)

sa_opts = sorted(df["Sales Area"].unique())
sel_sa = st.sidebar.multiselect("Sales Area", sa_opts, default=sa_opts)

jahre = sorted(df["Jahr"].unique())
sel_jahr = st.sidebar.multiselect("Jahr", jahre, default=jahre)

monate = list(range(1, 13))
sel_monat = st.sidebar.select_slider("Monatsbereich", options=monate,
                                      value=(min(monate), max(monate)))

# ============================================================
# Filter anwenden
# ============================================================
mask = (
    df["Primary Product Group"].isin(sel_primary)
    & df["Main Product Group"].isin(sel_main)
    & df["Product Group"].isin(sel_pg)
    & df["Item"].isin(sel_item)
    & df["Sales Area"].isin(sel_sa)
    & df["Jahr"].isin(sel_jahr)
    & df["Monat"].between(sel_monat[0], sel_monat[1])
)
df_f = df.loc[mask].copy()

if df_f.empty:
    st.warning("Keine Daten für die aktuelle Filterauswahl.")
    st.stop()

# ============================================================
# KPI + Donut
# ============================================================
col_kpi, col_pie = st.columns([3, 2], gap="large")

with col_kpi:
    k1, k2 = st.columns(2)
    k3, k4 = st.columns(2)
    k1.metric("Datensätze",    f"{len(df_f):,}".replace(",", "."))
    k2.metric("Absatz gesamt", f"{df_f['Absatzmenge'].sum():,.0f}".replace(",", "."))
    k3.metric("Ø pro Monat",   f"{df_f.groupby('Datum')['Absatzmenge'].sum().mean():,.0f}".replace(",", "."))
    k4.metric("Items",         df_f["Item"].nunique())

with col_pie:
    sa_agg = (df_f.groupby("Sales Area", as_index=False)["Absatzmenge"]
              .sum().sort_values("Absatzmenge", ascending=False))

    if len(sa_agg) > 10:
        top10 = sa_agg.head(10).copy()
        rest = sa_agg.iloc[10:]["Absatzmenge"].sum()
        sonstige = pd.DataFrame([{"Sales Area": "Sonstige", "Absatzmenge": rest}])
        sa_agg_display = pd.concat([top10, sonstige], ignore_index=True)
    else:
        sa_agg_display = sa_agg

    fig_pie = px.pie(sa_agg_display, names="Sales Area", values="Absatzmenge",
                     hole=0.5, color_discrete_sequence=WP_PALETTE)
    fig_pie.update_traces(textposition="inside", textinfo="percent",
                          textfont=dict(color="#FFFFFF", size=11),
                          marker=dict(line=dict(color="#FFFFFF", width=2)))
    fig_pie.update_layout(height=200, margin=dict(l=0, r=0, t=0, b=0),
                          legend=dict(orientation="v", yanchor="middle", y=0.5,
                                      xanchor="left", x=1.02,
                                      font=dict(color=PRIMARY, family="Arial", size=10)),
                          font=dict(color=PRIMARY, family="Arial"))
    st.plotly_chart(fig_pie, use_container_width=True)

st.markdown("---")

# ============================================================
# Liniendiagramm
# ============================================================
st.subheader("Absatzverlauf nach Main Product Group")

agg = (df_f.groupby(["Datum", "Main Product Group"], as_index=False)["Absatzmenge"]
       .sum().sort_values("Datum"))

fig = px.line(agg, x="Datum", y="Absatzmenge", color="Main Product Group",
              markers=True, color_discrete_sequence=WP_PALETTE)

fig.update_layout(
    height=420, margin=dict(l=10, r=10, t=20, b=10),
    legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="left", x=0,
                font=dict(color=PRIMARY, family="Arial", size=11)),
    xaxis_title=None, yaxis_title="Menge", plot_bgcolor="white",
    font=dict(color=PRIMARY, family="Arial"),
)
fig.update_xaxes(showgrid=True, gridcolor=GRAY_LIGHT,
                 tickfont=dict(color=PRIMARY), title_font=dict(color=PRIMARY))
fig.update_yaxes(showgrid=True, gridcolor=GRAY_LIGHT,
                 tickfont=dict(color=PRIMARY), title_font=dict(color=PRIMARY))

if not agg.empty:
    x_min, x_max = agg["Datum"].min(), agg["Datum"].max()
    x_axis_min = min(x_min, heute) - pd.Timedelta(days=15)
    x_axis_max = max(x_max, heute) + pd.Timedelta(days=15)
    fig.update_xaxes(range=[x_axis_min, x_axis_max])

    fig.add_shape(
        type="line", x0=heute, x1=heute, y0=0, y1=1,
        yref="paper", xref="x",
        line=dict(color=ACCENT_RED, width=2, dash="dash"), layer="above",
    )
    fig.add_annotation(
        x=heute, y=1.02, xref="x", yref="paper",
        text="Heute", showarrow=False,
        font=dict(color=ACCENT_RED, size=11, family="Arial"),
        bgcolor="rgba(255,255,255,0.9)", bordercolor=ACCENT_RED,
        borderwidth=1, borderpad=3,
    )

st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# ============================================================
# Prognose-Editor
# ============================================================
st.subheader("Prognose (editierbar)")
st.caption(
    f"Sichtbar: letzte {LOOKBACK_MONTHS} Monate ({pd.Timestamp(EDITOR_DATES[0]).strftime('%b %Y') if EDITOR_DATES else '-'} "
    f"– {pd.Timestamp(EDITOR_DATES[-1]).strftime('%b %Y') if EDITOR_DATES else '-'})"
)

df_editor_scope = df_f[df_f["Datum"].isin(EDITOR_DATES)].copy()

if df_editor_scope.empty:
    st.info("Keine Daten im sichtbaren Zeitraum.")
else:
    pivot = (df_editor_scope.pivot_table(index=KEY_COLS, columns="Datum",
                                          values="Absatzmenge", aggfunc="sum", fill_value=0)
             .astype(int))
    pivot.columns = [pd.Timestamp(d).strftime("%Y-%m") for d in pivot.columns]
    pivot = pivot.reindex(columns=EDITOR_MONAT_STRS, fill_value=0)

    db_df = load_forecast()
    if not db_df.empty:
        for _, r in db_df.iterrows():
            idx = (r["product_group"], r["item"], r["sales_area"])
            if idx in pivot.index and r["monat"] in pivot.columns:
                pivot.at[idx, r["monat"]] = r["menge"]

    pivot = pivot.reset_index()

    state_key = "forecast_edit"
    reset_key = f"{state_key}_signature"
    signature = (tuple(sel_primary), tuple(sel_main), tuple(sel_pg), tuple(sel_item),
                 tuple(sel_sa), tuple(sel_jahr), sel_monat,
                 tuple(EDITOR_MONAT_STRS),
                 st.session_state["cache_buster"])

    if state_key not in st.session_state or st.session_state.get(reset_key) != signature:
        st.session_state[state_key] = pivot.copy()
        st.session_state[reset_key] = signature

    disabled = (rolle == "Reviewer")

    column_config = {c: st.column_config.TextColumn(c, disabled=True) for c in KEY_COLS}

    edited = st.data_editor(
        st.session_state[state_key],
        num_rows="fixed",
        use_container_width=True,
        height=480,
        disabled=disabled,
        key="editor_main",
        column_config=column_config,
    )
    st.session_state[state_key] = edited

    # Workflow-Aktionen
    st.markdown("##### Aktionen")
    kommentar = st.text_input("Kommentar (optional)", key="kommentar_input")

    if rolle == "Planer":
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("💾 Als Entwurf speichern", use_container_width=True, disabled=not user):
                df_long = edited.melt(id_vars=KEY_COLS, var_name="Monat", value_name="Menge")
                n = save_forecast_with_diff(df_long, ORIGINAL_LOOKUP,
                                            user=user, status="ENTWURF", kommentar=kommentar)
                st.success(f"Als Entwurf gespeichert ({n} Wertänderungen).")
                st.rerun()
        with col_b:
            if st.button("📤 Zur Prüfung einreichen", use_container_width=True, disabled=not user):
                df_long = edited.melt(id_vars=KEY_COLS, var_name="Monat", value_name="Menge")
                n = save_forecast_with_diff(df_long, ORIGINAL_LOOKUP,
                                            user=user, status="ZUR_PRUEFUNG", kommentar=kommentar)
                st.success(f"Zur Prüfung eingereicht ({n} Wertänderungen).")
                st.rerun()
    else:
        if current_status != "ZUR_PRUEFUNG":
            st.info("Aktuell liegt keine Einreichung zur Prüfung vor.")
        else:
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("✅ Freigeben", use_container_width=True, disabled=not user):
                    n = update_status("FREIGEGEBEN", user=user, kommentar=kommentar)
                    st.success(f"{n} Einträge freigegeben.")
                    st.rerun()
            with col_b:
                if st.button("❌ Ablehnen", use_container_width=True, disabled=not user):
                    n = update_status("ABGELEHNT", user=user, kommentar=kommentar)
                    st.warning(f"{n} Einträge abgelehnt.")
                    st.rerun()

    # Excel-Export
    st.markdown("##### Export")
    if current_status == "FREIGEGEBEN":
        pivot_full = (df.groupby(KEY_COLS + ["Datum"], as_index=False)["Absatzmenge"].sum()
                        .pivot_table(index=KEY_COLS, columns="Datum",
                                     values="Absatzmenge", aggfunc="sum", fill_value=0)
                        .astype(int))
        pivot_full.columns = [pd.Timestamp(d).strftime("%Y-%m") for d in pivot_full.columns]

        db_df_full = load_forecast()
        if not db_df_full.empty:
            for _, r in db_df_full.iterrows():
                if r["status"] == "FREIGEGEBEN":
                    idx = (r["product_group"], r["item"], r["sales_area"])
                    if idx in pivot_full.index and r["monat"] in pivot_full.columns:
                        pivot_full.at[idx, r["monat"]] = r["menge"]

        pivot_full = pivot_full.reset_index().sort_values(KEY_COLS)
        excel_bytes = build_styled_excel(pivot_full)
        ts = datetime.now().strftime("%y%m%d_%H_%M")
        st.download_button(
            "⬇️ Freigegebene Prognose als Excel herunterladen",
            data=excel_bytes,
            file_name=f"S&OP_Prognose_Edit_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.caption(f"Excel-Export ist erst nach Freigabe verfügbar. Aktueller Status: **{STATUS_LABEL.get(current_status, 'Keine Daten')}**")

# ============================================================
# Änderungshistorie
# ============================================================
with st.expander("Änderungshistorie"):
    hist = load_change_history()
    if hist.empty:
        st.caption("Noch keine Wertänderungen protokolliert.")
    else:
        hist_view = hist.copy()
        hist_view["Zeitpunkt"]  = pd.to_datetime(hist_view["timestamp"]).dt.strftime("%d.%m.%Y %H:%M")
        hist_view["Bearbeiter"] = hist_view["user"]
        hist_view["Aktion"]     = hist_view["aktion"]
        for db_col, disp_col in zip(["product_group", "item", "sales_area"], KEY_COLS):
            hist_view[disp_col] = hist_view[db_col]
        hist_view["Monat"]      = hist_view["monat"]
        hist_view["Vorher"]     = hist_view["menge_alt"].astype("Int64")
        hist_view["Nachher"]    = hist_view["menge_neu"].astype("Int64")
        hist_view["Δ"]          = hist_view["differenz"].astype("Int64")
        hist_view["Status"]     = hist_view["status"].map(STATUS_LABEL).fillna(hist_view["status"])
        hist_view["Kommentar"]  = hist_view["kommentar"].fillna("")

        cols = ["Zeitpunkt", "Bearbeiter", "Aktion"] + KEY_COLS + [
                "Monat", "Vorher", "Nachher", "Δ", "Status", "Kommentar"]
        st.dataframe(hist_view[cols], use_container_width=True, height=350, hide_index=True)

with st.expander("Rohdaten anzeigen"):
    st.dataframe(df_f, use_container_width=True, height=300)