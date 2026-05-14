"""
Dashboard Planificación & Stock — UnionX
Ejecutar con:  streamlit run dashboard_stock.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import glob, os, re, sqlite3, io
from datetime import datetime

# ─── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Planificación Stock — UnionX",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── GOOGLE DRIVE (definir antes del sidebar que las usa) ─────────────────────
def _drive_secrets():
    """Retorna dict con google_credentials desde st.secrets, o None si no hay."""
    try:
        return dict(st.secrets["google_credentials"])
    except Exception:
        return None


def _drive_service():
    """Crea cliente Drive API usando secrets (cloud) o credentials.json (local)."""
    from drive_utils import get_service
    return get_service(rw=False, secrets=_drive_secrets())


@st.cache_data(ttl=300, show_spinner=False)
def _drive_list_analisis(folder_id: str) -> list:
    """Lista archivos analisis_planificacion_*.xlsx en la carpeta Drive."""
    try:
        from drive_utils import list_folder
        svc = _drive_service()
        files = list_folder(svc, folder_id)
        return [f for f in files if "analisis_planificacion" in f["name"]]
    except Exception as e:
        return []


@st.cache_data(ttl=300, show_spinner=False)
def _drive_bytes(file_id: str) -> bytes:
    """Descarga un archivo de Drive y retorna sus bytes. Cached 5 min."""
    from drive_utils import download_bytes
    svc = _drive_service()
    return download_bytes(svc, file_id)


# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 1.2rem; padding-bottom: 0; }
    h1 { font-size: 1.5rem !important; margin-bottom: 0 !important; }
    h3 { font-size: 1rem !important; font-weight: 600; margin-bottom: 4px; }
    [data-testid="metric-container"] {
        background: #f7f8fa;
        border-radius: 8px;
        padding: 10px 14px;
        border: 1px solid #e5e7eb;
    }
    [data-testid="stMetricValue"] { font-size: 1.6rem !important; }
    footer { visibility: hidden; }
    .legend-pill {
        display:inline-block; padding:2px 8px; border-radius:4px;
        color:white; font-size:11px; margin:2px;
    }
    /* Sidebar nav radio — make it look like a menu */
    [data-testid="stSidebar"] .stRadio > label {
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: .05em;
        text-transform: uppercase;
        color: #888;
        margin-bottom: 4px;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] {
        gap: 2px;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
        background: transparent;
        border-radius: 6px;
        padding: 7px 12px;
        width: 100%;
        cursor: pointer;
        font-size: 0.9rem;
        transition: background 0.15s;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover {
        background: rgba(255,255,255,0.08);
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label[data-checked="true"],
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label[aria-checked="true"] {
        background: rgba(125,60,152,0.18);
        font-weight: 700;
    }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTS ────────────────────────────────────────────────────────────────
BASE    = r"C:\Users\felip\Desktop\UNIONX\FORECAST FINAL SKU\Analisis Planificacion"
DB_PATH    = r"C:\Users\felip\Desktop\UnionX Cloude\planificacion.db"
METAS_PATH = r"C:\Users\felip\Desktop\UNIONX\PPTO 2026\Metas oficiales 1SEM Nuevo.xlsx"
METAS_SHEET = "Cómo Vamos Mayo"
C_MORADO   = "#7D3C98"
C_VERDE    = "#1E8449"
C_AMARILLO = "#D4AC0D"
C_NARANJA  = "#CA6F1E"
C_ROJO     = "#CB4335"
C_AZUL     = "#1A5276"
C_DARK     = "#1C2833"
C_TOTAL1   = "#2C3E50"
C_TOTAL2   = "#0B2341"
C_CRIT_BG  = "#921717"
C_SOB_BG   = "#4A235A"


def pct_color(v):
    """Retorna (bg, fg) suaves para % vs lineal (decimal: 0.85 → 85%)."""
    try:
        v = float(v)
    except Exception:
        return "#f0f0f0", "#888"
    if v >= 1.10: return "#EDE7F6", "#5E35B1"   # pastel morado
    if v >= 0.90: return "#E8F5E9", "#2E7D32"   # pastel verde
    if v >= 0.70: return "#FFF9C4", "#827717"   # pastel amarillo
    if v >= 0.50: return "#FFF3E0", "#BF360C"   # pastel naranja
    return "#FFEBEE", "#B71C1C"                  # pastel rojo


def cob_color(v):
    """Hex color for coverage in months."""
    try:
        v = float(v)
    except Exception:
        return "#999"
    if np.isnan(v): return "#999"
    if v < 1:  return C_ROJO
    if v < 2:  return C_NARANJA
    if v < 4:  return C_VERDE
    if v < 6:  return C_AZUL
    return C_MORADO


def fmt_mm(v, prefix="$", decimals=1):
    """Format number as $XXX.XMM."""
    try:
        f = float(v)
        if np.isnan(f):
            return "—"
        return f"{prefix}{f/1e6:.{decimals}f}MM"
    except Exception:
        return "—"


def fmt_pct(v, already_decimal=True):
    """Format as percentage string."""
    try:
        f = float(v)
        if already_decimal:
            f *= 100
        return f"{f:.1f}%"
    except Exception:
        return "—"


def auto_col_config(df):
    """
    Pre-formatea columnas numéricas de un DataFrame para display con separadores de miles.
    Devuelve una copia con las columnas numéricas convertidas a string formateado.
    Intenta también convertir columnas object que contienen floats.
    """
    df_out = df.copy()

    def _fmt_val(v, fmt):
        """Aplica formato a un valor. Retorna '—' si es nulo o no convertible."""
        try:
            if v is None or (isinstance(v, float) and np.isnan(v)):
                return "—"
            f = float(v)
            if np.isnan(f):
                return "—"
            return fmt(f)
        except Exception:
            return str(v) if v is not None else "—"

    # Columnas que son identificadores y nunca deben formatearse numéricamente
    ID_COLS = {"SKU", "Sku", "sku", "EAN", "Cod", "Código", "Codigo", "PI", "OC"}

    for col in df.columns:
        c = str(col)

        # Saltar columnas identificadoras (SKU, códigos, etc.)
        if c.strip() in ID_COLS or c.strip().upper() == "SKU":
            # Forzar a string para que no aparezca con formato numérico
            df_out[col] = df[col].apply(
                lambda v: str(int(v)) if pd.notna(v) and isinstance(v, (int, float)) and not (isinstance(v, float) and np.isnan(v))
                else (str(v) if pd.notna(v) else "—")
            )
            continue

        series = df[col]

        # Si no es numérico, intentar convertir (puede ser object con floats)
        if not pd.api.types.is_numeric_dtype(series):
            try:
                series = pd.to_numeric(series, errors="coerce")
                if series.isna().all():
                    continue   # columna completamente no numérica → saltar
            except Exception:
                continue

        try:
            # Meses / cobertura → "X.XXm"
            if any(k in c for k in ["Cobert", "(m)", "Meses", "Cob."]):
                df_out[col] = series.apply(lambda v: _fmt_val(v, lambda f: f"{f:.2f}m"))

            # Unidades — evaluar ANTES que moneda para evitar $ en "Stock Unid", "Venta PPTO Unid"
            elif any(k in c for k in ["Unid", "Cantidad", "Unidades", "Ranking"]):
                df_out[col] = series.apply(lambda v: _fmt_val(v, lambda f: f"{int(f):,}"))

            # Moneda CLP → "$ X,XXX"
            elif any(k in c for k in ["($)", "CST", "Venta", "Stock", "Capital", "Inmovil", "Óptimo", "Optimo"]):
                df_out[col] = series.apply(lambda v: _fmt_val(v, lambda f: f"$ {f:,.0f}"))

            # Moneda USD → "USD X,XXX"
            elif any(k in c for k in ["USD", "Valor USD"]):
                df_out[col] = series.apply(lambda v: _fmt_val(v, lambda f: f"USD {f:,.0f}"))

            # Porcentaje → "XX.X%"
            elif any(k in c for k in ["%", "Pct", "vs Meta"]):
                df_out[col] = series.apply(lambda v: _fmt_val(v, lambda f: f"{f*100:.1f}%"))

            # Enteros con separador de miles (Llegadas sin Unid → puede ser CST o unid genérico)
            elif any(k in c for k in ["Llegadas"]):
                df_out[col] = series.apply(lambda v: _fmt_val(v, lambda f: f"{int(f):,}"))

            else:
                # Genérico: si magnitud ≥ 1000 → separador miles; si pequeño → 2 decimales
                max_v = series.dropna().abs().max()
                if pd.isna(max_v):
                    continue
                if max_v >= 1000:
                    df_out[col] = series.apply(lambda v: _fmt_val(v, lambda f: f"{f:,.0f}"))
                elif max_v > 0:
                    df_out[col] = series.apply(lambda v: _fmt_val(v, lambda f: f"{f:.2f}"))

        except Exception:
            pass  # Si algo falla, dejar la columna como estaba

    # Reemplazar None / NaN / "None" / "nan" por "—" en todas las columnas
    for col in df_out.columns:
        df_out[col] = df_out[col].apply(
            lambda v: "—" if (
                v is None
                or (isinstance(v, float) and np.isnan(v))
                or str(v).strip().lower() in ("none", "nan", "nat", "")
            ) else v
        )

    return df_out


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
SECCIONES = [
    "📈  Cómo Vamos",
    "🔴  Stock Crítico",
    "⚠️  Capital Inmovilizado",
    "🚢  Tránsitos",
]

with st.sidebar:
    st.markdown(
        "<div style='padding:6px 0 2px'>"
        "<span style='font-size:1.35rem;font-weight:800;letter-spacing:-.5px'>📊 Análisis de</span><br>"
        "<span style='font-size:1.35rem;font-weight:800;letter-spacing:-.5px'>Planificación</span>"
        "<hr style='margin:10px 0 8px;border-color:#444'>"
        "</div>",
        unsafe_allow_html=True,
    )

    # ── Detectar modo: Drive (cloud) / SQLite / Excel local ──────────────
    _HAS_DRIVE_SECRETS = "dashboard" in st.secrets and "google_credentials" in st.secrets

    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

    if _HAS_DRIVE_SECRETS:
        # ── Modo Google Drive (Streamlit Cloud) ───────────────────────
        USE_SQLITE = False
        _folder_id = st.secrets["dashboard"]["drive_folder_id"]
        _drive_files = _drive_list_analisis(_folder_id)

        if not _drive_files:
            st.error("No se encontró ningún archivo de análisis en Drive.\nEjecuta el script del lunes primero.")
            st.stop()

        # Selector entre los archivos disponibles en Drive
        _file_names = [f["name"] for f in _drive_files]
        _sel = st.selectbox(
            "Análisis:",
            range(len(_file_names)),
            format_func=lambda i: _file_names[i],
            index=0,
        )
        _sel_file = _drive_files[_sel]
        FILE = _sel_file["name"]
        _mtime = _sel_file.get("modifiedTime", "")[:10]
        st.caption(f"☁️ Google Drive · 📅 {_mtime}")

    else:
        # ── Modo local: SQLite o Excel ────────────────────────────────
        db_existe = os.path.exists(DB_PATH)
        modo = st.radio(
            "Fuente de datos",
            ["📡 Tiempo Real", "📋 Análisis Excel"],
            index=0 if db_existe else 1,
            key="modo_fuente",
            label_visibility="visible",
            horizontal=True,
        )
        USE_SQLITE = (modo == "📡 Tiempo Real")

        if USE_SQLITE:
            if not db_existe:
                st.error("planificacion.db no encontrada.\nEjecuta: python fetcher.py --init-db")
                st.stop()
            fresco, lag_str = _db_freshness()
            color_dot = "🟢" if fresco else "🟡"
            st.caption(f"{color_dot} Último cálculo: **{lag_str}**")
            ts_stock = _db_meta("stock_fetched_at", "—")
            ts_trans = _db_meta("transitos_fetched_at", "—")
            st.caption(
                f"Stock: `{ts_stock[11:16] if len(ts_stock) > 10 else ts_stock}`  "
                f"· Tránsitos: `{ts_trans[11:16] if len(ts_trans) > 10 else ts_trans}`"
            )
            FILE = DB_PATH
        else:
            files = sorted(glob.glob(os.path.join(BASE, "analisis_planificacion_*.xlsx")))
            if not files:
                st.error("No se encontró ningún archivo de análisis.\nEjecuta analisis_stock_critico.py primero.")
                st.stop()
            file_names = [os.path.basename(f) for f in files]
            sel_idx = st.selectbox(
                "Archivo:",
                range(len(file_names)),
                format_func=lambda i: file_names[i],
                index=len(files) - 1,
            )
            FILE = files[sel_idx]
            mod_ts = datetime.fromtimestamp(os.path.getmtime(FILE))
            st.caption(f"📅 {mod_ts.strftime('%d/%m/%Y %H:%M')}")

    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

    seccion = st.radio(
        "Sección",
        SECCIONES,
        index=0,
        label_visibility="collapsed",
    )

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        if st.button("🔄 Refrescar", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    with c2:
        if USE_SQLITE:
            if st.button("🔁 Fetch", use_container_width=True,
                         help="Corre un ciclo de fetch + recálculo ahora"):
                with st.spinner("Fetching..."):
                    try:
                        from fetcher import run_once
                        run_once(DB_PATH)
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

    if USE_SQLITE:
        st.caption("Cache 1 min · Sheets cada 10 min")
    else:
        st.caption("Los datos se cachean 5 min.")


# ─── DATA LOADING ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner="Cargando análisis…")
def load_data(source):
    """source: ruta local (str) o bytes descargados desde Drive."""
    src = io.BytesIO(source) if isinstance(source, bytes) else source
    xl = pd.ExcelFile(src)
    sheets = xl.sheet_names
    result = {}

    def read(name, skiprows=1):
        if name not in sheets:
            return None
        df = xl.parse(name, skiprows=skiprows, header=0)
        df = df.dropna(how="all").reset_index(drop=True)
        # Normalize column names: strip newlines
        df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
        return df

    # ── VTA x Marca ───────────────────────────────────────────────────────────
    # The sheet title spans rows 1-2 in Excel; with skiprows=1 the real headers
    # land in data-row 0 as strings. Use skiprows=2 to get them as the actual header.
    vta_name = next((s for s in sheets if "VTA x Marca" in s), None)
    if vta_name:
        df = read(vta_name, skiprows=2)          # skip title row(s)
        if df is not None and df.shape[1] >= 6:
            base_cols = ["Marca", "Meta", "MetaLineal", "VentaAcum", "vsLineal", "PctLineal"]
            if df.shape[1] >= 7:
                df = df.iloc[:, :7].copy()
                df.columns = base_cols + ["PctMeta"]
            else:
                df.columns = base_cols
            df["Marca"] = df["Marca"].astype(str).str.strip()
            # Convert numeric columns (may have residual header strings)
            num_cols = [c for c in df.columns if c != "Marca"]
            for c in num_cols:
                df[c] = pd.to_numeric(df[c], errors="coerce")
            df = df.dropna(subset=["Meta"]).reset_index(drop=True)
            TOTALS = ["TOTAL PROPIA", "P. Nacionales", "TOTAL EMPRESA"]
            result["vta_brands"] = df[~df["Marca"].isin(TOTALS)].copy()
            result["vta_totals"] = df[df["Marca"].isin(TOTALS)].copy()
            row_emp = df[df["Marca"] == "TOTAL EMPRESA"]
            result["vta_empresa"] = row_emp.iloc[0] if len(row_emp) else None

    # ── Cómo Vamos Mayo (canal + marca para lineal 2ª tabla) ──────────────────
    cv_name = next((s for s in sheets if "Vamos" in s or "vamos" in s), None)
    if cv_name:
        result["como_vamos_raw"] = read(cv_name)

    # ── Critico x Marca ───────────────────────────────────────────────────────
    df = read("Critico x Marca")
    if df is not None:
        col_map = {}
        for c in df.columns:
            if "Marca" in c:              col_map[c] = "Marca"
            elif "SKUs" in c and "Sin" not in c and "PPTO" not in c: col_map[c] = "SKUs"
            elif "Cob" in c:              col_map[c] = "CobProm"
            elif "Sin" in c:              col_map[c] = "SinStock"
            elif "Stock Hoy" in c and "CST" not in c: col_map[c] = "StockUnid"
            elif "Venta PPTO" in c:       col_map[c] = "VentaPptoUnid"
            elif "Stock Hoy CST" in c or ("Stock" in c and "CST" in c and "Venta" not in c):
                                          col_map[c] = "StockCST"
            elif "Venta CST" in c:        col_map[c] = "VentaCST"
            elif "Llegadas" in c and "Unid" in c: col_map[c] = "LlegadasUnid"
            elif "Detalle" in c:          col_map[c] = "Detalle"
        df = df.rename(columns=col_map)
        df["Marca"] = df["Marca"].astype(str).str.strip()
        result["critico_marca"] = df[df["Marca"] != "TOTAL"].copy()
        tot_rows = df[df["Marca"] == "TOTAL"]
        result["critico_total"] = tot_rows.iloc[0] if len(tot_rows) else None

    # ── Detalle Critico ───────────────────────────────────────────────────────
    df = read("Detalle Critico")
    if df is not None:
        result["detalle_critico"] = df.copy()

    # ── Capital Inmovilizado ──────────────────────────────────────────────────
    df = read("Capital Inmovilizado")
    if df is not None:
        first_col = df.columns[0]
        df = df.rename(columns={first_col: "Entidad"})
        df["Entidad"] = df["Entidad"].astype(str)
        df["Nivel"] = df["Entidad"].apply(
            lambda x: "Marca"    if "▶" in x else
                      "CatPadre" if "▸" in x else
                      "CatHijo"  if "▹" in x else
                      "SKU"      if "↳" in x else
                      "TOTAL"    if x.strip() == "TOTAL" else "Other"
        )
        df["Nombre"] = df["Entidad"].str.replace(r"[▶▸▹↳]", "", regex=True).str.strip()
        result["capital_full"]  = df.copy()
        result["capital_marca"] = df[df["Nivel"] == "Marca"].copy()
        tot_rows = df[df["Nivel"] == "TOTAL"]
        result["capital_total"] = tot_rows.iloc[0] if len(tot_rows) else None

    # ── Sobrestock x Marca (resumen) ─────────────────────────────────────────
    df = read("Sobrestock x Marca")
    if df is not None:
        first_col = df.columns[0]
        df = df.rename(columns={first_col: "Marca"})
        df["Marca"] = df["Marca"].astype(str).str.strip()
        result["sob_marca"]  = df[df["Marca"] != "TOTAL"].copy()
        tot_rows = df[df["Marca"] == "TOTAL"]
        result["sob_total"]  = tot_rows.iloc[0] if len(tot_rows) else None

    # ── Sobrestock x Marca y Cat ──────────────────────────────────────────────
    df = read("Sobrestock x Marca y Cat")
    if df is not None:
        result["sob_marca_cat"] = df.copy()

    # ── Sobrestock c-Llegada (detalle SKU con llegadas encima) ────────────────
    df = read("Sobrestock c-Llegada")
    if df is not None:
        result["sobrestock"] = df.copy()

    # ── Tránsitos por Embarque ────────────────────────────────────────────────
    df = read("Tránsitos por Embarque")
    if df is not None:
        pi_col   = df.columns[0]
        marc_col = next((c for c in df.columns if "Marcas" in c), None)
        if marc_col:
            df_pi   = df[df[marc_col].notna() & (df[pi_col].astype(str) != "TOTAL")].copy()
            tot_r   = df[df[pi_col].astype(str) == "TOTAL"]
            result["transitos_total"] = tot_r.iloc[0] if len(tot_r) else None
        else:
            df_pi = df.copy()
        result["transitos"] = df_pi
        # Full sheet (PI rows + SKU detail rows), excluding TOTAL
        result["transitos_full"] = df[df[pi_col].astype(str) != "TOTAL"].reset_index(drop=True)

    # ── Nuevos en Tránsito ────────────────────────────────────────────────────
    df = read("Nuevos en Tránsito")
    if df is not None:
        mes_col = next((c for c in df.columns if "Mes" in c), None)
        sku_col = df.columns[0]
        if mes_col:
            df = df[df[mes_col].notna() & (df[sku_col].astype(str) != "TOTAL")].copy()
        result["nuevos_transito"] = df.reset_index(drop=True)

    return result


# ─── SQLITE LOADER ────────────────────────────────────────────────────────────
@st.cache_data(ttl=60, show_spinner="Cargando datos en tiempo real…")
def load_data_sqlite(db_path):
    """
    Lee desde planificacion.db (generado por fetcher.py).
    Devuelve el mismo dict que load_data() para que el dashboard funcione sin cambios.
    TTL=60s → el dashboard se refresca automáticamente cada minuto.
    """
    if not os.path.exists(db_path):
        return {"_error": "DB no encontrada. Ejecuta: python fetcher.py --init-db"}

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    result = {"source": "sqlite"}

    # ── Meta (timestamps de cada fuente) ──────────────────────────────
    meta_rows = conn.execute("SELECT key, value FROM meta").fetchall()
    result["meta"] = {r[0]: r[1] for r in meta_rows}

    # ── Verificar que hay datos calculados ────────────────────────────
    n_rt = conn.execute("SELECT COUNT(*) FROM analisis_rt").fetchone()[0]
    if n_rt == 0:
        conn.close()
        result["_warn"] = "analisis_rt vacío. Ejecuta: python fetcher.py"
        return result

    rt = pd.read_sql("SELECT * FROM analisis_rt", conn)
    fc = pd.read_sql("SELECT sku, venta_unid_m0, venta_unid_m1, venta_unid_m2 FROM forecast_sku", conn)

    _cur_m   = datetime.now().month
    _mes_abbr = {4:'ABR',5:'MAY',6:'JUN',7:'JUL',8:'AGO',9:'SEP',10:'OCT'}
    _mes_str  = _mes_abbr.get(_cur_m, 'MES')

    # ── Críticos ──────────────────────────────────────────────────────
    criticos = rt[rt["es_critico"] == 1].copy()
    if len(criticos):
        rt_fc = criticos.merge(fc, on="sku", how="left")
        grp = (rt_fc.groupby("marca", as_index=False).agg(
            SKUs          = ("sku",           "count"),
            CobProm       = ("cobert_actual",  "mean"),
            SinStock      = ("stock_unid",     lambda x: int((pd.to_numeric(x, errors="coerce").fillna(0) <= 0).sum())),
            StockUnid     = ("stock_unid",     "sum"),
            VentaPptoUnid = ("venta_unid_m0",  "sum"),
            StockCST      = ("stock_cst",      "sum"),
            VentaCST      = ("venta_cst_m0",   "sum"),
            LlegadasUnid  = ("qty_transito",   "sum"),
        ).rename(columns={"marca": "Marca"}))
        grp = grp.sort_values("StockCST", ascending=False)
        result["critico_marca"] = grp

        tot = grp[["SKUs","StockUnid","StockCST"]].sum()
        tot.name = "TOTAL"
        result["critico_total"] = tot

        # Detalle crítico
        det = criticos.merge(fc, on="sku", how="left").copy()
        det = det.rename(columns={
            "marca":         "Marca",
            "cat_com":       "Cat. Comercial",
            "cat_padre":     "Cat. Padre",
            "cat_hijo":      "Cat. Hijo",
            "ranking":       "Ranking Comercial",
            "sku":           "SKU",
            "descripcion":   "Descripcion",
            "cobert_actual": f"Cobert. ACT (m)",
            "stock_unid":    "Stock Unid",
            "venta_unid_m0": f"Venta PPTO {_mes_str} Unid",
            "stock_cst":     "Stock CST ($)",
            "venta_cst_m0":  f"Venta CST {_mes_str} ($)",
            "qty_transito":  f"Llegadas {_mes_str} Unid",
            "prox_eta":      "ETA Bodega",
            "prox_pi":       "PI Embarque",
            "prox_mes":      "Mes Llegada",
        })
        det = det.sort_values(["Marca", f"Cobert. ACT (m)"], na_position="last")
        result["detalle_critico"] = det.reset_index(drop=True)

    # ── Sobrestock ────────────────────────────────────────────────────
    sobre = rt[rt["es_sobrestock"] == 1].copy()
    if len(sobre):
        # Resumen por marca
        grp_s = (sobre.groupby("marca", as_index=False).agg(
            SKUs       = ("sku",           "count"),
            StockCST   = ("stock_cst",     "sum"),
            CapExceso  = ("capital_exceso","sum"),
        ).rename(columns={"marca": "Marca"})
          .sort_values("StockCST", ascending=False))
        result["sob_marca"] = grp_s
        tot_s = grp_s[["SKUs","StockCST","CapExceso"]].sum(); tot_s.name = "TOTAL"
        result["sob_total"] = tot_s

        # Detalle sobrestock con llegadas (c-Llegada)
        sob_ll = sobre[sobre["qty_transito"] > 0].copy()
        if len(sob_ll):
            sob_ll = sob_ll.rename(columns={
                "marca":         "Marca",
                "cat_com":       "Cat. Comercial",
                "sku":           "SKU",
                "descripcion":   "Descripción",
                "cobert_actual": f"Cobert. ACT (m)",
                "stock_cst":     "Stock CST ($)",
                "venta_cst_m0":  f"Venta Prom. {_mes_str} ($)",
                "qty_transito":  "Llegadas Unid",
                "prox_eta":      "ETA Próx. Llegada",
            })
            result["sobrestock"] = sob_ll.reset_index(drop=True)

        # Capital Inmovilizado jerárquico (Marca → CatPadre → SKU)
        rows_cap = []
        total_stock = sobre["stock_cst"].sum()
        total_cap   = sobre["capital_exceso"].sum()
        total_skus  = len(sobre)
        rows_cap.append({
            "Entidad": "TOTAL", "Nivel": "TOTAL", "Nombre": "TOTAL",
            "SKUs": total_skus,
            "Stock CST ($)": total_stock,
            "Capital Inmovil. ($)": total_cap,
            "Cobert. ACT (m)": sobre["cobert_actual"].mean(),
        })
        for marca, gm in sobre.groupby("marca"):
            rows_cap.append({
                "Entidad": f"▶  {marca}", "Nivel": "Marca", "Nombre": marca,
                "SKUs": len(gm),
                "Stock CST ($)": gm["stock_cst"].sum(),
                "Capital Inmovil. ($)": gm["capital_exceso"].sum(),
                "Cobert. ACT (m)": gm["cobert_actual"].mean(),
            })
            for cat_p, gcp in gm.groupby("cat_padre"):
                rows_cap.append({
                    "Entidad": f"   ▸  {cat_p}", "Nivel": "CatPadre", "Nombre": cat_p,
                    "SKUs": len(gcp),
                    "Stock CST ($)": gcp["stock_cst"].sum(),
                    "Capital Inmovil. ($)": gcp["capital_exceso"].sum(),
                    "Cobert. ACT (m)": gcp["cobert_actual"].mean(),
                })
                for cat_h, gch in gcp.groupby("cat_hijo"):
                    rows_cap.append({
                        "Entidad": f"         ▹  {cat_h}", "Nivel": "CatHijo", "Nombre": cat_h,
                        "SKUs": len(gch),
                        "Stock CST ($)": gch["stock_cst"].sum(),
                        "Capital Inmovil. ($)": gch["capital_exceso"].sum(),
                        "Cobert. ACT (m)": gch["cobert_actual"].mean(),
                    })
                    for _, sk in gch.iterrows():
                        rows_cap.append({
                            "Entidad": f"               ↳  {sk['sku']}", "Nivel": "SKU", "Nombre": sk["sku"],
                            "SKUs": 1,
                            "Stock CST ($)": sk["stock_cst"],
                            "Capital Inmovil. ($)": sk["capital_exceso"],
                            "Cobert. ACT (m)": sk["cobert_actual"],
                        })
        cap_df = pd.DataFrame(rows_cap)
        result["capital_full"]  = cap_df
        result["capital_marca"] = cap_df[cap_df["Nivel"] == "Marca"].copy()
        tot_cap = cap_df[cap_df["Nivel"] == "TOTAL"]
        result["capital_total"] = tot_cap.iloc[0] if len(tot_cap) else None

    # ── Tránsitos ─────────────────────────────────────────────────────
    trans = pd.read_sql("SELECT * FROM transitos", conn)
    if len(trans):
        trans_full = trans.rename(columns={
            "pi":         "PI",
            "sku":        "SKU",
            "descripcion":"Descripción",
            "cantidad":   "Unidades",
            "eta_bodega": "ETA Bodega",
            "mes":        "Mes",
            "marca":      "Marcas",
            "valor_usd":  "Valor USD",
            "tipo_cat":   "Tipo Cat.",
        })
        result["transitos_full"] = trans_full
        pi_grp = (trans.groupby("pi", as_index=False).agg(
            Marcas   = ("marca",     lambda x: " · ".join(sorted(set(str(v) for v in x if v)))),
            SKUs     = ("sku",       "nunique"),
            Unidades = ("cantidad",  "sum"),
            ValorUSD = ("valor_usd", "sum"),
            ETA      = ("eta_bodega","min"),
            Mes      = ("mes",       "min"),
        ).rename(columns={"pi": "PI"}))
        result["transitos"] = pi_grp

        nuevos = trans[trans["tipo_cat"].str.contains("nuevo", case=False, na=False)].copy()
        if len(nuevos):
            nuevos = nuevos.rename(columns={
                "sku": "SKU", "descripcion": "Descripción",
                "marca": "Marca", "mes": "Mes Llegada",
                "eta_bodega": "Fecha ETA Bodega", "cantidad": "Cantidad",
            })
            result["nuevos_transito"] = nuevos.reset_index(drop=True)

    # ── Ventas acumuladas ─────────────────────────────────────────────
    vta = pd.read_sql(
        "SELECT marca, mes, venta_cst FROM vta_acum_marca WHERE ano=? AND mes=?",
        conn, params=(datetime.now().year, _cur_m)
    )
    if len(vta):
        result["vta_acum_mes"] = vta  # usado para Cómo Vamos si está disponible

    conn.close()
    return result


# ── Helpers para obtener timestamps de la DB ──────────────────────────────────
def _db_meta(key: str, default="—") -> str:
    if not os.path.exists(DB_PATH):
        return default
    try:
        conn = sqlite3.connect(DB_PATH)
        r = conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
        conn.close()
        return r[0] if r else default
    except Exception:
        return default


def _db_freshness() -> tuple[bool, str]:
    """Devuelve (es_fresco, descripcion). Fresco = último cálculo hace < 20 min."""
    ts_str = _db_meta("last_computed_at")
    if ts_str == "—":
        return False, "Sin datos"
    try:
        ts  = datetime.fromisoformat(ts_str)
        lag = (datetime.now() - ts).total_seconds() / 60
        if lag < 5:    return True,  f"hace {lag:.0f}m"
        if lag < 20:   return True,  f"hace {lag:.0f}m"
        if lag < 60:   return False, f"hace {lag:.0f}m ⚠️"
        return False, f"hace {lag/60:.1f}h ❌"
    except Exception:
        return False, ts_str


@st.cache_data(ttl=300, show_spinner=False)
def load_metas(source=None, sheet=METAS_SHEET):
    """
    Lee 'Cómo Vamos Mayo' de Metas oficiales.
    source: bytes (desde Drive) | None (usa ruta local METAS_PATH)
    Devuelve dict con 4 DataFrames: vn_marca, cf_marca, vn_canal, cf_canal
    + sus respectivas filas de total (vn_marca_tot, etc.).
    Columnas de cada df: Entidad | Meta | MetaLineal | Real | vsLineal | PctLineal | PctMeta
    """
    if source is None:
        if not os.path.exists(METAS_PATH):
            return None
        src = METAS_PATH
    else:
        src = io.BytesIO(source)
    try:
        df_raw = pd.read_excel(src, sheet_name=sheet, header=None, dtype=object)
    except Exception:
        return None

    COLS = ["Entidad", "Meta", "MetaLineal", "Real", "vsLineal", "PctLineal", "PctMeta"]
    TOTAL_LABELS = {"total", "total propia", "total empresa",
                    "total canal", "total marcas"}

    def _parse_block(search_from, search_to):
        """Busca la fila de cabecera (col 0 = MARCA/CANAL) y extrae datos hasta blank."""
        header_row = None
        for i in range(search_from, min(search_to, len(df_raw))):
            v = str(df_raw.iat[i, 0]).strip().upper()
            if v in ("MARCA", "CANAL"):
                header_row = i
                break
        if header_row is None:
            return pd.DataFrame(columns=COLS), None

        rows = []
        for i in range(header_row + 1, min(len(df_raw), header_row + 25)):
            name = str(df_raw.iat[i, 0]).strip()
            if not name or name.lower() in ("nan", "none", ""):
                break
            vals = []
            for j in range(1, 7):
                try:
                    vals.append(float(df_raw.iat[i, j]))
                except Exception:
                    vals.append(np.nan)
            rows.append([name] + vals)

        if not rows:
            return pd.DataFrame(columns=COLS), None

        df = pd.DataFrame(rows, columns=COLS)
        is_tot = df["Entidad"].str.strip().str.lower().isin(TOTAL_LABELS)
        df_data = df[~is_tot].reset_index(drop=True)
        df_tot  = df[is_tot]
        return df_data, (df_tot.iloc[0] if len(df_tot) else None)

    # Posiciones de búsqueda (0-indexed, holgura ±3 filas)
    vn_marca_df,  vn_marca_tot  = _parse_block(1,  15)
    cf_marca_df,  cf_marca_tot  = _parse_block(13, 30)
    vn_canal_df,  vn_canal_tot  = _parse_block(26, 42)
    cf_canal_df,  cf_canal_tot  = _parse_block(38, 55)

    return {
        "vn_marca":    vn_marca_df,  "vn_marca_tot":  vn_marca_tot,
        "cf_marca":    cf_marca_df,  "cf_marca_tot":  cf_marca_tot,
        "vn_canal":    vn_canal_df,  "vn_canal_tot":  vn_canal_tot,
        "cf_canal":    cf_canal_df,  "cf_canal_tot":  cf_canal_tot,
    }


# ─── LOAD DATA ────────────────────────────────────────────────────────────────
if _HAS_DRIVE_SECRETS:
    _raw = _drive_bytes(_sel_file["id"])
    data = load_data(_raw)
elif USE_SQLITE:
    data = load_data_sqlite(DB_PATH)
else:
    data = load_data(FILE)

# ─── METAS desde Drive (si está en cloud) ─────────────────────────────────────
_METAS_FILE_ID = st.secrets.get("dashboard", {}).get("metas_file_id", "") if _HAS_DRIVE_SECRETS else ""

# ─── HEADER ───────────────────────────────────────────────────────────────────
if USE_SQLITE:
    _ts_c = data.get("meta", {}).get("last_computed_at", "—") if isinstance(data, dict) else "—"
    _period_label = f"Tiempo Real  `{_ts_c[11:16] if len(_ts_c) > 10 else _ts_c}`"
else:
    _pm = re.search(r"analisis_planificacion_(\w+)\.xlsx", os.path.basename(FILE))
    _period_label = _pm.group(1) if _pm else "—"
st.markdown(f"# 📦 Dashboard Planificación — {_period_label}")

# ─── KPIs ─────────────────────────────────────────────────────────────────────
crit_marcas   = data.get("critico_marca")
crit_tot      = data.get("critico_total")
cap_tot       = data.get("capital_total")
emp_row       = data.get("vta_empresa")
trans_df      = data.get("transitos")
trans_tot     = data.get("transitos_total")
nuevos_df     = data.get("nuevos_transito")

# Críticos
n_crit      = int(crit_tot["SKUs"]) if crit_tot is not None else 0
n_sin_stock = int(crit_marcas["SinStock"].fillna(0).sum()) if crit_marcas is not None and "SinStock" in crit_marcas.columns else 0

# Capital inmovilizado
def _find_val(row, keywords):
    if row is None: return 0.0
    for k in row.index:
        if any(kw in str(k) for kw in keywords):
            try: return float(row[k])
            except: pass
    return 0.0

cap_inmovil    = _find_val(cap_tot, ["Inmovil", "Capital"])
stock_sob_tot  = _find_val(cap_tot, ["Stock CST"])
cap_skus       = int(_find_val(cap_tot, ["SKUs"]))

# Ventas neta + contribución — leer desde metas (vn_canal_tot / cf_canal_tot)
_metas_kpi = load_metas(_drive_bytes(_METAS_FILE_ID) if _METAS_FILE_ID else None)
_vn_tot    = _metas_kpi["vn_canal_tot"]  if _metas_kpi else None
_cf_tot    = _metas_kpi["cf_canal_tot"]  if _metas_kpi else None

if _vn_tot is not None:
    try:
        pct_mayo   = float(_vn_tot["PctLineal"])
        venta_mayo = float(_vn_tot["Real"])
        meta_mayo  = float(_vn_tot["Meta"])
    except Exception:
        pct_mayo = venta_mayo = meta_mayo = 0.0
elif emp_row is not None:
    pct_mayo   = float(emp_row["PctLineal"])
    venta_mayo = float(emp_row["VentaAcum"])
    meta_mayo  = float(emp_row["Meta"])
else:
    pct_mayo = venta_mayo = meta_mayo = 0.0

if _cf_tot is not None:
    try:
        pct_cf    = float(_cf_tot["PctLineal"])
        venta_cf  = float(_cf_tot["Real"])
        meta_cf   = float(_cf_tot["Meta"])
    except Exception:
        pct_cf = venta_cf = meta_cf = 0.0
else:
    pct_cf = venta_cf = meta_cf = 0.0

# Tránsitos
n_emb = len(trans_df) if trans_df is not None else 0
usd_col_t = next((c for c in (trans_df.columns if trans_df is not None else [])), None)
if trans_tot is not None:
    usd_col_t = next((c for c in trans_tot.index if "USD" in c or "Valor" in c), None)
    usd_val   = float(trans_tot[usd_col_t]) if usd_col_t and not pd.isna(trans_tot[usd_col_t]) else 0
else:
    usd_val = 0
n_nuevos = len(nuevos_df) if nuevos_df is not None else 0

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric(f"📈 Venta Neta {_period_label}", fmt_mm(venta_mayo),
          delta=f"{pct_mayo*100:.1f}% vs lineal",
          delta_color="normal" if pct_mayo >= 1 else "inverse",
          help=f"Venta neta total empresa acum. | Meta: {fmt_mm(meta_mayo)}")
k2.metric(f"📊 Contrib. Front {_period_label}", fmt_mm(venta_cf),
          delta=f"{pct_cf*100:.1f}% vs lineal",
          delta_color="normal" if pct_cf >= 1 else "inverse",
          help=f"Contribución frontal total empresa acum. | Meta: {fmt_mm(meta_cf)}")
k3.metric("⚠️ Capital Inmovilizado", fmt_mm(cap_inmovil),
          delta=f"{cap_skus} SKUs",       delta_color="off",
          help="Exceso sobre cobertura óptima de 4 meses")
k4.metric("🔴 SKUs Críticos",        f"{n_crit}",
          delta=f"{n_sin_stock} sin stock", delta_color="inverse",
          help="Cobertura < 1 mes")
k5.metric("🚢 Embarques en Tránsito", f"{n_emb}",
          delta=f"USD ${usd_val:,.0f}", delta_color="off",
          help="Embarques activos con ETA vigente")
k6.metric("🆕 Nuevos SKUs en Tránsito", f"{n_nuevos}",
          help="SKUs que ingresan por primera vez")

st.divider()

# ─── SECCIÓN ACTIVA ───────────────────────────────────────────────────────────

# ── Helper: tabla HTML reutilizable para las 3 vistas de Cómo Vamos ──────────
def _render_cv_table(df_data, tot_row, dim_col="Marca",
                     real_col="Real", lbl_meta="Meta Mes",
                     lbl_real="Real Acum.", show_pct_meta=True):
    """
    Renderiza la tabla de Cómo Vamos en HTML.
    df_data / tot_row usan columnas: Entidad|Meta|MetaLineal|Real|vsLineal|PctLineal[|PctMeta]
    """
    if df_data is None or len(df_data) == 0:
        st.info("Sin datos.")
        return

    df_s = df_data.dropna(subset=["Meta"]).sort_values("PctLineal", ascending=False).copy()

    has_pm = show_pct_meta and "PctMeta" in df_data.columns

    def _row(r, is_total=False, bg_override=None):
        bg  = bg_override or ("white" if not is_total else C_TOTAL1)
        fg  = "white" if is_total else "#222"
        fw  = "800" if is_total else "700"
        pad = "7px 10px" if is_total else "6px 10px"

        name = r.get("Entidad", r.get("Marca", "—"))
        meta = fmt_mm(r.get("Meta", np.nan))
        ml   = fmt_mm(r.get("MetaLineal", np.nan))
        real = fmt_mm(r.get(real_col, r.get("Real", r.get("VentaAcum", np.nan))))
        vs   = r.get("vsLineal", np.nan)
        try:
            vs_f = float(vs)
            vs_s = (f'+{fmt_mm(vs_f)}' if vs_f >= 0 else fmt_mm(vs_f))
            vs_fg = C_VERDE if vs_f >= 0 else C_ROJO
        except Exception:
            vs_s = "—"; vs_fg = "#999"

        pc  = r.get("PctLineal", np.nan)
        try:    pc_f = float(pc)
        except: pc_f = np.nan
        pcbg, pcfg = pct_color(pc_f) if not np.isnan(pc_f) else ("#f0f0f0", "#888")
        pc_s = fmt_pct(pc_f) if not np.isnan(pc_f) else "—"

        vs_color = vs_fg if is_total else vs_fg

        cells = (
            f'<td style="background:{bg};color:{fg};padding:{pad};font-weight:{fw};white-space:nowrap;font-size:13px">{name}</td>'
            f'<td style="background:{bg};color:{fg};padding:{pad};text-align:right;font-size:13px">{meta}</td>'
            f'<td style="background:{bg};color:{fg};padding:{pad};text-align:right;font-size:13px">{ml}</td>'
            f'<td style="background:{bg};color:{fg};padding:{pad};text-align:right;font-weight:700;font-size:13px">{real}</td>'
            f'<td style="background:{bg};color:{vs_color};padding:{pad};text-align:right;font-weight:700;font-size:13px">{vs_s}</td>'
            f'<td style="background:{pcbg};color:{pcfg};padding:{pad};text-align:center;font-weight:600;font-size:13px;border-radius:4px">{pc_s}</td>'
        )
        if has_pm:
            pm = r.get("PctMeta", np.nan)
            try:    pm_f = float(pm); pm_s = fmt_pct(pm_f)
            except: pm_s = "—"
            cells += f'<td style="background:{bg};color:{fg};padding:{pad};text-align:center;font-size:13px">{pm_s}</td>'
        return f'<tr>{cells}</tr>'

    rows_html = ""
    for i, (_, r) in enumerate(df_s.iterrows()):
        bg = "white" if i % 2 == 0 else "#F8F9FA"
        rows_html += _row(r, bg_override=bg)

    if tot_row is not None:
        # tot_row puede ser Series o dict-like
        try:
            tr = tot_row.to_dict() if hasattr(tot_row, "to_dict") else dict(tot_row)
        except Exception:
            tr = {}
        rows_html += _row(tr, is_total=True, bg_override=C_TOTAL2)

    hdrs = [dim_col, lbl_meta, "Meta Lineal", lbl_real, "vs Lineal $", "% vs Lineal"]
    if has_pm:
        hdrs.append("% vs Meta")

    html_tbl = (
        '<table style="border-collapse:collapse;width:100%;font-family:Arial">'
        '<thead><tr>'
        + "".join(
            f'<th style="background:{C_DARK};color:white;padding:8px 10px;'
            f'text-align:{"left" if i==0 else "center" if "%" in h else "right"};'
            f'white-space:nowrap;font-size:12px">{h}</th>'
            for i, h in enumerate(hdrs)
        )
        + "</tr></thead><tbody>"
        + rows_html
        + "</tbody></table>"
    )
    st.markdown(html_tbl, unsafe_allow_html=True)

_LEGEND_HTML = " ".join([
    f'<span class="legend-pill" style="background:{bg};color:{fg}">{lbl}</span>'
    for (bg, fg), lbl in [
        (("#EDE7F6","#5E35B1"), "≥110%"),
        (("#E8F5E9","#2E7D32"), "90–110%"),
        (("#FFF9C4","#827717"), "70–90%"),
        (("#FFF3E0","#BF360C"), "50–70%"),
        (("#FFEBEE","#B71C1C"), "<50%"),
    ]
])

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 1 — CÓMO VAMOS
# ══════════════════════════════════════════════════════════════════════════════
if seccion == SECCIONES[0]:

    tab_vn, tab_mg, tab_cst = st.tabs([
        "📊  Venta Neta",
        "📈  Contribución Front",
        "💰  A Costo",
    ])

    # ── Tab 1: Venta Neta ─────────────────────────────────────────────────────
    with tab_vn:
        metas = load_metas(_drive_bytes(_METAS_FILE_ID) if _METAS_FILE_ID else None)
        if metas is None:
            st.warning(f"Archivo no encontrado:\n`{METAS_PATH}`")
        elif metas["vn_canal"] is None or len(metas["vn_canal"]) == 0:
            st.info("No se encontraron datos de Venta Neta.")
        else:
            st.markdown("##### Por Canal")
            _render_cv_table(metas["vn_canal"], metas["vn_canal_tot"],
                             dim_col="Canal", lbl_meta="Meta Mes", lbl_real="Real Acum.")
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("##### Por Marca")
            _render_cv_table(metas["vn_marca"], metas["vn_marca_tot"],
                             dim_col="Marca", lbl_meta="Meta Mes", lbl_real="Real Acum.")
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(_LEGEND_HTML, unsafe_allow_html=True)

    # ── Tab 2: Contribución Front ─────────────────────────────────────────────
    with tab_mg:
        metas = load_metas(_drive_bytes(_METAS_FILE_ID) if _METAS_FILE_ID else None)
        if metas is None:
            st.warning(f"Archivo no encontrado:\n`{METAS_PATH}`")
        elif metas["cf_canal"] is None or len(metas["cf_canal"]) == 0:
            st.info("No se encontraron datos de Contribución Front.")
        else:
            st.markdown("##### Por Canal")
            _render_cv_table(metas["cf_canal"], metas["cf_canal_tot"],
                             dim_col="Canal", lbl_meta="Meta Mes", lbl_real="Real Acum.")
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("##### Por Marca")
            _render_cv_table(metas["cf_marca"], metas["cf_marca_tot"],
                             dim_col="Marca", lbl_meta="Meta Mes", lbl_real="Real Acum.")
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(_LEGEND_HTML, unsafe_allow_html=True)

    # ── Tab 3: A Costo ────────────────────────────────────────────────────────
    with tab_cst:
        vta     = data.get("vta_brands")
        vta_tot = data.get("vta_totals")
        if vta is not None and len(vta):
            df_cst = vta.rename(columns={"Marca": "Entidad", "VentaAcum": "Real"}).copy()
            # Calcular % vs Meta si no existe o está vacío
            if "PctMeta" not in df_cst.columns or df_cst["PctMeta"].isna().all():
                df_cst["PctMeta"] = df_cst["Real"] / df_cst["Meta"].replace(0, np.nan)
            tot_cst = None
            if vta_tot is not None and len(vta_tot):
                tot_s = vta_tot[vta_tot["Marca"] == "TOTAL EMPRESA"]
                if len(tot_s) == 0:
                    tot_s = vta_tot.tail(1)
                tot_cst = tot_s.iloc[0].rename({"Marca": "Entidad", "VentaAcum": "Real"}).copy()
                if pd.isna(tot_cst.get("PctMeta", np.nan)):
                    try:
                        tot_cst["PctMeta"] = float(tot_cst["Real"]) / float(tot_cst["Meta"])
                    except Exception:
                        pass
            _render_cv_table(df_cst, tot_cst, dim_col="Marca",
                             lbl_meta="Meta Mes", lbl_real="Venta Acum.", show_pct_meta=True)
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(_LEGEND_HTML, unsafe_allow_html=True)
        else:
            st.info("No se encontraron datos de venta a costo.")

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 2 — STOCK CRÍTICO (resumen + detalle unificado)
# ══════════════════════════════════════════════════════════════════════════════
elif seccion == SECCIONES[1]:
    crit = data.get("critico_marca")
    det  = data.get("detalle_critico")

    if crit is not None and len(crit):
        crit_s = crit.copy()
        if "SKUs" in crit_s.columns:
            crit_s = crit_s.sort_values("SKUs", ascending=False)
        crit_s["SinStock"] = crit_s.get("SinStock", pd.Series(dtype=float)).fillna(0)

        # ── Tabla resumen por marca ───────────────────────────────────────────
        rows_h = ""
        for i, (_, r) in enumerate(crit_s.iterrows()):
            bg     = "white" if i % 2 == 0 else "#F8F9FA"
            cob    = r.get("CobProm", np.nan)
            cob_bg = cob_color(cob)
            ss     = int(r.get("SinStock", 0) or 0)
            ss_bg  = "#FADBD8" if ss > 0 else bg
            ss_fg  = C_ROJO    if ss > 0 else "#222"
            det_str= str(r.get("Detalle","")) if not pd.isna(r.get("Detalle", np.nan)) else "—"
            con_s  = int(r["SKUs"]) - ss
            rows_h += (
                f"<tr>"
                f'<td style="background:{bg};padding:6px 10px;font-weight:700;white-space:nowrap;font-size:13px">{r["Marca"]}</td>'
                f'<td style="background:{bg};padding:6px 10px;text-align:center;font-weight:700;font-size:13px">{int(r["SKUs"])}</td>'
                f'<td style="background:{ss_bg};color:{ss_fg};padding:6px 10px;text-align:center;font-weight:700;font-size:13px">{ss}</td>'
                f'<td style="background:{bg};padding:6px 10px;text-align:center;font-size:13px">{con_s}</td>'
                f'<td style="background:{cob_bg};color:white;padding:6px 10px;text-align:center;font-weight:700;font-size:13px">{f"{cob:.2f}m" if not pd.isna(cob) else "—"}</td>'
                f'<td style="background:{bg};padding:6px 10px;text-align:right;font-size:13px">{fmt_mm(r.get("StockCST",0))}</td>'
                f'<td style="background:{bg};padding:6px 10px;text-align:right;font-size:13px">{fmt_mm(r.get("VentaCST",0))}</td>'
                f'<td style="background:{bg};padding:6px 10px;font-size:12px">{det_str}</td>'
                f"</tr>"
            )
        if crit_tot is not None:
            rows_h += (
                f'<tr>'
                f'<td style="background:{C_CRIT_BG};color:white;padding:7px 10px;font-weight:800;font-size:13px">TOTAL</td>'
                f'<td style="background:{C_CRIT_BG};color:white;padding:7px 10px;text-align:center;font-weight:800;font-size:13px">{int(crit_tot["SKUs"])}</td>'
                f'<td style="background:{C_CRIT_BG};color:white;padding:7px 10px;text-align:center;font-weight:800;font-size:13px">{n_sin_stock}</td>'
                f'<td style="background:{C_CRIT_BG};color:white;padding:7px 10px;text-align:center;font-size:13px">{int(crit_tot["SKUs"])-n_sin_stock}</td>'
                f'<td style="background:{C_CRIT_BG};color:white;padding:7px 10px;text-align:center;font-size:13px">—</td>'
                f'<td style="background:{C_CRIT_BG};color:white;padding:7px 10px;text-align:right;font-weight:700;font-size:13px">{fmt_mm(crit_tot.get("StockCST",0))}</td>'
                f'<td style="background:{C_CRIT_BG};color:white;padding:7px 10px;text-align:right;font-weight:700;font-size:13px">{fmt_mm(crit_tot.get("VentaCST",0))}</td>'
                f'<td style="background:{C_CRIT_BG};color:white;padding:7px 10px"></td>'
                f'</tr>'
            )
        tbl_h = (
            '<table style="border-collapse:collapse;width:100%;font-family:Arial">'
            "<thead><tr>"
            + "".join(
                f'<th style="background:{C_CRIT_BG};color:white;padding:8px 10px;'
                f'text-align:{"left" if h=="Marca" else "right" if h in ["Stock CST","Venta CST"] else "center"}'
                f';white-space:nowrap;font-size:12px">{h}</th>'
                for h in ["Marca","Total SKUs","Sin Stock","Con Stock","Cob. Prom.","Stock CST","Venta CST","Próx. Llegadas"]
            )
            + "</tr></thead><tbody>" + rows_h + "</tbody></table>"
        )
        st.markdown(tbl_h, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            " ".join([
                f'<span class="legend-pill" style="background:{c}">Cob {lbl}</span>'
                for c, lbl in [(C_ROJO,"<1m"),(C_NARANJA,"1-2m"),(C_VERDE,"2-4m"),(C_AZUL,"4-6m"),(C_MORADO,">6m")]
            ]),
            unsafe_allow_html=True,
        )

        # ── Detalle por SKU ───────────────────────────────────────────────────
        st.divider()
        st.markdown("### 🔍 Detalle por SKU")

        if det is not None and len(det):
            dcols   = det.columns.tolist()
            marca_c = dcols[0]
            cat_c   = dcols[1] if len(dcols) > 1 else None
            stock_u = next((c for c in dcols if "Stock" in c and "Unid" in c), None)
            cob_c   = next((c for c in dcols if "Cobert" in c), None)
            rank_c  = next((c for c in dcols if "Ranking" in c), None)
            skip_cols = {"Llegadas ABR Unid", "Llegadas MAY Unid"}
            show_cols = [c for c in dcols if c not in skip_cols]

            fd1, fd2, fd3, fd4 = st.columns([2, 2, 1, 2])
            with fd1:
                marcas_disp = sorted(det[marca_c].dropna().unique().tolist())
                sel_marcas = st.multiselect("Filtrar marcas:", marcas_disp, placeholder="Todas las marcas", key="det2_m")
            with fd2:
                cats_disp = sorted(det[cat_c].dropna().unique().tolist()) if cat_c else []
                sel_cats  = st.multiselect("Cat. Comercial:", cats_disp, placeholder="Todas", key="det2_c")
            with fd3:
                solo_ss = st.checkbox("Solo sin stock", key="det2_ss")
            with fd4:
                sort_by = st.selectbox("Ordenar por:", ["Cobertura ↑","Ranking Comercial","Marca A-Z"], key="det2_sort")

            df_f = det.copy()
            if sel_marcas:          df_f = df_f[df_f[marca_c].isin(sel_marcas)]
            if sel_cats and cat_c:  df_f = df_f[df_f[cat_c].isin(sel_cats)]
            if solo_ss and stock_u: df_f = df_f[df_f[stock_u].fillna(0) == 0]
            if sort_by == "Cobertura ↑" and cob_c:        df_f = df_f.sort_values(cob_c)
            elif sort_by == "Ranking Comercial" and rank_c: df_f = df_f.sort_values(rank_c)
            elif sort_by == "Marca A-Z":                   df_f = df_f.sort_values(marca_c)

            total_f = len(df_f)
            ss_f    = int(df_f[stock_u].fillna(0).eq(0).sum()) if stock_u else 0
            st.caption(f"**{total_f} SKUs**" + (f" ({ss_f} sin stock)" if ss_f > 0 else "") + (f" — filtrado de {len(det)}" if total_f != len(det) else ""))

            marcas_en_resultado = df_f[marca_c].dropna().unique().tolist()
            if len(marcas_en_resultado) == 1:
                _dff = df_f[show_cols].reset_index(drop=True)
                st.dataframe(auto_col_config(_dff), hide_index=True, use_container_width=True, height=500)
            else:
                for marca_exp in crit_s["Marca"].tolist():
                    if marca_exp not in marcas_en_resultado: continue
                    df_m  = df_f[df_f[marca_c] == marca_exp]
                    n_m   = len(df_m)
                    ss_m  = int(df_m[stock_u].fillna(0).eq(0).sum()) if stock_u else 0
                    row_m = crit_s[crit_s["Marca"] == marca_exp]
                    cob_m = float(row_m["CobProm"].values[0]) if len(row_m) and "CobProm" in row_m else np.nan
                    cob_m_str = f"{cob_m:.2f}m" if not np.isnan(cob_m) else "—"
                    ss_badge  = f" · 🔴 {ss_m} sin stock" if ss_m > 0 else ""
                    label     = f"**{marca_exp}** — {n_m} SKUs · Cob. {cob_m_str}{ss_badge}"
                    with st.expander(label, expanded=False):
                        _dfm = df_m[show_cols].reset_index(drop=True)
                        st.dataframe(auto_col_config(_dfm), hide_index=True, use_container_width=True, height=min(35*n_m+38, 520))
        else:
            st.info("No hay datos de detalle disponibles.")
    else:
        st.info("No se encontraron datos de stock crítico.")

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 3 — CAPITAL INMOVILIZADO
# ══════════════════════════════════════════════════════════════════════════════
elif seccion == SECCIONES[2]:
    cap_full      = data.get("capital_full")
    cap_total_row = data.get("capital_total")
    sob_det       = data.get("sobrestock")

    # ── KPIs totales (siempre visibles, sobre los tabs) ───────────────────────
    if cap_total_row is not None:
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Stock Total CST",       fmt_mm(_find_val(cap_total_row, ["Stock CST"])))
        m2.metric("Stock Óptimo (4m)",     fmt_mm(_find_val(cap_total_row, ["Óptimo","Optimo"])))
        m3.metric("Capital Inmovilizado",  fmt_mm(_find_val(cap_total_row, ["Inmovil","Capital"])))
        m4.metric("SKUs en Sobrestock",    str(int(_find_val(cap_total_row, ["SKUs"]))) if _find_val(cap_total_row, ["SKUs"]) else "—")
        st.divider()

    # Contar SKUs con llegadas para el label del tab
    n_sob_ll = len(sob_det) if sob_det is not None else 0
    lleg_badge = f" ({n_sob_ll})" if n_sob_ll else ""

    tab_cap, tab_lleg = st.tabs([
        "💜  Capital Inmovilizado",
        f"🚨  Con Llegadas Encima{lleg_badge}",
    ])

    # ── Tab 1: Jerarquía Capital Inmovilizado ────────────────────────────────
    with tab_cap:
        if cap_full is not None and len(cap_full):
            cap_cols  = cap_full.columns.tolist()
            inmovil_c = next((c for c in cap_cols if "Inmovil" in c or ("Capital" in c and "Llegadas" not in c and "Nivel" not in c)), None)
            stock_c   = next((c for c in cap_cols if "Stock CST" in c), None)
            optimo_c  = next((c for c in cap_cols if "Óptimo" in c or "Optimo" in c), None)
            cob_c2    = next((c for c in cap_cols if "Cobert" in c), None)
            skus_c2   = next((c for c in cap_cols if c.strip() == "SKUs"), None)
            meses_c   = next((c for c in cap_cols if "Meses" in c and "Exceso" in c), None)

            # Columnas de datos (excluir metadatos de jerarquía)
            data_cols = [c for c in cap_cols if c not in ("Nivel", "Entidad", "Nombre")]

            NIVEL_STYLE = {
                "CatPadre": {"bg": "#E8D5F5", "fg": "#4A235A", "pl": "14px",  "fw": "700", "fs": "12px"},
                "CatHijo":  {"bg": "#F4ECF7", "fg": "#6C3483", "pl": "28px",  "fw": "600", "fs": "12px"},
                "SKU":      {"bg": "#FFFFFF", "fg": "#222222", "pl": "44px",  "fw": "400", "fs": "11px"},
            }
            ALT_SKU = "#FAFAFA"

            df = cap_full.reset_index(drop=True)
            all_m_pos = df[df["Nivel"] == "Marca"].index.tolist()

            if inmovil_c:
                m_vals = {mi: pd.to_numeric(df.loc[mi, inmovil_c], errors="coerce") for mi in all_m_pos}
                marca_order = sorted(all_m_pos, key=lambda x: m_vals.get(x, 0) or 0, reverse=True)
            else:
                marca_order = all_m_pos

            for mi in marca_order:
                nxt     = [p for p in all_m_pos if p > mi]
                m_end   = nxt[0] if nxt else len(df)
                m_row   = df.loc[mi]
                m_block = df.iloc[mi + 1 : m_end]

                inmovil_v = fmt_mm(m_row.get(inmovil_c, 0)) if inmovil_c else "—"
                skus_v    = int(float(m_row.get(skus_c2, 0))) if skus_c2 and not pd.isna(m_row.get(skus_c2, np.nan)) else "—"
                cob_v     = f"{float(m_row.get(cob_c2, 0)):.1f}m" if cob_c2 and not pd.isna(m_row.get(cob_c2, np.nan)) else "—"
                meses_v   = f"{float(m_row.get(meses_c, 0)):.1f}m exceso" if meses_c and not pd.isna(m_row.get(meses_c, np.nan)) else ""

                m_label = f"{m_row['Nombre']}   ·   {skus_v} SKUs   ·   Cob {cob_v}   ·   {meses_v}   ·   💜 {inmovil_v}"

                with st.expander(m_label, expanded=False):
                    if len(m_block) == 0:
                        st.caption("Sin detalle disponible.")
                        continue

                    th_style = f"background:{C_SOB_BG};color:white;padding:6px 8px;white-space:nowrap;font-size:11px;text-align:right"
                    th_left  = f"background:{C_SOB_BG};color:white;padding:6px 8px;white-space:nowrap;font-size:11px;text-align:left"
                    thead = (
                        "<thead><tr>"
                        + f'<th style="{th_left}">Categoría / SKU</th>'
                        + "".join(f'<th style="{th_style}">{c}</th>' for c in data_cols)
                        + "</tr></thead>"
                    )

                    tbody = ""
                    sku_alt = False
                    for _, row in m_block.iterrows():
                        nivel = row["Nivel"]
                        if nivel not in NIVEL_STYLE:
                            continue
                        s  = NIVEL_STYLE[nivel]
                        bg = s["bg"]
                        if nivel == "SKU":
                            bg = ALT_SKU if sku_alt else "#FFFFFF"
                            sku_alt = not sku_alt
                        else:
                            sku_alt = False

                        nombre_cell = (
                            f'<td style="background:{bg};color:{s["fg"]};padding:5px 8px;'
                            f'padding-left:{s["pl"]};font-weight:{s["fw"]};font-size:{s["fs"]};'
                            f'white-space:nowrap">{row["Nombre"]}</td>'
                        )

                        def _meses_bg(m):
                            if m > 8: return ("#E8DAEF", "#6C3483")
                            if m > 4: return ("#FADBD8", "#CB4335")
                            return           ("#FDEBD0", "#E67E22")

                        def _is_empty(v):
                            try:
                                return pd.isna(v) or str(v).strip() in ("", "nan")
                            except Exception:
                                return False

                        def _fmt_cell(val, col):
                            base_td = (f'background:{bg};color:{s["fg"]};padding:5px 8px;'
                                       f'font-size:{s["fs"]}')
                            if _is_empty(val):
                                return f'<td style="{base_td};text-align:right">—</td>'
                            if col == cob_c2:
                                try:
                                    f = float(val)
                                    cbg = cob_color(f)
                                    return (f'<td style="background:{cbg};color:white;padding:5px 8px;'
                                            f'text-align:center;font-weight:700;font-size:{s["fs"]}">'
                                            f'{f:.1f}m</td>')
                                except Exception:
                                    return f'<td style="{base_td};text-align:center">—</td>'
                            if col == skus_c2:
                                try:
                                    return (f'<td style="{base_td};text-align:center;font-weight:600">'
                                            f'{int(float(val))}</td>')
                                except Exception:
                                    return f'<td style="{base_td};text-align:center">—</td>'
                            if meses_c and col == meses_c:
                                try:
                                    f = float(val)
                                    mbg, mfg = _meses_bg(f)
                                    return (f'<td style="background:{mbg};color:{mfg};padding:5px 8px;'
                                            f'text-align:center;font-weight:700;font-size:{s["fs"]}">'
                                            f'{f:.1f}m</td>')
                                except Exception:
                                    return f'<td style="{base_td};text-align:center">—</td>'
                            if "Llegadas" in str(col):
                                return (f'<td style="{base_td};text-align:center">{val}</td>')
                            try:
                                f = float(val)
                                if np.isnan(f):
                                    return f'<td style="{base_td};text-align:right">—</td>'
                                return (f'<td style="{base_td};text-align:right">{fmt_mm(f)}</td>')
                            except Exception:
                                return f'<td style="{base_td};text-align:right">{val}</td>'

                        data_cells = "".join(_fmt_cell(row.get(c, ""), c) for c in data_cols)
                        tbody += f"<tr>{nombre_cell}{data_cells}</tr>"

                    html_tbl = (
                        '<div style="overflow-x:auto">'
                        '<table style="border-collapse:collapse;width:100%;font-family:Arial">'
                        + thead + "<tbody>" + tbody + "</tbody></table></div>"
                    )
                    st.markdown(html_tbl, unsafe_allow_html=True)
        else:
            st.info("No se encontraron datos de capital inmovilizado.")

    # ── Tab 2: SKUs en sobrestock con llegadas encima ────────────────────────
    with tab_lleg:
        st.caption("Productos con más de 6 meses de cobertura que tienen embarques en camino — requieren revisión comercial.")
        if sob_det is not None and len(sob_det):
            sd_cols  = sob_det.columns.tolist()
            marca_sc = sd_cols[0]
            cat_sc   = next((c for c in sd_cols if "Cat" in c and "Com" in c), sd_cols[1] if len(sd_cols) > 1 else None)
            cob_sc   = next((c for c in sd_cols if "Cobert" in c), None)
            stock_sc = next((c for c in sd_cols if "Stock" in c and "CST" in c), None)
            eta_sc   = next((c for c in sd_cols if "ETA" in c), None)
            lleg_sc  = next((c for c in sd_cols if "Llegadas" in c and "Unid" in c), None)

            f1, f2, f3 = st.columns([2, 2, 2])
            with f1:
                sel_ms = st.multiselect("Filtrar marcas:", sorted(sob_det[marca_sc].dropna().unique()), placeholder="Todas", key="sob_lleg_m")
            with f2:
                cats = sorted(sob_det[cat_sc].dropna().unique()) if cat_sc else []
                sel_cs = st.multiselect("Cat. Comercial:", cats, placeholder="Todas", key="sob_lleg_c")
            with f3:
                sort_sob = st.selectbox("Ordenar por:", ["Stock CST ↓", "Cobertura ↓", "ETA Próx."], key="sob_lleg_sort")

            df_sob = sob_det.copy()
            if sel_ms:            df_sob = df_sob[df_sob[marca_sc].isin(sel_ms)]
            if sel_cs and cat_sc: df_sob = df_sob[df_sob[cat_sc].isin(sel_cs)]
            if sort_sob == "Stock CST ↓" and stock_sc:  df_sob = df_sob.sort_values(stock_sc, ascending=False)
            elif sort_sob == "Cobertura ↓" and cob_sc:  df_sob = df_sob.sort_values(cob_sc, ascending=False)
            elif sort_sob == "ETA Próx." and eta_sc:     df_sob = df_sob.sort_values(eta_sc)

            st.caption(f"**{len(df_sob)} SKUs**" + (f" — filtrado de {len(sob_det)}" if len(df_sob) != len(sob_det) else ""))

            if stock_sc:
                orden_marcas = df_sob.groupby(marca_sc)[stock_sc].sum().sort_values(ascending=False).index.tolist()
            else:
                orden_marcas = sorted(df_sob[marca_sc].dropna().unique().tolist())

            for marca in orden_marcas:
                df_m     = df_sob[df_sob[marca_sc] == marca]
                n_m      = len(df_m)
                stock_v  = fmt_mm(df_m[stock_sc].sum()) if stock_sc else "—"
                cob_avg  = f"{df_m[cob_sc].mean():.1f}m prom" if cob_sc and n_m else "—"
                lleg_tot = f"{int(df_m[lleg_sc].sum())} un." if lleg_sc else ""
                n_cats   = df_m[cat_sc].nunique() if cat_sc else 0
                cat_lbl  = f"   ·   {n_cats} cats" if n_cats > 1 else ""
                lleg_lbl = f"   ·   🚢 {lleg_tot}" if lleg_tot else ""

                with st.expander(
                    f"**{marca}**   ·   {n_m} SKUs{cat_lbl}   ·   Stock CST: {stock_v}   ·   Cob: {cob_avg}{lleg_lbl}",
                    expanded=False
                ):
                    df_show = df_m.copy()
                    if cat_sc:
                        sort_cols = [cat_sc]
                        if sort_sob == "Stock CST ↓" and stock_sc:
                            sort_cols += [stock_sc]; df_show = df_show.sort_values(sort_cols, ascending=[True, False])
                        elif sort_sob == "Cobertura ↓" and cob_sc:
                            sort_cols += [cob_sc];   df_show = df_show.sort_values(sort_cols, ascending=[True, False])
                        elif sort_sob == "ETA Próx." and eta_sc:
                            sort_cols += [eta_sc];   df_show = df_show.sort_values(sort_cols, ascending=True)
                        else:
                            sku_c = next((c for c in df_show.columns if "SKU" in str(c).upper()), None)
                            if sku_c: sort_cols += [sku_c]
                            df_show = df_show.sort_values(sort_cols, ascending=True)
                    _dfsm = df_show.reset_index(drop=True)
                    st.dataframe(auto_col_config(_dfsm), hide_index=True, use_container_width=True,
                                 height=min(35 * n_m + 38, 600))
        else:
            st.info("No hay SKUs en sobrestock con llegadas registradas.")

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 4 — TRÁNSITOS
# ══════════════════════════════════════════════════════════════════════════════
elif seccion == SECCIONES[3]:
    sub1, sub2 = st.tabs(["📦 Por Embarque", "🆕 Nuevos SKUs"])

    with sub1:
        trans_full = data.get("transitos_full")
        trans      = data.get("transitos")        # PI-level only (para totales)
        if trans_full is not None and len(trans_full):
            tc         = trans_full.columns.tolist()
            pi_col_t   = tc[0]
            desc_col_t = tc[1] if len(tc) > 1 else None
            eta_bod_col= next((c for c in tc if "ETA" in c and "Bod" in c), None)
            mes_col    = next((c for c in tc if "Mes" in c), None)
            marc_col_t = next((c for c in tc if "Marcas" in c), None)
            skus_col_t = next((c for c in tc if "SKUs" in c or "Distintos" in c), None)
            crit_col_t = next((c for c in tc if "Crít" in c or "Crit" in c), None)
            inq_col_t  = next((c for c in tc if "Inquiet" in c), None)
            unid_col   = next((c for c in tc if "Unidades" in c), None)
            usd_col    = next((c for c in tc if "USD" in c or "Valor" in c), None)
            riesgo_col = next((c for c in tc if "Riesgo" in c), None)
            cob_col_t  = crit_col_t  # en filas SKU, col Críticos almacena cobertura

            # Marcar tipo de fila
            df_tf = trans_full.reset_index(drop=True).copy()
            df_tf["_es_pi"] = df_tf[marc_col_t].notna() if marc_col_t else True

            # PI rows únicos para filtros
            pi_rows = df_tf[df_tf["_es_pi"]].copy()
            all_pi_idx = df_tf[df_tf["_es_pi"]].index.tolist()

            # Filtros
            fc1, fc2 = st.columns(2)
            with fc1:
                meses_opts = ["Todos"] + (sorted(pi_rows[mes_col].dropna().unique().tolist()) if mes_col else [])
                sel_mes = st.selectbox("Mes llegada:", meses_opts, key="tr_mes")
            with fc2:
                riesgo_opts = ["Todos"] + (sorted(pi_rows[riesgo_col].dropna().unique().tolist()) if riesgo_col else [])
                sel_riesgo  = st.selectbox("Nivel riesgo:", riesgo_opts, key="tr_riesgo")

            # PI filtrados (por posición en df_tf)
            pi_filtered_idx = pi_rows.index.tolist()
            if sel_mes    != "Todos" and mes_col:    pi_filtered_idx = [i for i in pi_filtered_idx if pi_rows.loc[i, mes_col] == sel_mes]
            if sel_riesgo != "Todos" and riesgo_col: pi_filtered_idx = [i for i in pi_filtered_idx if pi_rows.loc[i, riesgo_col] == sel_riesgo]

            n_tot = len(pi_rows)
            n_fil = len(pi_filtered_idx)
            st.caption(f"**{n_fil} embarques**" + (f" — filtrado de {n_tot}" if n_fil != n_tot else ""))

            # Colores de riesgo
            RIESGO_COLOR = {
                "ALTO":    ("#FADBD8", "#CB4335"),
                "MEDIO":   ("#FDEBD0", "#E67E22"),
                "BAJO":    ("#FEF9E7", "#B7950B"),
                "VIGILAR": ("#FEF9E7", "#B7950B"),
                "OK":      ("#D5F5E3", "#1E8449"),
            }
            MES_COLOR_DASH = {
                "ABR26": "#1E8449", "MAY26": "#1E8449",
                "JUN26": "#CA6F1E", "JUL26": "#CA6F1E",
                "AGO26": "#CB4335", "SEP26": "#CB4335",
                "OCT26": "#6C3483",
            }
            COB_COLOR_DET = {
                # (bg, fg) para celdas de cobertura en fila SKU
            }
            def _cob_det_color(v_str):
                """Color badge para cobertura en fila SKU del embarque."""
                if v_str in ("Sin stock", "Sin venta", "—", "", None):
                    return ("#FADBD8", "#CB4335") if v_str == "Sin stock" else ("#EBF5FB", "#1A5276")
                try:
                    f = float(v_str)
                    if f < 1:   return ("#FADBD8", "#CB4335")
                    if f < 2:   return ("#FDEBD0", "#E67E22")
                    if f < 4:   return ("#D5F5E3", "#1E8449")
                    if f < 6:   return ("#D6EAF8", "#1A5276")
                    return             ("#E8DAEF", "#6C3483")
                except Exception:
                    return ("#F4F6F7", "#95A5A6")

            for pi_idx in pi_filtered_idx:
                pi_row = df_tf.loc[pi_idx]

                # Bloque SKU de este PI
                nxt = [p for p in all_pi_idx if p > pi_idx]
                end_idx  = nxt[0] if nxt else len(df_tf)
                sku_block = df_tf.iloc[pi_idx + 1 : end_idx]
                sku_rows  = sku_block[~sku_block["_es_pi"]].copy()

                # Datos del PI
                pi_name  = str(pi_row.get(pi_col_t, "?"))
                eta_bod  = str(pi_row.get(eta_bod_col, "—")) if eta_bod_col else "—"
                mes_v    = str(pi_row.get(mes_col, "—"))     if mes_col else "—"
                marcas_v = str(pi_row.get(marc_col_t, "—"))  if marc_col_t else "—"
                n_skus   = pi_row.get(skus_col_t, "?")       if skus_col_t else "?"
                n_crit   = pi_row.get(crit_col_t, 0)         if crit_col_t else 0
                n_inq    = pi_row.get(inq_col_t, 0)          if inq_col_t else 0
                unid_v   = pi_row.get(unid_col, 0)           if unid_col else 0
                usd_v    = pi_row.get(usd_col, 0)            if usd_col else 0
                riesgo_v = str(pi_row.get(riesgo_col, "—"))  if riesgo_col else "—"

                try: n_crit = int(float(n_crit))
                except: n_crit = 0
                try: n_inq = int(float(n_inq))
                except: n_inq = 0
                try: unid_fmt = f"{int(float(unid_v)):,}"
                except: unid_fmt = str(unid_v)
                try: usd_fmt = f"USD ${float(usd_v):,.0f}"
                except: usd_fmt = str(usd_v)

                rc, rf  = RIESGO_COLOR.get(riesgo_v, ("#F4F6F7", "#555"))
                mes_clr = MES_COLOR_DASH.get(mes_v, "#555")

                crit_badge = (f' · 🔴 {n_crit} críticos' if n_crit > 0 else '') + \
                             (f' · 🟠 {n_inq} inquiet.' if n_inq > 0 else '')

                exp_label = (
                    f"**{pi_name}**"
                    f"   ·   ETA Bodega: {eta_bod}"
                    f"   ·   {mes_v}"
                    f"   ·   {n_skus} SKUs"
                    + crit_badge +
                    f"   ·   {unid_fmt} unid."
                    f"   ·   {usd_fmt}"
                    f"   ·   Riesgo: **{riesgo_v}**"
                )

                with st.expander(exp_label, expanded=False):
                    # Marcas del embarque
                    st.markdown(
                        f'<div style="font-size:12px;color:#555;margin-bottom:6px">'
                        f'🏷️ <b>Marcas:</b> {marcas_v}</div>',
                        unsafe_allow_html=True,
                    )

                    if len(sku_rows) == 0:
                        st.caption("Sin detalle de SKUs disponible.")
                    else:
                        # Tabla HTML de SKUs
                        sku_col    = pi_col_t        # col 0: "  ↳  {sku}" en filas detalle
                        sku_desc_c = desc_col_t      # col 1: descripción
                        sku_eta_c  = eta_bod_col     # col 2: ETA Bodega individual
                        sku_mes_c  = mes_col         # col 3: Mes
                        sku_cob_c  = cob_col_t       # col 6: cobertura individual
                        sku_qty_c  = unid_col        # col 8: cantidad
                        sku_usd_c  = usd_col         # col 9: valor USD

                        th = f"background:{C_AZUL};color:white;padding:5px 8px;font-size:11px;white-space:nowrap"
                        thead = (
                            "<thead><tr>"
                            f'<th style="{th};text-align:left">SKU</th>'
                            f'<th style="{th};text-align:left">Descripción</th>'
                            f'<th style="{th};text-align:center">ETA Bodega</th>'
                            f'<th style="{th};text-align:center">Mes</th>'
                            f'<th style="{th};text-align:center">Cob. Proyect.</th>'
                            f'<th style="{th};text-align:right">Unidades</th>'
                            f'<th style="{th};text-align:right">Valor USD</th>'
                            "</tr></thead>"
                        )

                        tbody = ""
                        for i, (_, sr) in enumerate(sku_rows.iterrows()):
                            bg = "#F4F6F7" if i % 2 == 0 else "#FFFFFF"
                            # SKU: strip leading spaces and ↳ symbol
                            sku_raw = str(sr.get(sku_col, "")).replace("↳", "").strip()
                            desc_v  = str(sr.get(sku_desc_c, "")) if sku_desc_c else ""
                            eta_v   = str(sr.get(sku_eta_c, ""))  if sku_eta_c else ""
                            mes_sv  = str(sr.get(sku_mes_c, ""))  if sku_mes_c else ""
                            cob_raw = sr.get(sku_cob_c, "")       if sku_cob_c else ""
                            cob_str = "" if pd.isna(cob_raw) else str(cob_raw).strip() if isinstance(cob_raw, str) else f"{float(cob_raw):.2f}" if cob_raw != "" else ""
                            qty_raw = sr.get(sku_qty_c, "")       if sku_qty_c else ""
                            usd_raw = sr.get(sku_usd_c, "")       if sku_usd_c else ""

                            try: qty_str = f"{int(float(qty_raw)):,}"
                            except: qty_str = str(qty_raw) if not pd.isna(qty_raw) else "—"
                            try: usd_str = f"${float(usd_raw):,.0f}"
                            except: usd_str = str(usd_raw) if not pd.isna(usd_raw) else "—"

                            cb, cf = _cob_det_color(cob_str)
                            mes_mc = MES_COLOR_DASH.get(mes_sv, "#555")

                            cob_cell = (
                                f'<td style="background:{cb};color:{cf};padding:5px 8px;'
                                f'text-align:center;font-weight:700;font-size:11px">{cob_str or "—"}</td>'
                            )
                            tbody += (
                                f"<tr>"
                                f'<td style="background:{bg};padding:5px 8px;font-size:11px;font-weight:600;color:{C_AZUL}">{sku_raw}</td>'
                                f'<td style="background:{bg};padding:5px 8px;font-size:11px">{desc_v}</td>'
                                f'<td style="background:{bg};padding:5px 8px;font-size:11px;text-align:center">{eta_v}</td>'
                                f'<td style="background:{bg};padding:5px 8px;font-size:11px;text-align:center;color:{mes_mc};font-weight:700">{mes_sv}</td>'
                                + cob_cell +
                                f'<td style="background:{bg};padding:5px 8px;font-size:11px;text-align:right">{qty_str}</td>'
                                f'<td style="background:{bg};padding:5px 8px;font-size:11px;text-align:right">{usd_str}</td>'
                                f"</tr>"
                            )

                        st.markdown(
                            f'<div style="overflow-x:auto"><table style="border-collapse:collapse;width:100%;font-family:Arial">'
                            f'{thead}<tbody>{tbody}</tbody></table></div>',
                            unsafe_allow_html=True,
                        )

            # Totales al final
            if data.get("transitos_total") is not None and trans is not None:
                tot = data["transitos_total"]
                unid_t = int(float(tot[unid_col])) if unid_col and not pd.isna(tot.get(unid_col)) else "—"
                usd_t  = float(tot[usd_col])        if usd_col  and not pd.isna(tot.get(usd_col))  else 0
                st.markdown(
                    f'<div style="background:{C_TOTAL2};color:white;padding:8px 14px;border-radius:4px;font-size:13px;margin-top:8px">'
                    f'<b>TOTAL {n_tot} embarques</b> &nbsp;|&nbsp; '
                    f'{f"{unid_t:,}" if isinstance(unid_t,int) else unid_t} unidades &nbsp;|&nbsp; '
                    f'USD ${usd_t:,.0f}</div>',
                    unsafe_allow_html=True,
                )
        else:
            st.info("No se encontraron datos de tránsitos.")

    with sub2:
        nuevos = data.get("nuevos_transito")
        if nuevos is not None and len(nuevos):
            nc      = nuevos.columns.tolist()
            mes_nc  = next((c for c in nc if "Mes" in c), None)
            marc_nc = next((c for c in nc if "Marca" in c), None)

            nf1, nf2 = st.columns(2)
            with nf1:
                meses_n = ["Todos"] + (sorted(nuevos[mes_nc].dropna().unique().tolist()) if mes_nc else [])
                sel_mn  = st.selectbox("Mes llegada:", meses_n, key="nt_m")
            with nf2:
                marcas_n = ["Todas"] + (sorted(nuevos[marc_nc].dropna().unique().tolist()) if marc_nc else [])
                sel_man  = st.selectbox("Marca:", marcas_n, key="nt_marc")

            df_nt = nuevos.copy()
            if sel_mn  != "Todos"  and mes_nc:  df_nt = df_nt[df_nt[mes_nc]  == sel_mn]
            if sel_man != "Todas"  and marc_nc: df_nt = df_nt[df_nt[marc_nc] == sel_man]

            st.caption(f"**{len(df_nt)} nuevos SKUs**" + (f" — filtrado de {len(nuevos)}" if len(df_nt) != len(nuevos) else ""))
            _dfnt = df_nt.reset_index(drop=True)
            st.dataframe(auto_col_config(_dfnt), hide_index=True, use_container_width=True, height=550)
        else:
            st.info("No se encontraron datos de nuevos tránsitos.")
