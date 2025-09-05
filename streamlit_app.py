# streamlit_app.py — versión consolidada (Inscritos + Egresados + Indicadores + PDF/Excel)
# Incluye: paginación en captura manual, comparativo vs metas, conteos y exportaciones.

import os
import streamlit as st
import pandas as pd
import numpy as np
import datetime
from io import BytesIO
import pyxlsb  # noqa: F401  # requerido por pandas engine

# ReportLab
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics

# ================= CONFIGURACIÓN GENERAL ================= #
st.set_page_config(page_title="Generador de Reportes", layout="wide")

# ======= ESTILO GLOBAL & UI HELPERS ======= #
PRIMARY   = "#264653"   # azul petróleo
SECONDARY = "#2A9D8F"   # verde azulado
ACCENT    = "#E9C46A"   # mostaza
DANGER    = "#E76F51"   # rojo coral
INFO      = "#457B9D"   # azul info
MUTED     = "#586069"   # gris texto

st.markdown(f"""
<style>
/* Fondo de la app y tarjetas */
[data-testid="stAppViewContainer"] {{
  background: #0E1117;
}}
.block-container {{
  padding-top: 0.6rem;
  padding-bottom: 2rem;
  max-width: 1400px;
}}
/* Header con logo */
.app-header {{
  display:flex; align-items:center; justify-content:space-between;
  gap:16px; padding:10px 4px 2px 4px; margin-bottom:10px;
  border-bottom: 1px solid rgba(255,255,255,0.08);
}}
.app-header h1 {{ margin:0; font-size:1.6rem; }}
.app-header p  {{ margin:.25rem 0 0 0; color:#c9d1d9; }}

/* “Cinta” de sección */
.section-band {{
  background: linear-gradient(90deg, {PRIMARY} 0%, {SECONDARY} 60%, transparent 100%);
  border-radius: 14px;
  color: white;
  padding: 14px 18px;
  margin: 8px 0 18px 0;
  border: 1px solid rgba(255,255,255,0.08);
}}
.section-band h2 {{
  font-size: 1.15rem;
  margin: 0;
  letter-spacing: .2px;
}}
/* Tarjeta simple */
.card {{
  border: 1px solid rgba(255,255,255,0.08);
  border-radius: 14px;
  padding: 14px 16px;
  background: rgba(255,255,255,0.02);
  margin-bottom: 10px;
}}
/* Chips */
.chip {{
  display:inline-block; padding:3px 10px; border-radius: 999px; font-size:.82rem;
  border:1px solid rgba(255,255,255,.15); margin-right:6px; color:#E6E6E6;
}}
/* Etiquetas de estado (semaforo) */
.badge {{
  padding: 4px 10px; border-radius: 999px; color: #0E1117; font-weight: 600; display:inline-block;
}}
.badge-green  {{ background:#52c41a; }}
.badge-red    {{ background:#ff4d4f; }}
.badge-amber  {{ background:#fadb14; }}
.badge-grey   {{ background:#d9d9d9; }}

/* Dataframe: reducir espacio y hacer cabezal sticky */
.dataframe thead tr th {{ position: sticky; top: 0; background: #1C212A; z-index: 2; }}
</style>
""", unsafe_allow_html=True)

import os
import streamlit as st

def app_header(
    title: str,
    subtitle: str,
    logo_path: str = "unaq_logo.png",
    logo_width: int = 120,
    logo_top_pad: int = 12,   # 👈 empuja el logo hacia abajo
):
    """Encabezado de la app con logo a la derecha (sin recorte)."""
    col1, col2 = st.columns([5, 1], vertical_alignment="center")

    with col1:
        st.markdown(
            f"""
            <div class="app-header">
              <div>
                <h1 style="margin:0">{title}</h1>
                <p style="margin-top:6px; opacity:.85;">{subtitle}</p>
              </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:
        # separador superior para evitar recorte visual del logo
        st.markdown(f"<div style='height:{logo_top_pad}px'></div>", unsafe_allow_html=True)
        if os.path.exists(logo_path):
            st.image(logo_path, width=logo_width)
        else:
            # si no hay logo, mantenemos el alto para no romper el layout
            st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def leer_excel_auto(file, sheet_name=0, **kw):
    """
    Lee .xlsx (openpyxl), .xls (xlrd) y .xlsb (pyxlsb) automáticamente.
    - file: st.uploaded_file_manager.UploadedFile o ruta
    - sheet_name: índice o nombre de hoja
    """
    # Detectar extensión
    name = getattr(file, "name", str(file)).lower()
    if name.endswith(".xlsb"):
        return pd.read_excel(file, engine="pyxlsb", sheet_name=sheet_name, **kw)
    elif name.endswith(".xls"):
        # requiere 'xlrd' instalado
        return pd.read_excel(file, engine="xlrd", sheet_name=sheet_name, **kw)
    else:
        # .xlsx (openpyxl por defecto)
        return pd.read_excel(file, sheet_name=sheet_name, **kw)


def section_header(title: str, subtitle: str = "", icon: str = "📦"):
    st.markdown(
        f"""<div class="section-band">
              <h2>{icon}&nbsp;&nbsp;{title}</h2>
              {'<div style="opacity:.85;margin-top:4px">'+subtitle+'</div>' if subtitle else ''}
            </div>""",
        unsafe_allow_html=True,
    )

def info_chips(pairs):
    # pairs = [("Cuatrimestre", "C2 2025"), ("Periodo", "May-Ago")]
    html = "".join([f'<span class="chip"><b>{k}:</b> {v}</span>' for k,v in pairs])
    st.markdown(html, unsafe_allow_html=True)

def status_badge(status: str) -> str:
    s = (status or "").lower().strip()
    if s == "verde":     return '<span class="badge badge-green">🟢 Verde</span>'
    if s == "rojo":      return '<span class="badge badge-red">🔴 Rojo</span>'
    if s == "pendiente": return '<span class="badge badge-amber">🟡 Pendiente</span>'
    return '<span class="badge badge-grey">⚪ Sin dato</span>'

# Header principal con logo
app_header(
    "Generador de Reportes de Alumnos e Indicadores",
    "Universidad Aeronáutica en Querétaro",
    logo_path="unaq_logo.png",
    logo_width=400
)

# ================= UTILIDADES ================= #
def norm_txt(s):
    return str(s).strip().lower() if pd.notna(s) else ""

def to_num(x):
    """'58.0%' -> 58.0 ; 'N/A' -> NaN ; '1830' -> 1830.0"""
    if pd.isna(x):
        return np.nan
    s = str(x).strip().replace(",", ".")
    if s.upper() in ["N/A", "NA", "NONE", ""]:
        return np.nan
    if s.endswith("%"):
        s = s[:-1].strip()
        try:
            return float(s)
        except Exception:
            return np.nan
    try:
        return float(s)
    except Exception:
        return np.nan

def is_percent_row(row, cols=("Ene-Abr", "May-Ago", "Sep-Dic")):
    vals = [str(row.get(c, "")) for c in cols]
    return any("%" in v for v in vals)

def elegir_meta_efectiva(row, preferida_col):
    """Regla: usa la del periodo preferido; si es NaN y sólo una de las otras tiene dato, usa esa; si no, NaN."""
    prefer = row.get(preferida_col, np.nan)
    if pd.notna(prefer):
        return prefer
    otras = [c for c in ["Ene-Abr", "May-Ago", "Sep-Dic"] if c != preferida_col]
    vals = [row.get(otras[0], np.nan), row.get(otras[1], np.nan)]
    con_dato = [v for v in vals if pd.notna(v)]
    return con_dato[0] if len(con_dato) == 1 else np.nan

def comparador(resultado, meta):
    if pd.isna(meta):
        return "pendiente"
    if pd.isna(resultado):
        return "sin dato"
    return "verde" if float(resultado) >= float(meta) else "rojo"

def fmt_val(v, is_pct):
    if pd.isna(v):
        return ""
    x = float(v)
    if is_pct:
        if 0 <= x <= 1:
            x *= 100
        return f"{x:.1f}%"
    return f"{x:.0f}" if abs(x - round(x)) < 1e-9 else f"{x:.1f}"

@st.cache_data(show_spinner=False)
def leer_excel_xlsx(file, **kw):
    return pd.read_excel(file, **kw)

@st.cache_data(show_spinner=False)
def leer_excel_xlsb(file, **kw):
    return pd.read_excel(file, engine="pyxlsb", **kw)

# ================= PERIODO / PARÁMETROS ================= #
from datetime import date

# Rango de años mostrado
YEARS = list(range(2020, 2036))
year_today = date.today().year
try:
    default_year_idx = YEARS.index(year_today)
except ValueError:
    default_year_idx = max(0, len(YEARS) - 1)  # último año si el actual no está

section_header("Panel de parámetros", "Selecciona el periodo de trabajo", "🧭")

with st.container():
    colA, colB = st.columns(2)
    with colA:
        cuatrimestre = st.selectbox(
            "📅 Selecciona el cuatrimestre:",
            ["C1", "C2", "C3"],
            index=1
        )
    with colB:
        anio = st.selectbox(
            "📅 Selecciona el año:",
            YEARS,
            index=default_year_idx,
        )

periodo_map = {"C1": "Ene-Abr", "C2": "May-Ago", "C3": "Sep-Dic"}
periodo_col = periodo_map.get(cuatrimestre, "Ene-Abr")
cuatrimestre_actual = f"{cuatrimestre} {anio}"

# Chips informativos
info_chips([("Cuatrimestre", cuatrimestre_actual), ("Periodo", periodo_col)])

# Persistir en session_state para uso posterior
st.session_state["cuatrimestre"] = cuatrimestre
st.session_state["anio"] = anio
st.session_state["periodo_col"] = periodo_col
st.session_state["cuatrimestre_actual"] = cuatrimestre_actual

# ================= SECCIÓN: INSCRITOS ================= #
section_header(
    "Análisis de Alumnos Inscritos",
    "Carga, filtra y explora la matrícula",
    "🧑‍🎓"
)

# --- Cargador ---
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    archivo_inscritos = st.file_uploader(
    "Sube tu archivo de inscripciones (.xlsx / .xls)",
    type=["xlsx", "xls"],
    key="inscritos"
)
st.markdown('</div>', unsafe_allow_html=True)

conteo_inscritos_por_carrera = pd.DataFrame()
conteo_inscritos_por_nivel = pd.DataFrame()

if archivo_inscritos:
    # Usa el lector auto–engine (.xlsx/.xls)
    df_ins = leer_excel_auto(archivo_inscritos, sheet_name=0)

    # --- Vista previa ---
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("📄 Vista previa")
    st.dataframe(df_ins.head(50), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)


    # --- Validación mínima ---
    req_cols = ["Carrera"]
    faltan = [c for c in req_cols if c not in df_ins.columns]
    if faltan:
        st.error(f"Faltan columnas mínimas en Inscritos: {faltan}")
    else:
        # --- Filtros ---
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("🧰 Filtros")
        columnas_filtro = [c for c in ["Carrera", "Sexo", "Periodo", "Grupo", "Ciclo"] if c in df_ins.columns]

        filtros = {}
        for c in columnas_filtro:
            vals = sorted([v for v in df_ins[c].dropna().unique().tolist()])
            sel = st.multiselect(f"Filtrar por {c}", vals, default=vals, key=f"fi_{c}")
            filtros[c] = sel
        st.markdown('</div>', unsafe_allow_html=True)

        # Aplicar filtros
        df_ins_f = df_ins.copy()
        for c, vals in filtros.items():
            if vals:
                df_ins_f = df_ins_f[df_ins_f[c].isin(vals)]

        # --- Clasificación de nivel educativo ---
        def clasificar_nivel_inscrito(carrera):
            txt = str(carrera).lower()
            if "técnico" in txt or "tsu" in txt:
                return "TSU"
            if "maestría" in txt or "posgrado" in txt:
                return "POS"
            if "ingeniería" in txt:
                return "ING"
            return "Otro"

        if "Carrera" in df_ins_f.columns:
            df_ins_f["Nivel"] = df_ins_f["Carrera"].apply(clasificar_nivel_inscrito)

        # --- Conteos por carrera ---
        if "Carrera" in df_ins_f.columns:
            conteo_inscritos_por_carrera = (
                df_ins_f["Carrera"].value_counts().reset_index()
                .rename(columns={"index": "Carrera", "Carrera": "Total de Alumnos"})
            )
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.subheader("📊 Total de alumnos por carrera (filtrado)")
            st.dataframe(conteo_inscritos_por_carrera, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # --- Conteos por nivel ---
        if "Nivel" in df_ins_f.columns:
            conteo_inscritos_por_nivel = (
                df_ins_f["Nivel"].value_counts().reset_index()
                .rename(columns={"index": "Nivel", "Nivel": "Alcanzado"})
            )
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.subheader("🏁 Total de alumnos por nivel educativo")
            # KPIs arriba (opcional)
            try:
                k1, k2, k3 = st.columns(3)
                k1.metric("TSU (alcanzado)", int(conteo_inscritos_por_nivel.set_index("Nivel").get("Alcanzado", {}).get("TSU", 0)))
                k2.metric("ING (alcanzado)", int(conteo_inscritos_por_nivel.set_index("Nivel").get("Alcanzado", {}).get("ING", 0)))
                k3.metric("POS (alcanzado)", int(conteo_inscritos_por_nivel.set_index("Nivel").get("Alcanzado", {}).get("POS", 0)))
            except Exception:
                pass
            st.dataframe(conteo_inscritos_por_nivel, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # --- KPIs automáticos para Indicadores (se guardan en session_state) ---
        niveles_obj = ["TSU", "ING", "POS"]
        conteo_por_nivel = df_ins_f["Nivel"].value_counts() if "Nivel" in df_ins_f.columns else pd.Series(dtype=int)
        df_metricas_auto_ins = pd.DataFrame([
            {
                "Indicador": "Matrícula por nivel Educativo",
                "Responsable": niv,
                "Resultado": int(conteo_por_nivel.get(niv, 0)),
            }
            for niv in niveles_obj
        ])
        st.session_state["metricas_auto_inscritos"] = df_metricas_auto_ins


# ================= SECCIÓN: EGRESADOS ================= #
section_header("Reporte de Egresados", "Carga, filtra y explora los egresados", "🎓")

section_header("Reporte de Egresados", "Carga, filtra y explora los egresados", "🎓")

with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    archivo_egresados = st.file_uploader(
        "Sube tu archivo de egresados (.xlsb / .xlsx / .xls)",
        type=["xlsb", "xlsx", "xls"],
        key="egresados"
    )
    st.markdown('</div>', unsafe_allow_html=True)

# Se mantiene este DF para exportaciones (Excel/PDF)
conteo_egresados_por_carrera = pd.DataFrame()

if archivo_egresados:
    # Detecta extensión y usa el lector apropiado
    fname = getattr(archivo_egresados, "name", "").lower()
    if fname.endswith(".xlsb"):
        df_eg = leer_excel_xlsb(archivo_egresados)
    else:
        # .xlsx o .xls (requiere xlrd para .xls)
        df_eg = leer_excel_auto(archivo_egresados, sheet_name=0)

    # --- Vista previa ---
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("📄 Vista previa")
    st.dataframe(df_eg.head(50), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)


    # ---- Clasificador de nivel (se conserva por si lo usas en filtros)
    def clasificar_nivel_eg(carrera):
        carrera = str(carrera).lower()
        if "maestría" in carrera:
            return "Maestría"
        elif "ingeniería" in carrera:
            return "Ingeniería"
        elif "técnico" in carrera or "tsu" in carrera:
            return "TSU"
        elif "movilidad" in carrera:
            return "Movilidad Académica"
        return "Otro"

    df_eg["Nivel"] = df_eg.get("Carrera", "").apply(clasificar_nivel_eg)

    # ---------------- Filtros ---------------- #
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("🧰 Filtros")
    cols_f = [c for c in ["Carrera", "Sexo", "Periodo", "Grupo", "Ciclo"] if c in df_eg.columns]
    filtros_eg = {}
    for col in cols_f:
        vals = sorted([v for v in df_eg[col].dropna().unique().tolist()])
        filtros_eg[col] = st.multiselect(f"Filtrar por {col}", vals, default=vals, key=f"eg_{col}")
    st.markdown('</div>', unsafe_allow_html=True)

    df_eg_f = df_eg.copy()
    for col, valores in filtros_eg.items():
        if valores:
            df_eg_f = df_eg_f[df_eg_f[col].isin(valores)]

    # ---------------- Generaciones ---------------- #
    generaciones_filtradas = {}
    if "Generación" in df_eg_f.columns and "Nivel" in df_eg_f.columns:
        for nivel in sorted(df_eg_f["Nivel"].dropna().unique().tolist()):
            gens = sorted(df_eg_f[df_eg_f["Nivel"] == nivel]["Generación"].dropna().unique().tolist())
            generaciones_filtradas[nivel] = st.multiselect(
                f"Selecciona generaciones para {nivel}",
                gens, default=gens, key=f"gen_{nivel}"
            )
        if generaciones_filtradas:
            mask = pd.Series(False, index=df_eg_f.index)
            for nivel, gens in generaciones_filtradas.items():
                mask = mask | ((df_eg_f["Nivel"] == nivel) & (df_eg_f["Generación"].isin(gens)))
            df_eg_f = df_eg_f[mask]

    # ---------------- Conteo por carrera (se mantiene) ---------------- #
    if not df_eg_f.empty and "Carrera" in df_eg_f.columns:
        conteo_egresados_por_carrera = df_eg_f["Carrera"].value_counts().reset_index()
        conteo_egresados_por_carrera.columns = ["Carrera", "Total de Egresados"]
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("📊 Total de egresados por carrera (filtrado)")
        st.dataframe(conteo_egresados_por_carrera, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # ---------------- Mapeo Carrera → Código de Programa ---------------- #
    # TSUA, TSUM, TSUF, IAM, IDMA, IECSA, IMA, MIA
    def map_program_code(carrera: str) -> str:
        t = norm_txt(carrera)
        # Posgrado
        if "maestría en ingeniería aeroespacial" in t:
            return "MIA"
        # Ingeniería
        if "ingeniería aeronáutica en manufactura" in t:
            return "IAM"
        if "ingeniería en diseño mecánico aeronáutico" in t:
            return "IDMA"
        if "electrónica y control de sistemas de aeronaves" in t:
            return "IECSA"
        if "ingeniería en mantenimiento aeronáutico" in t:
            return "IMA"
        # TSU (Técnico)
        if ("técnico" in t or "tsu" in t) and "aviónica" in t:
            return "TSUA"
        if ("técnico" in t or "tsu" in t) and ("mantenimiento" in t or "planeador y motor" in t):
            return "TSUM"
        if ("técnico" in t or "tsu" in t) and ("manufactura" in t or "maquinados de precisión" in t or "manufactura de aeronaves" in t):
            return "TSUF"
        # Casos no mapeados (Esp. Valuación, Maestría en Ciencias, etc.)
        return ""

    if "Carrera" in df_eg_f.columns:
        df_eg_f["_prog"] = df_eg_f["Carrera"].map(map_program_code)
    else:
        df_eg_f["_prog"] = ""

    codigos_obj = ["TSUA", "TSUM", "TSUF", "IAM", "IDMA", "IECSA", "IMA", "MIA"]

    # Conteo de egresados por código de programa (sólo códigos de interés)
    conteo_prog = (
        df_eg_f[df_eg_f["_prog"].isin(codigos_obj)]["_prog"]
        .value_counts()
        .reindex(codigos_obj)
        .fillna(0)
        .astype(int)
        .to_dict()
    )

    # ---------------- Ingresos manuales por código y eficiencia ---------------- #
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("🧮 Ingresos por programa y eficiencia terminal (egresados / ingresos)")

    ingresos_manuales = {}
    cols = st.columns(4)
    for i, cod in enumerate(codigos_obj):
        with cols[i % 4]:
            ingresos_manuales[cod] = st.number_input(
                f"Ingresos {cod}",
                min_value=0,
                value=int(st.session_state.get(f"ingresos_{cod}", 0)),
                step=1,
                key=f"ingresos_{cod}"
            )

    # Calcula eficiencia por programa
    resultados_et = {}
    for cod in codigos_obj:
        egresados = int(conteo_prog.get(cod, 0))
        ingresos = int(ingresos_manuales.get(cod, 0))
        resultados_et[cod] = (egresados / ingresos) if ingresos > 0 else np.nan

    # DataFrame para mostrar (incluye % bonito)
    df_et_programas = pd.DataFrame({
        "Programa": codigos_obj,
        "Egresados": [int(conteo_prog.get(c, 0)) for c in codigos_obj],
        "Ingresos":  [int(ingresos_manuales.get(c, 0)) for c in codigos_obj],
        "Eficiencia": [resultados_et[c] for c in codigos_obj],
    })
    # Columna formateada en %
    def _fmt_pct(x):
        if pd.isna(x): return ""
        return f"{x*100:.1f}%"
    df_et_programas["Eficiencia (%)]"] = df_et_programas["Eficiencia"].map(_fmt_pct)

    st.dataframe(df_et_programas, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ---------------- MÉTRICAS AUTOMÁTICAS PARA INDICADORES ---------------- #
    # Indicador: "Eficiencia Terminal por cohorte por Programa Educativo"
    df_metricas_auto_eg = pd.DataFrame([
        {
            "Indicador": "Eficiencia Terminal por cohorte por Programa Educativo",
            "Responsable": cod,
            "Resultado": resultados_et[cod],   # proporción (0..1)
        }
        for cod in codigos_obj
    ])
    st.session_state["metricas_auto_egresados"] = df_metricas_auto_eg


# ================= SECCIÓN: INDICADORES ================= #
section_header("Comparativo de Indicadores vs Metas",
               "Captura variables, calcula resultados y compara contra metas",
               "📈")

# --- Cargador ---
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    archivo_indicadores = st.file_uploader(
        "Sube tu archivo de indicadores (.xlsx)",
        type="xlsx",
        key="indicadores"
    )
    st.markdown('</div>', unsafe_allow_html=True)

comp_out = pd.DataFrame()
captura_manual_df = pd.DataFrame()

if archivo_indicadores:
    # ---------- Hoja 0: base para captura manual (con paginación y búsqueda)
    df_manual = leer_excel_xlsx(archivo_indicadores, sheet_name=0)

    if "captura_manual" not in st.session_state:
        st.session_state["captura_manual"] = {}

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("📝 Captura manual por indicador")
    colf1, colf2 = st.columns([2, 1])
    with colf1:
        filtro_texto = st.text_input("Buscar por Indicador o Responsable", "")
    with colf2:
        page_size = st.number_input(
            "Indicadores por página", min_value=5, max_value=50, value=20, step=5
        )

    if filtro_texto:
        mask = (
            df_manual.get("Indicador", "").astype(str).str.contains(filtro_texto, case=False, na=False)
            | df_manual.get("Responsable", "").astype(str).str.contains(filtro_texto, case=False, na=False)
        )
        df_manual_filtrado = df_manual[mask].reset_index(drop=True)
    else:
        df_manual_filtrado = df_manual.reset_index(drop=True)

    import math
    n_total = len(df_manual_filtrado)
    max_pages = max(1, math.ceil(n_total / page_size))
    page = st.number_input("Página", min_value=1, max_value=max_pages, value=1, step=1)
    ini = int((page - 1) * page_size)
    fin = int(min(n_total, ini + page_size))
    df_page = df_manual_filtrado.iloc[ini:fin].copy()
    st.caption(f"Mostrando {ini+1}–{fin} de {n_total} indicadores")
    st.markdown('</div>', unsafe_allow_html=True)

    # ---------- Form de captura con cálculo y toggle de porcentaje por indicador
    with st.form("frm_captura_manual"):
        def _parse_val(txt: str, use_pct: bool):
            """
            Convierte a número. Si use_pct=True:
              - '50' o '50%' -> 0.5
              - valores 0..1 se dejan como están
            """
            v = to_num(txt)
            if pd.isna(v):
                return np.nan
            s = str(txt).strip()
            if use_pct and (s.endswith("%") or float(v) > 1):
                return float(v) / 100.0
            return float(v)

        registros = []
        pct_flags = {}

        for idx, row in df_page.iterrows():
            nom_ind = row.get("Indicador", f"Indicador {idx+1}")
            resp = row.get("Responsable", "")
            key_base = f"ind::{norm_txt(nom_ind)}::{norm_txt(resp)}"

            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown(f"#### {nom_ind}")

            # estado previo
            v1_prev  = st.session_state["captura_manual"].get(key_base+"::v1", "")
            v2_prev  = st.session_state["captura_manual"].get(key_base+"::v2", "")
            com_prev = st.session_state["captura_manual"].get(key_base+"::com", "")
            pct_prev = bool(st.session_state["captura_manual"].get(key_base+"::pct", False))

            # Toggle por indicador (porcentaje)
            pct_mode = st.checkbox(
                "Escribir variables como porcentaje (50 → 0.5)",
                value=pct_prev,
                key=f"{key_base}::pct_ui"
            )
            pct_flags[key_base] = pct_mode

            col1, col2 = st.columns(2)
            with col1:
                v1 = st.text_input("Variable 1", value=str(v1_prev), key=f"{key_base}::v1_ui")
            with col2:
                v2 = st.text_input("Variable 2", value=str(v2_prev), key=f"{key_base}::v2_ui")

            # Parseo con modo porcentaje por indicador
            v1_num = _parse_val(v1, pct_mode)
            v2_num = _parse_val(v2, pct_mode)

            # Cálculo de resultado = v2 / v1
            if pd.notna(v1_num) and float(v1_num) != 0 and pd.notna(v2_num):
                res_calc = float(v2_num) / float(v1_num)
                res_txt = f"{res_calc:.6f}"
            else:
                res_calc = np.nan
                res_txt = ""

            com = st.text_input("Comentarios", value=com_prev, key=f"{key_base}::com_ui")

            st.caption("Resultado = Variable 2 ÷ Variable 1. "
                       "Con el toggle activo puedes escribir '50' o '50%' y se interpreta como 0.5.")
            st.text_input("Resultado (calculado)", value=res_txt, key=f"{key_base}::res_ui", disabled=True)

            registros.append({
                "Indicador": nom_ind,
                "Responsable": resp,
                "Variable 1": v1,        # texto tal cual
                "Variable 2": v2,        # texto tal cual
                "Resultado": res_calc,   # numérico (proporción)
                "Comentarios": com,
                "_key": key_base,
            })
            st.markdown('</div>', unsafe_allow_html=True)
            st.divider()

        colsb1, colsb2 = st.columns([1, 3])
        with colsb1:
            submitted = st.form_submit_button("Guardar esta página")
        with colsb2:
            limpiar = st.form_submit_button("Limpiar campos de esta página")

    if submitted:
        for r in registros:
            kb = r["_key"]
            st.session_state["captura_manual"][kb+"::v1"]  = r["Variable 1"]
            st.session_state["captura_manual"][kb+"::v2"]  = r["Variable 2"]
            st.session_state["captura_manual"][kb+"::com"] = r["Comentarios"]
            # Guarda toggle por indicador
            st.session_state["captura_manual"][kb+"::pct"] = bool(pct_flags.get(kb, False))
            # Guarda resultado numérico si existe
            if pd.notna(r["Resultado"]):
                st.session_state["captura_manual"][kb+"::res"] = r["Resultado"]
            else:
                st.session_state["captura_manual"].pop(kb+"::res", None)
        st.success("Datos guardados para los indicadores mostrados.")

    if limpiar:
        for r in registros:
            kb = r["_key"]
            for suf in ("::v1", "::v2", "::res", "::com", "::pct"):
                st.session_state["captura_manual"].pop(kb+suf, None)
        st.info("Campos limpiados en esta página.")

    # ---------- Construcción del DataFrame completo con resultado calculado (usando el toggle por indicador)
    rows = []
    for _, row in df_manual_filtrado.iterrows():
        nom_ind = row.get("Indicador", "")
        resp = row.get("Responsable", "")
        key_base = f"ind::{norm_txt(nom_ind)}::{norm_txt(resp)}"

        v1_txt = st.session_state["captura_manual"].get(key_base+"::v1", "")
        v2_txt = st.session_state["captura_manual"].get(key_base+"::v2", "")
        com     = st.session_state["captura_manual"].get(key_base+"::com", "")
        pct_ind = bool(st.session_state["captura_manual"].get(key_base+"::pct", False))

        def _parse_val2(txt: str, use_pct: bool):
            v = to_num(txt)
            if pd.isna(v): return np.nan
            s = str(txt).strip()
            if use_pct and (s.endswith("%") or float(v) > 1):
                return float(v) / 100.0
            return float(v)

        v1_num = _parse_val2(v1_txt, pct_ind)
        v2_num = _parse_val2(v2_txt, pct_ind)
        if pd.notna(v1_num) and float(v1_num) != 0 and pd.notna(v2_num):
            res_calc = float(v2_num) / float(v1_num)
        else:
            prev_res = st.session_state["captura_manual"].get(key_base+"::res", "")
            res_calc = to_num(prev_res)

        rows.append({
            "Indicador": nom_ind,
            "Responsable": resp,
            "Variable 1": v1_txt,
            "Variable 2": v2_txt,
            "Resultado": res_calc,
            "Comentarios": com,
        })
    captura_manual_df = pd.DataFrame(rows)

    # ---------- Hoja2: metas
    try:
        df_metas = leer_excel_xlsx(archivo_indicadores, sheet_name="Hoja2")
    except Exception:
        df_metas = leer_excel_xlsx(archivo_indicadores, sheet_name=1)

    requeridas = ["Indicador", "proceso", "Periodicidad", "Responsable", "Ene-Abr", "May-Ago", "Sep-Dic"]
    faltantes = [c for c in requeridas if c not in df_metas.columns]
    if faltantes:
        st.error(f"En 'Hoja2' faltan columnas requeridas: {faltantes}")
    else:
        # detectar si el indicador es de porcentaje por presencia de "%" en metas
        metas = df_metas.copy()
        metas["_es_pct"] = metas.apply(is_percent_row, axis=1)

        # convertir metas a número
        for col in ["Ene-Abr", "May-Ago", "Sep-Dic"]:
            metas[col] = metas[col].map(to_num)

        # elegir meta efectiva según cuatrimestre elegido
        periodo_col = st.session_state.get("periodo_col", periodo_map.get(st.session_state.get("cuatrimestre", "C2"), "May-Ago"))
        metas["MetaEfectiva"] = metas.apply(lambda r: elegir_meta_efectiva(r, periodo_col), axis=1)

        # ---------- Resultados: captura manual + automáticos de Inscritos + automáticos de Egresados
        df_metricas_auto_ins = st.session_state.get(
            "metricas_auto_inscritos",
            pd.DataFrame(columns=["Indicador", "Responsable", "Resultado"]),
        )
        df_metricas_auto_eg = st.session_state.get(
            "metricas_auto_egresados",
            pd.DataFrame(columns=["Indicador", "Responsable", "Resultado"]),
        )

        resultados_all = pd.concat(
            [
                captura_manual_df[["Indicador", "Responsable", "Resultado"]],
                df_metricas_auto_ins[["Indicador", "Responsable", "Resultado"]],
                df_metricas_auto_eg[["Indicador", "Responsable", "Resultado"]],
            ],
            ignore_index=True
        )

        # claves normalizadas para hacer join
        resultados_all["_ind"] = resultados_all["Indicador"].map(norm_txt)
        resultados_all["_resp"] = resultados_all["Responsable"].map(norm_txt)
        resultados_all["_resultado_num"] = resultados_all["Resultado"].map(to_num)

        metas["_ind"]  = metas["Indicador"].map(norm_txt)
        metas["_resp"] = metas["Responsable"].map(norm_txt)

        # LEFT JOIN desde metas
        comp = metas.merge(
            resultados_all[["_ind", "_resp", "_resultado_num"]],
            on=["_ind", "_resp"],
            how="left",
        )
        comp["Estatus"] = [comparador(r, m) for r, m in zip(comp["_resultado_num"], comp["MetaEfectiva"])]

        # ---------- Salida formateada
        out = comp[[
            "Indicador", "proceso", "Periodicidad", "Responsable",
            "Ene-Abr", "May-Ago", "Sep-Dic", "MetaEfectiva", "_resultado_num", "_es_pct", "Estatus",
        ]].rename(columns={
            "proceso": "Proceso",
            "MetaEfectiva": "Meta efectiva",
            "_resultado_num": "Resultado",
        })

        # Formateos: si es porcentaje, se muestra como % (0.8 -> 80.0%)
        out["Meta Ene-Abr"] = [fmt_val(v, p) for v, p in zip(out["Ene-Abr"], out["_es_pct"])]
        out["Meta May-Ago"] = [fmt_val(v, p) for v, p in zip(out["May-Ago"], out["_es_pct"])]
        out["Meta Sep-Dic"] = [fmt_val(v, p) for v, p in zip(out["Sep-Dic"], out["_es_pct"])]
        out["Meta efectiva"] = [fmt_val(v, p) for v, p in zip(out["Meta efectiva"], out["_es_pct"])]
        out["Resultado"]     = [fmt_val(v, p) for v, p in zip(out["Resultado"],     out["_es_pct"])]

        comp_out = out[[
            "Indicador", "Proceso", "Periodicidad", "Responsable",
            "Meta Ene-Abr", "Meta May-Ago", "Meta Sep-Dic", "Meta efectiva", "Resultado", "Estatus",
        ]]

        # === Semáforo (emojis) ===
        SEMAFORO = {
            "verde": "🟢 Verde",
            "rojo": "🔴 Rojo",
            "pendiente": "🟡 Pendiente",
            "sin dato": "⚪ Sin dato",
        }
        if not comp_out.empty:
            comp_out.insert(
                comp_out.columns.get_loc("Estatus") + 1,
                "Semáforo",
                comp_out["Estatus"].map(SEMAFORO).fillna("⚪ Sin dato")
            )

        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Metas (Hoja2) completas y comparación")
        st.dataframe(comp_out, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)


# ================= PDF / EXCEL ================= #

def _table_col_widths(df, max_total_width):
    if df is None or df.empty:
        return []
    font_name, font_size = "Helvetica", 7
    cols = df.columns.tolist()
    widths = []
    for col in cols:
        header_w = pdfmetrics.stringWidth(str(col), font_name, font_size + 1)
        sample_rows = df[col].astype(str).head(30).tolist()
        body_w = max([pdfmetrics.stringWidth(s, font_name, font_size) for s in ([""] + sample_rows)])
        widths.append(max(header_w, body_w) + 12)
    total = sum(widths)
    if total <= 0:
        return [max_total_width / max(1, len(cols))] * len(cols)
    ratio = min(1.0, max_total_width / total)
    widths = [w * ratio for w in widths]
    diff = max_total_width - sum(widths)
    if widths:
        widths[-1] += diff
    return widths


def generar_reporte_pdf(df_indicadores, df_inscritos, df_egresados, cuatri_texto):
    """PDF con bloques por Proceso y estilo corporativo."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(letter),
        leftMargin=24, rightMargin=24, topMargin=36, bottomMargin=36,
    )
    elementos = []
    estilos = getSampleStyleSheet()

    # Estilos
    estilo_title = estilos['Title']
    estilo_sub = estilos['Heading2']
    estilo_celda = ParagraphStyle(name='TablaNormal', fontSize=7, leading=8)
    estilo_header = ParagraphStyle(name='TablaHeader', fontSize=7, leading=8, textColor=colors.white, fontName='Helvetica-Bold')

    azul_rey = HexColor("#0B2E59")
    azul_process = HexColor("#2E75B6")
    gris_process = HexColor("#5F7383")
    gris_zebra = HexColor("#F4F6F8")

    fecha = datetime.date.today().strftime("%d/%m/%Y")
    elementos.append(Paragraph("METAS — Comparativo y Conteos", estilo_title))
    elementos.append(Paragraph(f"Periodo: {cuatri_texto} — Generado el {fecha}", estilos['Normal']))
    elementos.append(Spacer(1, 10))

    # ---- Tabla de indicadores por bloque de Proceso
    if df_indicadores is not None and not df_indicadores.empty:
        elementos.append(Paragraph("Indicadores (Comparativo)", estilo_sub))

        # Orden de columnas mostradas
        cols_show = [
            "Indicador", "Responsable", "Periodicidad",
            "Meta Ene-Abr", "Meta May-Ago", "Meta Sep-Dic",
            "Meta efectiva", "Resultado", "Estatus", "Semáforo"
        ]
        df_src = df_indicadores.copy()
        # Si falta alguna, la ignoramos
        cols_show = [c for c in cols_show if c in df_src.columns]

        # Agrupar por Proceso si existe
        if "Proceso" in df_src.columns:
            grupos = df_src.groupby("Proceso")
            color_toggle = True
            for proceso, df_g in grupos:
                # Barra de proceso
                barra = Table([[f"Proceso: {proceso}"]],
                              colWidths=[landscape(letter)[0] - (doc.leftMargin + doc.rightMargin)])
                barra.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), azul_process if color_toggle else gris_process),
                    ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('LEFTPADDING', (0, 0), (-1, -1), 6),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                    ('TOPPADDING', (0, 0), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ]))
                color_toggle = not color_toggle
                elementos.append(barra)
                elementos.append(Spacer(1, 4))

                # Tabla del bloque
                df_g2 = df_g[cols_show].copy()
                data = [[Paragraph(str(c), estilo_header) for c in df_g2.columns]]
                for _, r in df_g2.iterrows():
                    data.append([Paragraph(str(r[c]), estilo_celda) for c in df_g2.columns])

                ancho_util = landscape(letter)[0] - (doc.leftMargin + doc.rightMargin)
                col_widths = _table_col_widths(df_g2, ancho_util)
                t = Table(data, repeatRows=1, colWidths=col_widths)
                stl = [
                    ('BACKGROUND', (0, 0), (-1, 0), azul_rey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ]
                for i in range(1, len(data)):
                    if i % 2 == 0:
                        stl.append(('BACKGROUND', (0, i), (-1, i), gris_zebra))
                t.setStyle(TableStyle(stl))
                elementos.append(t)
                elementos.append(Spacer(1, 10))
        else:
            # Fallback: todo en una sola tabla
            df_g2 = df_src[cols_show].copy()
            data = [[Paragraph(str(c), estilo_header) for c in df_g2.columns]]
            for _, r in df_g2.iterrows():
                data.append([Paragraph(str(r[c]), estilo_celda) for c in df_g2.columns])
            ancho_util = landscape(letter)[0] - (doc.leftMargin + doc.rightMargin)
            col_widths = _table_col_widths(df_g2, ancho_util)
            t = Table(data, repeatRows=1, colWidths=col_widths)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), azul_rey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ]))
            elementos.append(t)
            elementos.append(Spacer(1, 10))

    # ---- Conteos extra
    if df_inscritos is not None and not df_inscritos.empty:
        elementos.append(Paragraph("Inscritos (conteo por carrera)", estilo_sub))
        data = [[Paragraph(str(c), estilo_header) for c in df_inscritos.columns]]
        for _, r in df_inscritos.iterrows():
            data.append([Paragraph(str(x), estilo_celda) for x in r])
        ancho_util = landscape(letter)[0] - (doc.leftMargin + doc.rightMargin)
        col_widths = _table_col_widths(df_inscritos, ancho_util)
        t = Table(data, repeatRows=1, colWidths=col_widths)
        stl = [
            ('BACKGROUND', (0, 0), (-1, 0), azul_rey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
        ]
        for i in range(1, len(data)):
            if i % 2 == 0:
                stl.append(('BACKGROUND', (0, i), (-1, i), gris_zebra))
        t.setStyle(TableStyle(stl))
        elementos.append(t)
        elementos.append(Spacer(1, 10))

    if df_egresados is not None and not df_egresados.empty:
        elementos.append(Paragraph("Egresados (conteo por carrera)", estilo_sub))
        data = [[Paragraph(str(c), estilo_header) for c in df_egresados.columns]]
        for _, r in df_egresados.iterrows():
            data.append([Paragraph(str(x), estilo_celda) for x in r])
        ancho_util = landscape(letter)[0] - (doc.leftMargin + doc.rightMargin)
        col_widths = _table_col_widths(df_egresados, ancho_util)
        t = Table(data, repeatRows=1, colWidths=col_widths)
        stl = [
            ('BACKGROUND', (0, 0), (-1, 0), azul_rey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
        ]
        for i in range(1, len(data)):
            if i % 2 == 0:
                stl.append(('BACKGROUND', (0, i), (-1, i), gris_zebra))
        t.setStyle(TableStyle(stl))
        elementos.append(t)

    # Footer
    def _footer(canvas, docx):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        text = f"{cuatri_texto} — Página {docx.page}"
        canvas.drawRightString(landscape(letter)[0] - docx.rightMargin, 18, text)
        canvas.restoreState()

    doc.build(elementos, onFirstPage=_footer, onLaterPages=_footer)
    buffer.seek(0)
    return buffer.read()


def exportar_excel_corporativo(comp_out, conteo_inscritos_por_carrera, conteo_egresados_por_carrera, cuatrimestre_actual):
    """
    Excel con estilo similar a tu plantilla:
    - Título principal
    - Bandas por Proceso
    - Cabeceras y bordes
    - Semáforo/Estatus coloreado
    """
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book

        # ---- Formatos
        f_title = wb.add_format({
            'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter',
            'font_color': 'white', 'bg_color': '#0B2E59'
        })
        f_hdr = wb.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'font_color': 'white', 'bg_color': '#0B2E59', 'border': 1
        })
        f_cell = wb.add_format({'align': 'left', 'valign': 'top', 'border': 1})
        f_num  = wb.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1})
        f_band_blue = wb.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#2E75B6', 'border': 1})
        f_band_gray = wb.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#5F7383', 'border': 1})

        f_est_verde = wb.add_format({'align': 'center', 'border': 1, 'bg_color': '#C6E0B4'})
        f_est_rojo  = wb.add_format({'align': 'center', 'border': 1, 'bg_color': '#F8CBAD'})
        f_est_pend  = wb.add_format({'align': 'center', 'border': 1, 'bg_color': '#FFE699'})
        f_est_sin   = wb.add_format({'align': 'center', 'border': 1, 'bg_color': '#D9D9D9'})

        # ---- Comparativo (corporativo)
        ws = wb.add_worksheet("Comparativo")
        # Columnas a mostrar
        base_cols = [
            "Indicador", "Responsable", "Periodicidad",
            "Meta Ene-Abr", "Meta May-Ago", "Meta Sep-Dic",
            "Meta efectiva", "Resultado", "Estatus", "Semáforo"
        ]
        cols = [c for c in base_cols if c in comp_out.columns]

        # Título
        ncols = len(cols)
        ws.merge_range(0, 0, 1, ncols-1, f"METAS — {cuatrimestre_actual}", f_title)

        # Cabeceras
        for j, c in enumerate(cols):
            ws.write(3, j, c, f_hdr)

        # Anchos sugeridos
        widths = {
            "Indicador": 48, "Responsable": 12, "Periodicidad": 12,
            "Meta Ene-Abr": 12, "Meta May-Ago": 12, "Meta Sep-Dic": 12,
            "Meta efectiva": 14, "Resultado": 12, "Estatus": 12, "Semáforo": 12
        }
        for j, c in enumerate(cols):
            ws.set_column(j, j, widths.get(c, 14))

        # Filas por Proceso en bandas
        df_cmp = comp_out.copy()
        band_blue = True
        row = 4
        if "Proceso" in df_cmp.columns:
            for proceso, df_g in df_cmp.groupby("Proceso"):
                ws.merge_range(row, 0, row, ncols-1, f"Proceso: {proceso}", f_band_blue if band_blue else f_band_gray)
                band_blue = not band_blue
                row += 1
                for _, r in df_g.iterrows():
                    for j, c in enumerate(cols):
                        val = r.get(c, "")
                        if c in ("Meta Ene-Abr", "Meta May-Ago", "Meta Sep-Dic", "Meta efectiva", "Resultado"):
                            ws.write(row, j, val, f_num)
                        elif c == "Estatus":
                            fmt = {
                                "verde": f_est_verde, "rojo": f_est_rojo,
                                "pendiente": f_est_pend, "sin dato": f_est_sin
                            }.get(str(r.get("Estatus", "")).strip(), f_cell)
                            ws.write(row, j, str(val), fmt)
                        else:
                            ws.write(row, j, val, f_cell)
                    row += 1
        else:
            for _, r in df_cmp.iterrows():
                for j, c in enumerate(cols):
                    ws.write(row, j, r.get(c, ""), f_cell)
                row += 1

        # ---- Hoja Inscritos
        if not conteo_inscritos_por_carrera.empty:
            conteo_inscritos_por_carrera.to_excel(writer, sheet_name="Inscritos", index=False)

        # ---- Hoja Egresados
        if not conteo_egresados_por_carrera.empty:
            conteo_egresados_por_carrera.to_excel(writer, sheet_name="Egresados", index=False)

    buf.seek(0)
    return buf


# ===== DESCARGAS ===== #
colL, colR = st.columns([3, 2])

with colL:
    if 'comp_out' in locals() and not comp_out.empty:
        # Excel corporativo
        excel_buffer = exportar_excel_corporativo(
            comp_out,
            conteo_inscritos_por_carrera if 'conteo_inscritos_por_carrera' in locals() else pd.DataFrame(),
            conteo_egresados_por_carrera if 'conteo_egresados_por_carrera' in locals() else pd.DataFrame(),
            cuatrimestre_actual
        )
        st.download_button(
            "📊 Descargar Excel (estilo corporativo)",
            data=excel_buffer,
            file_name=f"Metas_{cuatrimestre_actual.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with colR:
    if 'comp_out' in locals() and not comp_out.empty \
       and 'conteo_inscritos_por_carrera' in locals() and not conteo_inscritos_por_carrera.empty \
       and 'conteo_egresados_por_carrera' in locals() and not conteo_egresados_por_carrera.empty:
        st.subheader("🖨️ Reporte PDF")
        pdf_bytes = generar_reporte_pdf(
            comp_out,
            conteo_inscritos_por_carrera,
            conteo_egresados_por_carrera,
            cuatrimestre_actual
        )
        st.download_button(
            "📥 Descargar PDF (estilo corporativo)",
            data=pdf_bytes,
            file_name=f"Reporte_{cuatrimestre_actual.replace(' ', '_')}.pdf",
            mime="application/pdf",
        )
    else:
        st.info("Carga Indicadores, Inscritos y Egresados y genera el comparativo para habilitar las descargas.")
