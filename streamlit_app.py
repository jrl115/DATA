# streamlit_app.py — versión consolidada (Inscritos + Egresados + Indicadores + PDF/Excel)
# Incluye: paginación en captura manual, comparativo vs metas, conteos y exportaciones.

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
st.title("Generador de Reportes de Alumnos e Indicadores")

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


# ================= SELECTORES DE PERIODO ================= #
colA, colB = st.columns(2)
with colA:
    cuatrimestre = st.selectbox("📅 Selecciona el cuatrimestre:", ["C1", "C2", "C3"], index=1)
with colB:
    anio = st.selectbox(
        "📅 Selecciona el año:",
        list(range(2020, 2036)),
        index=list(range(2020, 2036)).index(datetime.date.today().year),
    )

periodo_map = {"C1": "Ene-Abr", "C2": "May-Ago", "C3": "Sep-Dic"}
periodo_col = periodo_map[cuatrimestre]
cuatrimestre_actual = f"{cuatrimestre} {anio}"
st.caption(f"Has seleccionado: **{cuatrimestre_actual} ({periodo_col})**")

# ================= SECCIÓN: INSCRITOS ================= #
st.header("Análisis de Alumnos Inscritos")
archivo_inscritos = st.file_uploader("Sube tu archivo de inscripciones (.xlsx)", type="xlsx", key="inscritos")

conteo_inscritos_por_carrera = pd.DataFrame()
conteo_inscritos_por_nivel = pd.DataFrame()

if archivo_inscritos:
    df_ins = leer_excel_xlsx(archivo_inscritos)

    st.subheader("Vista previa de inscripciones")
    st.dataframe(df_ins.head(50), use_container_width=True)

    req_cols = ["Carrera"]
    faltan = [c for c in req_cols if c not in df_ins.columns]
    if faltan:
        st.error(f"Faltan columnas mínimas en Inscritos: {faltan}")
    else:
        columnas_filtro = [c for c in ["Carrera", "Sexo", "Periodo", "Grupo", "Ciclo"] if c in df_ins.columns]
        with st.expander("Filtros", expanded=True):
            filtros = {}
            for c in columnas_filtro:
                vals = sorted([v for v in df_ins[c].dropna().unique().tolist()])
                sel = st.multiselect(f"Filtrar por {c}", vals, default=vals, key=f"fi_{c}")
                filtros[c] = sel

        df_ins_f = df_ins.copy()
        for c, vals in filtros.items():
            if vals:
                df_ins_f = df_ins_f[df_ins_f[c].isin(vals)]

        def clasificar_nivel_inscrito(carrera):
            txt = str(carrera).lower()
            if "técnico" in txt or "tsu" in txt:
                return "TSU"
            if "maestría" in txt or "posgrado" in txt:
                return "POS"
            if "ingeniería" in txt:
                return "ING"
            return "Otro"

        df_ins_f["Nivel"] = df_ins_f["Carrera"].apply(clasificar_nivel_inscrito)

        if "Carrera" in df_ins_f.columns:
            conteo_inscritos_por_carrera = df_ins_f["Carrera"].value_counts().reset_index()
            conteo_inscritos_por_carrera.columns = ["Carrera", "Total de Alumnos"]
            st.subheader("Total de alumnos por carrera (filtrado)")
            st.dataframe(conteo_inscritos_por_carrera, use_container_width=True)

        conteo_inscritos_por_nivel = df_ins_f["Nivel"].value_counts().reset_index()
        conteo_inscritos_por_nivel.columns = ["Nivel", "Alcanzado"]
        st.subheader("Total de alumnos por nivel educativo")
        st.dataframe(conteo_inscritos_por_nivel, use_container_width=True)

        # KPIs automáticos para Indicadores
        niveles_obj = ["TSU", "ING", "POS"]
        conteo_por_nivel = df_ins_f["Nivel"].value_counts()
        df_metricas_auto_ins = pd.DataFrame([
            {"Indicador": "Matrícula por nivel Educativo", "Responsable": niv, "Resultado": int(conteo_por_nivel.get(niv, 0))}
            for niv in niveles_obj
        ])
        st.session_state["metricas_auto_inscritos"] = df_metricas_auto_ins

        with st.expander("Programado por nivel (opcional)"):
            programados = {fila["Nivel"]: st.number_input(f"Programado para {fila['Nivel']}", min_value=0, value=0, key=f"prog_{fila['Nivel']}") for _, fila in conteo_inscritos_por_nivel.iterrows()}
        if not conteo_inscritos_por_nivel.empty:
            conteo_inscritos_por_nivel["Programado"] = conteo_inscritos_por_nivel["Nivel"].map(programados)
            total_row = pd.DataFrame({
                "Nivel": ["TOTAL"],
                "Alcanzado": [conteo_inscritos_por_nivel["Alcanzado"].sum()],
                "Programado": [conteo_inscritos_por_nivel["Programado"].sum()],
            })
            resumen_final = pd.concat([conteo_inscritos_por_nivel, total_row], ignore_index=True)
            st.subheader("Resumen final")
            st.dataframe(resumen_final, use_container_width=True)

# ================= SECCIÓN: EGRESADOS ================= #
st.header("Reporte de Egresados")
archivo_egresados = st.file_uploader("Sube tu archivo de egresados (.xlsb)", type="xlsb")

conteo_egresados_por_carrera = pd.DataFrame()

if archivo_egresados:
    df_eg = leer_excel_xlsb(archivo_egresados)

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

    cols_f = [c for c in ["Carrera", "Sexo", "Periodo", "Grupo", "Ciclo"] if c in df_eg.columns]
    with st.expander("Filtros", expanded=True):
        filtros_eg = {}
        for col in cols_f:
            vals = sorted([v for v in df_eg[col].dropna().unique().tolist()])
            filtros_eg[col] = st.multiselect(f"Filtrar por {col}", vals, default=vals, key=f"eg_{col}")

    df_eg_f = df_eg.copy()
    for col, valores in filtros_eg.items():
        if valores:
            df_eg_f = df_eg_f[df_eg_f[col].isin(valores)]

    generaciones_filtradas = {}
    if "Generación" in df_eg_f.columns and "Nivel" in df_eg_f.columns:
        for nivel in sorted(df_eg_f["Nivel"].dropna().unique().tolist()):
            gens = sorted(df_eg_f[df_eg_f["Nivel"] == nivel]["Generación"].dropna().unique().tolist())
            generaciones_filtradas[nivel] = st.multiselect(f"Selecciona generaciones para {nivel}", gens, default=gens, key=f"gen_{nivel}")
        if generaciones_filtradas:
            mask = pd.Series(False, index=df_eg_f.index)
            for nivel, gens in generaciones_filtradas.items():
                mask = mask | ((df_eg_f["Nivel"] == nivel) & (df_eg_f["Generación"].isin(gens)))
            df_eg_f = df_eg_f[mask]

    if not df_eg_f.empty and "Carrera" in df_eg_f.columns:
        conteo_egresados_por_carrera = df_eg_f["Carrera"].value_counts().reset_index()
        conteo_egresados_por_carrera.columns = ["Carrera", "Total de Egresados"]
        st.subheader("Total de egresados por carrera (filtrado)")
        st.dataframe(conteo_egresados_por_carrera, use_container_width=True)

# ================= SECCIÓN: INDICADORES ================= #
st.header("Comparativo de Indicadores vs Metas")
archivo_indicadores = st.file_uploader("Sube tu archivo de indicadores (.xlsx)", type="xlsx", key="indicadores")

comp_out = pd.DataFrame()
captura_manual_df = pd.DataFrame()

if archivo_indicadores:
    # Hoja 0: base para captura manual (con paginación y búsqueda)
    df_manual = leer_excel_xlsx(archivo_indicadores, sheet_name=0)

    if "captura_manual" not in st.session_state:
        st.session_state["captura_manual"] = {}

    st.markdown("**Captura manual por indicador**")
    colf1, colf2, colf3 = st.columns([2, 1, 1])
    with colf1:
        filtro_texto = st.text_input("Buscar por Indicador o Responsable", "")
    with colf2:
        page_size = st.number_input("Indicadores por página", min_value=5, max_value=50, value=20, step=5)

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

    with st.form("frm_captura_manual"):
        registros = []
        for idx, row in df_page.iterrows():
            nom_ind = row.get("Indicador", f"Indicador {idx+1}")
            resp = row.get("Responsable", "")
            st.markdown(f"#### {nom_ind}")
            key_base = f"ind::{norm_txt(nom_ind)}::{norm_txt(resp)}"
            col1, col2 = st.columns(2)
            with col1:
                v1 = st.text_input("Variable 1", value=st.session_state["captura_manual"].get(key_base+"::v1", ""), key=f"{key_base}::v1_ui")
                res = st.text_input("Resultado",  value=st.session_state["captura_manual"].get(key_base+"::res", ""), key=f"{key_base}::res_ui")
            with col2:
                v2 = st.text_input("Variable 2", value=st.session_state["captura_manual"].get(key_base+"::v2", ""), key=f"{key_base}::v2_ui")
                com = st.text_input("Comentarios", value=st.session_state["captura_manual"].get(key_base+"::com", ""), key=f"{key_base}::com_ui")
            registros.append({
                "Indicador": nom_ind,
                "Responsable": resp,
                "Variable 1": v1,
                "Variable 2": v2,
                "Resultado": res,
                "Comentarios": com,
                "_key": key_base,
            })
            st.divider()

        colsb1, colsb2 = st.columns([1, 3])
        with colsb1:
            submitted = st.form_submit_button("Guardar esta página")
        with colsb2:
            limpiar = st.form_submit_button("Limpiar campos de esta página")

    if submitted:
        for r in registros:
            st.session_state["captura_manual"][r["_key"]+"::v1"] = r["Variable 1"]
            st.session_state["captura_manual"][r["_key"]+"::v2"] = r["Variable 2"]
            st.session_state["captura_manual"][r["_key"]+"::res"] = r["Resultado"]
            st.session_state["captura_manual"][r["_key"]+"::com"] = r["Comentarios"]
        st.success("Datos guardados para los indicadores mostrados.")

    if limpiar:
        for r in registros:
            for suf in ("::v1", "::v2", "::res", "::com"):
                st.session_state["captura_manual"].pop(r["_key"]+suf, None)
        st.info("Campos limpiados en esta página.")

    rows = []
    for _, row in df_manual_filtrado.iterrows():
        nom_ind = row.get("Indicador", "")
        resp = row.get("Responsable", "")
        key_base = f"ind::{norm_txt(nom_ind)}::{norm_txt(resp)}"
        rows.append({
            "Indicador": nom_ind,
            "Responsable": resp,
            "Variable 1": st.session_state["captura_manual"].get(key_base+"::v1", ""),
            "Variable 2": st.session_state["captura_manual"].get(key_base+"::v2", ""),
            "Resultado": st.session_state["captura_manual"].get(key_base+"::res", ""),
            "Comentarios": st.session_state["captura_manual"].get(key_base+"::com", ""),
        })
    captura_manual_df = pd.DataFrame(rows)

    # Hoja2: metas
    try:
        df_metas = leer_excel_xlsx(archivo_indicadores, sheet_name="Hoja2")
    except Exception:
        df_metas = leer_excel_xlsx(archivo_indicadores, sheet_name=1)

    requeridas = ["Indicador", "proceso", "Periodicidad", "Responsable", "Ene-Abr", "May-Ago", "Sep-Dic"]
    faltantes = [c for c in requeridas if c not in df_metas.columns]
    if faltantes:
        st.error(f"En 'Hoja2' faltan columnas requeridas: {faltantes}")
    else:
        metas = df_metas.copy()
        metas["_es_pct"] = metas.apply(is_percent_row, axis=1)
        for col in ["Ene-Abr", "May-Ago", "Sep-Dic"]:
            metas[col] = metas[col].map(to_num)
        metas["MetaEfectiva"] = metas.apply(lambda r: elegir_meta_efectiva(r, periodo_col), axis=1)

        # Resultados: captura manual + métricas automáticas de Inscritos
        df_metricas_auto = st.session_state.get(
            "metricas_auto_inscritos",
            pd.DataFrame(columns=["Indicador", "Responsable", "Resultado"]),
        )
        resultados_all = pd.concat([
            captura_manual_df[["Indicador", "Responsable", "Resultado"]],
            df_metricas_auto[["Indicador", "Responsable", "Resultado"]],
        ], ignore_index=True)

        resultados_all["_ind"] = resultados_all["Indicador"].map(norm_txt)
        resultados_all["_resp"] = resultados_all["Responsable"].map(norm_txt)
        resultados_all["_resultado_num"] = resultados_all["Resultado"].map(to_num)

        metas["_ind"] = metas["Indicador"].map(norm_txt)
        metas["_resp"] = metas["Responsable"].map(norm_txt)

        comp = metas.merge(
            resultados_all[["_ind", "_resp", "_resultado_num"]],
            on=["_ind", "_resp"],
            how="left",
        )
        comp["Estatus"] = [comparador(r, m) for r, m in zip(comp["_resultado_num"], comp["MetaEfectiva"])]

        out = comp[[
            "Indicador", "proceso", "Periodicidad", "Responsable",
            "Ene-Abr", "May-Ago", "Sep-Dic", "MetaEfectiva", "_resultado_num", "_es_pct", "Estatus",
        ]].rename(columns={
            "proceso": "Proceso",
            "MetaEfectiva": "Meta efectiva",
            "_resultado_num": "Resultado",
        })

        out["Meta Ene-Abr"] = [fmt_val(v, p) for v, p in zip(out["Ene-Abr"], out["_es_pct"])]
        out["Meta May-Ago"] = [fmt_val(v, p) for v, p in zip(out["May-Ago"], out["_es_pct"])]
        out["Meta Sep-Dic"] = [fmt_val(v, p) for v, p in zip(out["Sep-Dic"], out["_es_pct"])]
        out["Meta efectiva"] = [fmt_val(v, p) for v, p in zip(out["Meta efectiva"], out["_es_pct"])]
        out["Resultado"] = [fmt_val(v, p) for v, p in zip(out["Resultado"], out["_es_pct"])]

        comp_out = out[[
            "Indicador", "Proceso", "Periodicidad", "Responsable",
            "Meta Ene-Abr", "Meta May-Ago", "Meta Sep-Dic", "Meta efectiva", "Resultado", "Estatus",
        ]]

        st.subheader("Metas (Hoja2) completas y comparación")
        st.dataframe(comp_out, use_container_width=True)

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
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        leftMargin=24, rightMargin=24, topMargin=36, bottomMargin=36,
    )
    elementos = []
    estilos = getSampleStyleSheet()

    estilo_celda = ParagraphStyle(name='TablaNormal', fontSize=7, leading=8)
    estilo_header = ParagraphStyle(name='TablaHeader', fontSize=7, leading=8, textColor=colors.white, fontName='Helvetica-Bold')

    azul_rey = HexColor("#003366")
    gris_zebra = HexColor("#f2f2f2")

    fecha = datetime.date.today().strftime("%d/%m/%Y")
    elementos.append(Paragraph("Reporte General - Indicadores, Inscritos y Egresados", estilos['Title']))
    elementos.append(Paragraph(f"Fecha de generación: {fecha} — Cuatrimestre: {cuatri_texto}", estilos['Normal']))
    elementos.append(Spacer(1, 12))

    def agregar_tabla(titulo, df):
        if df is None or df.empty:
            return
        elementos.append(Paragraph(titulo, estilos['Heading2']))

        data = [[Paragraph(str(col), estilo_header) for col in df.columns]]
        for _, row in df.iterrows():
            fila = [Paragraph(str(cell), estilo_celda) for cell in row]
            data.append(fila)

        ancho_util = landscape(letter)[0] - (doc.leftMargin + doc.rightMargin)
        col_widths = _table_col_widths(df, ancho_util)

        tabla = Table(data, repeatRows=1, colWidths=col_widths)
        estilo_tabla = [
            ('BACKGROUND', (0, 0), (-1, 0), azul_rey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ]
        for i in range(1, len(data)):
            if i % 2 == 0:
                estilo_tabla.append(('BACKGROUND', (0, i), (-1, i), gris_zebra))
        tabla.setStyle(TableStyle(estilo_tabla))

        elementos.append(tabla)
        elementos.append(Spacer(1, 12))

    agregar_tabla("Indicadores (comparativo mostrado)", df_indicadores)
    agregar_tabla("Inscritos (conteo por carrera)", df_inscritos)
    agregar_tabla("Egresados (conteo por carrera)", df_egresados)

    def _footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        text = f"{cuatri_texto} — Página {doc.page}"
        canvas.drawRightString(landscape(letter)[0] - doc.rightMargin, 18, text)
        canvas.restoreState()

    doc.build(elementos, onFirstPage=_footer, onLaterPages=_footer)
    buffer.seek(0)
    return buffer.read()

# ===== DESCARGAS ===== #
colL, colR = st.columns([3, 2])
with colL:
    if 'comp_out' in locals() and not comp_out.empty:
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            comp_out.to_excel(writer, index=False, sheet_name='Comparativo')
            if not conteo_inscritos_por_carrera.empty:
                conteo_inscritos_por_carrera.to_excel(writer, index=False, sheet_name='Inscritos')
            if not conteo_egresados_por_carrera.empty:
                conteo_egresados_por_carrera.to_excel(writer, index=False, sheet_name='Egresados')
        excel_buffer.seek(0)
        st.download_button(
            "📊 Descargar Excel (comparativo y conteos)",
            data=excel_buffer,
            file_name="salida_comparativos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with colR:
    if 'comp_out' in locals() and not comp_out.empty \
        and not conteo_inscritos_por_carrera.empty \
        and not conteo_egresados_por_carrera.empty:
        st.subheader("🖨️ Generar reporte PDF completo")
        pdf_bytes = generar_reporte_pdf(comp_out, conteo_inscritos_por_carrera, conteo_egresados_por_carrera, cuatrimestre_actual)
        st.download_button(
            "📥 Descargar reporte PDF",
            data=pdf_bytes,
            file_name="reporte_general.pdf",
            mime="application/pdf",
        )
    else:
        st.info("Carga Indicadores, Inscritos y Egresados y genera el comparativo para habilitar las descargas.")