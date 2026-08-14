import io
import os
import re
import zipfile
from datetime import datetime
import google.generativeai as genai
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor
import pandas as pd
import streamlit as st

MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}

st.set_page_config(
    page_title="Monitoreo Político Oficial",
    page_icon="📊",
    layout="centered"
)

st.title("📊 Generador de Monitoreo de Actores Políticos")
st.write(
    "Sube tu reporte para generar los informes oficiales en Word "
    "con extracción automática de datos y temas reales."
)

# API Key Pre-integrada
GEMINI_API_KEY = "AQ.Ab8RN6LoOHgBblHSIETp2LjyBofO48YsSqSeojXYFAAKGvFa0w"
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")

def cargar_archivo_seguro(file):
    try:
        return pd.read_excel(file, sheet_name=None)
    except Exception:
        pass

    try:
        file.seek(0)
        df_csv = pd.read_csv(file, on_bad_lines="skip")
        return {"Reporte": df_csv}
    except Exception:
        pass

    try:
        file.seek(0)
        df_csv = pd.read_csv(file, encoding="latin1", on_bad_lines="skip")
        return {"Reporte": df_csv}
    except Exception:
        pass

    raise Exception("No se pudo leer el archivo. Asegúrate de que sea un archivo Excel (.xlsx/.xls) o CSV válido.")

def obtener_campo(row, lista_cols):
    for c in lista_cols:
        if c in row.index:
            v = str(row[c]).strip()
            if v and v != "nan" and v != "None":
                return v
        for col_existente in row.index:
            if str(col_existente).strip().lower() == c.lower():
                v = str(row[col_existente]).strip()
                if v and v != "nan" and v != "None":
                    return v
    return ""

def obtener_columna_serie(df_data, lista_posibles_cols):
    for c in lista_posibles_cols:
        if c in df_data.columns:
            return df_data[c]
        for col_existente in df_data.columns:
            if str(col_existente).strip().lower() == c.lower():
                return df_data[col_existente]
    return pd.Series([""] * len(df_data), index=df_data.index)

def parsear_fecha_perfecta(val):
    if not val or pd.isna(val) or str(val) == "nan":
        return pd.NaT

    s = str(val).strip()
    if "," in s:
        s = s.split(",")[0].strip()
    s_date = s.split(" ")[0].strip()

    # Caso YYYY-MM-DD
    if "-" in s_date:
        try:
            parts = s_date.split("-")
            if len(parts) == 3 and len(parts[0]) == 4:
                return datetime(int(parts[0]), int(parts), int(parts))
        except Exception:
            pass

    # Caso DD/MM/YYYY
    if "/" in s_date:
        try:
            parts = s_date.split("/")
            if len(parts) == 3:
                if len(parts[0]) == 4:
                    return datetime(int(parts[0]), int(parts), int(parts))
                else:
                    d = int(parts[0])
                    m = int(parts)
                    y = int(parts)
                    if y < 100: y += 2000
                    return datetime(y, m, d)
        except Exception:
            pass

    try:
        return pd.to_datetime(s_date, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT

def limpiar_texto(texto):
    if not isinstance(texto, str) or texto == "nan":
        return ""
    texto_sin_urls = re.sub(r"https?://\S+", "", texto)
    texto_sin_emojis = re.sub(r":[a-zA-Z0-9_\-|]+:", "", texto_sin_urls)
    texto_sin_emojis = re.sub(r'[\U00010000-\U0010ffff]', '', texto_sin_emojis)
    lineas = [
        re.sub(r"[ \t]+", " ", line).strip()
        for line in texto_sin_emojis.split("\n")
        if line.strip()
    ]
    return "\n".join(lineas)

def clasificar_con_ia(row, actor_nombre_target):
    detalle = obtener_campo(
        row,
        ["Contenido", "Detail", "Titulo", "Summary", "Síntesis", "Title"]
    )

    if not detalle:
        return "NEUTRA"

    prompt = f"""
ROL: Analista de comunicación política.
Determina el sentimiento para el actor "{actor_nombre_target}".
Categorías: POSITIVA, NEUTRA o NEGATIVA.

TEXTO:
"{detalle}"

Responde ÚNICAMENTE: POSITIVA, NEUTRA o NEGATIVA.
"""
    try:
        res = model.generate_content(prompt).text.strip().upper()
        if "NEGATIVA" in res:
            return "NEGATIVA"
        elif "NEUTRA" in res:
            return "NEUTRA"
        elif "POSITIVA" in res:
            return "POSITIVA"
        return "NEUTRA"
    except Exception:
        return "NEUTRA"

def determinar_sentimiento(row, actor_nombre_target, es_tradicionales):
    if es_tradicionales:
        sent_val = obtener_campo(row, ["Sentimiento", "Sentiment", "Sentimiento de la Nota"]).lower()
        if sent_val and sent_val != "no sentiment":
            if "negat" in sent_val:
                return "NEGATIVA"
            elif "positi" in sent_val or "neutr" in sent_val:
                return "POSITIVA"

    return clasificar_con_ia(row, actor_nombre_target)

def extraer_resumen_temas_real(df_data, actor_nombre):
    pos_df = df_data[df_data["sentimiento_final"].isin(["POSITIVA", "NEUTRA"])]
    neg_df = df_data[df_data["sentimiento_final"] == "NEGATIVA"]

    pos_textos = []
    for col in ["Titulo", "Contenido", "Detail", "Summary"]:
        if col in pos_df.columns:
            pos_textos = pos_df[col].dropna().astype(str).tolist()
            break

    neg_textos = []
    for col in ["Titulo", "Contenido", "Detail", "Summary"]:
        if col in neg_df.columns:
            neg_textos = [t for t in neg_df[col].dropna().astype(str).tolist() if "teaser" not in t.lower()]
            break

    # 1. Intentar con IA validando que no devuelva falsos "no registrados"
    if len(pos_textos) > 0 or len(neg_textos) > 0:
        prompt = f"""
Eres un analista de comunicación política.
Redacta un RESUMEN EJECUTIVO para "{actor_nombre}".
Debes basarte ESTRICTAMENTE en estos hechos del archivo:

NOTICIAS POSITIVAS ({len(pos_textos)} notas):
{str(pos_textos[:20])}

NOTICIAS NEGATIVAS ({len(neg_textos)} notas):
{str(neg_textos[:15])}

REGLAS:
1. Puntos Positivos: De 1 a 3 temas basados en las notas positivas.
2. Puntos Negativos:
   - Si hay {len(neg_textos)} notas negativas, redacta de 1 a 3 temas NEGATIVOS explicando las controversias reales.
   - Si NO hay notas negativas (0 notas negativas), pon exactamente:
     1. No se registraron temas negativos en el periodo analizado.
3. NO incluyas frases estadísticas.

FORMATO:
1. [Tema positivo 1]
2. [Tema positivo 2]

TEMAS NEGATIVOS:
1. [Tema negativo 1 o leyenda si no hay]
"""
        try:
            res_ia = model.generate_content(prompt).text.strip()
            if len(neg_textos) > 0 and "no se registraron temas negativos" in res_ia.lower():
                pass
            elif res_ia and len(res_ia) > 25:
                return res_ia
        except Exception:
            pass

    # 2. Extractor Directo de las Notas Reales (Sin texto genérico)
    lineas_res = []
    if len(pos_textos) > 0:
        titulos_unicos_pos = list(dict.fromkeys([
            re.sub(r'^\d+:\d+\s*(hrs|am|pm|\.)\s*', '', t, flags=re.I).strip() 
            for t in pos_textos if t.strip() and "teaser" not in t.lower()
        ]))
        if not titulos_unicos_pos:
            titulos_unicos_pos = [pos_textos[0]]
        for i, t in enumerate(titulos_unicos_pos[:3], 1):
            lineas_res.append(f"{i}. {t}")
    else:
        lineas_res.append("1. Difusión de actividades y agenda institucional de trabajo.")

    lineas_res.append("\nTEMAS NEGATIVOS:")

    if len(neg_textos) > 0:
        titulos_unicos_neg = list(dict.fromkeys([
            re.sub(r'^\d+:\d+\s*(hrs|am|pm|\.)\s*', '', t, flags=re.I).strip() 
            for t in neg_textos if t.strip() and "teaser" not in t.lower()
        ]))
        if not titulos_unicos_neg:
            titulos_unicos_neg = [neg_textos[0]]
        for i, t in enumerate(titulos_unicos_neg[:3], 1):
            lineas_res.append(f"{i}. {t}")
    else:
        lineas_res.append("1. No se registraron temas negativos en el periodo analizado.")

    return "\n".join(lineas_res)

def crear_doc_desde_hoja(df_hoja, nombre_hoja, es_redes_sociales):
    pc_mask = df_hoja.apply(
        lambda r: (
            "pcgobpue" in str(r.get("Author handle (@username)", "")).lower()
            or "protección civil" in str(r.get("Author name", "")).lower()
            or "protección civil" in str(r.get("Source", "")).lower()
            or "protección civil" in str(r.get("Autor", "")).lower()
            or "protección civil" in str(r.get("Nombre del Medio", "")).lower()
        ),
        axis=1
    )
    df_filtrado = df_hoja[~pc_mask].copy()

    if len(df_filtrado) == 0:
        return None

    if "sin notas" in str(df_filtrado.iloc[0].values).lower():
        return None

    serie_fechas_raw = obtener_columna_serie(
        df_filtrado, ["Publish date", "Fecha", "Date", "Fecha de publicación"]
    )
    df_filtrado["fecha_dt"] = serie_fechas_raw.apply(parsear_fecha_perfecta)

    df_filtrado = (
        df_filtrado
        .dropna(subset=["fecha_dt"])
        .sort_values(by="fecha_dt", ascending=True)
    )

    if len(df_filtrado) == 0:
        return None

    df_filtrado["fecha_str"] = df_filtrado["fecha_dt"].dt.strftime("%d.%m.%2y")

    fechas_validas = df_filtrado["fecha_dt"]
    min_d = fechas_validas.min()
    max_d = fechas_validas.max()

    if min_d.strftime("%d.%m") == max_d.strftime("%d.%m"):
        periodo_texto = f"{min_d.strftime('%d')} de {MESES_ES[max_d.month]} de {max_d.year}"
    else:
        periodo_texto = f"{min_d.strftime('%d')} al {max_d.strftime('%d')} de {MESES_ES[max_d.month]} de {max_d.year}"

    # Asignar sentimientos
    sentimientos = []
    progreso = st.progress(0)
    total_filas = len(df_filtrado)

    for i, (_, row) in enumerate(df_filtrado.iterrows()):
        s_val = determinar_sentimiento(row, nombre_hoja, es_tradicionales=not es_redes_sociales)
        sentimientos.append(s_val)
        progreso.progress((i + 1) / total_filas)

    df_filtrado["sentimiento_final"] = sentimientos
    progreso.empty()

    positivas_cnt = len(df_filtrado[df_filtrado["sentimiento_final"].isin(["POSITIVA", "NEUTRA"])])
    negativas_cnt = len(df_filtrado[df_filtrado["sentimiento_final"] == "NEGATIVA"])
    total_cnt = len(df_filtrado)

    serie_media = obtener_columna_serie(
        df_filtrado,
        ["Tipo de Medio", "Fuente", "Media type", "Media Type", "Medio", "Nombre del Medio", "Canal"]
    )

    redes_cnt = total_cnt if es_redes_sociales else 0
    tv_cnt = serie_media.astype(str).str.contains("Tele", case=False, na=False).sum() if not es_redes_sociales else 0
    rad_cnt = serie_media.astype(str).str.contains("Rad", case=False, na=False).sum() if not es_redes_sociales else 0
    portales_cnt = serie_media.astype(str).str.contains("Portal|Web|Online|Internet", case=False, na=False).sum() if not es_redes_sociales else 0
    prensa_cnt = serie_media.astype(str).str.contains("Prensa|Diario|Periódico", case=False, na=False).sum() if not es_redes_sociales else 0
    columnas_cnt = serie_media.astype(str).str.contains("Columna|Opinión", case=False, na=False).sum() if not es_redes_sociales else 0

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    doc.styles["Normal"].font.name = "Verdana"
    doc.styles["Normal"].font.size = Pt(10)

    def add_run_verdana(p, text, bold=False, italic=False, size_pt=10, color_rgb=None, underline=False):
        run = p.add_run(text)
        run.font.name = "Verdana"
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(size_pt)
        run.underline = underline
        if color_rgb:
            run.font.color.rgb = color_rgb
        return run

    def fondo_celda(cell, fill_hex):
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

    # 1. Encabezado
    p_title = doc.add_paragraph()
    add_run_verdana(p_title, nombre_hoja.upper(), bold=True, size_pt=12, color_rgb=RGBColor(0, 51, 102))

    p_per = doc.add_paragraph()
    add_run_verdana(p_per, f"PERIODO DE MEDICIÓN: {periodo_texto}", bold=True, size_pt=10)

    p_can = doc.add_paragraph()
    p_can.paragraph_format.space_after = Pt(8)
    add_run_verdana(p_can, "CANALES: PRENSA, TV, RADIO, PORTALES, REDES SOCIALES Y COLUMNAS.", bold=True, size_pt=9.5)

    # 2. Balance de Impactos
    p_bal = doc.add_paragraph()
    add_run_verdana(p_bal, "BALANCE DE IMPACTOS", bold=True, size_pt=10.5)

    t_imp = doc.add_table(rows=2, cols=3)
    t_imp.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = ["POSITIVA / INFORMATIVA", "NEGATIVA", "TOTAL DE IMPACTOS"]
    fills = ["E2EFDA", "FCE4D6", "D9E1F2"]

    for col_idx, (h_text, fill_color) in enumerate(zip(headers, fills)):
        cell = t_imp.cell(0, col_idx)
        fondo_celda(cell, fill_color)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run_verdana(p, h_text, bold=True, size_pt=9.5)

    val_counts = [str(positivas_cnt), str(negativas_cnt), str(total_cnt)]
    for col_idx, val_text in enumerate(val_counts):
        cell = t_imp.cell(1, col_idx)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run_verdana(p, val_text, bold=True, size_pt=11)

    p_tot = doc.add_paragraph()
    p_tot.paragraph_format.space_before = Pt(10)
    p_tot.paragraph_format.space_after = Pt(10)

    if es_redes_sociales:
        texto_totales = f"TOTAL NOTAS INFORMATIVAS: {total_cnt}\nREDES SOCIALES: {total_cnt}\nPORTALES DIGITALES: 0\nPRENSA LOCAL: 0\nCOLUMNAS: 0"
    else:
        partes = [f"TOTAL NOTAS INFORMATIVAS: {total_cnt}"]
        if tv_cnt > 0: partes.append(f"TELEVISIÓN: {tv_cnt}")
        if rad_cnt > 0: partes.append(f"RADIO: {rad_cnt}")
        if portales_cnt > 0: partes.append(f"PORTALES DIGITALES: {portales_cnt}")
        if prensa_cnt > 0: partes.append(f"PRENSA LOCAL: {prensa_cnt}")
        if columnas_cnt > 0: partes.append(f"COLUMNAS: {columnas_cnt}")
        if len(partes) == 1:
            partes.extend(["PORTALES DIGITALES: 0", "TELEVISIÓN: 0", "RADIO: 0", "PRENSA LOCAL: 0", "COLUMNAS: 0"])
        texto_totales = "\n".join(partes)

    add_run_verdana(p_tot, texto_totales, bold=True, size_pt=10)

    # 3. Resumen
    p_res = doc.add_paragraph()
    p_res.paragraph_format.space_before = Pt(10)
    add_run_verdana(p_res, "RESUMEN", bold=True, size_pt=11)

    p_temas = doc.add_paragraph()
    p_temas.paragraph_format.space_after = Pt(4)
    add_run_verdana(p_temas, "Temas relevantes informativos", bold=True, size_pt=10)

    temas_3x3 = extraer_resumen_temas_real(df_filtrado, nombre_hoja)
    for linea in temas_3x3.split("\n"):
        linea_clean = linea.strip()
        if linea_clean:
            p_t = doc.add_paragraph()
            p_t.paragraph_format.space_before = Pt(1)
            p_t.paragraph_format.space_after = Pt(2)
            if "NEGATIVO" in linea_clean.upper():
                add_run_verdana(p_t, linea_clean, bold=True, size_pt=10, color_rgb=RGBColor(180, 0, 0))
            else:
                add_run_verdana(p_t, linea_clean, size_pt=9.5)

    # 4. Desglose
    p_des = doc.add_paragraph()
    p_des.paragraph_format.space_before = Pt(12)
    add_run_verdana(p_des, "DESGLOSE", bold=True, size_pt=11)

    for fecha_dt_val, sub_df in df_filtrado.groupby("fecha_dt", sort=True):
        fecha_item = fecha_dt_val.strftime("%d.%m.%2y")

        p_f = doc.add_paragraph()
        p_f.paragraph_format.space_before = Pt(10)
        p_f.paragraph_format.space_after = Pt(2)
        add_run_verdana(p_f, fecha_item, bold=True, size_pt=10.5, color_rgb=RGBColor(0, 51, 102))

        pos_df = sub_df[sub_df["sentimiento_final"].isin(["POSITIVA", "NEUTRA"])]
        neg_df = sub_df[sub_df["sentimiento_final"] == "NEGATIVA"]

        if es_redes_sociales:
            if len(pos_df) > 0:
                p_m = doc.add_paragraph()
                p_m.paragraph_format.space_before = Pt(4)
                p_m.paragraph_format.space_after = Pt(4)
                add_run_verdana(p_m, f"REDES SOCIALES: {len(pos_df)}", bold=True, size_pt=10)

                for _, row in pos_df.iterrows():
                    autor = obtener_campo(row, ["Autor", "Author name", "Fuente", "Media name", "Programa"])
                    handle = obtener_campo(row, ["Author handle (@username)", "Handle", "Username"])
                    detalle = obtener_campo(row, ["Contenido", "Detail", "Summary", "Síntesis", "Titulo", "Title"])
                    link = obtener_campo(row, ["Link URL Medio", "Link a Testigo", "Link de Nota", "URL", "Link", "Enlace"])

                    p_a = doc.add_paragraph()
                    p_a.paragraph_format.space_before = Pt(4)
                    p_a.paragraph_format.space_after = Pt(1)
                    if handle and not handle.startswith("@"): handle = f"@{handle}"
                    add_run_verdana(p_a, f"{autor} {handle}".strip() if handle else autor, bold=True, size_pt=10)

                    p_d = doc.add_paragraph()
                    p_d.paragraph_format.space_after = Pt(2)
                    add_run_verdana(p_d, limpiar_texto(detalle), bold=False, size_pt=9.5)

                    if link:
                        p_l = doc.add_paragraph()
                        p_l.paragraph_format.space_after = Pt(6)
                        add_run_verdana(p_l, link, bold=False, size_pt=9, color_rgb=RGBColor(0, 102, 204), underline=True)

        else:
            if len(pos_df) > 0:
                grupos = pos_df.groupby(lambda i: obtener_campo(pos_df.loc[i], ["Tipo de Medio", "Tipo de Nota", "Fuente"]) or "MEDIOS")
                for m_type, grupo_m in grupos:
                    p_m = doc.add_paragraph()
                    p_m.paragraph_format.space_before = Pt(4)
                    p_m.paragraph_format.space_after = Pt(4)
                    heading_base = "PORTALES DIGITALES" if ("Común" in m_type or "Internet" in m_type) else m_type.upper()
                    add_run_verdana(p_m, f"{heading_base}: {len(grupo_m)}", bold=True, size_pt=10)

                    for _, row in grupo_m.iterrows():
                        medio = obtener_campo(row, ["Nombre del Medio", "Fuente", "Media name"])
                        autor = obtener_campo(row, ["Autor", "Author name", "Programa"])
                        titulo = obtener_campo(row, ["Titulo", "Contenido", "Detail", "Summary"])
                        link = obtener_campo(row, ["Link URL Medio", "Link a Testigo", "Link de Nota", "URL", "Link", "Enlace"])

                        p_a = doc.add_paragraph()
                        p_a.paragraph_format.space_before = Pt(4)
                        p_a.paragraph_format.space_after = Pt(1)
                        cabecera = f"{medio} - {autor}" if (autor and autor not in ["Redacción", "Staff", "Online"]) else medio
                        add_run_verdana(p_a, cabecera, bold=True, size_pt=10)

                        p_d = doc.add_paragraph()
                        p_d.paragraph_format.space_after = Pt(2)
                        add_run_verdana(p_d, limpiar_texto(titulo), bold=False, size_pt=9.5)

                        if link:
                            p_l = doc.add_paragraph()
                            p_l.paragraph_format.space_after = Pt(6)
                            add_run_verdana(p_l, link, bold=False, size_pt=9, color_rgb=RGBColor(0, 102, 204), underline=True)

        if len(neg_df) > 0:
            p_neg_hdr = doc.add_paragraph()
            p_neg_hdr.paragraph_format.space_before = Pt(6)
            p_neg_hdr.paragraph_format.space_after = Pt(4)
            add_run_verdana(p_neg_hdr, f"NEGATIVAS: {len(neg_df)}", bold=True, size_pt=10, color_rgb=RGBColor(180, 0, 0))

            for _, row in neg_df.iterrows():
                medio = obtener_campo(row, ["Nombre del Medio", "Fuente", "Media name"])
                autor = obtener_campo(row, ["Autor", "Author name", "Programa"])
                titulo = obtener_campo(row, ["Titulo", "Contenido", "Detail", "Summary"])
                link = obtener_campo(row, ["Link URL Medio", "Link a Testigo", "Link de Nota", "URL", "Link", "Enlace"])

                p_a = doc.add_paragraph()
                p_a.paragraph_format.space_before = Pt(4)
                p_a.paragraph_format.space_after = Pt(1)
                cabecera = f"{medio} - {autor}" if (autor and autor not in ["Redacción", "Staff", "Online"]) else medio
                add_run_verdana(p_a, cabecera, bold=True, size_pt=10)

                p_d = doc.add_paragraph()
                p_d.paragraph_format.space_after = Pt(2)
                add_run_verdana(p_d, limpiar_texto(titulo), bold=False, size_pt=9.5)

                if link:
                    p_l = doc.add_paragraph()
                    p_l.paragraph_format.space_after = Pt(6)
                    add_run_verdana(p_l, link, bold=False, size_pt=9, color_rgb=RGBColor(0, 102, 204), underline=True)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ STREAMLIT ---

tipo_analisis = st.radio(
    "¿Qué tipo de archivo vas a analizar?",
    [
        "Redes Sociales",
        "Medios Tradicionales / Portales / TV y Radio"
    ],
    index=0
)

if tipo_analisis == "Redes Sociales":
    uploaded_file = st.file_uploader(
        "Sube tu archivo Excel o CSV de Redes Sociales",
        type=["xlsx", "xls", "csv"]
    )
    actor_nombre_in = st.text_input(
        "Nombre y Partido del Actor Político",
        placeholder="ej. LAURA ARTEMISA GARCÍA CHÁVEZ (MORENA)",
    ).strip().upper()

    if uploaded_file and actor_nombre_in:
        if st.button("Generar Reporte Oficial", type="primary"):
            with st.spinner("Analizando publicaciones con IA..."):
                try:
                    dict_h = cargar_archivo_seguro(uploaded_file)
                    df_redes = list(dict_h.values())[0]
                    buf = crear_doc_desde_hoja(df_redes, actor_nombre_in, es_redes_sociales=True)
                    if buf is not None:
                        st.success(f"¡Reporte generado exitosamente para '{actor_nombre_in}'!")
                        st.download_button(
                            label=f"📥 Descargar Reporte Word de {actor_nombre_in}",
                            data=buf,
                            file_name=f"Reporte_{actor_nombre_in.replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    else:
                        st.warning("El archivo no contiene notas válidas registradas.")
                except Exception as e:
                    st.error(f"Error procesando el archivo: {str(e)}")

else:
    uploaded_file = st.file_uploader(
        "Sube tu archivo Excel de Medios Tradicionales (TV, Radio, Portales o Prensa)",
        type=["xlsx", "xls", "csv"]
    )

    if uploaded_file:
        try:
            dict_hojas = cargar_archivo_seguro(uploaded_file)
            hojas_disponibles = list(dict_hojas.keys())
            
            st.subheader("Hojas / Candidatos detectados en el archivo:")
            st.write(", ".join(hojas_disponibles))

            modo = st.radio(
                "Selecciona la modalidad de descarga:",
                ["Generar un Candidato Específico", "Generar TODOS los Candidatos en un archivo .ZIP (Masivo)"],
                index=0
            )

            if modo == "Generar un Candidato Específico":
                hoja_sel = st.selectbox("Selecciona la hoja a procesar:", hojas_disponibles)
                if st.button("Generar Reporte de esta Hoja", type="primary"):
                    with st.spinner("Procesando notas y generando documento..."):
                        df_h = dict_hojas[hoja_sel]
                        if 'Menu' in df_h.columns and len(df_h['Menu'].dropna()) > 0:
                            nombre_h = str(df_h['Menu'].dropna().iloc[0]).strip()
                        else:
                            nombre_h = hoja_sel

                        buf = crear_doc_desde_hoja(df_h, nombre_h, es_redes_sociales=False)
                        if buf is not None:
                            st.success(f"¡Reporte generado exitosamente para '{nombre_h}'!")
                            st.download_button(
                                label=f"📥 Descargar Reporte Word de {nombre_h}",
                                data=buf,
                                file_name=f"Reporte_{nombre_h.replace(' ', '_')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        else:
                            st.warning(f"La hoja '{hoja_sel}' no contiene notas registradas.")

            else:
                if st.button("Generar y Descargar TODOS los Reportes en .ZIP", type="primary"):
                    with st.spinner("Generando reportes individuales para todas las hojas..."):
                        zip_buffer = io.BytesIO()
                        cnt_generados = 0
                        
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for h_key in hojas_disponibles:
                                df_h = dict_hojas[h_key]
                                if 'Menu' in df_h.columns and len(df_h['Menu'].dropna()) > 0:
                                    nombre_h = str(df_h['Menu'].dropna().iloc[0]).strip()
                                else:
                                    nombre_h = h_key

                                buf = crear_doc_desde_hoja(df_h, nombre_h, es_redes_sociales=False)
                                if buf is not None:
                                    doc_bytes = buf.getvalue()
                                    fname = f"Reporte_{nombre_h.replace(' ', '_')}.docx"
                                    zip_file.writestr(fname, doc_bytes)
                                    cnt_generados += 1

                        zip_buffer.seek(0)
                        if cnt_generados > 0:
                            st.success(f"¡Se generaron con éxito {cnt_generados} reportes individuales!")
                            st.download_button(
                                label="📦 Descargar Archivo .ZIP con TODOS los Reportes",
                                data=zip_buffer,
                                file_name="Reportes_Tradicionales_Completos.zip",
                                mime="application/zip"
                            )
                        else:
                            st.warning("No se encontraron hojas con notas activas para generar los reportes.")

        except Exception as e:
            st.error(f"Error procesando el archivo: {str(e)}")
