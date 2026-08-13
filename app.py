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

# Diccionario de meses en español
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
    "con clasificación de IA."
)

# API Key Pre-integrada
GEMINI_API_KEY = "AQ.Ab8RN6LoOHgBblHSIETp2LjyBofO48YsSqSeojXYFAAKGvFa0w"
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")

# FUNCIÓN DE CARGA SEGURO EN CASCADA
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

    raise Exception(
        "No se pudo leer el archivo. "
        "Asegúrate de que sea un archivo Excel (.xlsx/.xls) "
        "o CSV válido."
    )

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

    # Caso DD/MM/YYYY (fuerza Día primero)
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
ROL:
Eres un analista experto en monitoreo de medios, comunicación política, reputación gubernamental y análisis de framing.

Tu tarea es analizar una noticia, columna o publicación de redes sociales sobre el actor político:
"{actor_nombre_target}"
y determinar su impacto reputacional.

============================================================
OBJETIVO PRINCIPAL
============================================================
La prioridad absoluta es detectar TODA publicación que pueda afectar negativamente la imagen, reputación, credibilidad, desempeño, gestión o percepción pública del actor político "{actor_nombre_target}".

============================================================
CATEGORÍAS
============================================================
POSITIVA / INFORMATIVA:
Clasifica como POSITIVA cuando la publicación presenta favorablemente al actor político, su gobierno, administración, acciones o resultados.

NEUTRA / INFORMATIVA:
Clasifica como NEUTRA cuando la publicación simplemente informa sobre un hecho y NO existe una afectación directa a la reputación del actor político.
- Reportes policiales o nota roja que solo relatan un suceso sin responsabilizar al actor = NEUTRA.

NEGATIVA / CRÍTICA:
Clasifica como NEGATIVA cuando el contenido afecta DIRECTAMENTE al actor político "{actor_nombre_target}", su gobierno, administración, gestión, equipo político o reputación.
Debe considerarse NEGATIVA cualquier contenido que:
- Critique directamente al actor político o lo señale como incompetente.
- Lo acuse de corrupción, nepotismo, favoritismo o falta de resultados.
- Denuncie despintado de bardas, clausura ciudadana o actos de campaña anticipada.
- Contenga ataques directos de opositores o columnas hostiles.

============================================================
TEXTO A ANALIZAR
============================================================
"{detalle}"

============================================================
RESPUESTA
============================================================
Responde ÚNICAMENTE con una palabra: POSITIVA, NEUTRA o NEGATIVA.
"""
    try:
        response = model.generate_content(prompt)
        resultado = response.text.strip().upper()

        if "NEGATIVA" in resultado:
            return "NEGATIVA"
        elif "NEUTRA" in resultado:
            return "NEUTRA"
        elif "POSITIVA" in resultado:
            return "POSITIVA"
        return "NEUTRA"
    except Exception:
        return "NEUTRA"

def obtener_3_temas_positivos_y_negativos(df_data, actor_p):
    pos_textos = str(
        df_data[df_data["sentimiento_ia"].isin(["POSITIVA", "NEUTRA"])]["Contenido"].dropna().head(30).tolist()
    ) if "Contenido" in df_data.columns else str(
        df_data[df_data["sentimiento_ia"].isin(["POSITIVA", "NEUTRA"])]["Titulo"].dropna().head(30).tolist()
    )

    neg_textos = str(
        df_data[df_data["sentimiento_ia"] == "NEGATIVA"]["Contenido"].dropna().head(15).tolist()
    ) if "Contenido" in df_data.columns else str(
        df_data[df_data["sentimiento_ia"] == "NEGATIVA"]["Titulo"].dropna().head(15).tolist()
    )

    prompt = f"""
Eres un analista de comunicación política. Analiza las noticias sobre "{actor_p}".

NOTICIAS POSITIVAS / INFORMATIVAS:
{pos_textos}

NOTICIAS NEGATIVAS / CRÍTICAS:
{neg_textos}

INSTRUCCIONES:
1. NO incluyas frases estadísticas como "Predominó la cobertura favorable...".
2. Redacta exactamente 3 temas positivos o informativos principales.
3. Redacta exactamente 3 temas negativos principales. Si no existen noticias negativas, escribe: "1. No se registraron temas negativos en el periodo analizado."

FORMATO:
1. [Primer tema positivo]
2. [Segundo tema positivo]
3. [Tercer tema positivo]

TEMAS NEGATIVOS:
1. [Primer tema negativo o 'No se registraron temas negativos en el periodo analizado.']
"""
    try:
        return model.generate_content(prompt).text.strip()
    except Exception:
        return (
            "1. Difusión de actividades públicas y agenda de trabajo.\n"
            "2. Presencia en medios informativos.\n"
            "3. Cobertura de posicionamientos.\n\n"
            "TEMAS NEGATIVOS:\n"
            "1. No se registraron temas negativos en el periodo analizado."
        )

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

    sentimientos = []
    progreso = st.progress(0)
    total_filas = len(df_filtrado)

    for i, (_, row) in enumerate(df_filtrado.iterrows()):
        sentimiento = clasificar_con_ia(row, nombre_hoja)
        sentimientos.append(sentimiento)
        progreso.progress((i + 1) / total_filas)

    df_filtrado["sentimiento_ia"] = sentimientos
    progreso.empty()

    positivas_cnt = len(df_filtrado[df_filtrado["sentimiento_ia"].isin(["POSITIVA", "NEUTRA"])])
    negativas_cnt = len(df_filtrado[df_filtrado["sentimiento_ia"] == "NEGATIVA"])
    total_cnt = len(df_filtrado)

    serie_media = obtener_columna_serie(
        df_filtrado,
        ["Fuente", "Media type", "Media Type", "Medio", "Nombre del Medio", "Tipo de Medio", "Canal"]
    )

    redes_cnt = total_cnt if es_redes_sociales else 0
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
        if color_rgb: run.font.color.rgb = color_rgb
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
        texto_totales = f"TOTAL NOTAS INFORMATIVAS: {total_cnt}\nPORTALES DIGITALES: {portales_cnt + redes_cnt}\nREDES SOCIALES: 0\nPRENSA LOCAL: {prensa_cnt}\nCOLUMNAS: {columnas_cnt}"

    add_run_verdana(p_tot, texto_totales, bold=True, size_pt=10)

    # 3. Resumen
    p_res = doc.add_paragraph()
    p_res.paragraph_format.space_before = Pt(10)
    add_run_verdana(p_res, "RESUMEN", bold=True, size_pt=11)

    p_temas = doc.add_paragraph()
    p_temas.paragraph_format.space_after = Pt(4)
    add_run_verdana(p_temas, "Temas relevantes informativos", bold=True, size_pt=10)

    temas_3x3 = obtener_3_temas_positivos_y_negativos(df_filtrado, nombre_hoja)
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

        pos_df = sub_df[sub_df["sentimiento_ia"].isin(["POSITIVA", "NEUTRA"])]
        neg_df = sub_df[sub_df["sentimiento_ia"] == "NEGATIVA"]

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
                    link = obtener_campo(row, ["URL", "Link de Nota", "Link", "Enlace"])

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
                grupos = pos_df.groupby(lambda i: obtener_campo(pos_df.loc[i], ["Tipo de Nota", "Tipo de Medio", "Fuente"]) or "PORTALES DIGITALES")
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
                        link = obtener_campo(row, ["Link de Nota", "URL", "Link"])

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
                link = obtener_campo(row, ["Link de Nota", "URL", "Link"])

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
        "Medios Tradicionales / Portales (Prensa, TV, Radio, Portales)"
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
    # Medios Tradicionales / Portales (Multiple Hojas Hanakuá)
    uploaded_file = st.file_uploader(
        "Sube tu archivo Excel de Medios Tradicionales (con múltiples candidatos/hojas)",
        type=["xlsx", "xls"]
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
                    with st.spinner("Analizando publicaciones con IA..."):
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
                    with st.spinner("Generando reportes individuales para todas las hojas con IA..."):
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
                                file_name="Reportes_Monitoreo_Completos.zip",
                                mime="application/zip"
                            )
                        else:
                            st.warning("No se encontraron hojas con notas activas para generar los reportes.")

        except Exception as e:
            st.error(f"Error procesando el archivo: {str(e)}")
