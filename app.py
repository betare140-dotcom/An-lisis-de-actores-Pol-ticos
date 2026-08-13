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
    page_title="Monitoreo Político Oficial", page_icon="📊", layout="centered"
)

st.title("📊 Generador de Monitoreo de Actores Políticos")
st.write(
    "Sube tu reporte para generar los informes oficiales en Word con clasificación de IA."
)

# Configuración de la API Key de Gemini
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

    raise Exception("No se pudo leer el archivo. Asegúrate de que sea un archivo Excel (.xlsx/.xls) o CSV válido.")

def obtener_campo(row, lista_cols):
    for c in lista_cols:
        if c in row.index:
            v = str(row[c]).strip()
            if v and v != "nan" and v != "None":
                return v
        for col_existente in row.index:
            if col_existente.strip().lower() == c.lower():
                v = str(row[col_existente]).strip()
                if v and v != "nan" and v != "None":
                    return v
    return ""

def obtener_columna_serie(df_data, lista_posibles_cols):
    for c in lista_posibles_cols:
        if c in df_data.columns:
            return df_data[c]
        for col_existente in df_data.columns:
            if col_existente.strip().lower() == c.lower():
                return df_data[col_existente]
    return pd.Series([""] * len(df_data), index=df_data.index)

# PARSER ESTRUCTURADO PARA GARANTIZAR DÍA, MES Y AÑO (DD.MM.YY)
def parsear_fecha_estricta_dia_mes_anio(val):
    if not val or pd.isna(val) or str(val) == "nan":
        return pd.NaT
    s = str(val).strip()
    if ',' in s:
        s = s.split(',')[0].strip()
    s_date = s.split(' ')[0].strip()
    
    # Caso YYYY-MM-DD
    if '-' in s_date:
        try:
            parts = s_date.split('-')
            if len(parts) == 3 and len(parts[0]) == 4:
                return datetime(int(parts[0]), int(parts), int(parts))
        except Exception:
            pass

    # Caso DD/MM/YYYY (fuerza a leer el primer número como DÍA)
    if '/' in s_date:
        try:
            parts = s_date.split('/')
            if len(parts) == 3:
                d, m, y = int(parts[0]), int(parts), int(parts)
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

# PROCESAR UNA HOJA Y GENERAR DOCUMENTO WORD
def crear_doc_desde_hoja(df_hoja, nombre_hoja, es_redes_sociales):
    pc_mask = df_hoja.apply(
        lambda r: (
            "pcgobpue" in str(r.get("Author handle (@username)", "")).lower()
            or "protección civil" in str(r.get("Author name", "")).lower()
            or "protección civil" in str(r.get("Source", "")).lower()
            or "protección civil" in str(r.get("Autor", "")).lower()
            or "protección civil" in str(r.get("Nombre del Medio", "")).lower()
        ),
        axis=1,
    )
    df_filtrado = df_hoja[~pc_mask].copy()

    if len(df_filtrado) == 0 or 'sin notas' in str(df_filtrado.iloc[0].values).lower():
        return None

    # Parseo estricto Día/Mes/Año y ordenamiento por objeto datetime
    serie_fechas_raw = obtener_columna_serie(df_filtrado, ["Publish date", "Fecha", "Date", "Fecha de publicación"])
    df_filtrado["fecha_dt"] = serie_fechas_raw.apply(parsear_fecha_estricta_dia_mes_anio)
    df_filtrado = df_filtrado.dropna(subset=["fecha_dt"]).sort_values(by="fecha_dt", ascending=True)

    if len(df_filtrado) == 0:
        return None

    # Formato strictly Día.Mes.Año (ej. 07.08.26 para 7 de agosto)
    df_filtrado["fecha_str"] = df_filtrado["fecha_dt"].dt.strftime("%d.%m.%2y")

    fechas_validas = df_filtrado["fecha_dt"]
    min_d = fechas_validas.min()
    max_d = fechas_validas.max()
    
    if min_d.strftime("%d.%m") == max_d.strftime("%d.%m"):
        periodo_texto = f"{min_d.strftime('%d')} de {MESES_ES[max_d.month]} de {max_d.year}"
    else:
        periodo_texto = f"{min_d.strftime('%d')} al {max_d.strftime('%d')} de {MESES_ES[max_d.month]} de {max_d.year}"

    # Clasificación con IA usando marco de Framing
    def clasificar_con_ia(row):
        detalle = obtener_campo(row, ["Contenido", "Detail", "Titulo", "Summary", "Síntesis", "Title"])
        prompt = (
            "Rol: Eres un analista experto en monitoreo de medios, comunicación política y relaciones públicas.\n"
            "Tu tarea es analizar el texto referente a '" + str(nombre_hoja) + "' para determinar su intencionalidad basándote en su encuadre (framing):\n\n"
            "- Positiva / Institucional: Notas que promueven la agenda, destacan inauguraciones, apoyos, programas sociales, buena imagen del actor político, respaldos o clima de gobernabilidad.\n"
            "- Neutra / Informativa: Cobertura estrictamente descriptiva, hechos sin sesgo, agendas del día, encuestas de posición o nota roja policial sin adjudicar culpa o responsabilidad directa al actor o administración.\n"
            "- Negativa / Crítica: Textos de crisis reputacional, escándalos de corrupción, ataques directos de opositores, columnas hostiles, protestas o acusaciones explícitas de inacción o mala gestión hacia el actor o administración.\n\n"
            "Texto a analizar: \"" + str(detalle) + "\"\n\n"
            "Responde ÚNICAMENTE con una palabra: POSITIVA, NEUTRA o NEGATIVA."
        )
        try:
            res = model.generate_content(prompt).text.strip().upper()
            if "NEGATIVA" in res:
                return "NEGATIVA"
            elif "NEUTRA" in res:
                return "NEUTRA"
            return "POSITIVA"
        except Exception:
            return "POSITIVA"

    df_filtrado["sentimiento_ia"] = df_filtrado.apply(clasificar_con_ia, axis=1)

    positivas_cnt = len(df_filtrado[df_filtrado["sentimiento_ia"].isin(["POSITIVA", "NEUTRA"])])
    negativas_cnt = len(df_filtrado[df_filtrado["sentimiento_ia"] == "NEGATIVA"])
    total_cnt = len(df_filtrado)

    serie_media = obtener_columna_serie(df_filtrado, ["Fuente", "Media type", "Media Type", "Medio", "Nombre del Medio", "Tipo de Medio", "Canal"])
    
    redes_cnt = total_cnt if es_redes_sociales else 0
    portales_cnt = serie_media.astype(str).str.contains("Portal|Web|Online|Internet", case=False, na=False).sum() if not es_redes_sociales else 0
    prensa_cnt = serie_media.astype(str).str.contains("Prensa|Diario|Periódico", case=False, na=False).sum() if not es_redes_sociales else 0
    columnas_cnt = serie_media.astype(str).str.contains("Columna|Opinión", case=False, na=False).sum() if not es_redes_sociales else 0

    def obtener_3_temas_positivos_y_negativos(df_data, actor_p):
        pos_textos = str(df_data[df_data["sentimiento_ia"].isin(["POSITIVA", "NEUTRA"])]["Contenido"].dropna().head(30).tolist()) if "Contenido" in df_data.columns else str(df_data[df_data["sentimiento_ia"].isin(["POSITIVA", "NEUTRA"])]["Titulo"].dropna().head(30).tolist())
        neg_textos = str(df_data[df_data["sentimiento_ia"] == "NEGATIVA"]["Contenido"].dropna().head(15).tolist()) if "Contenido" in df_data.columns else str(df_data[df_data["sentimiento_ia"] == "NEGATIVA"]["Titulo"].dropna().head(15).tolist())

        prompt = (
            "Eres un analista de comunicación política. Analiza las noticias sobre '" + str(actor_p) + "'.\n\n"
            "NOTICIAS POSITIVAS/INFORMATIVAS:\n" + pos_textos + "\n\n"
            "NOTICIAS NEGATIVAS/CRÍTICAS:\n" + neg_textos + "\n\n"
            "INSTRUCCIONES ESTRICTAS:\n"
            "1. NO incluyas frases estadísticas como 'Predominó la cobertura favorable...'.\n"
            "2. Redacta ESTRICTAMENTE 3 temas positivos o informativos principales numerados (1., 2., 3.).\n"
            "3. Redacta ESTRICTAMENTE 3 temas negativos principales numerados (1., 2., 3.). Si no existen noticias negativas, escribe simplemente: '1. No se registraron temas negativos en el periodo analizado.'\n\n"
            "FORMATO EXACTO DE SALIDA:\n"
            "1. [Primer tema positivo]\n"
            "2. [Segundo tema positivo]\n"
            "3. [Tercer tema positivo]\n\n"
            "TEMAS NEGATIVOS:\n"
            "1. [Primer tema negativo o 'No se registraron temas negativos en el periodo analizado.']"
        )
        try:
            return model.generate_content(prompt).text.strip()
        except Exception:
            return "1. Difusión de actividades públicas y agenda de trabajo.\n2. Presencia en medios informativos.\n3. Cobertura de posicionamientos.\n\nTEMAS NEGATIVOS:\n1. No se registraron temas negativos en el periodo analizado."

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

    headers = ["POSITIVA", "NEGATIVA", "TOTAL DE IMPACTOS"]
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
                add_run_verdana(p_m, "REDES SOCIALES", bold=True, size_pt=10)

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
                for m_type, grupo_m in pos_df.groupby(lambda i: obtener_campo(pos_df.loc[i], ["Tipo de Nota", "Tipo de Medio", "Fuente"]) or "PORTALES DIGITALES"):
                    p_m = doc.add_paragraph()
                    p_m.paragraph_format.space_before = Pt(4)
                    p_m.paragraph_format.space_after = Pt(4)
                    
                    heading_m = "PORTALES DIGITALES" if ("Común" in m_type or "Internet" in m_type) else m_type.upper()
                    add_run_verdana(p_m, heading_m, bold=True, size_pt=10)

                    for _, row in grupo_m.iterrows():
                        medio = obtener_campo(row, ["Nombre del Medio", "Fuente", "Media name"])
                        autor = obtener_campo(row, ["Autor", "Author name", "Programa"])
                        titulo = obtener_campo(row, ["Titulo", "Contenido", "Detail", "Summary"])
                        link = obtener_campo(row, ["Link de Nota", "URL", "Link"])

                        p_a = doc.add_paragraph()
                        p_a.paragraph_format.space_before = Pt(4)
                        p_a.paragraph_format.space_after = Pt(1)
                        cabecera = f"{medio} - {autor}" if (autor and autor != "Redacción" and autor != "Staff" and autor != "Online") else medio
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
            add_run_verdana(p_neg_hdr, "NEGATIVAS", bold=True, size_pt=10, color_rgb=RGBColor(180, 0, 0))

            for _, row in neg_df.iterrows():
                medio = obtener_campo(row, ["Nombre del Medio", "Fuente", "Media name"])
                autor = obtener_campo(row, ["Autor", "Author name", "Programa"])
                titulo = obtener_campo(row, ["Titulo", "Contenido", "Detail", "Summary"])
                link = obtener_campo(row, ["Link de Nota", "URL", "Link"])

                p_a = doc.add_paragraph()
                p_a.paragraph_format.space_before = Pt(4)
                p_a.paragraph_format.space_after = Pt(1)
                cabecera = f"{medio} - {autor}" if (autor and autor != "Redacción" and autor != "Staff" and autor != "Online") else medio
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
    ["Redes Sociales", "Medios Tradicionales / Portales (Prensa, TV, Radio, Portales)"],
    index=0
)

uploaded_file = st.file_uploader(
    "Sube tu archivo Excel o CSV de Monitoreo", type=["xlsx", "xls", "csv"]
)

if uploaded_file:
    try:
        dict_hojas = cargar_archivo_seguro(uploaded_file)
        hojas_disponibles = list(dict_hojas.keys())
        es_redes = (tipo_analisis == "Redes Sociales")
        
        st.subheader("Hojas / Secciones detectadas en el archivo:")
        st.write(", ".join(hojas_disponibles))

        modo = st.radio(
            "Selecciona la modalidad de descarga:",
            ["Generar una Hoja Específica", "Generar TODAS las Hojas en un archivo .ZIP (Masivo)"],
            index=0
        )

        if modo == "Generar una Hoja Específica":
            hoja_sel = st.selectbox("Selecciona la hoja a procesar:", hojas_disponibles)
            if st.button("Generar Reporte de esta Hoja", type="primary"):
                with st.spinner("Analizando publicaciones con IA..."):
                    df_h = dict_hojas[hoja_sel]
                    
                    if 'Menu' in df_h.columns and len(df_h['Menu'].dropna()) > 0:
                        nombre_h = str(df_h['Menu'].dropna().iloc[0]).strip()
                    else:
                        nombre_h = hoja_sel

                    buf = crear_doc_desde_hoja(df_h, nombre_h, es_redes)
                    if buf is not None:
                        st.success(f"¡Reporte generado exitosamente para '{nombre_h}'!")
                        st.download_button(
                            label=f"📥 Descargar Reporte Word de {nombre_h}",
                            data=buf,
                            file_name=f"Reporte_{nombre_h.replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    else:
                        st.warning(f"La hoja '{hoja_sel}' no contiene notas registradas para el día de hoy.")

        else:
            if st.button("Generar y Descargar TODAS las Hojas en .ZIP", type="primary"):
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

                            buf = crear_doc_desde_hoja(df_h, nombre_h, es_redes)
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
