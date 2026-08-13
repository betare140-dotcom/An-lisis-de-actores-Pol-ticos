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

def parsear_fecha_todas_las_formas(val):
    if not val or pd.isna(val) or str(val) == "nan":
        return pd.NaT
    s = str(val).strip()
    if ',' in s:
        s = s.split(',')[0].strip()
    s_date = s.split(' ')[0].strip()
    
    if '-' in s_date:
        try:
            parts = s_date.split('-')
            if len(parts) == 3 and len(parts[0]) == 4:
                return datetime(int(parts[0]), int(parts), int(parts))
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

# PROCESAR UNA HOJA Y GENERAR DOCUMENTO WORD CON MARCO DE FRAMING
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

    # Mapeo y orden cronológico de fechas por objeto datetime
    serie_fechas_raw = obtener_columna_serie(df_filtrado, ["Publish date", "Fecha", "Date", "Fecha de publicación"])
    df_filtrado["fecha_dt"] = serie_fechas_raw.apply(parsear_fecha_todas_las_formas)
    df_filtrado = df_filtrado.dropna(subset=["fecha_dt"]).sort_values(by="fecha_dt", ascending=True)

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

    # Clasificación con IA usando tu marco de Framing
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

    # Conteo oficial
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
        neg_textos = str(df_data[df_data["sentimiento_ia"] == "NEG
