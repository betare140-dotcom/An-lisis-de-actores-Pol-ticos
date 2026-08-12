import io
import os
import re
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

# Configuración de la página
st.set_page_config(
    page_title="Monitoreo Político Oficial", page_icon="📊", layout="centered"
)

st.title("📊 Generador de Monitoreo de Actores Políticos")
st.write(
    "Sube tu reporte de Onclusive para generar el informe oficial en Word con"
    " clasificación de IA."
)

# Configuración de la API Key
GEMINI_API_KEY = "AQ.Ab8RN6LoOHgBblHSIETp2LjyBofO48YsSqSeojXYFAAKGvFa0w"
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")

# Interfaz de usuario para carga de archivos
uploaded_file = st.file_uploader(
    "Sube tu archivo Excel o CSV de Onclusive", type=["xlsx", "xls", "csv"]
)
actor_nombre = st.text_input(
    "Nombre y Partido del Actor Político",
    placeholder="ej. FEDRHA SURIANO (MOVIMIENTO CIUDADANO)",
).strip().upper()

if uploaded_file and actor_nombre:
    if st.button("Generar Reporte Oficial", type="primary"):
        with st.spinner("Analizando publicaciones con Inteligencia Artificial..."):
            # Leer archivo según extensión
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            elif uploaded_file.name.endswith(".xls"):
                df = pd.read_excel(uploaded_file, engine="xlrd")
            else:
                df = pd.read_excel(uploaded_file)

            # Filtrar Protección Civil
            pc_mask = df.apply(
                lambda r: (
                    "pcgobpue" in str(r.get("Author handle (@username)", "")).lower()
                    or "protección civil" in str(r.get("Author name", "")).lower()
                    or "protección civil" in str(r.get("Source", "")).lower()
                ),
                axis=1,
            )
            df_filtrado = df[~pc_mask].copy()

            # Detección automática de columna de fecha
            posibles_cols_fecha = ["Publish date", "Date", "Fecha", "Fecha de publicación"]
            col_fecha = None
            for c in posibles_cols_fecha:
                if c in df_filtrado.columns:
                    col_fecha = c
                    break

            if col_fecha is None:
                st.error(f"No se encontró una columna de fecha válida. Columnas: {list(df_filtrado.columns)}")
                st.stop()

            df_filtrado["fecha_dt"] = pd.to_datetime(
                df_filtrado[col_fecha],
                format="mixed",
                errors="coerce",
            )
            # Formato Día.Mes.Año (07.08.26)
            df_filtrado["fecha_str"] = df_filtrado["fecha_dt"].dt.strftime("%d.%m.%2y")
            df_filtrado = df_filtrado.sort_values(by="fecha_dt", ascending=True)

            # Periodo de medición en español
            fechas_validas = df_filtrado["fecha_dt"].dropna()
            if len(fechas_validas) > 0:
                min_d = fechas_validas.min()
                max_d = fechas_validas.max()
                periodo_texto = f"{min_d.strftime('%d')} al {max_d.strftime('%d')} de {MESES_ES[max_d.month]} de {max_d.year}"
            else:
                periodo_texto = "Periodo de Monitoreo"

            def obtener_campo(row, lista_cols):
                for c in lista_cols:
                    if c in row.index:
                        v = str(row[c]).strip()
                        if v and v != "nan" and v != "None":
                            return v
                return ""

            def limpiar_texto(texto):
                if not isinstance(texto, str) or texto == "nan":
                    return ""
                texto_sin_urls = re.sub(r"https?://\S+", "", texto)
                texto_sin_emojis = re.sub(r":[a-zA-Z0-9_\-|]+:", "", texto_sin_urls)
                lineas = [
                    re.sub(r"[ \t]+", " ", line).strip()
                    for line in texto_sin_emojis.split("\n")
                    if line.strip()
                ]
                return "\n".join(lineas)

            # Clasificación de sentimiento
            def clasificar_con_ia(row):
                detalle = obtener_campo(row, ["Detail", "Summary", "Síntesis", "Title"])
                prompt = f"""
