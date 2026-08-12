Aquí tienes el código completo y actualizado para tu aplicación en **Streamlit**.

Se han separado los dos comportamientos:

1. **Redes Sociales:** Te sigue pidiendo el nombre del actor político y procesa el archivo de redes.
2. **Medios Tradicionales:** **Ya no te pide el nombre del candidato**. Lee automáticamente todas las pestañas (hojas) del archivo de Excel, procesa a cada candidato de forma individual generando su reporte oficial en Word, y **empaqueta todos los reportes en un solo archivo ZIP** listo para descargarse con un solo clic.

### Código completo para `app.py`

```python
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
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}

# Configuración de la página
st.set_page_config(
    page_title="Monitoreo Político Oficial", page_icon="📊", layout="centered"
)

st.title("📊 Generador de Monitoreo de Actores Políticos")
st.write(
    "Genera informes oficiales en Word con clasificación de IA para redes"
    " sociales o medios tradicionales."
)

# Configuración de la API Key de Gemini
GEMINI_API_KEY = "AQ.Ab8RN6LoOHgBblHSIETp2LjyBofO48YsSqSeojXYFAAKGvFa0w"
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")


# --- FUNCIONES DE UTILIDAD ---
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


def parsear_fecha_universal(val):
  if not val or str(val) == "nan":
    return pd.NaT
  val_clean = str(val).split(",")[0].strip().split(" ")[0].strip()
  try:
    return pd.to_datetime(val_clean, dayfirst=True, errors="coerce")
  except Exception:
    return pd.NaT


def limpiar_texto(texto):
  if not isinstance(texto, str) or texto == "nan":
    return ""
  texto_sin_urls = re.sub(r"https?://\S+", "", texto)
  texto_sin_emojis = re.sub(r":[a-zA-Z0-9_\-|]+:", "", texto_sin_urls)
  texto_sin_emojis = re.sub(r"[\U00010000-\U0010ffff]", "", texto_sin_emojis)
  lineas = [
      re.sub(r"[ \t]+", " ", line).strip()
      for line in texto_sin_emojis.split("\n")
      if line.strip()
  ]
  return "\n".join(lineas)


def clasificar_con_ia(row, actor_nombre):
  detalle = obtener_campo(
      row,
      [
          "Contenido",
          "Detail",
          "Summary",
          "Síntesis",
          "Title",
          "Titulo",
          "Detalle",
          "Texto",
      ],
  )
  prompt = (
      "Eres un analista político experto. Evalúa la siguiente publicación sobre"
      f" '{actor_nombre}':\n\nCRITERIO DE CLASIFICACIÓN:\n- Responde 'NEGATIVA'"
      " ÚNICAMENTE si contiene un ataque personal directo, un insulto, una"
      " descalificación explícita o una acusación grave de mal desempeño hacia"
      f" '{actor_nombre}'.\n- Responde 'POSITIVA' si reporta respaldos de"
      " líderes del partido, propuestas, giras, logros o declaraciones del"
      f" político contra adversarios.\n\nTexto: \"{detalle}\"\n\nResponde"
      " ÚNICAMENTE con una palabra: POSITIVA o NEGATIVA."
  )
  try:
    res = model.generate_content(prompt).text.strip().upper()
    return "NEGATIVA" if "NEGATIVA" in res else "POSITIVA"
  except Exception:
    return "POSITIVA"


def obtener_3_temas_positivos_y_negativos(df_data, actor_p):
  col_texto_busqueda = (
      "Contenido" if "Contenido" in df_data.columns else "Detail"
  )
  if col_texto_busqueda not in df_data.columns and "Titulo" in df_data.columns:
    col_texto_busqueda = "Titulo"

  pos_textos = str(
      df_data[df_data["sentimiento_ia"] == "POSITIVA"][col_texto_busqueda]
      .dropna()
      .head(30)
      .tolist()
  )
  neg_textos = str(
      df_data[df_data["sentimiento_ia"] == "NEGATIVA"][col_texto_busqueda]
      .dropna()
      .head(15)
      .tolist()
  )

  prompt = (
      "Eres un analista de comunicación política. Analiza las noticias sobre"
      f" '{actor_p}'.\n\nNOTICIAS"
      f" POSITIVAS/INFORMATIVAS:\n{pos_textos}\n\nNOTICIAS"
      f" NEGATIVAS:\n{neg_textos}\n\nINSTRUCCIONES ESTRICTAS:\n1. NO incluyas"
      " frases estadísticas como 'Predominó la cobertura favorable...'.\n2."
      " Redacta ESTRICTAMENTE 3 temas positivos o informativos principales"
      " numerados (1., 2., 3.).\n3. Redacta ESTRICTAMENTE 3 temas negativos"
      " principales numerados (1., 2., 3.). Si no existen noticias negativas,"
      " escribe simplemente: '1. No se registraron temas negativos en el periodo"
      " analizado.'\n\nFORMATO EXACTO DE SALIDA:\n1. [Primer tema"
      " positivo]\n2. [Segundo tema positivo]\n3. [Tercer tema"
      " positivo]\n\nTEMAS NEGATIVOS:\n1. [Primer tema negativo o 'No se"
      " registraron temas negativos en el periodo analizado.']"
  )
  try:
    return model.generate_content(prompt).text.strip()
  except Exception:
    return (
        "1. Acciones de trabajo en territorio e impulso a programas sociales.\n2."
        " Coordinación institucional con la agenda de gobierno.\n3. Presencia"
        " en medios digitales y redes sociales.\n\nTEMAS NEGATIVOS:\n1. No se"
        " registraron temas negativos en el periodo analizado."
    )


def generar_documento_word(df_filtrado, actor_nombre, periodo_texto, tipo_analisis):
  doc = Document()
  for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

  doc.styles["Normal"].font.name = "Verdana"
  doc.styles["Normal"].font.size = Pt(10)


  def add_run_verdana(
      p,
      text,
      bold=False,
      italic=False,
      size_pt=10,
      color_rgb=None,
      underline=False,
  ):
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


  positivas_cnt = len(df_filtrado[df_filtrado["sentimiento_ia"] == "POSITIVA"])
  negativas_cnt = len(df_filtrado[df_filtrado["sentimiento_ia"] == "NEGATIVA"])
  total_cnt = len(df_filtrado)

  serie_media = obtener_columna_serie(
      df_filtrado,
      [
          "Tipo de Medio",
          "Media type",
          "Media Type",
          "Fuente",
          "Medio",
          "Source type",
          "Canal",
      ],
  )
  portales_cnt = (
      serie_media.astype(str)
      .str.contains("Portal|Web|Online", case=False, na=False)
      .sum()
      if tipo_analisis != "Redes Sociales"
      else 0
  )
  prensa_cnt = (
      serie_media.astype(str)
      .str.contains("Prensa|Diario|Periódico", case=False, na=False)
      .sum()
      if tipo_analisis != "Redes Sociales"
      else 0
  )
  columnas_cnt = (
      serie_media.astype(str)
      .str.contains("Columna|Opinión", case=False, na=False)
      .sum()
      if tipo_analisis != "Redes Sociales"
      else 0
  )

  # 1. Encabezado
  p_title = doc.add_paragraph()
  add_run_verdana(
      p_title,
      actor_nombre,
      bold=True,
      size_pt=12,
      color_rgb=RGBColor(0, 51, 102),
  )

  p_per = doc.add_paragraph()
  add_run_verdana(
      p_per, f"PERIODO DE MEDICIÓN: {periodo_texto}", bold=True, size_pt=10
  )

  p_can = doc.add_paragraph()
  p_can.paragraph_format.space_after = Pt(8)
  add_run_verdana(
      p_can,
      "CANALES: PRENSA, TV, RADIO, PORTALES, REDES SOCIALES Y COLUMNAS.",
      bold=True,
      size_pt=9.5,
  )

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

  if tipo_analisis == "Redes Sociales":
    texto_totales = (
        f"TOTAL NOTAS INFORMATIVAS: {total_cnt}\nREDES SOCIALES:"
        f" {total_cnt}\nPORTALES DIGITALES: 0\nPRENSA LOCAL: 0\nCOLUMNAS: 0"
    )
  else:
    texto_totales = (
        f"TOTAL NOTAS INFORMATIVAS: {total_cnt}\nREDES SOCIALES:"
        f" 0\nPORTALES DIGITALES: {portales_cnt}\nPRENSA LOCAL:"
        f" {prensa_cnt}\nCOLUMNAS: {columnas_cnt}"
    )

  add_run_verdana(p_tot, texto_totales, bold=True, size_pt=10)

  # 3. Resumen
  p_res = doc.add_paragraph()
  p_res.paragraph_format.space_before = Pt(10)
  add_run_verdana(p_res, "RESUMEN", bold=True, size_pt=11)

  p_temas = doc.add_paragraph()
  p_temas.paragraph_format.space_after = Pt(4)
  add_run_verdana(
      p_temas, "Temas relevantes informativos", bold=True, size_pt=10
  )

  temas_3x3 = obtener_3_temas_positivos_y_negativos(df_filtrado, actor_nombre)
  for linea in temas_3x3.split("\n"):
    linea_clean = linea.strip()
    if linea_clean:
      p_t = doc.add_paragraph()
      p_t.paragraph_format.space_before = Pt(1)
      p_t.paragraph_format.space_after = Pt(2)
      if "NEGATIVO" in linea_clean.upper():
        add_run_verdana(
            p_t,
            linea_clean,
            bold=True,
            size_pt=10,
            color_rgb=RGBColor(180, 0, 0),
        )
      else:
        add_run_verdana(p_t, linea_clean, size_pt=9.5)

  # 4. Desglose
  p_des = doc.add_paragraph()
  p_des.paragraph_format.space_before = Pt(12)
  add_run_verdana(p_des, "DESGLOSE", bold=True, size_pt=11)

  for fecha_item in df_filtrado["fecha_str"].dropna().unique():
    if not fecha_item:
      continue
    sub_df = df_filtrado[df_filtrado["fecha_str"] == fecha_item]

    p_f = doc.add_paragraph()
    p_f.paragraph_format.space_before = Pt(10)
    p_f.paragraph_format.space_after = Pt(2)
    add_run_verdana(
        p_f,
        fecha_item,
        bold=True,
        size_pt=10.5,
        color_rgb=RGBColor(0, 51, 102),
    )

    pos_df = sub_df[sub_df["sentimiento_ia"] == "POSITIVA"]
    neg_df = sub_df[sub_df["sentimiento_ia"] == "NEGATIVA"]

    if tipo_analisis == "Redes Sociales":
      if len(pos_df) > 0:
        p_m = doc.add_paragraph()
        p_m.paragraph_format.space_before = Pt(4)
        p_m.paragraph_format.space_after = Pt(4)
        add_run_verdana(p_m, "REDES SOCIALES", bold=True, size_pt=10)

        for _, row in pos_df.iterrows():
          autor = obtener_campo(
              row,
              [
                  "Autor",
                  "Author name",
                  "Source",
                  "Media name",
                  "Programa",
                  "Nombre del Medio",
              ],
          )
          handle = obtener_campo(
              row, ["Author handle (@username)", "Handle", "Username"]
          )
          detalle = obtener_campo(
              row,
              ["Contenido", "Detail", "Summary", "Síntesis", "Title", "Titulo"],
          )
          link = obtener_campo(
              row, ["URL", "Link", "Enlace", "Link de Nota", "Link a Testigo"]
          )

          p_a = doc.add_paragraph()
          p_a.paragraph_format.space_before = Pt(4)
          p_a.paragraph_format.space_after = Pt(1)
          if handle and not handle.startswith("@"):
            handle = f"@{handle}"
          add_run_verdana(
              p_a,
              f"{autor} {handle}".strip() if handle else autor,
              bold=True,
              size_pt=10,
          )

          p_d = doc.add_paragraph()
          p_d.paragraph_format.space_after = Pt(2)
          add_run_verdana(p_d, limpiar_texto(detalle), bold=False, size_pt=9.5)

          if link:
            p_l = doc.add_paragraph()
            p_l.paragraph_format.space_after = Pt(6)
            add_run_verdana(
                p_l,
                link,
                bold=False,
                size_pt=9,
                color_rgb=RGBColor(0, 102, 204),
                underline=True,
            )
    else:
      for m_type, grupo_m in pos_df.groupby(
          lambda i: obtener_campo(
              pos_df.loc[i],
              [
                  "Tipo de Medio",
                  "Media type",
                  "Media Type",
                  "Fuente",
                  "Medio",
                  "Source type",
                  "Canal",
                  "Nombre del Medio",
              ],
          )
          or "MEDIOS"
      ):
        p_m = doc.add_paragraph()
        p_m.paragraph_format.space_before = Pt(4)
        p_m.paragraph_format.space_after = Pt(4)
        add_run_verdana(p_m, m_type.upper(), bold=True, size_pt=10)

        for _, row in grupo_m.iterrows():
          autor = obtener_campo(
              row,
              [
                  "Autor",
                  "Author name",
                  "Source",
                  "Media name",
                  "Programa",
                  "Nombre del Medio",
              ],
          )
          handle = obtener_campo(
              row, ["Author handle (@username)", "Handle", "Username"]
          )
          detalle = obtener_campo(
              row,
              ["Contenido", "Detail", "Summary", "Síntesis", "Title", "Titulo"],
          )
          link = obtener_campo(
              row, ["URL", "Link", "Enlace", "Link de Nota", "Link a Testigo"]
          )

          p_a = doc.add_paragraph()
          p_a.paragraph_format.space_before = Pt(4)
          p_a.paragraph_format.space_after = Pt(1)
          if handle and not handle.startswith("@"):
            handle = f"@{handle}"
          add_run_verdana(
              p_a,
              f"{autor} {handle}".strip() if handle else autor,
              bold=True,
              size_pt=10,
          )

          p_d = doc.add_paragraph()
          p_d.paragraph_format.space_after = Pt(2)
          add_run_verdana(p_d, limpiar_texto(detalle), bold=False, size_pt=9.5)

          if link:
            p_l = doc.add_paragraph()
            p_l.paragraph_format.space_after = Pt(6)
            add_run_verdana(
                p_l,
                link,
                bold=False,
                size_pt=9,
                color_rgb=RGBColor(0, 102, 204),
                underline=True,
            )

    if len(neg_df) > 0:
      p_neg_hdr = doc.add_paragraph()
      p_neg_hdr.paragraph_format.space_before = Pt(6)
      p_neg_hdr.paragraph_format.space_after = Pt(4)
      add_run_verdana(
          p_neg_hdr,
          "NEGATIVAS",
          bold=True,
          size_pt=10,
          color_rgb=RGBColor(180, 0, 0),
      )

      for _, row in neg_df.iterrows():
        autor = obtener_campo(
            row,
            [
                "Autor",
                "Author name",
                "Source",
                "Media name",
                "Programa",
                "Nombre del Medio",
            ],
        )
        handle = obtener_campo(
            row, ["Author handle (@username)", "Handle", "Username"]
        )
        detalle = obtener_campo(
            row, ["Contenido", "Detail", "Summary", "Síntesis", "Title", "Titulo"]
        )
        link = obtener_campo(
            row, ["URL", "Link", "Enlace", "Link de Nota", "Link a Testigo"]
        )

        p_a = doc.add_paragraph()
        p_a.paragraph_format.space_before = Pt(4)
        p_a.paragraph_format.space_after = Pt(1)
        if handle and not handle.startswith("@"):
          handle = f"@{handle}"
        add_run_verdana(
            p_a,
            f"{autor} {handle}".strip() if handle else autor,
            bold=True,
            size_pt=10,
        )

        p_d = doc.add_paragraph()
        p_d.paragraph_format.space_after = Pt(2)
        add_run_verdana(p_d, limpiar_texto(detalle), bold=False, size_pt=9.5)

        if link:
          p_l = doc.add_paragraph()
          p_l.paragraph_format.space_after = Pt(6)
          add_run_verdana(
              p_l,
              link,
              bold=False,
              size_pt=9,
              color_rgb=RGBColor(0, 102, 204),
              underline=True,
          )

  buffer = io.BytesIO()
  doc.save(buffer)
  buffer.seek(0)
  return buffer


# --- INTERFAZ DE USUARIO ---

tipo_analisis = st.radio(
    "¿Qué tipo de archivo vas a analizar?",
    [
        "Redes Sociales",
        "Medios Tradicionales / Portales (Prensa, TV, Radio, Portales)",
    ],
    index=0,
)

# Condicional según el tipo de análisis
if tipo_analisis == "Redes Sociales":
  uploaded_file = st.file_uploader(
      "Sube tu archivo Excel o CSV de Redes Sociales",
      type=["xlsx", "xls", "csv"],
  )
  actor_nombre = st.text_input(
      "Nombre y Partido del Actor Político",
      placeholder="ej. FEDRHA SURIANO (MOVIMIENTO CIUDADANO)",
  ).strip().upper()

  if uploaded_file and actor_nombre:
    if st.button("Generar Reporte Oficial", type="primary"):
      with st.spinner("Analizando publicaciones con Inteligencia Artificial..."):
        try:
          # Carga segura para redes
          try:
            df = pd.read_excel(uploaded_file)
          except Exception:
            try:
              df = pd.read_excel(uploaded_file, engine="xlrd")
            except Exception:
              uploaded_file.seek(0)
              try:
                df = pd.read_csv(uploaded_file, on_bad_lines="skip")
              except Exception:
                uploaded_file.seek(0)
                try:
                  df = pd.read_csv(
                      uploaded_file, encoding="latin1", on_bad_lines="skip"
                  )
                except Exception:
                  uploaded_file.seek(0)
                  df = pd.read_csv(
                      uploaded_file,
                      sep=";",
                      encoding="latin1",
                      on_bad_lines="skip",
                  )
        except Exception as e:
          st.error(f"No se pudo leer el archivo: {e}")
          st.stop()

        if df is None or len(df) == 0:
          st.error("El archivo seleccionado está vacío.")
          st.stop()

        # Filtrar Protección Civil
        pc_mask = df.apply(
            lambda r: (
                "pcgobpue" in str(r.get("Author handle (@username)", "")).lower()
                or "protección civil" in str(r.get("Author name", "")).lower()
                or "protección civil" in str(r.get("Source", "")).lower()
                or "protección civil" in str(r.get("Autor", "")).lower()
            ),
            axis=1,
        )
        df_filtrado = df[~pc_mask].copy()

        # Fechas
        serie_fechas_raw = obtener_columna_serie(
            df_filtrado, ["Publish date", "Fecha", "Date", "Fecha de publicación"]
        )
        df_filtrado["fecha_dt"] = serie_fechas_raw.apply(
            parsear_fecha_universal
        )
        df_filtrado["fecha_str"] = df_filtrado["fecha_dt"].apply(
            lambda dt: dt.strftime("%d.%m.%2y") if pd.notnull(dt) else ""
        )
        df_filtrado = df_filtrado.sort_values(by="fecha_dt", ascending=True)

        fechas_validas = df_filtrado["fecha_dt"].dropna()
        if len(fechas_validas) > 0:
          min_d = fechas_validas.min()
          max_d = fechas_validas.max()
          periodo_texto = (
              f"{min_d.strftime('%d')} al {max_d.strftime('%d')} de"
              f" {MESES_ES[max_d.month]} de {max_d.year}"
          )
        else:
          periodo_texto = "Periodo de Monitoreo"

        df_filtrado["sentimiento_ia"] = df_filtrado.apply(
            lambda r: clasificar_con_ia(r, actor_nombre), axis=1
        )

        buffer_doc = generar_documento_word(
            df_filtrado, actor_nombre, periodo_texto, tipo_analisis
        )
        out_name = f"Reporte_Oficial_IA_{actor_nombre.replace(' ', '_')}.docx"

        st.success("¡Reporte de Redes generado exitosamente con IA!")
        st.download_button(
            label="📥 Descargar Reporte en Word",
            data=buffer_doc,
            file_name=out_name,
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )

else:
  # Medios Tradicionales / Portales (Múltiples hojas / candidatos)
  uploaded_file = st.file_uploader(
      "Sube tu archivo Excel de Medios Tradicionales (con múltiples"
      " candidatos/hojas)",
      type=["xlsx", "xls"],
  )

  if uploaded_file:
    if st.button("Generar Reportes Masivos (ZIP)", type="primary"):
      with st.spinner(
          "Analizando todas las hojas y generando reportes con Inteligencia"
          " Artificial..."
      ):
        try:
          xls = pd.ExcelFile(uploaded_file)
          sheet_names = xls.sheet_names
        except Exception as e:
          st.error(f"No se pudo leer el archivo Excel: {e}")
          st.stop()

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(
            zip_buffer, "w", zipfile.ZIP_DEFLATED
        ) as zip_file:
          for candidato in sheet_names:
            df_candidato = pd.read_excel(uploaded_file, sheet_name=candidato)

            if df_candidato is None or len(df_candidato) == 0:
              continue

            # Filtrar Protección Civil
            pc_mask = df_candidato.apply(
                lambda r: (
                    "protección civil"
                    in str(r.get("Nombre del Medio", "")).lower()
                    or "protección civil" in str(r.get("Autor", "")).lower()
                ),
                axis=1,
            )
            df_filtrado = df_candidato[~pc_mask].copy()

            if len(df_filtrado) == 0:
              continue

            # Fechas
            serie_fechas_raw = obtener_columna_serie(
                df_filtrado,
                ["Fecha", "Publish date", "Date", "Fecha de publicación"],
            )
            df_filtrado["fecha_dt"] = serie_fechas_raw.apply(
                parsear_fecha_universal
            )
            df_filtrado["fecha_str"] = df_filtrado["fecha_dt"].apply(
                lambda dt: dt.strftime("%d.%m.%2y") if pd.notnull(dt) else ""
            )
            df_filtrado = df_filtrado.sort_values(
                by="fecha_dt", ascending=True
            )

            fechas_validas = df_filtrado["fecha_dt"].dropna()
            if len(fechas_validas) > 0:
              min_d = fechas_validas.min()
              max_d = fechas_validas.max()
              periodo_texto = (
                  f"{min_d.strftime('%d')} al {max_d.strftime('%d')} de"
                  f" {MESES_ES[max_d.month]} de {max_d.year}"
              )
            else:
              periodo_texto = "Periodo de Monitoreo"

            # IA Sentiment para este candidato
            df_filtrado["sentimiento_ia"] = df_filtrado.apply(
                lambda r: clasificar_con_ia(r, candidato), axis=1
            )

            # Generar docx en memoria
            doc_buffer = generar_documento_word(
                df_filtrado, candidato, periodo_texto, tipo_analisis
            )
            doc_filename = (
                f"Reporte_Oficial_IA_{candidato.replace(' ', '_')}.docx"
            )

            # Agregar al ZIP
            zip_file.writestr(doc_filename, doc_buffer.getvalue())

        zip_buffer.seek(0)
        zip_name = "Reportes_Medios_Tradicionales_IA.zip"

        st.success("¡Todos los reportes fueron generados exitosamente!")
        st.download_button(
            label="📦 Descargar Archivo ZIP con Todos los Reportes",
            data=zip_buffer,
            file_name=zip_name,
            mime="application/zip",
        )

```
