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

# Interfaz de usuario para carga de archivos (con soporte para xlsx, xls y csv)
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
      # Leer archivo según su extensión
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

      # Detección automática de la columna de fecha (soporta inglés y español)
      posibles_cols_fecha = [
          "Publish date",
          "Date",
          "Fecha",
          "Fecha de publicación",
      ]
      col_fecha = None
      for c in posibles_cols_fecha:
        if c in df_filtrado.columns:
          col_fecha = c
          break

      if col_fecha is None:
        st.error(
            "No se encontró una columna de fecha válida en el archivo. Las"
            f" columnas detectadas son: {list(df_filtrado.columns)}"
        )
        st.stop()

      df_filtrado["fecha_dt"] = pd.to_datetime(
          df_filtrado[col_fecha],
          format="mixed",
          errors="coerce",
      )
      df_filtrado["fecha_str"] = df_filtrado["fecha_dt"].dt.strftime("%d.%m.%2y")
      df_filtrado = df_filtrado.sort_values(by="fecha_dt", ascending=True)

      fechas_validas = df_filtrado["fecha_dt"].dropna()
      periodo_texto = (
          f"{fechas_validas.min().strftime('%d')} al"
          f" {fechas_validas.max().strftime('%d de %B de %Y')}"
          if len(fechas_validas) > 0
          else "Periodo de Monitoreo"
      )


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


      def clasificar_con_ia(row):
        detalle = obtener_campo(row, ["Detail", "Summary", "Síntesis", "Title"])
        prompt = f"""
                Eres un analista político experto. Analiza el texto sobre '{actor_nombre}'.
                CRITERIO:
                - Responde 'NEGATIVA' ÚNICAMENTE si contiene un ataque personal directo, insulto o acusación grave de mal desempeño hacia '{actor_nombre}'.
                - Responde 'POSITIVA' si reporta respaldos de líderes, propuestas, giras, declaraciones o logros.
                Texto: "{detalle}"
                Responde ÚNICAMENTE: POSITIVA o NEGATIVA.
                """
        try:
          res = model.generate_content(prompt).text.strip().upper()
          return "NEGATIVA" if "NEGATIVA" in res else "POSITIVA"
        except Exception:
          return "POSITIVA"


      df_filtrado["sentimiento_ia"] = df_filtrado.apply(
          clasificar_con_ia, axis=1
      )

      positivas_cnt = len(
          df_filtrado[df_filtrado["sentimiento_ia"] == "POSITIVA"]
      )
      negativas_cnt = len(
          df_filtrado[df_filtrado["sentimiento_ia"] == "NEGATIVA"]
      )
      total_cnt = len(df_filtrado)

      redes_cnt = (
          df_filtrado["Media type"]
          .astype(str)
          .str.contains(
              "Facebook|X|Twitter|Instagram|TikTok", case=False, na=False
          )
          .sum()
      )
      portales_cnt = (
          df_filtrado["Media type"]
          .astype(str)
          .str.contains("Portal|Web|Online", case=False, na=False)
          .sum()
      )
      prensa_cnt = (
          df_filtrado["Media type"]
          .astype(str)
          .str.contains("Prensa|Diario|Periódico", case=False, na=False)
          .sum()
      )
      columnas_cnt = (
          df_filtrado["Media type"]
          .astype(str)
          .str.contains("Columna|Opinión", case=False, na=False)
          .sum()
      )

      # Crear Documento Word
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

      # 2. Tabla de Impactos
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
      add_run_verdana(
          p_tot,
          f"TOTAL NOTAS INFORMATIVAS: {total_cnt}\nREDES SOCIALES:"
          f" {redes_cnt}\nPORTALES DIGITALES: {portales_cnt}\nPRENSA LOCAL:"
          f" {prensa_cnt}\nCOLUMNAS: {columnas_cnt}",
          bold=True,
          size_pt=10,
      )

      # 3. Resumen
      p_res = doc.add_paragraph()
      p_res.paragraph_format.space_before = Pt(10)
      add_run_verdana(p_res, "RESUMEN", bold=True, size_pt=11)

      p_temas = doc.add_paragraph()
      p_temas.paragraph_format.space_after = Pt(4)
      add_run_verdana(
          p_temas, "Temas relevantes informativos", bold=True, size_pt=10
      )

      p_r1 = doc.add_paragraph()
      add_run_verdana(
          p_r1,
          f"• Predominó la cobertura favorable: {positivas_cnt} de {total_cnt}"
          f" impactos fueron positivos o informativos, frente a {negativas_cnt}"
          " impactos negativos.",
          size_pt=9.5,
      )

      # 4. Desglose
      p_des = doc.add_paragraph()
      p_des.paragraph_format.space_before = Pt(12)
      add_run_verdana(p_des, "DESGLOSE", bold=True, size_pt=11)

      for fecha_item in df_filtrado["fecha_str"].dropna().unique():
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

        for m_type, grupo_m in pos_df.groupby(
            lambda i: obtener_campo(
                pos_df.loc[i], ["Media type", "Source type"]
            )
            or "REDES SOCIALES"
        ):
          p_m = doc.add_paragraph()
          p_m.paragraph_format.space_before = Pt(4)
          p_m.paragraph_format.space_after = Pt(4)
          add_run_verdana(p_m, m_type.upper(), bold=True, size_pt=10)

          for _, row in grupo_m.iterrows():
            autor = obtener_campo(
                row, ["Author name", "Source", "Media name", "Programa"]
            )
            handle = obtener_campo(row, ["Author handle (@username)", "Handle"])
            detalle = obtener_campo(
                row, ["Detail", "Summary", "Síntesis", "Title"]
            )
            link = obtener_campo(row, ["Link", "URL", "Enlace"])

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
                row, ["Author name", "Source", "Media name", "Programa"]
            )
            handle = obtener_campo(row, ["Author handle (@username)", "Handle"])
            detalle = obtener_campo(
                row, ["Detail", "Summary", "Síntesis", "Title"]
            )
            link = obtener_campo(row, ["Link", "URL", "Enlace"])

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

      # Guardar en memoria para descarga en Streamlit
      buffer = io.BytesIO()
      doc.save(buffer)
      buffer.seek(0)

      out_name = f"Reporte_Oficial_IA_{actor_nombre.replace(' ', '_')}.docx"

      st.success("¡Reporte generado exitosamente con IA!")
      st.download_button(
          label="📥 Descargar Reporte en Word",
          data=buffer,
          file_name=out_name,
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )
