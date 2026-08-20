import io
import json
import os
import re
import unicodedata
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
    "Sistema universal de monitoreo político para cualquier perfil o cargo público. "
    "Genera reportes oficiales en Word con clasificación de IA, filtrado automático y síntesis ejecutiva."
)

# API Key Pre-integrada
GEMINI_API_KEY = "AQ.Ab8RN6LoOHgBblHSIETp2LjyBofO48YsSqSeojXYFAAKGvFa0w"
genai.configure(api_key=GEMINI_API_KEY)

# ==============================================================================
# BASE DE CONOCIMIENTO Y CRITERIOS UNIVERSALES DE CLASIFICACIÓN POLÍTICA
# ==============================================================================
SYSTEM_PROMPT_UNIVERSAL = """
Eres un analista senior de inteligencia política, comunicación gubernamental y monitoreo de medios.
Tu objetivo es clasificar publicaciones para CUALQUIER actor político (presidente municipal, gobernador, legislador, secretario de estado, dirigente partidista o candidato).

Determina el impacto reputacional del texto para el actor político objetivo en una de dos categorías: 'POSITIVA' o 'NEGATIVA'.

============================================================
REGLAS UNIVERSALES DE CLASIFICACIÓN:
============================================================

1. DEBES CLASIFICAR COMO 'POSITIVA' (O INFORMATIVA FAVORABLE):
   - LOGROS DE GESTIÓN Y CIFRAS A FAVOR: Informes de actividades, rendición de cuentas, inauguraciones de obras públicas, infraestructura, pavimentaciones, alumbrado, equipamiento y reportes de DISMINUCIÓN o BAJA en índices delictivos o rezago social.
   - BIENESTAR Y PROGRAMAS SOCIALES: Entrega de apoyos directos, becas, despensas, kits escolares, jornadas de salud, vacunación, atención a grupos vulnerables y ferias del empleo.
   - AGENDA INSTITUCIONAL Y RELACIONES PÚBLICAS: Convenios de colaboración, acuerdos con sindicatos o sectores empresariales, reuniones de trabajo, comparecencias, iniciativas legislativas aprobadas o presentadas, foros, eventos culturales y festividades tradicionales.
   - COBERTURA INFORMATIVA NEUTRAL: Notas descriptivas de medios de comunicación sobre las actividades, posicionamientos o giras de trabajo del actor.
   - POSICIONAMIENTO POLÍTICO: Resultados favorables o neutrales en encuestas de opinión pública, aprobación ciudadana o respaldos políticos.

2. DEBES CLASIFICAR COMO 'NEGATIVA' (CRISIS O AFECTACIÓN REPUTACIONAL):
   - INSEGURIDAD Y HECHOS DELICTIVOS: Cobertura de asaltos, homicidios, robos de vehículos, autopartes, balaceras o delitos ocurridos en su territorio/área de responsabilidad atribuibles a falta de vigilancia.
   - PROTESTAS Y RECLAMOS CIUDADANOS: Manifestaciones, paros, bloqueos viales, quejas ciudadanas por deficiencia de servicios públicos (baches, basura, agua potable, drenaje, alumbrado) o señalamientos de abandono gubernamental.
   - CORRUPCIÓN, AUDITORÍAS Y FISCALIZACIÓN: Observaciones de órganos de control o auditorías superiores por presunto daño patrimonial, desvío de recursos, enriquecimiento ilícito, nepotismo o falta de transparencia.
   - CONDUCTA INDEBIDA Y ESCÁNDALOS: Funcionarios involucrados en detenciones, incidentes viales, prepotencia, abuso de poder, uso indebido de recursos públicos o tráfico de influencias ('charolazos').
   - CRÍTICA POLÍTICA DIRECTA: Señalamientos de adversarios políticos, acusaciones de coacción/acarreo a eventos, 'fuego amigo', divisiones internas o columnas de opinión que descalifiquen su desempeño, origen o legitimidad.
"""

model_clasificador = genai.GenerativeModel(
    model_name="gemini-1.5-flash",
    system_instruction=SYSTEM_PROMPT_UNIVERSAL,
    generation_config={"temperature": 0.0}
)

model_redactor = genai.GenerativeModel(
    model_name="gemini-1.5-flash",
    generation_config={"temperature": 0.2}
)

def quitar_acentos(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', str(texto)) if unicodedata.category(c) != 'Mn').lower()

def normalizar_cadena(texto):
    t = quitar_acentos(texto)
    return re.sub(r'[^a-z0-9]', '', t)

# --- PATRONES INSTITUCIONALES UNIVERSALES ---
PATRONES_INSTITUCIONALES_GENERICOS = [
    r'ayuntamiento', r'gobierno', r'gobpue', r'gobiernoded', r'ayto',
    r'secretaria', r'dependencia', r'organismo', r'sindicatura', r'presidencia', r'comunicacionsocial',
    r'seguridadciudadana', r'seguridadpublica', r'proteccioncivil', r'policiamunicipal', r'policiastat',
    r'serviciospublicos', r'servicios_pub', r'obraspublicas', r'desarrollourbano', r'desarrolloeconomico',
    r'difmunicipal', r'sistemadif', r'difestatal', r'institutodelamujer', r'institutodelajuventud',
    r'organismodeagua', r'serviciodelimpia', r'organismolimpia', r'derechoshumanos'
]

def es_cuenta_del_actor_universal(autor, handle, actor_target):
    aut = normalizar_cadena(autor)
    hnd = normalizar_cadena(handle)
    act = normalizar_cadena(actor_target)
    
    if not act:
        return False
        
    if act in aut or act in hnd or aut in act:
        return True
        
    tokens_actor = [t for t in re.findall(r'\w+', quitar_acentos(actor_target)) if len(t) > 3]
    
    if len(tokens_actor) >= 2:
        coincidencias_aut = sum(1 for t in tokens_actor if t in aut)
        coincidencias_hnd = sum(1 for t in tokens_actor if t in hnd)
        if coincidencias_aut >= 2 or coincidencias_hnd >= 2:
            return True
            
    return False

def es_institucional_universal(autor, handle):
    texto = (str(autor) + " " + str(handle))
    texto_norm = normalizar_cadena(texto)
    for p in PATRONES_INSTITUCIONALES_GENERICOS:
        if re.search(p, texto_norm):
            return True
    return False

def limpiar_dataframe_redes_automatico(df_raw, actor_nombre_target):
    def es_descartable(row):
        autor = str(row.get('Autor', row.get('Author name', ''))).strip()
        handle = str(row.get('Author handle (@username)', row.get('Handle', ''))).strip()
        detalle = str(row.get('Contenido', row.get('Detail', row.get('Titulo', '')))).strip()
        
        if es_cuenta_del_actor_universal(autor, handle, actor_nombre_target):
            return True
            
        if es_institucional_universal(autor, handle):
            return True
        
        det_lower = quitar_acentos(detalle)
        if any(p in det_lower for p in ['mis ahijados', 'con toda la actitud #graciasdios', 'primeracomunion', 'en familia festejando']):
            return True
            
        return False

    mask_descarte = df_raw.apply(es_descartable, axis=1)
    df_limpio = df_raw[~mask_descarte].copy()
    total_descartadas = mask_descarte.sum()
    
    return df_limpio, total_descartadas

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

    if "-" in s_date:
        try:
            parts = s_date.split("-")
            if len(parts) == 3 and len(parts[0]) == 4:
                return datetime(int(parts[0]), int(parts), int(parts))
        except Exception:
            pass

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

def clasificar_lote_con_ia(lista_notas, actor_nombre):
    prompt = f"""
Clasifica las siguientes publicaciones respecto al actor político: "{actor_nombre}".
Aplica estrictamente las reglas universales de clasificación política del sistema.

NOTAS A EVALUAR:
{json.dumps(lista_notas, ensure_ascii=False)}

Responde ÚNICAMENTE un JSON válido con este formato exacto:
[
  {{"id": 0, "sentimiento": "POSITIVA"}},
  {{"id": 1, "sentimiento": "NEGATIVA"}}
]
"""
    try:
        response = model_clasificador.generate_content(prompt)
        raw_txt = response.text.strip()
        raw_txt = re.sub(r"^```json\s*", "", raw_txt, flags=re.I)
        raw_txt = re.sub(r"^```\s*", "", raw_txt)
        raw_txt = re.sub(r"\s*```$", "", raw_txt)
        
        datos = json.loads(raw_txt)
        res_map = {}
        for item in datos:
            sent = item.get("sentimiento", "POSITIVA").upper()
            res_map[item["id"]] = "NEGATIVA" if "NEGAT" in sent else "POSITIVA"
        return res_map
    except Exception:
        res_map = {}
        for item in lista_notas:
            t = item["texto"].lower()
            es_logro_disminucion = any(k in t for k in ["disminuye", "disminución", "bajan", "reduccion", "cae", "a la baja"]) and any(k in t for k in ["delito", "robo", "incidencia", "homicidio"])
            es_acuerdo_laboral = any(k in t for k in ["sindicato", "convenio", "acuerdo laboral", "prestaciones", "condiciones de trabajo"])
            
            if not es_logro_disminucion and not es_acuerdo_laboral and any(k in t for k in [
                "asalto", "asesinato", "homicidio", "robo de ", "cristalazo", "daño patrimonial", "desvío de recursos",
                "auditoría", "nepotismo", "desfalco", "desplante", "se manifestaron", "protestan", "exigen",
                "ebrio", "borracho", "charolazo", "prepotencia", "prepotente", "bloqueo", "baches", "acarreados"
            ]):
                res_map[item["id"]] = "NEGATIVA"
            else:
                res_map[item["id"]] = "POSITIVA"
        return res_map

def determinar_sentimiento_df(df_data, actor_nombre_target, es_tradicionales):
    if es_tradicionales:
        if any(c in df_data.columns for c in ["Sentimiento", "Sentiment", "Sentimiento de la Nota"]):
            sent_col = obtener_columna_serie(df_data, ["Sentimiento", "Sentiment", "Sentimiento de la Nota"])
            return sent_col.apply(lambda s: "NEGATIVA" if "negat" in str(s).lower() else "POSITIVA").tolist()

    sentimientos_finales = ["POSITIVA"] * len(df_data)
    lote_tamano = 15
    indices = df_data.index.tolist()
    
    progreso = st.progress(0)
    total_lotes = (len(indices) + lote_tamano - 1) // lote_tamano

    for l_idx in range(total_lotes):
        sub_indices = indices[l_idx * lote_tamano : (l_idx + 1) * lote_tamano]
        lista_lote = []
        for local_id, idx in enumerate(sub_indices):
            row = df_data.loc[idx]
            texto = obtener_campo(row, ["Contenido", "Detail", "Titulo", "Summary", "Síntesis", "Title"])
            lista_lote.append({"id": local_id, "texto": texto})
            
        res_map = clasificar_lote_con_ia(lista_lote, actor_nombre_target)
        for local_id, idx in enumerate(sub_indices):
            real_pos = df_data.index.get_loc(idx)
            sentimientos_finales[real_pos] = res_map.get(local_id, "POSITIVA")
            
        progreso.progress((l_idx + 1) / total_lotes)
        
    progreso.empty()
    return sentimientos_finales

# --- GENERADOR UNIVERSAL DEL RESUMEN EJECUTIVO EN PÁRRAFOS CONCISOS ---
def extraer_resumen_temas_real(df_data, actor_nombre):
    pos_df = df_data[df_data["sentimiento_final"].isin(["POSITIVA", "NEUTRA"])]
    neg_df = df_data[df_data["sentimiento_final"] == "NEGATIVA"]

    pos_textos = []
    for col in ["Titulo", "Contenido", "Detail", "Summary"]:
        if col in pos_df.columns:
            pos_textos = [t for t in pos_df[col].dropna().astype(str).tolist() if "teaser" not in t.lower() and len(t.strip()) > 10]
            break

    neg_textos = []
    for col in ["Titulo", "Contenido", "Detail", "Summary"]:
        if col in neg_df.columns:
            neg_textos = [t for t in neg_df[col].dropna().astype(str).tolist() if "teaser" not in t.lower() and len(t.strip()) > 10]
            break

    prompt = f"""
Eres un analista senior de comunicación política y redacción ejecutiva.
Redacta el "RESUMEN" ejecutivo oficial para el actor político: "{actor_nombre}".

REGLAS DE FORMATO (OBLIGATORIAS):
1. Cada punto DEBE ser exactamente de 1 solo párrafo fluido (de 2 a 3 líneas).
2. Estructura exacta de cada viñeta:
   [Número]. [Título del Eje Temático Específico en Mayúsculas y Minúsculas]: [Síntesis ejecutiva explicando los hechos con nombres de programas, lugares, obras o instituciones mencionados en las notas].
3. Prohibido copiar o pegar párrafos completos de las noticias originales. Debes resumir y redactar con estilo formal e institucional.
4. Genera de 1 a 3 puntos en 'Temas relevantes informativos' y de 1 a 3 puntos en 'Temas negativos'.
5. Si no hay notas negativas ({len(neg_textos)} negativas registradas), escribe obligatoriamente:
   1. No se registraron temas negativos en el periodo analizado.

EJEMPLO DE ESTRUCTURA:
Temas relevantes informativos
1. [Eje Temático Positivo 1]: [Síntesis ejecutiva de 2 a 3 líneas de los hechos].
2. [Eje Temático Positivo 2]: [Síntesis ejecutiva de 2 a 3 líneas de los hechos].

Temas negativos
1. [Eje Temático Negativo 1 o Leyenda si no hay]: [Síntesis ejecutiva de 2 a 3 líneas de las controversias].

NOTICIAS POSITIVAS DISPONIBLES ({len(pos_textos)} notas):
{str(pos_textos[:30])}

NOTICIAS NEGATIVAS DISPONIBLES ({len(neg_textos)} notas):
{str(neg_textos[:25])}
"""
    try:
        res_ia = model_redactor.generate_content(prompt).text.strip()
        if len(neg_textos) > 0 and "no se registraron temas negativos" in res_ia.lower():
            pass
        elif res_ia and len(res_ia) > 30 and len(res_ia) < 1500:
            if not res_ia.startswith("Temas relevantes informativos"):
                res_ia = "Temas relevantes informativos\n" + res_ia
            return res_ia
    except Exception:
        pass

    lineas_res = ["Temas relevantes informativos"]
    if len(pos_textos) > 0:
        titulos_unicos_pos = list(dict.fromkeys([
            re.sub(r'^\d+:\d+\s*(hrs|am|pm|\.)\s*', '', t, flags=re.I).strip() 
            for t in pos_textos if t.strip() and "teaser" not in t.lower()
        ]))
        for i, t in enumerate(titulos_unicos_pos[:3], 1):
            t_clean = t.split('.')[0].strip()
            lineas_res.append(f"{i}. Gestión y Agenda Institucional: Difusión y seguimiento a {t_clean.lower()}.")
    else:
        lineas_res.append("1. Agenda Institucional: Difusión de actividades públicas y agenda de trabajo.")

    lineas_res.append("\nTemas negativos")
    if len(neg_textos) > 0:
        titulos_unicos_neg = list(dict.fromkeys([
            re.sub(r'^\d+:\d+\s*(hrs|am|pm|\.)\s*', '', t, flags=re.I).strip() 
            for t in neg_textos if t.strip() and "teaser" not in t.lower()
        ]))
        for i, t in enumerate(titulos_unicos_neg[:3], 1):
            t_clean = t.split('.')[0].strip()
            lineas_res.append(f"{i}. Controversias y Señalamientos Públicos: Cobertura mediática sobre {t_clean.lower()}.")
    else:
        lineas_res.append("1. No se registraron temas negativos en el periodo analizado.")

    return "\n".join(lineas_res)

def crear_doc_desde_hoja(df_hoja, nombre_hoja, es_redes_sociales):
    if es_redes_sociales:
        df_filtrado, total_descartadas = limpiar_dataframe_redes_automatico(df_hoja, nombre_hoja)
        if total_descartadas > 0:
            st.info(f"ℹ️ Se limpiaron automáticamente {total_descartadas} publicaciones institucionales/personales. Analizando {len(df_filtrado)} menciones ciudadanas reales.")
    else:
        df_filtrado = df_hoja.copy()

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

    df_filtrado["sentimiento_final"] = determinar_sentimiento_df(df_filtrado, nombre_hoja, es_tradicionales=not es_redes_sociales)

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

    temas_texto = extraer_resumen_temas_real(df_filtrado, nombre_hoja)
    for linea in temas_texto.split("\n"):
        linea_clean = linea.strip()
        if linea_clean:
            p_t = doc.add_paragraph()
            p_t.paragraph_format.space_before = Pt(1)
            p_t.paragraph_format.space_after = Pt(2)
            if "TEMAS RELEVANTES INFORMATIVOS" in linea_clean.upper():
                add_run_verdana(p_t, "Temas relevantes informativos", bold=True, size_pt=10)
            elif "TEMAS NEGATIVOS" in linea_clean.upper():
                p_t.paragraph_format.space_before = Pt(4)
                add_run_verdana(p_t, "Temas negativos", bold=True, size_pt=10, color_rgb=RGBColor(180, 0, 0))
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
        placeholder="ej. ALEJANDRO ARMENTA (MORENA), EDUARDO RIVERA (PAN), etc.",
    ).strip().upper()

    if uploaded_file and actor_nombre_in:
        if st.button("Generar Reporte Oficial", type="primary"):
            with st.spinner("Limpiando cuentas oficiales y procesando todas las notas con IA universal..."):
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
                        st.warning("El archivo no contiene notas válidas tras la limpieza automática.")
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
