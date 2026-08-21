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
    "Permite procesar archivos individuales o fusionar múltiples archivos (Radio/TV, Portales Web y Redes) en un solo reporte oficial consolidado."
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
    if texto is None or pd.isna(texto):
        return ""
    return ''.join(c for c in unicodedata.normalize('NFD', str(texto)) if unicodedata.category(c) != 'Mn').lower()

def normalizar_cadena(texto):
    t = quitar_acentos(texto)
    return re.sub(r'[^a-z0-9]', '', t)

def normalizar_nombre_candidato(nombre):
    t = quitar_acentos(str(nombre).strip())
    t_clean = re.sub(r'[^a-z0-9]', '', t)
    
    alias_map = {
        "pepechedraui": "JOSÉ CHEDRAUI BUDIB",
        "josechedraui": "JOSÉ CHEDRAUI BUDIB",
        "rmvb": "RAFAEL MORENO VALLE BUITRÓN",
        "rafamorenovalle": "RAFAEL MORENO VALLE BUITRÓN",
        "rafaelmorenovalle": "RAFAEL MORENO VALLE BUITRÓN",
        "carolinabeau": "CAROLINA BEAUREGARD MARTÍNEZ",
        "carolinabeauregard": "CAROLINA BEAUREGARD MARTÍNEZ",
        "genovevahuerta": "GENOVEVA HUERTA VILLEGAS",
        "gabrielasanchez": "GABRIELA SÁNCHEZ SAAVEDRA",
        "lauraartemisa": "LAURA ARTEMISA GARCÍA CHÁVEZ",
        "lupitacuautle": "GUADALUPE CUAUTLE TORRES",
        "guadalupecuautle": "GUADALUPE CUAUTLE TORRES",
        "tonantzinfernandez": "TONANTZIN FERNÁNDEZ DÍAZ",
        "lizsanchez": "LIZETH SÁNCHEZ GARCÍA",
        "lizethsanchez": "LIZETH SÁNCHEZ GARCÍA",
        "nestorcamarillo": "NESTOR CAMARILLO MEDINA",
        "celinapena": "CELINA PEÑA GUZMÁN",
        "rodrigoabdala": "RODRIGO ABDALA DARTIGUES",
        "delfinapozos": "DELFINA POZOS VERGARA",
        "xitlalicceja": "XITLALIC CEJA GARCÍA",
        "blancaalcala": "BLANCA ALCALÁ RUIZ"
    }
    for alias, canon in alias_map.items():
        if alias in t_clean or t_clean in alias:
            return canon
    return str(nombre).strip().upper()

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
        file_bytes = file.read()
        file.seek(0)
    except Exception:
        file_bytes = file

    # 1. Intentar como Excel usando openpyxl
    try:
        return pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
    except Exception:
        pass

    # 2. Intentar como CSV UTF-8
    try:
        df_csv = pd.read_csv(io.BytesIO(file_bytes), on_bad_lines="skip")
        return {"Reporte": df_csv}
    except Exception:
        pass

    # 3. Intentar como CSV Latin-1
    try:
        df_csv = pd.read_csv(io.BytesIO(file_bytes), encoding="latin1", on_bad_lines="skip")
        return {"Reporte": df_csv}
    except Exception:
        pass

    return {}

def obtener_campo(row, lista_cols):
    for c in lista_cols:
        if c in row.index:
            val = row[c]
            if val is not None and not pd.isna(val):
                v = str(val).strip()
                if v and v.lower() not in ["nan", "none", "null"]:
                    return v
        for col_existente in row.index:
            col_str = str(col_existente).strip()
            if col_str.lower() == c.lower() or quitar_acentos(col_str) == quitar_acentos(c):
                val = row[col_existente]
                if val is not None and not pd.isna(val):
                    v = str(val).strip()
                    if v and v.lower() not in ["nan", "none", "null"]:
                        return v
    return ""

def obtener_columna_serie(df_data, lista_posibles_cols):
    for c in lista_posibles_cols:
        if c in df_data.columns:
            return df_data[c]
        for col_existente in df_data.columns:
            col_str = str(col_existente).strip()
            if col_str.lower() == c.lower() or quitar_acentos(col_str) == quitar_acentos(c):
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
            elif len(parts) == 3 and len(parts) == 4:
                return datetime(int(parts), int(parts), int(parts[0]))
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

def estandarizar_categoria_medio(m_type_raw):
    if m_type_raw is None or pd.isna(m_type_raw):
        return "PORTALES DIGITALES"
    t_norm = quitar_acentos(str(m_type_raw).strip())
    if "tele" in t_norm or "tv" in t_norm:
        return "TELEVISIÓN"
    elif "rad" in t_norm or "fm" in t_norm or "am" in t_norm:
        return "RADIO"
    elif any(k in t_norm for k in ["portal", "web", "online", "internet", "digital", "comun"]):
        return "PORTALES DIGITALES"
    elif any(k in t_norm for k in ["prensa", "periodico", "diario", "impreso"]):
        return "PRENSA LOCAL"
    elif any(k in t_norm for k in ["columna", "opinion"]):
        return "COLUMNAS"
    elif any(k in t_norm for k in ["redes", "social", "twitter", "facebook", "instagram", "tiktok", "youtube", "x"]):
        return "REDES SOCIALES"
    return "PORTALES DIGITALES"

def reparar_desfase_columnas_excel(df):
    if df is None or df.empty:
        return df
        
    cols = [str(c).strip() for c in df.columns]
    
    es_desfasado_por_hora = False
    if "Hora" in cols:
        sample_hora = df["Hora"].dropna().astype(str).head(10).tolist()
        no_son_horas = any(re.match(r'^(Puebla|M[eé]xico|Tlaxcala|Veracruz|CDMX|Hidalgo|Nacional|Internacional)$', v.strip(), re.I) for v in sample_hora)
        if no_son_horas:
            es_desfasado_por_hora = True
            
    if "Alcance" in cols and not es_desfasado_por_hora:
        sample_alcance = df["Alcance"].dropna().astype(str).head(10).tolist()
        if any(v.startswith("http") for v in sample_alcance):
            es_desfasado_por_hora = True

    if es_desfasado_por_hora:
        columnas_reales_ordenadas = [
            "ID Nota", "Menu", "Titulo", "Autor", "Fecha", 
            "Estado", "Pais", "Nombre del Medio", "Tipo de Medio", 
            "Tipo de Nota", "Sentimiento", "Costo", "Alcance", 
            "Link URL Medio", "Link de Nota"
        ]
        
        df_reparado = pd.DataFrame()
        raw_matrix = df.values
        
        for idx_col, col_name in enumerate(columnas_reales_ordenadas):
            if idx_col < raw_matrix.shape:
                df_reparado[col_name] = raw_matrix[:, idx_col]
            else:
                df_reparado[col_name] = ""
                
        df_reparado["Hora"] = ""
        return df_reparado
        
    return df

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
            t = str(item.get("texto", "")).lower()
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
        sent_col_name = None
        for col_name in df_data.columns:
            if quitar_acentos(str(col_name)) in ["sentimiento", "sentiment", "sentimiento de la nota", "tono", "sentimiento nota"]:
                sent_col_name = col_name
                break
        if sent_col_name:
            sent_series = df_data[sent_col_name].fillna("").astype(str).str.lower()
            return sent_series.apply(lambda s: "NEGATIVA" if any(k in s for k in ["negat", "critica", "contra"]) else "POSITIVA").tolist()

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
            texto = obtener_campo(row, ["Contenido", "Detail", "Titulo", "Título", "Summary", "Síntesis", "Sintesis", "Title", "Encabezado", "Tema", "Nota"])
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
    columnas_posibles_texto = ["Titulo", "Título", "Contenido", "Detail", "Summary", "Síntesis", "Sintesis", "Encabezado", "Tema", "Nota", "Title"]
    
    for col in columnas_posibles_texto:
        if col in pos_df.columns or any(quitar_acentos(str(c)) == quitar_acentos(col) for c in pos_df.columns):
            serie_t = obtener_columna_serie(pos_df, [col])
            pos_textos = [t for t in serie_t.dropna().astype(str).tolist() if "teaser" not in t.lower() and len(t.strip()) > 10]
            if len(pos_textos) > 0:
                break

    neg_textos = []
    for col in columnas_posibles_texto:
        if col in neg_df.columns or any(quitar_acentos(str(c)) == quitar_acentos(col) for c in neg_df.columns):
            serie_t = obtener_columna_serie(neg_df, [col])
            neg_textos = [t for t in serie_t.dropna().astype(str).tolist() if "teaser" not in t.lower() and len(t.strip()) > 10]
            if len(neg_textos) > 0:
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

def obtener_link_inteligente(row):
    # 1. Prioridad para Portales / Web / Redes: Link directo del medio si existe
    link_url_medio = obtener_campo(row, ["Link URL Medio", "URL Medio", "Link Medio", "URL"])
    if link_url_medio and str(link_url_medio).startswith("http") and "hanakua.mx/Testigo" not in str(link_url_medio):
        return str(link_url_medio).strip()
        
    # 2. Prioridad para Radio / TV: Link de la nota en Hanakua ([https://next.hanakua.mx/Notas?id=](https://next.hanakua.mx/Notas?id=)...)
    link_de_nota = obtener_campo(row, ["Link de Nota", "Link Nota", "URL Nota"])
    if link_de_nota and str(link_de_nota).startswith("http") and "hanakua.mx/Testigo" not in str(link_de_nota):
        return str(link_de_nota).strip()
        
    # 3. Respaldo general (excluyendo siempre Testigo)
    link_generico = obtener_campo(row, ["Enlace", "Link"])
    if link_generico and str(link_generico).startswith("http") and "hanakua.mx/Testigo" not in str(link_generico):
        return str(link_generico).strip()
        
    return ""

def crear_doc_desde_hoja(df_hoja, nombre_hoja, es_redes_sociales):
    if df_hoja is None or df_hoja.empty:
        return None

    # Auto-reparar posibles desfasamientos humanos de columnas en Excel
    df_hoja = reparar_desfase_columnas_excel(df_hoja)

    texto_primer_renglon = " ".join([str(v) for v in df_hoja.iloc[0].dropna().values]).strip() if len(df_hoja) > 0 else ""
    texto_columnas = " ".join([str(c) for c in df_hoja.columns]).strip()

    if "sin notas" in texto_primer_renglon.lower() or "sin notas" in texto_columnas.lower():
        return None

    if es_redes_sociales:
        df_filtrado, total_descartadas = limpiar_dataframe_redes_automatico(df_hoja, nombre_hoja)
        if total_descartadas > 0:
            st.info(f"ℹ️ Se limpiaron automáticamente {total_descartadas} publicaciones institucionales/personales. Analizando {len(df_filtrado)} menciones ciudadanas reales.")
    else:
        df_filtrado = df_hoja.copy()

    if len(df_filtrado) == 0:
        return None

    serie_fechas_raw = obtener_columna_serie(
        df_filtrado, ["Publish date", "Fecha", "Date", "Fecha de publicación", "Fecha de publicacion"]
    )
    df_filtrado["fecha_dt"] = serie_fechas_raw.apply(parsear_fecha_perfecta)

    df_filtrado = (
        df_filtrado
        .dropna(subset=["fecha_dt"])
        .sort_values(by="fecha_dt", ascending=True)
    )

    if len(df_filtrado) == 0:
        return None

    # Descartar duplicados si se fusionaron múltiples archivos
    subset_dup = [c for c in ["ID Nota", "Titulo", "Link de Nota", "Link URL Medio"] if c in df_filtrado.columns]
    if len(subset_dup) > 0:
        df_filtrado = df_filtrado.drop_duplicates(subset=subset_dup)

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

    # Identificación estandarizada de medio para tradicionales
    serie_media_raw = obtener_columna_serie(
        df_filtrado,
        ["Tipo de Medio", "Fuente", "Media type", "Media Type", "Medio", "Nombre del Medio", "Canal", "Tipo de Nota"]
    )
    df_filtrado["categoria_medio_std"] = serie_media_raw.apply(estandarizar_categoria_medio) if not es_redes_sociales else "REDES SOCIALES"

    tv_cnt = (df_filtrado["categoria_medio_std"] == "TELEVISIÓN").sum() if not es_redes_sociales else 0
    rad_cnt = (df_filtrado["categoria_medio_std"] == "RADIO").sum() if not es_redes_sociales else 0
    portales_cnt = (df_filtrado["categoria_medio_std"] == "PORTALES DIGITALES").sum() if not es_redes_sociales else 0
    prensa_cnt = (df_filtrado["categoria_medio_std"] == "PRENSA LOCAL").sum() if not es_redes_sociales else 0
    columnas_cnt = (df_filtrado["categoria_medio_std"] == "COLUMNAS").sum() if not es_redes_sociales else 0

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

    orden_medios_oficial = ["TELEVISIÓN", "RADIO", "PORTALES DIGITALES", "PRENSA LOCAL", "COLUMNAS", "REDES SOCIALES"]

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
                    detalle = obtener_campo(row, ["Contenido", "Detail", "Summary", "Síntesis", "Sintesis", "Titulo", "Título", "Title", "Encabezado"])
                    link = obtener_link_inteligente(row)

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

            if len(neg_df) > 0:
                p_neg_hdr = doc.add_paragraph()
                p_neg_hdr.paragraph_format.space_before = Pt(6)
                p_neg_hdr.paragraph_format.space_after = Pt(4)
                add_run_verdana(p_neg_hdr, f"NEGATIVAS: {len(neg_df)}", bold=True, size_pt=10, color_rgb=RGBColor(180, 0, 0))

                for _, row in neg_df.iterrows():
                    autor = obtener_campo(row, ["Autor", "Author name", "Fuente", "Media name", "Programa"])
                    handle = obtener_campo(row, ["Author handle (@username)", "Handle", "Username"])
                    detalle = obtener_campo(row, ["Contenido", "Detail", "Summary", "Síntesis", "Sintesis", "Titulo", "Título", "Title", "Encabezado"])
                    link = obtener_link_inteligente(row)

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
            # MEDIOS TRADICIONALES - POSITIVAS AGRUPADAS POR CANAL
            if len(pos_df) > 0:
                for cat_nombre in orden_medios_oficial:
                    sub_pos_cat = pos_df[pos_df["categoria_medio_std"] == cat_nombre]
                    if len(sub_pos_cat) > 0:
                        p_m = doc.add_paragraph()
                        p_m.paragraph_format.space_before = Pt(4)
                        p_m.paragraph_format.space_after = Pt(4)
                        add_run_verdana(p_m, f"{cat_nombre}: {len(sub_pos_cat)}", bold=True, size_pt=10)

                        for _, row in sub_pos_cat.iterrows():
                            medio = obtener_campo(row, ["Nombre del Medio", "Fuente", "Media name", "Medio"])
                            autor = obtener_campo(row, ["Autor", "Author name", "Programa", "Conductor"])
                            hora = obtener_campo(row, ["Hora", "Hour", "Time", "Hora de Transmisión", "Hora de Transmision"])
                            titulo = obtener_campo(row, ["Titulo", "Título", "Contenido", "Detail", "Summary", "Síntesis", "Sintesis", "Encabezado", "Nota"])
                            link = obtener_link_inteligente(row)

                            p_a = doc.add_paragraph()
                            p_a.paragraph_format.space_before = Pt(4)
                            p_a.paragraph_format.space_after = Pt(1)
                            cabecera = f"{medio} - {autor}" if (autor and autor not in ["Redacción", "Staff", "Online", medio]) else medio
                            add_run_verdana(p_a, cabecera, bold=True, size_pt=10)

                            cuerpo_texto = limpiar_texto(titulo)
                            if hora and not cuerpo_texto.lower().startswith(hora.lower()):
                                cuerpo_texto = f"{hora} {cuerpo_texto}".strip()

                            p_d = doc.add_paragraph()
                            p_d.paragraph_format.space_after = Pt(2)
                            add_run_verdana(p_d, cuerpo_texto, bold=False, size_pt=9.5)

                            if link:
                                p_l = doc.add_paragraph()
                                p_l.paragraph_format.space_after = Pt(6)
                                add_run_verdana(p_l, link, bold=False, size_pt=9, color_rgb=RGBColor(0, 102, 204), underline=True)

            # MEDIOS TRADICIONALES - NEGATIVAS AGRUPADAS POR CANAL
            if len(neg_df) > 0:
                p_neg_hdr = doc.add_paragraph()
                p_neg_hdr.paragraph_format.space_before = Pt(6)
                p_neg_hdr.paragraph_format.space_after = Pt(4)
                add_run_verdana(p_neg_hdr, f"NEGATIVAS: {len(neg_df)}", bold=True, size_pt=10, color_rgb=RGBColor(180, 0, 0))

                for cat_nombre in orden_medios_oficial:
                    sub_neg_cat = neg_df[neg_df["categoria_medio_std"] == cat_nombre]
                    if len(sub_neg_cat) > 0:
                        p_sub_neg = doc.add_paragraph()
                        p_sub_neg.paragraph_format.space_before = Pt(4)
                        p_sub_neg.paragraph_format.space_after = Pt(4)
                        add_run_verdana(p_sub_neg, f"{cat_nombre}: {len(sub_neg_cat)}", bold=True, size_pt=10, color_rgb=RGBColor(180, 0, 0))

                        for _, row in sub_neg_cat.iterrows():
                            medio = obtener_campo(row, ["Nombre del Medio", "Fuente", "Media name", "Medio"])
                            autor = obtener_campo(row, ["Autor", "Author name", "Programa", "Conductor"])
                            hora = obtener_campo(row, ["Hora", "Hour", "Time", "Hora de Transmisión", "Hora de Transmision"])
                            titulo = obtener_campo(row, ["Titulo", "Título", "Contenido", "Detail", "Summary", "Síntesis", "Sintesis", "Encabezado", "Nota"])
                            link = obtener_link_inteligente(row)

                            p_a = doc.add_paragraph()
                            p_a.paragraph_format.space_before = Pt(4)
                            p_a.paragraph_format.space_after = Pt(1)
                            cabecera = f"{medio} - {autor}" if (autor and autor not in ["Redacción", "Staff", "Online", medio]) else medio
                            add_run_verdana(p_a, cabecera, bold=True, size_pt=10)

                            cuerpo_texto = limpiar_texto(titulo)
                            if hora and not cuerpo_texto.lower().startswith(hora.lower()):
                                cuerpo_texto = f"{hora} {cuerpo_texto}".strip()

                            p_d = doc.add_paragraph()
                            p_d.paragraph_format.space_after = Pt(2)
                            add_run_verdana(p_d, cuerpo_texto, bold=False, size_pt=9.5)

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
        "Medios Tradicionales / Portales / TV y Radio (Multi-Archivo)"
    ],
    index=1
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
    uploaded_files = st.file_uploader(
        "Sube uno o varios archivos Excel (ej. Archivo de TV/Radio y Archivo de Portales Web)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True
    )

    if uploaded_files and len(uploaded_files) > 0:
        candidatos_unificados = {}

        for file_item in uploaded_files:
            try:
                dict_hojas = cargar_archivo_seguro(file_item)
                for h_name, df_h in dict_hojas.items():
                    if df_h is None or df_h.empty:
                        continue
                    
                    if 'Menu' in df_h.columns and len(df_h['Menu'].dropna()) > 0:
                        nombre_raw = str(df_h['Menu'].dropna().iloc[0]).strip()
                    else:
                        nombre_raw = h_name
                    
                    candidato_canon = normalizar_nombre_candidato(nombre_raw)
                    
                    if candidato_canon not in candidatos_unificados:
                        candidatos_unificados[candidato_canon] = []
                    candidatos_unificados[candidato_canon].append(df_h)
            except Exception as e:
                st.warning(f"No se pudo procesar una de las hojas de {file_item.name}: {str(e)}")

        candidatos_disponibles = sorted([c for c in candidatos_unificados.keys() if c and "SIN NOTAS" not in c])

        if len(candidatos_disponibles) > 0:
            st.success(f"✅ Se cargaron exitosamente {len(uploaded_files)} archivo(s) y se detectaron {len(candidatos_disponibles)} candidatos.")
            st.markdown(f"**Candidatos detectados:** {', '.join(candidatos_disponibles)}")

            st.write("---")
            st.subheader("Opciones de Descarga:")

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("### 📦 Descarga Masiva")
                if st.button("Generar y Descargar TODOS los Reportes en .ZIP", type="primary", use_container_width=True):
                    with st.spinner("Generando reportes consolidados para todos los candidatos..."):
                        zip_buffer = io.BytesIO()
                        cnt_generados = 0
                        
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for cand_name in candidatos_disponibles:
                                lista_dfs = candidatos_unificados[cand_name]
                                df_total_candidato = pd.concat(lista_dfs, ignore_index=True)

                                buf = crear_doc_desde_hoja(df_total_candidato, cand_name, es_redes_sociales=False)
                                if buf is not None:
                                    doc_bytes = buf.getvalue()
                                    fname = f"Reporte_{cand_name.replace(' ', '_')}.docx"
                                    zip_file.writestr(fname, doc_bytes)
                                    cnt_generados += 1

                        zip_buffer.seek(0)
                        if cnt_generados > 0:
                            st.success(f"¡Se generaron con éxito {cnt_generados} reportes consolidados!")
                            st.download_button(
                                label="📥 Descargar Archivo .ZIP",
                                data=zip_buffer,
                                file_name="Reportes_Monitoreo_Consolidados_Completos.zip",
                                mime="application/zip",
                                use_container_width=True
                            )
                        else:
                            st.warning("No se encontraron candidatos con notas activas.")

            with col2:
                st.markdown("### 📄 Descarga Individual")
                cand_sel = st.selectbox("Selecciona un candidato:", candidatos_disponibles)
                if st.button(f"Generar Reporte de {cand_sel}", use_container_width=True):
                    with st.spinner(f"Consolidando notas para {cand_sel}..."):
                        lista_dfs = candidatos_unificados[cand_sel]
                        df_total_candidato = pd.concat(lista_dfs, ignore_index=True)
                        
                        buf = crear_doc_desde_hoja(df_total_candidato, cand_sel, es_redes_sociales=False)
                        if buf is not None:
                            st.success(f"¡Reporte consolidado listo para '{cand_sel}'!")
                            st.download_button(
                                label=f"📥 Descargar Word de {cand_sel}",
                                data=buf,
                                file_name=f"Reporte_{cand_sel.replace(' ', '_')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                        else:
                            st.warning(f"El candidato '{cand_sel}' no contiene notas registradas.")
        else:
            st.warning("No se detectaron candidatos con notas válidas en los archivos seleccionados.")
