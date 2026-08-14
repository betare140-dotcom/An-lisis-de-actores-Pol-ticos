import json
import re
import google.generativeai as genai

# 1. Configuración del Modelo con Temperatura CERO e Instrucción de Sistema
SYSTEM_PROMPT_POLITICO = """
Eres un analista senior de inteligencia política, reputación gubernamental y monitoreo de medios en el Estado de Puebla.

Tu misión es evaluar el impacto reputacional de cada nota para el actor político objetivo.

REGLAS ESTRICTAS DE CLASIFICACIÓN POLÍTICA:
1. NEGATIVA:
   - Seguridad y Nota Roja: Asaltos, cristalazos, robo de autopartes, homicidios o balaceras ocurridas en su municipio o atribuibles a falta de vigilancia.
   - Corrupción y Fiscalización: Auditorías de la ASE, desvíos de recursos, presunto daño patrimonial, señalamientos de nepotismo o falta de transparencia.
   - Crisis Política: Desplantes institucionales, confrontaciones, controversias por bardas/lonas, o quejas ciudadanas por baches, basura, agua o socavones.
   - Críticas Ciudadanas: Comentarios de usuarios, periodistas o columnistas cuestionando su desempeño, capacidad o ética.

2. POSITIVA / INFORMATIVA:
   - Programas y Bienestar: Entrega de apoyos alimentarios, medicamentos, kits escolares, becas, desayunadores o atención médica.
   - Obras y Gobierno: Pavimentación, luminarias, convenios de colaboración (CANACO, OCDE, universidades), hermanamientos y sesiones de cabildo.
   - Agenda Cultural y Deportiva: Foros ('Ser Mujer'), ferias tradicionales ('Feria del Queso'), carreras y eventos juveniles.
   - Posicionamiento Favorable: Encuestas de aprobación, liderazgo interno y notas informativas descriptivas de su gestión.
"""

genai.configure(api_key="AQ.Ab8RN6LoOHgBblHSIETp2LjyBofO48YsSqSeojXYFAAKGvFa0w")

# Modelo con System Instruction y Temperature 0.0
model_politico = genai.GenerativeModel(
    model_name="gemini-1.5-flash",
    system_instruction=SYSTEM_PROMPT_POLITICO,
    generation_config={"temperature": 0.0}
)

def clasificar_lote_con_ia(lista_notas, actor_nombre):
    """
    Clasifica un lote de hasta 20 notas en una sola llamada rápida y precisa.
    lista_notas es una lista de diccionarios: [{'id': 0, 'texto': '...'}, ...]
    """
    prompt = f"""
Clasifica las siguientes notas respecto al actor político: "{actor_nombre}".

NOTAS A EVALUAR:
{json.dumps(lista_notas, ensure_ascii=False)}

Responde ÚNICAMENTE un JSON válido con la lista de resultados en este formato exacto:
[
  {{"id": 0, "sentimiento": "POSITIVA"}},
  {{"id": 1, "sentimiento": "NEGATIVA"}}
]
"""
    try:
        response = model_politico.generate_content(prompt)
        raw_txt = response.text.strip()
        # Limpiar posibles bloques ```json ```
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
        # Fallback de respaldo basado en palabras críticas clave
        res_map = {}
        for item in lista_notas:
            t = item["texto"].lower()
            if any(k in t for k in ["asalto", "robo", "daño patrimonial", "317 millones", "ase", "nepotismo", "desfalco", "desplante", "inseguridad", "bache"]):
                res_map[item["id"]] = "NEGATIVA"
            else:
                res_map[item["id"]] = "POSITIVA"
        return res_map
