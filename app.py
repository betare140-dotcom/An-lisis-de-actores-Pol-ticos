# --- PATRONES INSTITUCIONALES EXPANDIDOS ---
PATRONES_INSTITUCIONALES_GENERICOS = [
    r'ayuntamiento', r'gobierno', r'gobpue', r'gobiernoded', r'ayto', r'snandresoficial', r'cholulaoficial',
    r'secretaria', r'dependencia', r'organismo', r'sindicatura', r'presidencia', r'comunicacionsocial',
    r'seguridadciudadana', r'seguridadpublica', r'proteccioncivil', r'policiamunicipal', r'policiastat',
    r'serviciospublicos', r'servicios_pub', r'obraspublicas', r'desarrollourbano', r'desarrolloeconomico',
    r'dif', r'sistemadif', r'difestatal', r'difmunicipal', r'institutodelamujer', r'institutodelajuventud',
    r'organismodeagua', r'serviciodelimpia', r'organismolimpia', r'derechoshumanos', r'casadecultura',
    r'bienestarsanandres', r'bienestarcholula', r'ssppc', r'ssppcsnandres'
]

ABREV_APELLIDOS = {
    'fdz': 'fernandez', 'hdez': 'hernandez', 'mtz': 'martinez', 'glez': 'gonzalez',
    'gcia': 'garcia', 'lpz': 'lopez', 'prz': 'perez', 'sdo': 'saavedra', 'chvz': 'chavez'
}

def es_cuenta_del_actor_universal(autor, handle, actor_target):
    aut = normalizar_cadena(autor)
    hnd = normalizar_cadena(handle)
    act = normalizar_cadena(actor_target)
    
    if not act:
        return False
        
    if act in aut or act in hnd or aut in act:
        return True
        
    tokens_actor = [t for t in re.findall(r'\w+', quitar_acentos(actor_target)) if len(t) > 2]
    for abrev, apellido_completo in ABREV_APELLIDOS.items():
        if apellido_completo in tokens_actor:
            tokens_actor.append(abrev)
            
    nombres_distintivos = [t for t in tokens_actor if len(t) >= 5 and t not in ['perez', 'lopez', 'garcia', 'martinez']]
    for n in nombres_distintivos:
        if n in aut or n in hnd:
            otros_tokens = [t for t in tokens_actor if t != n]
            if any(ot in aut or ot in hnd for ot in otros_tokens):
                return True
            if hnd == n or aut == n:
                return True
                
    coincidencias_aut = sum(1 for t in tokens_actor if t in aut)
    coincidencias_hnd = sum(1 for t in tokens_actor if t in hnd)
    if coincidencias_aut >= 2 or coincidencias_hnd >= 2:
        return True
        
    return False

def es_institucional_universal(autor, handle, texto_post=""):
    texto = (str(autor) + " " + str(handle))
    texto_norm = normalizar_cadena(texto)
    for p in PATRONES_INSTITUCIONALES_GENERICOS:
        if re.search(p, texto_norm):
            return True
            
    t_clean = quitar_acentos(texto_post).lower()
    frases_institucionales = [
        'nuestra presidenta municipal', 'nuestra presidenta honoraria', 'nuestro presidente municipal',
        'mi mas sincero agradecimiento', 'nuestro compromiso de seguir', 'desde el gobierno municipal',
        'agradecemos a nuestra presidenta', 'agradecemos a nuestro presidente', 'los invitamos a participar en nuestro'
    ]
    if any(f in t_clean for f in frases_institucionales):
        return True
        
    return False

def limpiar_dataframe_redes_automatico(df_raw, actor_nombre_target):
    def es_descartable(row):
        autor = obtener_campo(row, ["Autor", "Author name", "Author", "User", "Username", "Handle", "Fuente", "Nombre del Medio", "Canal"])
        handle = obtener_campo(row, ["Author handle (@username)", "Handle", "Username", "Screen Name", "Account", "Perfil"])
        detalle = obtener_campo(row, ["Contenido", "Detail", "Titulo", "Título", "Summary", "Síntesis", "Nota"])
        
        if es_cuenta_del_actor_universal(autor, handle, actor_nombre_target):
            return True
            
        if es_institucional_universal(autor, handle, detalle):
            return True
        
        det_lower = quitar_acentos(detalle)
        if any(p in det_lower for p in ['mis ahijados', 'con toda la actitud #graciasdios', 'primeracomunion', 'en familia festejando']):
            return True
            
        return False

    mask_descarte = df_raw.apply(es_descartable, axis=1)
    df_limpio = df_raw[~mask_descarte].copy()
    total_descartadas = mask_descarte.sum()
    
    return df_limpio, total_descartadas
