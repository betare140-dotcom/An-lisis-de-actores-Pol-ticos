def limpiar_dataframe_redes_automatico(df_raw, actor_nombre_target):
    def es_descartable(row):
        autor = obtener_campo(row, ["Autor", "Author name", "Author", "User", "Username", "Handle", "Fuente", "Nombre del Medio", "Canal"])
        handle = obtener_campo(row, ["Author handle (@username)", "Handle", "Username", "Screen Name", "Account", "Perfil"])
        detalle = obtener_campo(row, ["Contenido", "Detail", "Titulo", "Título", "Summary", "Síntesis", "Nota"])
        link = obtener_campo(row, ["Link URL Medio", "URL", "Enlace", "Link", "Link de Nota"])
        
        # 1. Comprobar cuenta del actor político
        if es_cuenta_del_actor_universal(autor, handle, actor_nombre_target):
            return True
            
        # 2. Comprobar cuenta institucional / DIF / Ayuntamiento
        if es_institucional_universal(autor, handle, detalle):
            return True
            
        # 3. Comprobar si el link del post apunta al perfil propio del candidato
        if link and actor_nombre_target:
            link_norm = normalizar_cadena(link)
            tokens_actor = [t for t in re.findall(r'\w+', quitar_acentos(actor_nombre_target)) if len(t) >= 5]
            for n in tokens_actor:
                if f"x.com/{n}" in link.lower() or f"twitter.com/{n}" in link.lower() or f"facebook.com/{n}" in link.lower() or f"{n}fdz" in link_norm or f"{n}t" in link_norm:
                    return True
        
        det_lower = quitar_acentos(detalle)
        if any(p in det_lower for p in ['mis ahijados', 'con toda la actitud #graciasdios', 'primeracomunion', 'en familia festejando']):
            return True
            
        return False

    mask_descarte = df_raw.apply(es_descartable, axis=1)
    df_limpio = df_raw[~mask_descarte].copy()
    total_descartadas = mask_descarte.sum()
    
    return df_limpio, total_descartadas


def crear_doc_desde_hoja(df_hoja, nombre_hoja, es_redes_sociales):
    if df_hoja is None or df_hoja.empty:
        return None

    # Auto-reparar posibles desfasamientos humanos de columnas en Excel
    df_hoja = reparar_desfase_columnas_excel(df_hoja)

    # 1. Eliminar filas con leyendas tipo 'sin notas'
    mask_sin_notas = df_hoja.apply(lambda r: any(k in str(v).lower() for v in r.values for k in ['sin notas', 'si notas', 'sin nota', 'sin registro']), axis=1)
    df_hoja = df_hoja[~mask_sin_notas].copy()
    if len(df_hoja) == 0:
        return None

    # 2. Limpieza universal de cuentas oficiales y personales del actor político
    df_filtrado, total_descartadas = limpiar_dataframe_redes_automatico(df_hoja, nombre_hoja)

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

    # Identificación estandarizada de medio
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

        if es_redes_sociales or (len(pos_df) > 0 and pos_df["categoria_medio_std"].iloc[0] == "REDES SOCIALES"):
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
            # MEDIOS TRADICIONALES
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
