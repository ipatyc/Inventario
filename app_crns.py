
# Las reglas específicas pendientes deberán incorporarse aquí.

def auditar_restricciones_internas(df):
    incidencias = []
    for i, fila in df.reset_index(drop=True).iterrows():
        periodo = limpiar_clave_texto(fila.get("Periodo"))
        nivel = normalizar_para_cruce(fila.get("Nivel"))
        estatus = normalizar_para_cruce(fila.get("Estatus"))

        if not periodo: incidencias.append({"Fila": i + 1, "Tipo": "Periodo", "Detalle": "Periodo vacío."})
        if nivel not in ("LICENCIATURA", "POSGRADO", "BACHILLERATO"):
            incidencias.append({"Fila": i + 1, "Tipo": "Nivel", "Detalle": f"Nivel no reconocido: {nivel}"})
        if estatus and estatus not in ("A", "R"):
            incidencias.append({"Fila": i + 1, "Tipo": "Estatus", "Detalle": f"Estatus no reconocido: {estatus}"})

    return pd.DataFrame(incidencias, columns=["Fila", "Tipo", "Detalle"])

# ============================================================
# 9. CONSERVAR EXCEL ORIGINALES Y MACROS
# ============================================================
def crear_solicitud_excel(df_grupo, archivos_originales):
    origenes = df_grupo["ArchivoOrigen"].dropna().unique().tolist()
    if len(origenes) != 1: raise ValueError("No se pueden fusionar libros con macros de distintos orígenes.")
    origen = origenes[0]
    if origen not in archivos_originales: raise ValueError(f"No se encontró el archivo original: {origen}")

    extension = ".xlsm" if origen.lower().endswith(".xlsm") else ".xlsx"
    wb = openpyxl.load_workbook(io.BytesIO(archivos_originales[origen]), keep_vba=(extension == ".xlsm"))
    hoja = next((h for h in wb.sheetnames if h.strip().upper() == HOJA_ALTAS), None)
    if hoja is None: raise ValueError(f"{origen}: falta la hoja ALTAS.")
    ws = wb[hoja]

    encabezados = {normalizar_para_busqueda(c.value): c.column for c in ws[1] if c.value is not None}

    campos = {
        "Subject": ["subj", "subject", "area"],
        "Course": ["course", "crse", "nocurso"],
        "Sección": ["seccion", "grupo"],
        "Tipo de Horario": ["tipohorario"],
        "Método Educativo": ["metodoeducativo"],
        "Modo de Calificar": ["modocalificar", "mododecalificar"],
        "Parte de Periodo": ["parteperiodo"],
        "Estatus": ["status", "estatus"],
        "Capacidad": ["capacidad", "cupo"],
        "Sesion": ["sesion"],
        "Campus": ["sede", "campus"]
    }


    for _, fila in df_grupo.iterrows():
        numero = fila.get("_FilaExcel")
        if pd.isna(numero): raise ValueError(f"{origen}: falta _FilaExcel.")
        for campo, alternativas in campos.items():
            columna = next((encabezados[a] for a in alternativas if a in encabezados), None)
            if columna is not None and campo in fila.index:
                valor = fila[campo]
                ws.cell(int(numero), columna).value = None if pd.isna(valor) else str(valor)

    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue(), extension

# ============================================================
# 10. GENERAR ZIP SOLICITUDES + CSV_ALTAS
# ============================================================
def generar_zip_altas(df_final, archivos_originales, funcion_csv):
    requeridas = ["Periodo", "Responsable", "Clúster", "ArchivoOrigen", "_FilaExcel"]
    faltantes = [c for c in requeridas if c not in df_final.columns]
    if faltantes: raise ValueError(f"Faltan columnas para exportar: {faltantes}")

    df = df_final.copy()
    df["Periodo"] = df["Periodo"].apply(limpiar_clave_texto)
    df["Responsable"] = df["Responsable"].apply(responsable_normalizado)
    if df["Periodo"].eq("").any() or df["Responsable"].eq("").any() or df["Clúster"].apply(limpiar_clave_texto).eq("").any():
        raise ValueError("Hay ALTAS sin Periodo, Responsable o Clúster.")

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        # Un Excel por archivo original. Evitar colisiones con consecutivos.
        nombres = set()
        for origen, grupo in df.groupby("ArchivoOrigen", sort=False):
            contenido, extension = crear_solicitud_excel(grupo, archivos_originales)
            nombre = nombre_solicitud(grupo, extension)
            ruta, consecutivo = f"SOLICITUDES/{nombre}", 2
            while ruta.casefold() in nombres:
                ruta = f"SOLICITUDES/{nombre[:-len(extension)]} ({consecutivo}){extension}"
                consecutivo += 1
            z.writestr(ruta, contenido)
            nombres.add(ruta.casefold())

        # CSV_ALTAS: un CSV por periodo y responsable, aunque provenga de varios Excel.
        for (periodo, responsable), grupo in df.groupby(["Periodo", "Responsable"], sort=False):
            nombre = nombre_csv_altas(periodo, responsable)
            contenido = funcion_csv(grupo)
            if isinstance(contenido, str): contenido = contenido.encode("utf-8")
            z.writestr(f"CSV_ALTAS/{nombre}", contenido)

    return buffer.getvalue()

# ============================================================
# 11. ESTADOS DE STREAMLIT
# ============================================================
ESTADOS_INICIALES = {
    "original_files_bytes": {}, "res_auditoria": None, "raw_altas": None,
    "ready_for_download": False, "zip_file_bytes": None, "csv_files_to_download": {},
    "csv_consolidado_bytes": None, "delta_files": {}, "final_argos_zip": None,
    "df_cruce_rapido": None, "df_delta_cache": None, "nombre_delta_cache": None,
    "llave_control_archivos": "", "archivo_final_bytes": None, "archivo_final_nombre": None,
    "cat_avanzado_cache": None, "cat_avanzado_firma": None, "indice_nombres_avanzado": None,
    "cat_pa_cache": None, "advertencias_pa": [], "errores_restricciones": None,
    "df_secciones_corregidas": None, "errores_secciones": None, "secciones_validadas": False,
    "zip_solicitudes_csv_altas": None,
    "df_manual_fijo": None, "manual_candidatos": [], "manual_busqueda_realizada": False
}
for clave, valor in ESTADOS_INICIALES.items():
    if clave not in st.session_state: st.session_state[clave] = valor

# ============================================================
# 12. CONFIGURACIÓN VISUAL
# ============================================================
st.set_page_config(page_title="Consola Iris Cavazos", page_icon="⚙️", layout="wide")
st.title("⚙️ Consola de Control de Materias e Inyección de NRCs")
st.markdown("---")

tab1, tab_err, tab3 = st.tabs([
    "1️⃣ Proceso: Validación y Generar CSVs",
    "⚠️ Reporte de Errores (Extraer Delta)",
    "2️⃣ Proceso: Inyección de NRCs Masiva (ARGOS)"


metodo = st.selectbox("Método educativo", opciones_metodo, key="manual_t1_metodo")
            sede = st.text_input("Sede", key="manual_t1_sede")
            estatus = st.text_input("Estatus", key="manual_t1_estatus")
            modo = st.selectbox("Modo de calificar", [""] + sorted(info["gmod"]), key="manual_t1_modo")
            sesion = st.text_input("Sesión", key="manual_t1_sesion")

        e1, e2 = st.columns(2)
        nivel = e1.selectbox("Nivel", ["LICENCIATURA", "BACHILLERATO", "POSGRADO"], key="manual_t1_nivel")
        cluster = e2.selectbox("Clúster", CLUSTERS_PERMITIDOS, index=None, key="manual_t1_cluster")

        if st.button("➕ Agregar materia", key="manual_t1_agregar"):
            if not confirmar or not cluster or not periodo or not responsable:
                st.warning("Confirma la materia, periodo, responsable y clúster.")
            else:
                nuevo = dict(zip(COLUMNAS_CSV, [
                    periodo, sede, llave[0], llave[1], parte, estatus, capacidad, "1",
                    seccion, horario, metodo, "D2L", modo, sesion,
                    "Bachillerato" if nivel == "BACHILLERATO" else cluster
                ]))
                nuevo.update({"Responsable": responsable_normalizado(responsable), "Nivel": nivel})
                st.session_state.df_manual_fijo = pd.concat(
                    [st.session_state.df_manual_fijo, pd.DataFrame([nuevo])], ignore_index=True)
                st.session_state.manual_t1_editor_version = st.session_state.get("manual_t1_editor_version", 0) + 1
                st.success("Materia agregada.")

    # 12. TABLA MANUAL EDITABLE
    st.subheader("📋 Tabla manual")
    if "manual_t1_editor_version" not in st.session_state: st.session_state.manual_t1_editor_version = 0
    df_manual = st.session_state.df_manual_fijo.copy()
    df_manual.insert(0, "Seleccionar", False)

    with st.form("form_tabla_manual_t1"):
        editado = st.data_editor(df_manual, hide_index=True, num_rows="fixed", use_container_width=True,
                                 key=f"manual_editor_t1_{st.session_state.manual_t1_editor_version}")
        guardar_manual = st.form_submit_button("💾 Guardar cambios de la tabla", use_container_width=True)

    if guardar_manual:
        st.session_state.df_manual_fijo = editado.drop(columns="Seleccionar").copy()
        st.session_state.manual_t1_seleccion = editado.index[editado["Seleccionar"]].tolist()
        st.session_state.manual_t1_editor_version += 1
        st.success("Cambios guardados.")
        st.rerun()

    seleccion = st.session_state.get("manual_t1_seleccion", [])
    a1, a2, a3 = st.columns(3)

    if a1.button("Nuevo renglón", key="manual_t1_nuevo"):
        nuevo = {c: "" for c in st.session_state.df_manual_fijo.columns}
        nuevo.update({"GRUPOS": "1", "SOCIODEINTEGRACION": "D2L"})
        st.session_state.df_manual_fijo = pd.concat(
            [st.session_state.df_manual_fijo, pd.DataFrame([nuevo])], ignore_index=True)
        st.session_state.manual_t1_editor_version += 1
        st.rerun()

    if a2.button("Copiar seleccionados", key="manual_t1_copiar"):
        if seleccion:
            copias = st.session_state.df_manual_fijo.iloc[seleccion].copy()
            st.session_state.df_manual_fijo = pd.concat(
                [st.session_state.df_manual_fijo, copias], ignore_index=True)
            st.session_state.manual_t1_editor_version += 1
            st.session_state.manual_t1_seleccion = []
            st.rerun()
        else: st.warning("Primero selecciona y guarda los renglones.")

    if a3.button("Eliminar seleccionados", key="manual_t1_eliminar"):
        if seleccion:
            st.session_state.df_manual_fijo = st.session_state.df_manual_fijo.drop(
                index=seleccion).reset_index(drop=True)
            st.session_state.manual_t1_editor_version += 1
            st.session_state.manual_t1_seleccion = []
            st.rerun()
        else: st.warning("Primero selecciona y guarda los renglones.")

    # 13. CSV MANUAL
    df_out = st.session_state.df_manual_fijo.copy()
    if not df_out.empty:
        if df_out["PERIODO"].apply(limpiar_clave_texto).eq("").any() or df_out["Responsable"].apply(limpiar_clave_texto).eq("").any():
            st.warning("Hay registros manuales sin Periodo o Responsable.")
        else:
            temporal = pd.DataFrame({
                "Periodo": df_out["PERIODO"], "Subject": df_out["SUBJ"], "Course": df_out["COURSE"],
                "Nivel": df_out["Nivel"], "Sección": df_out["SECCION"]
            })
            temporal, errores_manual = asignar_secciones_globales(temporal, st.session_state.get("cat_pa_cache"))
            df_out["SECCION"] = temporal["Sección"].values

            if not errores_manual.empty:
                st.error("Hay errores de secciones en los registros manuales.")
                st.dataframe(errores_manual, hide_index=True)
            else:
                df_out["Responsable"] = df_out["Responsable"].apply(responsable_normalizado)
                df_out["GRUPOS"], df_out["SOCIODEINTEGRACION"] = "1", "D2L"
                for c in ["SUBJ", "COURSE", "TIPODEHORARIO", "METODO_EDUCATIVO"]:
                    df_out[c] = df_out[c].apply(sin_espacios)

                buffer = io.BytesIO()
                with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
                    for (periodo, resp), grupo in df_out.groupby(["PERIODO", "Responsable"], sort=False):
                        nombre = nombre_csv_altas(periodo, resp)
                        z.writestr(f"CSV_ALTAS/{nombre}",
                                   grupo[COLUMNAS_CSV].to_csv(**CSV_KWARGS_R).encode("utf-8"))

                st.download_button("📥 Descargar CSV_ALTAS Manual (.ZIP)", buffer.getvalue(),
                                   file_name="CSV_ALTAS_MANUAL.zip", mime="application/zip",
                                   type="primary", use_container_width=True)

# ============================================================
# FIN DE PESTAÑA 1
# ============================================================



st.session_state.t2_delta_firma = firma

    if base_ext and errores_ext and st.button("🔍 Extraer líneas con errores", key="t2_extraer", use_container_width=True):
        try:
            base = leer_csv_t2(base_ext)
            indices = leer_reporte_t2(errores_ext, len(base))
            delta = base.iloc[indices].copy().reset_index(drop=True)
            delta.insert(0, COL_FILA_T2, [i + 2 for i in indices])
            st.session_state.t2_delta = delta.copy()
            st.session_state.t2_delta_editado = delta.copy()
            st.session_state.t2_delta_indices = indices
            st.session_state.t2_delta_nombre = f"{nombre_base_t2(base_ext.name)}_{version}"
            st.session_state.t2_editor_version = st.session_state.get("t2_editor_version", 0) + 1
            st.success(f"Se extrajeron {len(delta)} registros con errores de {len(base)} registros.")
        except Exception as e:
            st.session_state.t2_delta = None
            st.error(f"Error al extraer: {e}")

    delta = st.session_state.get("t2_delta")
    if isinstance(delta, pd.DataFrame):
        st.markdown("#### 📋 Fragmento de errores")
        st.caption(f"La columna {COL_FILA_T2} identifica la fila original. No la cambies ni la elimines al corregir fuera de Streamlit.")
        modo = st.radio("¿Cómo deseas corregirlo?", ["Excel (.xlsx)", "CSV (.csv)", "Editar en Streamlit"], horizontal=True, key="t2_modo")
        nombre_delta = st.session_state.t2_delta_nombre
        if modo == "Excel (.xlsx)":
            st.download_button("📥 Descargar fragmento Excel", exportar_delta_t2(st.session_state.t2_delta_editado, "xlsx"),
                               file_name=f"{nombre_delta}.xlsx", use_container_width=True)
        elif modo == "CSV (.csv)":
            st.download_button("📥 Descargar fragmento CSV", exportar_delta_t2(st.session_state.t2_delta_editado, "csv"),
                               file_name=f"{nombre_delta}.csv", mime="text/csv", use_container_width=True)
        else:
            with st.form("t2_form_edicion"):
                editado = st.data_editor(st.session_state.t2_delta_editado, hide_index=True, num_rows="fixed",
                                         use_container_width=True, disabled=[COL_FILA_T2],
                                         key=f"t2_editor_datos_{st.session_state.t2_editor_version}")
                guardar = st.form_submit_button("💾 Guardar correcciones", use_container_width=True)
            if guardar:
                st.session_state.t2_delta_editado = editado.copy()
                st.session_state.t2_editor_version += 1
                st.success("Cambios guardados correctamente.")
                st.rerun()
            st.download_button("📥 Descargar fragmento corregido", exportar_delta_t2(st.session_state.t2_delta_editado, "csv"),
                               file_name=f"{nombre_delta}.csv", mime="text/csv", use_container_width=True)

    # 2. ENSAMBLAR ARCHIVO FINAL
    st.divider()
    st.subheader("💉 2. Inyectar correcciones y generar CSV final")
    c1, c2, c3 = st.columns(3)
    base_iny = c1.file_uploader("📁 CSV original", type=["csv"], key="t2_base_iny")
    errores_iny = c2.file_uploader("📊 Reporte de errores", type=["xlsx"], key="t2_errores_iny")
    corregido_iny = c3.file_uploader("📝 Fragmento corregido", type=["csv", "xlsx"], key="t2_corregido_iny")
    etiqueta = st.text_input("Etiqueta del archivo final", value="final", key="t2_etiqueta")
    firma_final = (firma_archivos_t2(base_iny, errores_iny, corregido_iny), etiqueta)
    if st.session_state.get("t2_final_firma") != firma_final:
        st.session_state.t2_final_bytes = None
        st.session_state.t2_final_nombre = None
        st.session_state.t2_final_firma = firma_final

    if base_iny and errores_iny and corregido_iny and st.button("🚀 Ensamblar Archivo Final", type="primary", key="t2_ensamblar", use_container_width=True):
        try:
            base = leer_csv_t2(base_iny)
            indices = leer_reporte_t2(errores_iny, len(base))
            if corregido_iny.name.lower().endswith(".xlsx"):
                corregidas = pd.read_excel(io.BytesIO(corregido_iny.getvalue()), dtype=str, keep_default_na=False)
            else:
                corregidas = pd.read_csv(io.BytesIO(corregido_iny.getvalue()), dtype=str,
                                         encoding="utf-8-sig", keep_default_na=False)
            corregidas = corregidas.fillna("").reset_index(drop=True)
            if corregidas.columns.duplicated().any(): raise ValueError("El fragmento contiene encabezados duplicados.")
            if COL_FILA_T2 not in corregidas.columns:
                raise ValueError(f"El fragmento debe conservar {COL_FILA_T2}. Vuelve a extraerlo con esta versión de la pestaña 2.")
            if len(corregidas) != len(indices):
                raise ValueError(f"El reporte tiene {len(indices)} líneas, pero el fragmento tiene {len(corregidas)}.")
            if set(corregidas.columns) - {COL_FILA_T2} != set(base.columns):
                raise ValueError("Las columnas del fragmento no coinciden con las del CSV original.")
            numeros = pd.to_numeric(corregidas[COL_FILA_T2], errors="coerce")
            if numeros.isna().any() or (numeros % 1 != 0).any(): raise ValueError("Hay identificadores de fila inválidos.")
            numeros = numeros.astype(int)
            if numeros.duplicated().any(): raise ValueError("Hay filas originales repetidas en el fragmento.")
            esperadas = {i + 2 for i in indices}
            if set(numeros) != esperadas: raise ValueError("Las filas identificadas no coinciden con las líneas del reporte de errores.")
            final = base.copy()
            for j, numero in enumerate(numeros):
                final.iloc[numero - 2] = corregidas.loc[j, list(base.columns)].to_numpy()
            if len(final) != len(base): raise ValueError("Cambió la cantidad de registros originales.")
            st.session_state.t2_final_bytes = final.to_csv(**CSV_KWARGS_R).encode("utf-8")
            etiqueta_segura = re.sub(r'[^a-zA-Z0-9_-]', '_', etiqueta.strip()) or "final"
            st.session_state.t2_final_nombre = f"{nombre_base_t2(base_iny.name)}_{etiqueta_segura}.csv"
            st.success(f"Archivo final generado: {len(indices)} correcciones, {len(final)} registros. Se respetó la posición original de cada fila.")
        except Exception as e:
            st.session_state.t2_final_bytes = None
            st.error(f"Error al ensamblar: {e}")

    if st.session_state.get("t2_final_bytes"):
        st.download_button("📥 DESCARGAR CSV FINAL", data=st.session_state.t2_final_bytes,
                           file_name=st.session_state.t2_final_nombre, mime="text/csv", type="primary",
                           use_container_width=True, key="t2_descargar_final")
# ============================================================
# FIN DE PESTAÑA 2
# ============================================================

encontrados += 1
                                    filas.append((fila_excel, valores, coincidencia["_NRC"]))
                            if not filas: continue
                            # Bloquear este libro si tiene coincidencias faltantes o ambiguas.
                            if encontrados != len(filas):
                                alertas.append({"Archivo": fx.name, "Estado": "Libro no exportado", "Detalle": f"{encontrados}/{len(filas)} filas con NRC unico"})
                                continue
                            if HOJA_SALIDA_NRC in wb.sheetnames: del wb[HOJA_SALIDA_NRC]
                            ws = wb.create_sheet(HOJA_SALIDA_NRC)
                            ws.append(["Fila ALTAS", "NRC"] + list(COLUMNAS_ALTAS_REALES))
                            for fila_excel, valores, codigo in filas:
                                valores_originales = {normalizar_para_busqueda(c.value): c.column for c in ws_altas[1] if c.value is not None}
                                ws.append([fila_excel, codigo] + [ws_altas.cell(fila_excel, valores_originales[normalizar_para_busqueda(c)]).value if normalizar_para_busqueda(c) in valores_originales else "" for c in COLUMNAS_ALTAS_REALES])
                            formatear_excel_t3(ws)
                            salida = io.BytesIO()
                            wb.save(salida)
                            nombre = fx.name.rsplit(".", 1)[0] + "_con_NRC" + ext
                            original_nombre, n = nombre, 2
                            while nombre in nombres_zip:
                                nombre = original_nombre[:-len(ext)] + f" ({n})" + ext
                                n += 1
                            nombres_zip.add(nombre)
                            z.writestr(nombre, salida.getvalue())
                            procesados += 1
                    st.session_state.final_argos_zip = buffer.getvalue() if procesados else None
                    st.session_state.t3_alertas = alertas
                    if procesados: st.success(f"Se generaron {procesados} Excel con NRC.")
                    else: st.error("No se generó ningún Excel. Revisa las incidencias de correspondencia.")
            except Exception as e:
                st.session_state.final_argos_zip = None
                st.error(f"Error de inyección: {e}")
        alertas = st.session_state.get("t3_alertas") or []
        if alertas:
            with st.expander(f"⚠️ Incidencias de cruce ({len(alertas)})", expanded=True):
                st.dataframe(pd.DataFrame(alertas), hide_index=True, use_container_width=True)
        if st.session_state.get("final_argos_zip"):
            st.download_button("📥 Descargar Excel con NRC (.ZIP)", st.session_state.final_argos_zip, file_name="Excels_Finales_con_NRC.zip", mime="application/zip", type="primary", use_container_width=True)

    # 4. MODO RAPIDO: tabla editable y descarga.
    else:
        st.subheader("⚡ Cruce rápido de NRC")
        c1, c2 = st.columns(2)
        argos = c1.file_uploader("📊 ARGOS actualizado (.csv)", type=["csv"], key="t3_argos_rapido")
        archivos_csv = c2.file_uploader("📝 CSV finales", type=["csv"], accept_multiple_files=True, key="t3_csv_rapido")
        firma = firma_t3(argos, *(archivos_csv or []))
        if st.session_state.get("t3_firma_rapido") != firma:
            st.session_state.df_cruce_rapido = None
            st.session_state.t3_resultado_editado = None
            st.session_state.t3_editor_version = 0
            st.session_state.t3_firma_rapido = firma
            st.session_state.pop("t3_columnas", None)
        if argos and archivos_csv and st.button("⚡ Cruzar NRC y generar tabla", type="primary", use_container_width=True, key="t3_procesar_rapido"):
            try:
                argos_df, mapa, avisos = leer_argos_t3(argos)
                resultados, alertas = [], list(avisos)
                for archivo in archivos_csv:
                    df = leer_csv_t3(archivo)
                    nrc, faltantes = cruzar_nrc_t3(df, argos_df, mapa, archivo.name)
                    df.insert(0, "NRC", nrc.values)
                    df = df.rename(columns={"TIPODEHORARIO": "TIPO DE HORARIO", "METODO_EDUCATIVO": "METODO_ED", "datocomplementario": "Cluster"})
                    columnas = ["NRC", "PERIODO", "SUBJ", "COURSE", "CAPACIDAD", "SECCION", "TIPO DE HORARIO", "METODO_ED", "Cluster"]
                    resultados.append(df[columnas])
                    alertas.extend(faltantes)
                resultado = pd.concat(resultados, ignore_index=True)
                st.session_state.df_cruce_rapido = resultado.copy()
                st.session_state.t3_resultado_editado = resultado.copy()
                st.session_state.t3_editor_version = st.session_state.get("t3_editor_version", 0) + 1
                st.session_state.t3_alertas = alertas
                st.success(f"Cruce completado: {len(resultado)} registros; {int(resultado['NRC'].ne('').sum())} con NRC.")
            except Exception as e:
                st.session_state.t3_resultado_editado = None
                st.error(f"Error en cruce rápido: {e}")
        alertas = st.session_state.get("t3_alertas") or []
        if alertas:
            with st.expander(f"⚠️ Incidencias ({len(alertas)})"):
                st.dataframe(pd.DataFrame(alertas), hide_index=True, use_container_width=True)
        resultado = st.session_state.get("t3_resultado_editado")
        if isinstance(resultado, pd.DataFrame):
            st.subheader("📋 Resultados del cruce")
            with st.form("t3_form_editor"):
                editado = st.data_editor(resultado, hide_index=True, num_rows="fixed", use_container_width=True, key=f"t3_editor_{st.session_state.t3_editor_version}")
                guardar = st.form_submit_button("💾 Guardar cambios", use_container_width=True)
            if guardar:
                st.session_state.t3_resultado_editado = editado.copy()
                st.session_state.t3_editor_version += 1
                st.rerun()
            columnas = st.multiselect("Columnas para copiar o descargar", options=list(resultado.columns), default=list(resultado.columns), key="t3_columnas")
            if columnas:
                mostrar = st.session_state.t3_resultado_editado[columnas].copy()
                st.dataframe(mostrar, hide_index=True, use_container_width=True)
                st.code(mostrar.to_csv(index=False, sep="\t"), language="text")
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    mostrar.to_excel(writer, index=False, sheet_name="Cruce_NRC")
                    formatear_excel_t3(writer.sheets["Cruce_NRC"])
                st.download_button("📥 Descargar Excel (.xlsx)", buffer.getvalue(), file_name="Cruce_Rapido_NRC.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
            else: st.warning("Selecciona al menos una columna.")
# ============================================================
# FIN DE PESTAÑA 3
# ============================================================


usadas.add(siguiente)
            clave = t4_num_clave(siguiente)
            g = g.assign(_tipo_orden=g["TIPO"].map(lambda x: 0 if t4_normal(x) == "MANDANTE" else 1)).sort_values(["_tipo_orden", "_fila"], kind="stable")
            for _, r in g.iterrows():
                filas.append({"Clave": clave, "NRC": r["NRC"], "Tipo": "Mandante" if t4_normal(r["TIPO"]) == "MANDANTE" else "Cruce",
                             "Responsable": r["RESPONSABLE"], "Periodo": periodo, "SUB": r["SUBJ"], "COURSE": r["COURSE"], "Comentarios": ""})
        full = pd.DataFrame(filas, columns=COL_COMPLETO)
        ti = full[COL_TI].copy()
        if not anteriores_ti.empty:
            cons_ti = pd.concat([anteriores_ti, ti], ignore_index=True).drop_duplicates()
        else: cons_ti = ti.copy()
        cons_full = None
        if anteriores_ti.empty: cons_full = full.copy()
        elif not anteriores_full.empty:
            # Solo es completo cuando todos los NRC anteriores tienen sus 8 campos.
            cubiertos = set(map(tuple, anteriores_full[COL_TI].to_numpy()))
            todos = set(map(tuple, anteriores_ti[COL_TI].to_numpy()))
            if cubiertos == todos: cons_full = pd.concat([anteriores_full, full], ignore_index=True).drop_duplicates()
        return full, ti, cons_full, cons_ti
    def t4_excel(df, hoja):
        salida = io.BytesIO()
        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=hoja, index=False)
            ws = writer.sheets[hoja]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for celda in ws[1]:
                celda.font = Font(bold=True, color="FFFFFF")
                celda.fill = PatternFill(fill_type="solid", fgColor="7030A0")
            for col in ws.columns:
                ancho = min(40, max(12, max(len(str(c.value or "")) for c in col)+2))
                ws.column_dimensions[col[0].column_letter].width = ancho
        return salida.getvalue()

    if st.button("🔄 Limpiar Cruces", key="t4_limpiar"):
        for k in list(st.session_state):
            if k.startswith("t4_"): st.session_state.pop(k, None)
        st.rerun()
    st.subheader("1. Cargar plantillas")
    archivos_nuevos = st.file_uploader("Plantillas con hoja Cruces", type=["xlsx", "xlsm"], accept_multiple_files=True, key="t4_nuevos")
    modo_t4 = st.radio("Continuación de claves", ["Empezar en AA", "Cargar cruces anteriores", "Indicar última clave manualmente"], key="t4_modo")
    archivos_anteriores = []
    if modo_t4 == "Cargar cruces anteriores":
        archivos_anteriores = st.file_uploader("Cruces anteriores: completo o TI (.xlsx, .xlsm, .csv)", type=["xlsx", "xlsm", "csv"], accept_multiple_files=True, key="t4_anteriores") or []
    if modo_t4 == "Indicar última clave manualmente":
        st.caption("Escribe una última clave distinta por periodo. Deja vacío si ese periodo comienza en AA.")
    st.subheader("2. Revisar y asignar claves")
    firma = tuple((f.name, hashlib.sha256(f.getvalue()).hexdigest()) for f in list(archivos_nuevos or []) + list(archivos_anteriores)) + (modo_t4,)
    if st.session_state.get("t4_firma") != firma:
        st.session_state.t4_firma = firma
        st.session_state.t4_resultados = None
        st.session_state.t4_editor_version = st.session_state.get("t4_editor_version", 0) + 1
    if archivos_nuevos:
        try:
            nuevos = t4_leer_nuevos(archivos_nuevos)
            periodos = sorted(p for p in nuevos["PERIODO"].unique() if p)
            ultimas = {}
            if modo_t4 == "Indicar última clave manualmente":
                cols = st.columns(min(3, max(1, len(periodos))))
                for i, p in enumerate(periodos):
                    ultimas[p] = cols[i % len(cols)].text_input(f"Última clave del periodo {p}", key=f"t4_ultima_{p}").strip().upper()
            if modo_t4 == "Cargar cruces anteriores" and not archivos_anteriores:
                st.info("Carga los archivos anteriores para continuar la numeración.")
            else:
                anteriores_full, anteriores_ti = t4_leer_anteriores(archivos_anteriores) if archivos_anteriores else (pd.DataFrame(columns=COL_COMPLETO), pd.DataFrame(columns=COL_TI))
                if modo_t4 == "Empezar en AA" and not anteriores_ti.empty: raise ValueError("No se pueden cargar anteriores en modo AA.")
                if st.button("🔍 Validar y generar claves", type="primary", key="t4_generar"):
                    full, ti, cons_full, cons_ti = t4_generar(nuevos, anteriores_full, anteriores_ti, ultimas)
                    st.session_state.t4_resultados = (full, ti, cons_full, cons_ti)
                    st.session_state.t4_editor_version += 1
                    st.success(f"Se asignaron {full['Clave'].groupby([full['Periodo'], full['Clave']]).ngroups} cruces nuevos.")
        except Exception as e:
            st.error(f"Revisión de cruces: {e}")
    resultado = st.session_state.get("t4_resultados")
    if resultado is not None:
        full, ti, cons_full, cons_ti = resultado
        st.subheader("3. Revisar resultados")
        st.caption("Las claves y NRC no son editables para evitar romper la relación entre Mandante e hijos. Puedes editar los datos descriptivos de los nuevos cruces.")
        with st.form("t4_edicion"):
            editado = st.data_editor(full, hide_index=True, num_rows="fixed", use_container_width=True,
                                     disabled=["Clave", "NRC", "Tipo", "Periodo"], key=f"t4_editor_{st.session_state.t4_editor_version}")
            guardar = st.form_submit_button("💾 Guardar cambios")
        if guardar:
            full = editado.copy()
            ti = full[COL_TI].copy()
            if cons_full is not None:
                anteriores = cons_full.iloc[:len(cons_full)-len(resultado[0])].copy()
                cons_full = pd.concat([anteriores, full], ignore_index=True)
            if len(cons_ti) >= len(resultado[1]):
                anteriores_ti = cons_ti.iloc[:len(cons_ti)-len(resultado[1])].copy()
                cons_ti = pd.concat([anteriores_ti, ti], ignore_index=True)
            st.session_state.t4_resultados = (full, ti, cons_full, cons_ti)
            st.session_state.t4_editor_version += 1
            st.rerun()
        salida = st.radio("¿Qué deseas descargar?", ["Solo nuevos", "Consolidado completo", "Ambas opciones"], index=2, horizontal=True, key="t4_salida")
        archivos = {}
        if salida in ("Solo nuevos", "Ambas opciones"):
            archivos["Cruces_Nuevos_Completo.xlsx"] = t4_excel(full[COL_COMPLETO], "Cruces")
            archivos["Cruces_Nuevos_TI.xlsx"] = t4_excel(ti[COL_TI], "TI")
        if salida in ("Consolidado completo", "Ambas opciones"):
            archivos["Cruces_Consolidado_TI.xlsx"] = t4_excel(cons_ti[COL_TI], "TI")
            if cons_full is not None: archivos["Cruces_Consolidado_Completo.xlsx"] = t4_excel(cons_full[COL_COMPLETO], "Cruces")
            else: st.warning("No se genera consolidado completo de 8 columnas porque los anteriores solo tienen Clave, NRC y Periodo.")
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
            for nombre, datos in archivos.items(): z.writestr(nombre, datos)
        st.download_button("📥 Descargar cruces (.ZIP)", buffer.getvalue(), file_name="CRUCES_PROCESADOS.zip",
                           mime="application/zip", type="primary", use_container_width=True, key="t4_descargar")
