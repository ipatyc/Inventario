# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import io
import unicodedata
import openpyxl
import zipfile
from difflib import SequenceMatcher

# ================= 1. CONFIGURACIÓN Y FUNCIONES ESTRUCTURALES =================
HOJA_ALTAS = "ALTAS"
UMBRAL_FUZZY = 0.82  # Precisión para búsqueda de nombres

def quitar_acentos(t):
    if pd.isna(t) or t is None: return ""
    return "".join(c for c in unicodedata.normalize("NFD", str(t)) if unicodedata.category(c) != "Mn")

def normalizar_para_cruce(t):
    return quitar_acentos(str(t).upper().strip())

def similitud(a, b): 
    return SequenceMatcher(None, a, b).ratio()

# 🧼 FUNCIÓN DE FORMATO ESTILO R: Copia exactamente el comportamiento de tu write.csv y mutate
def format_r_style(val, is_seccion=False):
    if pd.isna(val) or val is None: return ""
    s = str(val).strip()
    if s.lower() == "nan": return ""
    if s.endswith(".0"): s = s[:-2] # Elimina el flotante automático de pandas (.0)
    if is_seccion:
        if s.isdigit(): s = str(int(s)) # Imita al as.numeric(SECCION) de tu R: quita ceros a la izquierda (ej: '01' -> '1')
    return s

# Limpiador interno exclusivo para que las llaves del cruce de la pestaña 2 no fallen jamás
def limpia_seccion_interna(x):
    if pd.isna(x): return ""
    s = str(x).strip()
    if s.lower() == "nan": return ""
    if s.endswith(".0"): s = s[:-2]
    if s.isdigit(): return f"{int(s):02d}"
    return s

# Inicialización de estados en memoria de Streamlit
if "original_files_bytes" not in st.session_state: st.session_state.original_files_bytes = {}
if "df_corregido" not in st.session_state: st.session_state.df_corregido = None
if "raw_altas" not in st.session_state: st.session_state.raw_altas = None
if "res_auditoria" not in st.session_state: st.session_state.res_auditoria = None
if "ready_for_download" not in st.session_state: st.session_state.ready_for_download = False
if "zip_file_bytes" not in st.session_state: st.session_state.zip_file_bytes = None
if "summary_csv_bytes" not in st.session_state: st.session_state.summary_csv_bytes = None
if "csv_files_to_download" not in st.session_state: st.session_state.csv_files_to_download = {}

# Configuración de la interfaz visual
st.set_page_config(page_title="Consola Iris Cavazos", page_icon="🎛️", layout="wide")
st.title("🎛️ Consola de Control de Materias e Inyección de NRCs")
st.markdown("---")

# Creación de las 3 pestañas solicitadas
tab1, tab_err, tab3 = st.tabs([
    "1️⃣ Proceso: Validación y Generar CSV", 
    "⚠️ Reporte de Errores (Filtro Delta)", 
    "2️⃣ Proceso: Inyección de NRCs (ARGOS)"
])

# ============================================================
# PESTAÑA 1: VALIDACIÓN Y GENERACIÓN DE CSV
# ============================================================
with tab1:
    st.header("Validación de Claves y Generación de CSV")
    
    col1, col2 = st.columns(2)
    with col1:
        file_cat = st.file_uploader("📑 Catálogo de Materias Estatales (Excel)", type=["xlsx"])
    with col2:
        files_altas = st.file_uploader("📁 Archivos de ALTAS (Puedes subir varios)", accept_multiple_files=True, type=["xlsx"])
    
    if files_altas and file_cat:
        if st.button("⚡ Ejecutar Validación Inteligente", type="primary"):
            st.session_state.ready_for_download = False 
            st.toast("Cargando Catálogo de Materias...", icon="📑")
            
            xls_cat = pd.ExcelFile(file_cat)
            indice_cat = {}
            indice_cat_claves = {} 
            
            for hoja in xls_cat.sheet_names:
                df_c = xls_cat.parse(hoja)
                if "Nivel" in df_c.columns and "Materia" in df_c.columns:
                    for _, f in df_c.iterrows():
                        niv = normalizar_para_cruce(f.get("Nivel"))
                        mat_o = str(f.get("Materia")).strip()
                        s_val = format_r_style(f.get("Subj"))
                        c_val = format_r_style(f.get("Crse"))
                        
                        indice_cat.setdefault(niv, []).append({
                            "mat_orig": mat_o,
                            "mat_norm": normalizar_para_cruce(f.get("Materia")), 
                            "subj": s_val, 
                            "crse": c_val
                        })
                        
                        if s_val and c_val:
                            s_norm = normalizar_para_cruce(s_val)
                            c_norm = c_val
                            indice_cat_claves[(s_norm, c_norm)] = mat_o
            
            piezas = []
            for f in files_altas:
                primera_palabra = f.name.split()[0]
                st.info(f"🔍 Checking / Revisando archivo: **{primera_palabra}**")
                st.session_state.original_files_bytes[f.name] = f.getvalue()
                
                xls_a = pd.ExcelFile(f)
                hojas_reales = [h for h in xls_a.sheet_names if h.strip().upper() == HOJA_ALTAS]
                if hojas_reales:
                    df_a = xls_a.parse(hojas_reales[0])
                    
                    # 🚨 FILTRO ANTI-FANTASMAS
                    essential_cols = [c for c in ["Periodo", "Campus", "Subject", "Course"] if c in df_a.columns]
                    if essential_cols:
                        df_a = df_a.dropna(subset=essential_cols, how="all")
                    df_a = df_a.dropna(how="all")
                    
                    if not df_a.empty:
                        df_a["ArchivoOrigen"] = f.name
                        piezas.append(df_a)
            
            if piezas:
                df_total = pd.concat(piezas, ignore_index=True)
                st.session_state.raw_altas = df_total.copy()
                
                resultados = []
                for idx, fila in df_total.iterrows():
                    niv_n = normalizar_para_cruce(fila.get("Nivel"))
                    mat_excel_orig = fila.get("Nombre de la Materia")
                    mat_n = normalizar_para_cruce(mat_excel_orig)
                    subj_orig = format_r_style(fila.get("Subject"))
                    crse_orig = format_r_style(fila.get("Course"))
                    
                    candidatos = indice_cat.get(niv_n, [])
                    matches_exactos = [c for c in candidatos if c["mat_norm"] == mat_n]
                    
                    match_elegido = None
                    
                    if matches_exactos:
                        coincidencia_perfecta = next((m for m in matches_exactos if m["subj"] == subj_orig and m["crse"] == crse_orig), None)
                        match_elegido = coincidencia_perfecta if coincidencia_perfecta else matches_exactos[0]
                    else:
                        mejor, mejor_s = None, -1.0
                        for c in candidatos:
                            s = similitud(mat_n, c["mat_norm"])
                            if s > mejor_s: mejor_s, mejor = s, c
                        if mejor and mejor_s >= UMBRAL_FUZZY:
                            matches_fuzzy = [c for c in candidatos if c["mat_norm"] == mejor["mat_norm"]]
                            coincidencia_perf_f = next((m for m in matches_fuzzy if m["subj"] == subj_orig and m["crse"] == crse_orig), None)
                            match_elegido = coincidencia_perf_f if coincidencia_perf_f else mejor
                    
                    if match_elegido:
                        subj_sug = match_elegido["subj"]
                        crse_sug = match_elegido["crse"]
                        mat_cat_nombre = match_elegido["mat_orig"]
                        
                        if subj_orig == subj_sug and crse_orig == crse_sug:
                            comentario = "Todo correcto"
                        elif subj_orig != subj_sug and crse_orig != crse_sug:
                            comentario = "Subj y Crse incorrectos"
                        elif subj_orig != subj_sug:
                            comentario = "Subject incorrecto"
                        else:
                            comentario = "Course incorrecto"
                    else:
                        s_excel_norm = normalizar_para_cruce(subj_orig)
                        c_excel_norm = crse_orig
                        
                        if (s_excel_norm, c_excel_norm) in indice_cat_claves:
                            mat_cat_nombre = indice_cat_claves[(s_excel_norm, c_excel_norm)]
                            comentario = "Nombre de materia incorrecto"
                        else:
                            mat_cat_nombre = mat_excel_orig
                            comentario = "No se encontró en catálogo"
                            
                        subj_sug = subj_orig
                        crse_sug = crse_orig
                    
                    resultados.append({
                        "Luz Verde": False, 
                        "idx": idx, 
                        "Archivo": fila.get("ArchivoOrigen"), 
                        "Materia Excel": mat_excel_orig, 
                        "Materia Catálogo": mat_cat_nombre, 
                        "Comentario": comentario,
                        "Subj Original": subj_orig, 
                        "Crse Original": crse_orig,
                        "Subj Sugerido": subj_sug, 
                        "Crse Sugerido": crse_sug
                    })
                
                st.session_state.res_auditoria = pd.DataFrame(resultados)
                st.success("¡Revisión terminada con éxito!")
            else:
                st.error(f"❌ Ninguno de los archivos subidos tiene filas válidas en la pestaña '{HOJA_ALTAS}'")

    if st.session_state.res_auditoria is not None:
        st.markdown("### ⚖️ Mesa de Control Interactiva por Excel")
        
        df_aud = st.session_state.res_auditoria
        archivos_subidos = df_aud["Archivo"].unique()
        
        st.markdown("#### 📊 Resumen Estadístico de Observaciones")
        resumen_errores = []
        for arch in archivos_subidos:
            df_file = df_aud[df_aud["Archivo"] == arch]
            total_err = len(df_file[df_file["Comentario"] != "Todo correcto"])
            resumen_errores.append({
                "Archivo Cargado": arch, 
                "Errores / Observaciones Detectadas": total_err,
                "Estado del Archivo": "⚠️ Requiere Revisión" if total_err > 0 else "✅ Todo Limpio"
            })
        st.dataframe(pd.DataFrame(resumen_errores), hide_index=True, use_container_width=True)
        st.markdown("---")
        
        for arch in archivos_subidos:
            df_file = df_aud[df_aud["Archivo"] == arch]
            errores_filas = df_file[df_file["Comentario"] != "Todo correcto"]
            total_detalles = len(errores_filas)
            
            if total_detalles == 0:
                st.success(f"✅ **{arch}** — ¡Todo perfecto, sin observaciones!")
            else:
                with st.expander(f"⚠️ **{arch}** — ({total_detalles} renglones con observaciones encontradas)", expanded=True):
                    quitar_rep = st.checkbox("🔍 Combinar repetidas (Ver solo 1 renglón por caso)", value=True, key=f"rep_{arch}")
                    
                    if quitar_rep:
                        df_vista = errores_filas.drop_duplicates(subset=["Materia Excel", "Materia Catálogo", "Subj Original", "Crse Original", "Comentario"])
                    else:
                        df_vista = errores_filas
                        
                    columnas_vista = ["Luz Verde", "Materia Excel", "Materia Catálogo", "Comentario", "Subj Original", "Crse Original", "Subj Sugerido", "Crse Sugerido"]
                    
                    df_editado_archivo = st.data_editor(
                        df_vista[columnas_vista],
                        hide_index=True,
                        disabled=["Materia Excel", "Materia Catálogo", "Comentario", "Subj Original", "Crse Original"],
                        column_config={
                            "Luz Verde": st.column_config.CheckboxColumn("¿Aplicar?", help="Marca para autorizar este cambio"),
                            "Materia Excel": st.column_config.TextColumn("Materia (Excel)", width="medium"),
                            "Materia Catálogo": st.column_config.TextColumn("Materia (Catálogo Oficial)", width="medium"),
                            "Comentario": st.column_config.TextColumn("Diagnóstico", width="medium"),
                            "Subj Original": st.column_config.TextColumn("Subj (Excel)", width="small"),
                            "Crse Original": st.column_config.TextColumn("Crse (Excel)", width="small"),
                            "Subj Sugerido": st.column_config.TextColumn("Subj Sugerido ✍️", width="small"),
                            "Crse Sugerido": st.column_config.TextColumn("Crse Sugerido ✍️", width="small"),
                        },
                        key=f"editor_{arch}",
                        use_container_width=True
                    )
                    
                    for _, row in df_editado_archivo.iterrows():
                        mascara = (st.session_state.res_auditoria["Archivo"] == arch) & \
                                  (st.session_state.res_auditoria["Materia Excel"] == row["Materia Excel"]) & \
                                  (st.session_state.res_auditoria["Materia Catálogo"] == row["Materia Catálogo"]) & \
                                  (st.session_state.res_auditoria["Subj Original"] == row["Subj Original"]) & \
                                  (st.session_state.res_auditoria["Crse Original"] == row["Crse Original"])
                        
                        st.session_state.res_auditoria.loc[mascara, "Luz Verde"] = row["Luz Verde"]
                        st.session_state.res_auditoria.loc[mascara, "Subj Sugerido"] = row["Subj Sugerido"]
                        st.session_state.res_auditoria.loc[mascara, "Crse Sugerido"] = row["Crse Sugerido"]
        
        st.markdown("---")
        if st.button("💾 Aplicar Cambios Autorizados y Procesar Todo", type="primary"):
            corregido = st.session_state.raw_altas.copy()
            
            for _, row in st.session_state.res_auditoria.iterrows():
                if row["Luz Verde"] and pd.notna(row["Subj Sugerido"]):
                    corregido.loc[row["idx"], "Subject"] = row["Subj Sugerido"]
                    corregido.loc[row["idx"], "Course"] = row["Crse Sugerido"]
            
            st.session_state.df_corregido = corregido
            
            summary_str = corregido.to_csv(index=False, encoding="utf-8-sig")
            st.session_state.summary_csv_bytes = summary_str.encode("utf-8-sig")
            
            st.session_state.csv_files_to_download = {}
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for name, sub in corregido.groupby("ArchivoOrigen"):
                    resultado_df = pd.DataFrame()
                    
                    resultado_df["PERIODO"] = sub["Periodo"].apply(lambda x: format_r_style(x))
                    resultado_df["SEDE"] = sub["Campus"].apply(lambda x: format_r_style(x))
                    resultado_df["SUBJ"] = sub["Subject"].apply(lambda x: format_r_style(x))
                    resultado_df["COURSE"] = sub["Course"].apply(lambda x: format_r_style(x))
                    resultado_df["PARTEPERIODO"] = sub["Parte de Periodo"].apply(lambda x: format_r_style(x))
                    resultado_df["STATUS"] = sub["Estatus"].apply(lambda x: format_r_style(x))
                    resultado_df["CAPACIDAD"] = sub["Capacidad"].apply(lambda x: format_r_style(x))
                    resultado_df["GRUPOS"] = 1
                    resultado_df["SECCION"] = sub["Sección"].apply(lambda x: format_r_style(x, is_seccion=True))
                    resultado_df["TIPODEHORARIO"] = sub["Tipo de Horario"].apply(lambda x: format_r_style(x))
                    resultado_df["METODO_EDUCATIVO"] = sub["Método Educativo"].apply(lambda x: format_r_style(x))
                    resultado_df["SOCIODEINTEGRACION"] = "D2L"
                    resultado_df["MODODECALIFICAR"] = sub["Modo de Calificar"].apply(lambda x: format_r_style(x))
                    resultado_df["SESION"] = sub["Sesion"].apply(lambda x: format_r_style(x))
                    
                    columnas_ordenadas = ["PERIODO", "SEDE", "SUBJ", "COURSE", "PARTEPERIODO", "STATUS",
                                          "CAPACIDAD", "GRUPOS", "SECCION", "TIPODEHORARIO",
                                          "METODO_EDUCATIVO", "SOCIODEINTEGRACION", "MODODECALIFICAR", "SESION"]
                    resultado_df = resultado_df[columnas_ordenadas]
                    
                    nombre_base = name.rsplit('.', 1)[0] if '.' in name else name
                    csv_filename = f"{nombre_base}.csv"
                    
                    csv_string = resultado_df.to_csv(index=False, encoding="utf-8-sig")
                    zip_file.writestr(csv_filename, csv_string)
                    st.session_state.csv_files_to_download[csv_filename] = csv_string.encode("utf-8-sig")
            
            st.session_state.zip_file_bytes = zip_buffer.getvalue()
            st.session_state.ready_for_download = True
            st.rerun()

        if st.session_state.ready_for_download:
            st.markdown("### 📥 Panel de Descarga de Resultados")
            
            st.download_button(
                label="📝 📥 DESCARGAR ARCHIVO DE RESUMEN PARA PESTAÑA 2 (.CSV)",
                data=st.session_state.summary_csv_bytes,
                file_name="resumen_proceso_1.csv",
                mime="text/csv",
                use_container_width=True,
                key="resumen_p1_btn"
            )
            
            st.download_button(
                label="💥 📥 DESCARGAR TODOS LOS CSVs JUNTOS (.ZIP)",
                data=st.session_state.zip_file_bytes,
                file_name="todos_los_csvs_estructurados.zip",
                mime="application/zip",
                use_container_width=True,
                type="primary",
                key="zip_p1_btn"
            )

# ============================================================
# PESTAÑA NUEVA: FILTRADO POR REPORTE DE ERRORES (BANNER)
# ============================================================
with tab_err:
    st.header("⚠️ Extracción Delta por Reporte de Errores (Banner)")
    st.markdown("Sube el archivo Excel de errores rebotado por Banner para extraer quirúrgicamente solo las filas que fallaron.")
    
    file_errores = st.file_uploader("📥 1. Cargar Reporte de Errores de Banner (.xlsx)", type=["xlsx"])
    
    st.markdown("---")
    st.markdown("#### 📄 2. Selecciona el archivo origen que deseas filtrar:")
    
    origen_csv = st.radio(
        "¿De dónde tomamos la base de datos original?",
        ["Utilizar los CSVs estructurados generados en la Pestaña 1 (En memoria)", "Subir un archivo CSV manualmente de mis carpetas"],
        key="origen_csv_radio"
    )
    
    dict_csvs_a_procesar = {}
    
    if origen_csv == "Utilizar los CSVs estructurados generados en la Pestaña 1 (En memoria)":
        if st.session_state.csv_files_to_download:
            st.success(f"✅ Se detectaron {len(st.session_state.csv_files_to_download)} archivos listos en la memoria.")
            opciones_nombres = list(st.session_state.csv_files_to_download.keys())
            seleccion_memoria = st.selectbox("Elige el CSV que deseas depurar:", opciones_nombres)
            if seleccion_memoria:
                bytes_csv = st.session_state.csv_files_to_download[seleccion_memoria]
                dict_csvs_a_procesar[seleccion_memoria] = pd.read_csv(io.BytesIO(bytes_csv), encoding="utf-8-sig")
        else:
            st.warning("⚠️ No se encontraron archivos en la memoria interna de la sesión. Primero ejecuta la Pestaña 1 o cambia a subida manual.")
    else:
        file_csv_manual = st.file_uploader("Subir el archivo CSV original cargado a Banner (.csv)", type=["csv"], key="csv_manual_err")
        if file_csv_manual:
            try:
                dict_csvs_a_procesar[file_csv_manual.name] = pd.read_csv(file_csv_manual, encoding="utf-8-sig")
            except:
                dict_csvs_a_procesar[file_csv_manual.name] = pd.read_csv(file_csv_manual, encoding="latin-1")
                
    st.markdown("---")
    num_version = st.number_input("🔢 3. Indica el número de corrección (Versión):", min_value=1, max_value=99, value=1, step=1)
    
    if file_errores and dict_csvs_a_procesar:
        if st.button("🔍 Extraer Filas con Error y Generar Delta", type="primary"):
            try:
                # Se descartan las primeras 2 filas (la fila 3 pasa a ser el Header del DataFrame)
                df_err_excel = pd.read_excel(file_errores, skiprows=2)
                
                if "Línea" not in df_err_excel.columns:
                    st.error("❌ Error de formato: No se encontró la columna llamada exactamente 'Línea' en la fila 3 del reporte.")
                else:
                    # Extraer los números de fila únicos de la columna Línea
                    renglones_errores = df_err_excel["Línea"].dropna().astype(int).unique().tolist()
                    st.info(f"📋 Renglones detectados con fallas: {renglones_errores}")
                    
                    for nombre_archivo, df_datos in dict_csvs_a_procesar.items():
                        # Mapeo matemático Banner: Línea 1 = Headers, Línea 2 = index 0 de Pandas. 
                        # Por ende: pandas_index = Línea - 2
                        indices_validos = [int(r) - 2 for r in renglones_errores if (int(r) - 2) >= 0 and (int(r) - 2) < len(df_datos)]
                        
                        df_delta_filtrado = df_datos.iloc[indices_validos]
                        
                        if not df_delta_filtrado.empty:
                            st.success(f"🎯 ¡Filtro completado! Se aislaron {len(df_delta_filtrado)} registros conflictivos de {nombre_archivo}.")
                            
                            nombre_limpio = nombre_archivo.rsplit('.', 1)[0] if '.' in nombre_archivo else "Horario"
                            nombre_final_out = f"{nombre_limpio}_Correccion_V{num_version}.csv"
                            
                            csv_bytes_out = df_delta_filtrado.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                            
                            st.download_button(
                                label=f"📥 Descargar {nombre_final_out}",
                                data=csv_bytes_out,
                                file_name=nombre_final_out,
                                mime="text/csv",
                                key=f"btn_err_dl_{nombre_archivo}",
                                use_container_width=True
                            )
                        else:
                            st.error(f"❌ Ninguno de los números de 'Línea' ({renglones_errores}) coincide con los rangos reales de {nombre_archivo}.")
            except Exception as e:
                st.error(f"❌ Ocurrió un error inesperado al procesar los archivos: {e}")

# ============================================================
# PESTAÑA 3: INYECTAR EN REPORTE ARGOS (CONSTRUIR HOJA "CRNs")
# ============================================================
with tab3:
    st.header("Inyección de NRCs desde Reporte de ARGOS")
    
    mismo_momento = st.session_state.df_corregido is not None
    procesar_cruce = False
    df_base_cruce = None
    dict_bytes_altas = {}
    file_argos = None
    
    if mismo_momento:
        st.success("🧠 **Modo Instantáneo:** La app recuerda tus correcciones actuales de la Pestaña 1. Solo necesitas subir el reporte de ARGOS.")
        file_argos = st.file_uploader("📊 Cargar Reporte de ARGOS (.csv)", type=["csv"], key="argos_directo")
        
        if file_argos:
            procesar_cruce = st.button("🚀 Cruzar Datos y Modificar Excels", type="primary", key="btn_directo")
            if procesar_cruce:
                df_base_cruce = st.session_state.df_corregido.copy()
                dict_bytes_altas = st.session_state.original_files_bytes
    else:
        st.info("🕒 **Modo Asincrónico (Trabajo de otro día o post-correcciones):** Sube el resumen final limpio/corregido y tus archivos de ALTAS actualizados.")
        
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            file_argos = st.file_uploader("📊 1. Cargar Reporte de ARGOS (.csv)", type=["csv"], key="argos_asinc")
        with col_b:
            file_resumen = st.file_uploader("📝 2. Cargar Archivo de Resumen (.csv)", type=["csv"])
        with col_c:
            files_altas_p2 = st.file_uploader("📁 3. Archivos de ALTAS (.xlsx)", accept_multiple_files=True, type=["xlsx"])
            
        if file_argos and file_resumen and files_altas_p2:
            procesar_cruce = st.button("🚀 Cruzar Datos Directo (Usando Archivo Resumen)", type="primary", key="btn_asinc")
            if procesar_cruce:
                try:
                    df_base_cruce = pd.read_csv(file_resumen, encoding="utf-8")
                except:
                    df_base_cruce = pd.read_csv(file_resumen, encoding="latin-1")
                
                for f in files_altas_p2:
                    dict_bytes_altas[f.name] = f.getvalue()

    if procesar_cruce and df_base_cruce is not None:
        try:
            argos_df = pd.read_csv(file_argos, encoding="utf-8")
        except:
            argos_df = pd.read_csv(file_argos, encoding="latin-1")
        
        st.info("Estandarizando llaves y realizando left_join de R...")
        solicitud_p2 = df_base_cruce.copy()
        
        solicitud_p2["_k_per"] = solicitud_p2["Periodo"].astype(str).str.strip().apply(lambda x: x[:-2] if x.endswith(".0") else x)
        solicitud_p2["_k_niv"] = solicitud_p2["Nivel"].apply(normalizar_para_cruce)
        solicitud_p2["_k_sub"] = solicitud_p2["Subject"].apply(normalizar_para_cruce)
        solicitud_p2["_k_crs"] = solicitud_p2["Course"].astype(str).str.strip().apply(lambda x: x[:-2] if x.endswith(".0") else x)
        solicitud_p2["_k_sec"] = solicitud_p2["Sección"].apply(limpia_seccion_interna)
        
        argos_df["_k_per"] = argos_df["Periodo"].astype(str).str.strip().apply(lambda x: x[:-2] if x.endswith(".0") else x)
        argos_df["_k_niv"] = argos_df["Nivel"].apply(normalizar_para_cruce)
        argos_df["_k_sub"] = argos_df["Área"].apply(normalizar_para_cruce)
        argos_df["_k_crs"] = argos_df["No..Curso"].astype(str).str.strip().apply(lambda x: x[:-2] if x.endswith(".0") else x)
        argos_df["_k_sec"] = argos_df["Grupo"].apply(limpia_seccion_interna)
        
        llaves_cruce = ["_k_per", "_k_niv", "_k_sub", "_k_crs", "_k_sec"]
        
        argos_subset = argos_df[llaves_cruce + ["NRC"]]
        fusion = solicitud_p2.merge(argos_subset, on=llaves_cruce, how="left")
        fusion.drop(columns=llaves_cruce, inplace=True)
        
        fusion = fusion.drop_duplicates(subset=["NRC"], keep="first")
        columnas_finales = ["NRC"] + [c for c in fusion.columns if c != "NRC" and c != "ArchivoOrigen"]
        
        st.success("¡Cruce completado con éxito!")
        st.markdown("#### 📥 Descarga tus Excels modificados con la pestaña 'CRNs':")
        
        for name, sub in fusion.groupby("ArchivoOrigen"):
            if name in dict_bytes_altas:
                df_escribir = sub[columnas_finales].copy()
                original_bytes = dict_bytes_altas[name]
                wb = openpyxl.load_workbook(io.BytesIO(original_bytes))
                
                if "CRNs" in wb.sheetnames:
                    del wb["CRNs"]
                
                ws = wb.create_sheet(title="CRNs")
                ws.append(list(df_escribir.columns))
                for r in df_escribir.values:
                    ws.append(list(r))
                
                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                excel_buffer.seek(0)
                
                nombre_base_excel = name.rsplit('.', 1)[0] if '.' in name else name
                excel_filename = f"{nombre_base_excel} con hoja CRNs.xlsx"
                
                st.download_button(
                    label=f"⬇️ Descargar {excel_filename}",
                    data=excel_buffer.getvalue(),
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{name}"
                )
            else:
                st.warning(f"⚠️ El archivo '{name}' se detectó en el resumen, pero no se subió en el bloque de ALTAS de esta pestaña.")

        # ============================================================
        # 🧩 GENERACIÓN DE CSV DE CLÚSTER UNIFICADO (TODOS LOS EXCEL EN UNO SOLO)
        # ============================================================
        st.markdown("---")
        st.markdown("#### 🧩 Archivo de Carga (Formato Clúster)")
        
        columnas_cluster = [
            "Periodo", "CRN", "Tipo.de.Reunión", "Fecha.Inicio", "Fecha.Fin", 
            "Dom", "Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "horarioIni", 
            "horarioFin", "Inicio.de.sesión", "edificio", "salon", 
            "Tipo.de.horario", "indCategoria", "idInstructor", 
            "responsabilidad", "Ind.principal", "ind.sobre.paso", 
            "datocomplementario"
        ]
        
        df_cluster = pd.DataFrame(columns=columnas_cluster)
        
        df_cluster["Periodo"] = fusion["Periodo"].apply(lambda x: format_r_style(x))
        df_cluster["CRN"] = fusion["NRC"].apply(lambda x: format_r_style(x))
        
        if "Clúster" in fusion.columns:
            df_cluster["datocomplementario"] = fusion["Clúster"]
        elif "Cluster" in fusion.columns:
            df_cluster["datocomplementario"] = fusion["Cluster"]
        else:
            st.warning("⚠️ No se detectó la columna 'Clúster' en los Excel. El campo 'datocomplementario' se exportará vacío.")
            
        df_cluster = df_cluster.fillna("")
        
        csv_cluster_bytes = df_cluster.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        
        st.download_button(
            label="🧩 📥 DESCARGAR CSV DE CARGA DE CLÚSTER UNIFICADO",
            data=csv_cluster_bytes,
            file_name="carga_cluster_unificado.csv",
            mime="text/csv",
            use_container_width=True,
            type="primary",
            key="btn_csv_cluster_final"
        )
