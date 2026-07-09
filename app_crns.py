# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import csv
import io
import unicodedata
import openpyxl
import zipfile
from difflib import SequenceMatcher

# ================= 1. CONFIGURACIÓN Y FUNCIONES ESTRUCTURALES =================
HOJA_ALTAS = "ALTAS"
UMBRAL_FUZZY = 0.82  

# Parámetros exactos de R para clonar su comportamiento de write.csv()
CSV_KWARGS_R = {
    'index': False,
    'encoding': 'utf-8',
    'quoting': csv.QUOTE_NONNUMERIC, 
    'lineterminator': '\r\n',        
    'na_rep': 'NA'                   
}

def quitar_acentos(t):
    if pd.isna(t) or t is None: return ""
    return "".join(c for c in unicodedata.normalize("NFD", str(t)) if unicodedata.category(c) != "Mn")

def normalizar_para_cruce(t):
    return quitar_acentos(str(t).upper().strip())

def similitud(a, b): 
    return SequenceMatcher(None, a, b).ratio()

def format_r_string(val):
    if pd.isna(val) or val is None:
        return np.nan
    s = str(val).strip()
    if s.lower() == "nan" or s == "":
        return np.nan
    if s.endswith(".0"): s = s[:-2]
    return s

def limpia_seccion_interna(x):
    if pd.isna(x): return ""
    s = str(x).strip()
    if s.lower() == "nan": return ""
    if s.endswith(".0"): s = s[:-2]
    if s.isdigit(): return f"{int(s):02d}"
    return s

def obtener_base_nombre(filename):
    """Extrae el nombre base de un archivo removiendo extensiones y sufijos cortos de versión (V1, V2, etc.)"""
    if not filename: return ""
    s = filename.upper().strip()
    for ext in [".CSV", ".XLSX", ".XLS"]:
        if s.endswith(ext): s = s[:-len(ext)]
    
    # Intentar remover de forma dinámica patrones como _V1, _V2, _V3... hasta _V20
    for v_num in range(1, 21):
        sufijo = f"_V{v_num}"
        if s.endswith(sufijo):
            s = s[:-len(sufijo)]
            break
            
    return s.strip()

# Inicialización de estados en memoria de Streamlit
if "original_files_bytes" not in st.session_state: st.session_state.original_files_bytes = {}
if "res_auditoria" not in st.session_state: st.session_state.res_auditoria = None
if "raw_altas" not in st.session_state: st.session_state.raw_altas = None
if "ready_for_download" not in st.session_state: st.session_state.ready_for_download = False
if "zip_file_bytes" not in st.session_state: st.session_state.zip_file_bytes = None
if "csv_files_to_download" not in st.session_state: st.session_state.csv_files_to_download = {}

st.set_page_config(page_title="Consola Iris Cavazos", page_icon="🎛️", layout="wide")
st.title("🎛️ Consola de Control de Materias e Inyección de NRCs (Flujo Multi-CSV)")
st.markdown("---")

tab1, tab_err, tab3 = st.tabs([
    "1️⃣ Proceso: Validación y Generar CSVs", 
    "⚠️ Reporte de Errores (Extraer Delta)", 
    "2️⃣ Proceso: Inyección de NRCs Masiva (ARGOS)"
])

# ============================================================
# PESTAÑA 1: VALIDACIÓN Y GENERACIÓN DE CSV INDIVIDUALES
# ============================================================
with tab1:
    st.header("Validación de Claves y Generación de bloques CSV")
    
    col1, col2 = st.columns(2)
    with col1:
        file_cat = st.file_uploader("📑 Catálogo de Materias Estatales (Excel)", type=["xlsx"])
    with col2:
        files_altas = st.file_uploader("📁 Archivos de ALTAS (Puedes subir varios Excel)", accept_multiple_files=True, type=["xlsx"])
    
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
                        s_val = format_r_string(f.get("Subj"))
                        c_val = format_r_string(f.get("Crse"))
                        
                        indice_cat.setdefault(niv, []).append({
                            "mat_orig": mat_o,
                            "mat_norm": normalizar_para_cruce(f.get("Materia")), 
                            "subj": s_val, 
                            "crse": c_val
                        })
                        
                        if pd.notna(s_val) and pd.notna(c_val):
                            s_norm = normalizar_para_cruce(s_val)
                            c_norm = c_val
                            indice_cat_claves[(s_norm, c_norm)] = mat_o
            
            piezas = []
            for f in files_altas:
                st.info(f"🔍 Revisando archivo original: **{f.name}**")
                st.session_state.original_files_bytes[f.name] = f.getvalue()
                
                xls_a = pd.ExcelFile(f)
                hojas_reales = [h for h in xls_a.sheet_names if h.strip().upper() == HOJA_ALTAS]
                if hojas_reales:
                    df_a = xls_a.parse(hojas_reales[0])
                    
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
                    subj_orig = format_r_string(fila.get("Subject"))
                    crse_orig = format_r_string(fila.get("Course"))
                    
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
                st.success("¡Revisión de catálogos finalizada!")
            else:
                st.error(f"❌ Ninguno de los archivos subidos tiene filas válidas en la pestaña '{HOJA_ALTAS}'")

    if st.session_state.res_auditoria is not None:
        st.markdown("### ⚖️ Mesa de Control Interactiva")
        df_aud = st.session_state.res_auditoria
        archivos_subidos = df_aud["Archivo"].unique()
        
        for arch in archivos_subidos:
            df_file = df_aud[df_aud["Archivo"] == arch]
            errores_filas = df_file[df_file["Comentario"] != "Todo correcto"]
            total_detalles = len(errores_filas)
            
            if total_detalles == 0:
                st.success(f"✅ **{arch}** — ¡Todo limpio, listo para procesar!")
            else:
                with st.expander(f"⚠️ **{arch}** — ({total_detalles} advertencias detectadas)", expanded=True):
                    quitar_rep = st.checkbox("🔍 Agrupar repetidas", value=True, key=f"rep_{arch}")
                    df_vista = errores_filas.drop_duplicates(subset=["Materia Excel", "Materia Catálogo", "Subj Original", "Crse Original", "Comentario"]) if quitar_rep else errores_filas
                    columnas_vista = ["Luz Verde", "Materia Excel", "Materia Catálogo", "Comentario", "Subj Original", "Crse Original", "Subj Sugerido", "Crse Sugerido"]
                    
                    df_editado_archivo = st.data_editor(
                        df_vista[columnas_vista],
                        hide_index=True,
                        disabled=["Materia Excel", "Materia Catálogo", "Comentario", "Subj Original", "Crse Original"],
                        column_config={"Luz Verde": st.column_config.CheckboxColumn("¿Aplicar?")},
                        key=f"editor_{arch}",
                        use_container_width=True
                    )
                    
                    for _, row in df_editado_archivo.iterrows():
                        mascara = (
                            (st.session_state.res_auditoria["Archivo"] == arch) & 
                            (st.session_state.res_auditoria["Materia Excel"] == row["Materia Excel"]) & 
                            (st.session_state.res_auditoria["Subj Original"] == row["Subj Original"]) & 
                            (st.session_state.res_auditoria["Crse Original"] == row["Crse Original"])
                        )
                        st.session_state.res_auditoria.loc[mascara, "Luz Verde"] = row["Luz Verde"]
                        st.session_state.res_auditoria.loc[mascara, "Subj Sugerido"] = row["Subj Sugerido"]
                        st.session_state.res_auditoria.loc[mascara, "Crse Sugerido"] = row["Crse Sugerido"]
        
        st.markdown("---")
        if st.button("💾 Generar Bloque de Archivos CSV", type="primary"):
            corregido = st.session_state.raw_altas.copy()
            for _, row in st.session_state.res_auditoria.iterrows():
                if row["Luz Verde"] and pd.notna(row["Subj Sugerido"]):
                    corregido.loc[row["idx"], "Subject"] = row["Subj Sugerido"]
                    corregido.loc[row["idx"], "Course"] = row["Crse Sugerido"]
            
            st.session_state.csv_files_to_download = {}
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for name, sub in corregido.groupby("ArchivoOrigen"):
                    resultado_df = pd.DataFrame()
                    
                    resultado_df["PERIODO"] = sub["Periodo"].apply(format_r_string)
                    resultado_df["SEDE"] = sub["Campus"].apply(format_r_string)
                    resultado_df["SUBJ"] = sub["Subject"].apply(format_r_string)
                    resultado_df["COURSE"] = sub["Course"].apply(format_r_string)
                    resultado_df["PARTEPERIODO"] = sub["Parte de Periodo"].apply(format_r_string)
                    resultado_df["STATUS"] = sub["Estatus"].apply(format_r_string)
                    
                    resultado_df["CAPACIDAD"] = pd.to_numeric(sub["Capacidad"], errors='coerce').astype('Int64')
                    resultado_df["GRUPOS"] = pd.to_numeric(1, errors='ignore').astype('Int64')
                    resultado_df["SECCION"] = pd.to_numeric(sub["Sección"], errors='coerce').astype('Int64')
                    
                    resultado_df["TIPODEHORARIO"] = sub["Tipo de Horario"].apply(format_r_string)
                    resultado_df["METODO_EDUCATIVO"] = sub["Método Educativo"].apply(format_r_string)
                    resultado_df["SOCIODEINTEGRACION"] = "D2L"
                    resultado_df["MODODECALIFICAR"] = sub["Modo de Calificar"].apply(format_r_string)
                    resultado_df["SESION"] = sub["Sesion"].apply(format_r_string)
                    
                    columnas_ordenadas = ["PERIODO", "SEDE", "SUBJ", "COURSE", "PARTEPERIODO", "STATUS",
                                          "CAPACIDAD", "GRUPOS", "SECCION", "TIPODEHORARIO",
                                          "METODO_EDUCATIVO", "SOCIODEINTEGRACION", "MODODECALIFICAR", "SESION"]
                    resultado_df = resultado_df[columnas_ordenadas]
                    
                    csv_filename = f"{name.rsplit('.', 1)[0] if '.' in name else name}.csv"
                    csv_string = resultado_df.to_csv(**CSV_KWARGS_R)
                    
                    zip_file.writestr(csv_filename, csv_string.encode('utf-8'))
                    st.session_state.csv_files_to_download[csv_filename] = csv_string.encode('utf-8')
            
            st.session_state.zip_file_bytes = zip_buffer.getvalue()
            st.session_state.ready_for_download = True
            st.rerun()

        if st.session_state.ready_for_download:
            st.markdown("### 📥 Panel de Descarga")
            st.download_button("💥 📥 DESCARGAR TODOS LOS CSVs (.ZIP)", data=st.session_state.zip_file_bytes, file_name="archivos_carga_banner.zip", mime="application/zip", use_container_width=True, type="primary")

# ============================================================
# PESTAÑA 2: REPORTE DE ERRORES (RECORTE DELTA)
# ============================================================
with tab_err:
    st.header("⚠️ Extracción de Líneas de Error (Delta)")
    st.markdown("Si Banner rechaza registros de un archivo CSV específico, sube ese CSV aquí junto con el reporte de errores para extraer las líneas que debes corregir a mano.")
    
    metodo_input = st.radio("🛠️ Ingreso de líneas fallidas:", ["Subir archivo Excel de Banner", "Escribir números de línea manualmente"], horizontal=True)
    
    file_errores = None
    txt_errores = ""
    input_valido = False
    
    if metodo_input == "Subir archivo Excel de Banner":
        file_errores = st.file_uploader("📥 Cargar Reporte de Errores de Banner (.xlsx)", type=["xlsx"])
        if file_errores: input_valido = True
    else:
        txt_errores = st.text_input("✍️ Números de línea separados por comas (Ejemplo: 4, 18, 55):")
        if txt_errores.strip(): input_valido = True
            
    st.markdown("---")
    origen_csv = st.radio("¿Qué CSV vas a recortar para corregir?", ["Utilizar un CSV de la Pestaña 1 (En memoria)", "Subir un CSV desde tu equipo"])
    
    dict_csvs_a_procesar = {}
    if origen_csv == "Utilizar un CSV de la Pestaña 1 (En memoria)":
        if st.session_state.csv_files_to_download:
            opciones = list(st.session_state.csv_files_to_download.keys())
            sel = st.selectbox("Elige el CSV a depurar:", opciones)
            if sel: dict_csvs_a_procesar[sel] = pd.read_csv(io.BytesIO(st.session_state.csv_files_to_download[sel]), encoding="utf-8")
        else:
            st.warning("⚠️ No hay CSVs generados en la memoria de la pestaña 1 aún.")
    else:
        file_csv_manual = st.file_uploader("Subir CSV original completo", type=["csv"], key="csv_manual_err")
        if file_csv_manual: 
            dict_csvs_a_procesar[file_csv_manual.name] = pd.read_csv(file_csv_manual, encoding="utf-8")
                
    num_version = st.number_input("🔢 Número de corrección (V):", min_value=1, value=1)
    
    if input_valido and dict_csvs_a_procesar:
        if st.button("🔍 Generar Segmento Corto de Errores", type="primary"):
            renglones_errores = []
            if metodo_input == "Subir archivo Excel de Banner":
                df_err_excel = pd.read_excel(file_errores, skiprows=2)
                renglones_errores = df_err_excel["Línea"].dropna().astype(int).unique().tolist()
            else:
                renglones_errores = [int(p.strip()) for p in txt_errores.split(",") if p.strip().isdigit()]
                
            for nombre_archivo, df_datos in dict_csvs_a_procesar.items():
                indices = [r - 2 for r in renglones_errores if 0 <= (r - 2) < len(df_datos)]
                df_delta = df_datos.iloc[indices].copy()
                if not df_delta.empty:
                    csv_bytes = df_delta.to_csv(**CSV_KWARGS_R).encode("utf-8")
                    out_name = f"{nombre_archivo.rsplit('.', 1)[0]}_V{num_version}.csv"
                    st.download_button(label=f"📥 Descargar fragmento: {out_name}", data=csv_bytes, file_name=out_name, mime="text/csv", use_container_width=True)

# ============================================================
# PESTAÑA 3: INYECTAR EN REPORTE ARGOS (CONTRUCCIÓN EN LOTE)
# ============================================================
with tab3:
    st.header("Inyección Masiva de NRCs e Inyección de Hojas 'CRNs'")
    st.markdown("Sube tu reporte de ARGOS, **todos tus archivos CSV finales** (tanto los que salieron limpios directos de la pestaña 1 como tus archivos corregidos a mano como `_V1, _V2`), y tus **Excels originales**.")
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        file_argos = st.file_uploader("📊 1. Cargar Reporte ARGOS (.csv)", type=["csv"])
    with col_b:
        files_csv_finales = st.file_uploader("📝 2. Sube TODOS los CSVs finales (Limpios + Versiones _V1/_V2)", type=["csv"], accept_multiple_files=True)
    with col_c:
        files_xlsx_originales = st.file_uploader("📁 3. Sube los archivos EXCEL originales (.xlsx)", type=["xlsx"], accept_multiple_files=True)
        
    if file_argos and files_csv_finales and files_xlsx_originales:
        if st.button("🚀 Ejecutar Cruce e Inyección de Lote", type="primary"):
            try:
                # 1. Cargar reporte de ARGOS limpiando espacios/comillas extrañas en los nombres de las columnas
                argos_df = pd.read_csv(file_argos, encoding="utf-8", on_bad_lines='skip')
                argos_df.columns = [str(c).replace('"', '').replace("'", "").strip() for c in argos_df.columns]
                
                # Mapeo inteligente de las columnas críticas de ARGOS por si cambian de puntos "."
                mapa_columnas_argos = {}
                for col in argos_df.columns:
                    col_upper = col.upper()
                    if col_upper in ["PERIODO", "ESTATUS", "NRC", "ÁREA", "GRUPO", "NIVEL", "CAMPUS", "CLUSTER", "CLÚSTER"]:
                        mapa_columnas_argos[col_upper] = col
                    elif "CURSO" in col_upper:
                        mapa_columnas_argos["CURSO"] = col  # Detectará "No..Curso" o "No.Curso" indistintamente
                
                # Validar que las columnas mínimas existan tras la limpieza
                columnas_requeridas = ["PERIODO", "ÁREA", "CURSO", "GRUPO", "NRC"]
                faltantes = [c for c in columnas_requeridas if c not in mapa_columnas_argos]
                if faltantes:
                    st.error(f"❌ No se pudieron encontrar mapeadas las columnas {faltantes} en tu archivo de ARGOS. Columnas leídas: {list(argos_df.columns)}")
                    st.stop()
                
                # Construcción de llaves de cruce estables en ARGOS
                argos_df["_k_per"] = argos_df[mapa_columnas_argos["PERIODO"]].astype(str).str.strip().apply(lambda x: x[:-2] if x.endswith(".0") else x)
                argos_df["_k_sub"] = argos_df[mapa_columnas_argos["ÁREA"]].apply(normalizar_para_cruce)
                argos_df["_k_crs"] = argos_df[mapa_columnas_argos["CURSO"]].astype(str).str.strip().apply(lambda x: x[:-2] if x.endswith(".0") else x)
                argos_df["_k_sec"] = argos_df[mapa_columnas_argos["GRUPO"]].apply(limpia_seccion_interna)
                if "NIVEL" in mapa_columnas_argos:
                    argos_df["_k_niv"] = argos_df[mapa_columnas_argos["NIVEL"]].apply(normalizar_para_cruce)
                
                # Diccionario para guardar las estructuras XLSX en bytes listas para descargar
                excels_inyectados_zip = io.BytesIO()
                listado_clusters = []
                xlsx_mapeados_correctamente = 0
                
                # Mapear los bytes de los excels cargados por su nombre base simplificado
                dict_excels_originales = {}
                for fx in files_xlsx_originales:
                    base_x = obtener_base_nombre(fx.name)
                    dict_excels_originales[base_x] = {"filename": fx.name, "bytes": fx.getvalue()}

                with zipfile.ZipFile(excels_inyectados_zip, "w", zipfile.ZIP_DEFLATED) as zip_out:
                    
                    # 2. Iterar sobre cada CSV subido (ya sea limpio o corregido por el usuario)
                    for fc in files_csv_finales:
                        base_c = obtener_base_nombre(fc.name)
                        df_csv = pd.read_csv(io.BytesIO(fc.getvalue()), encoding="utf-8")
                        
                        # Construcción de llaves de cruce del CSV
                        df_csv["_k_per"] = df_csv["PERIODO"].astype(str).str.strip().apply(lambda x: x[:-2] if x.endswith(".0") else x)
                        df_csv["_k_sub"] = df_csv["SUBJ"].apply(normalizar_para_cruce)
                        df_csv["_k_crs"] = df_csv["COURSE"].astype(str).str.strip().apply(lambda x: x[:-2] if x.endswith(".0") else x)
                        df_csv["_k_sec"] = df_csv["SECCION"].apply(limpia_seccion_interna)
                        
                        llaves_match = ["_k_per", "_k_sub", "_k_crs", "_k_sec"]
                        if "_k_niv" in argos_df.columns and "Nivel" in df_csv.columns:
                            df_csv["_k_niv"] = df_csv["Nivel"].apply(normalizar_para_cruce)
                            llaves_match.append("_k_niv")
                            
                        # Hacer el cruce con el segmento de ARGOS para pescar el NRC
                        col_nrc_real = mapa_columnas_argos["NRC"]
                        fusion = df_csv.merge(argos_df[llaves_match + [col_nrc_real]], on=llaves_match, how="left")
                        
                        # Renombrar la columna del NRC a un estándar limpio 'NRC' si es necesario
                        if col_nrc_real != "NRC" and col_nrc_real in fusion.columns:
                            fusion["NRC"] = fusion[col_nrc_real]
                            fusion = fusion.drop(columns=[col_nrc_real])
                            
                        fusion = fusion.drop_duplicates(subset=["NRC"], keep="first") if "NRC" in fusion.columns else fusion
                        
                        # Limpiar llaves temporales
                        columnas_limpias = [c for c in fusion.columns if not c.startswith("_k_")]
                        fusion_final = fusion[columnas_limpias].copy()
                        
                        if "NRC" in fusion_final.columns:
                            # Reordenar para colocar el NRC en primer lugar
                            cols_orden = ["NRC"] + [c for c in fusion_final.columns if c != "NRC"]
                            fusion_final = fusion_final[cols_orden]
                        
                        # 3. Emparejar e inyectar en el Excel original
                        if base_c in dict_excels_originales:
                            excel_info = dict_excels_originales[base_c]
                            wb = openpyxl.load_workbook(io.BytesIO(excel_info["bytes"]))
                            
                            if "CRNs" in wb.sheetnames: 
                                del wb["CRNs"]
                            ws = wb.create_sheet(title="CRNs")
                            
                            # Escribir dataframe en la nueva pestaña
                            ws.append(list(fusion_final.columns))
                            for r in fusion_final.values:
                                ws.append(list(r))
                                
                            excel_buffer = io.BytesIO()
                            wb.save(excel_buffer)
                            
                            # Guardar en el ZIP final conservando el nombre original del excel
                            zip_out.writestr(excel_info["filename"], excel_buffer.getvalue())
                            xlsx_mapeados_correctamente += 1
                        else:
                            st.warning(f"⚠️ El archivo CSV `{fc.name}` no encontró su contraparte Excel `.xlsx` vinculada con el nombre `{base_c}`.")

                        # 4. Extraer datos para el Clúster Unificado
                        df_cl_temp = pd.DataFrame()
                        df_cl_temp["Periodo"] = df_csv["PERIODO"].apply(format_r_string)
                        df_cl_temp["CRN"] = pd.to_numeric(fusion["NRC"], errors='coerce').astype('Int64') if "NRC" in fusion.columns else np.nan
                        
                        # Intentar jalar la columna Clúster dinámicamente según lo que venga de ARGOS o del CSV
                        col_cl_origen = "Clúster" if "Clúster" in df_csv.columns else ("Cluster" if "Cluster" in df_csv.columns else None)
                        df_cl_temp["datocomplementario"] = df_csv[col_cl_origen] if col_cl_origen else np.nan
                        listado_clusters.append(df_cl_temp)
                
                st.success(f"🎉 ¡Procesamiento completado con éxito! Se inyectaron las pestañas CRNs en {xlsx_mapeados_correctamente} archivos Excel.")
                
                # 5. Generar descargas en la interfaz
                st.markdown("### 📥 Panel de Resultados Listos")
                
                # Descarga de todos los Excels parchados en un único ZIP
                if xlsx_mapeados_correctamente > 0:
                    st.download_button(
                        label="📁 📥 DESCARGAR TODOS LOS EXCELES CON CRNs (.ZIP)",
                        data=excels_inyectados_zip.getvalue(),
                        file_name="archivos_excel_con_CRNs.zip",
                        mime="application/zip",
                        use_container_width=True,
                        type="primary"
                    )
                
                # Consolidación del archivo clúster final de todos los archivos procesados
                if listado_clusters:
                    df_cluster_total = pd.concat(listado_clusters, ignore_index=True)
                    # Columnas reglamentarias vacías requeridas por tu layout de R
                    columnas_cluster_completas = [
                        "Periodo", "CRN", "Tipo.de.Reunión", "Fecha.Inicio", "Fecha.Fin", "Dom", "Lun", "Mar", 
                        "Mie", "Jue", "Vie", "Sab", "horarioIni", "horarioFin", "Inicio.de.sesión", "edificio", 
                        "salon", "Tipo.de.horario", "indCategoria", "idInstructor", "responsabilidad", 
                        "Ind.principal", "ind.sobre.paso", "datocomplementario"
                    ]
                    for col in columnas_cluster_completas:
                        if col not in df_cluster_total.columns:
                            df_cluster_total[col] = np.nan
                            
                    df_cluster_total = df_cluster_total[columnas_cluster_completas]
                    csv_cluster_bytes = df_cluster_total.to_csv(**CSV_KWARGS_R).encode("utf-8")
                    
                    st.download_button(
                        label="🧩 📥 DESCARGAR CLÚSTER UNIFICADO TOTAL",
                        data=csv_cluster_bytes,
                        file_name="cluster_unificado.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                    
            except Exception as e:
                st.error(f"❌ Error crítico durante el cruce en lote con ARGOS: {str(e)}")
