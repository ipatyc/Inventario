# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import csv
import io
import unicodedata
import openpyxl
import zipfile
import re
from difflib import SequenceMatcher

# ================= 1. CONFIGURACIÓN Y FUNCIONES ESTRUCTURALES =================
HOJA_ALTAS = "ALTAS"
HOJA_SALIDA_NRC = "NRC"  # Pestaña en mayúsculas estrictas
UMBRAL_FUZZY = 0.82  

CSV_KWARGS_R = {
    'index': False,
    'encoding': 'utf-8',
    'quoting': csv.QUOTE_NONNUMERIC, 
    'lineterminator': '\r\n',        
    'na_rep': 'NA'                   
}

# Plantilla estricta de 24 columnas para el Clúster
COLUMNAS_CLUSTER_FINAL = [
    "Periodo", "CRN", "Tipo.de.Reunión", "Fecha.Inicio", "Fecha.Fin", "Dom", "Lun", 
    "Mar", "Mie", "Jue", "Vie", "Sab", "horarioIni", "horarioFin", "Inicio.de.sesión", 
    "edificio", "salon", "Tipo.de.horario", "indCategoria", "idInstructor", 
    "responsabilidad", "Ind.principal", "ind.sobre.paso", "datocomplementario"
]

def quitar_acentos(t):
    if pd.isna(t) or t is None: 
        return ""
    return "".join(c for c in unicodedata.normalize("NFD", str(t)) if unicodedata.category(c) != "Mn")

def normalizar_para_cruce(t):
    if pd.isna(t) or t is None:
        return ""
    s = str(t).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return quitar_acentos(s).upper().strip()

def similitud(a, b): 
    return SequenceMatcher(None, a, b).ratio()

def limpiar_clave_texto(val):
    """Garantiza la conversión de cualquier tipo (float, int, NaN) a un string limpio para llaves de cruce"""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.lower() == "nan" or s == "":
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    return s

def format_r_string(val):
    if pd.isna(val) or val is None:
        return np.nan
    s = str(val).strip()
    if s.lower() == "nan" or s == "":
        return np.nan
    if s.endswith(".0"): 
        s = s[:-2]
    return s

def limpia_seccion_interna(x):
    if pd.isna(x): 
        return ""
    s = str(x).strip()
    if s.lower() == "nan" or s == "": 
        return ""
    if s.endswith(".0"): 
        s = s[:-2]
    if s.isdigit(): 
        return f"{int(s):02d}"
    return s

def obtener_base_y_version(filename):
    """Extrae el nombre base de un archivo y su número de versión (V1, V2...)"""
    if not filename: 
        return "", 0
    s = str(filename).upper().strip()
    for ext in [".CSV", ".XLSX", ".XLS"]:
        if s.endswith(ext): 
            s = s[:-len(ext)]
    
    match = re.search(r'_V(\d+)$', s)
    if match:
        return s[:match.start()].strip(), int(match.group(1))
    match = re.search(r'V(\d+)$', s)
    if match:
        return s[:match.start()].strip('_ '), int(match.group(1))
    return s.strip(), 0

# Inicialización de estados en memoria global de Streamlit
if "original_files_bytes" not in st.session_state: 
    st.session_state.original_files_bytes = {}
if "res_auditoria" not in st.session_state: 
    st.session_state.res_auditoria = None
if "raw_altas" not in st.session_state: 
    st.session_state.raw_altas = None
if "ready_for_download" not in st.session_state: 
    st.session_state.ready_for_download = False
if "zip_file_bytes" not in st.session_state: 
    st.session_state.zip_file_bytes = None
if "csv_files_to_download" not in st.session_state: 
    st.session_state.csv_files_to_download = {}
if "delta_files" not in st.session_state: 
    st.session_state.delta_files = {}
if "final_argos_zip" not in st.session_state: 
    st.session_state.final_argos_zip = None

st.set_page_config(page_title="Consola Iris Cavazos", page_icon="️", layout="wide")
st.title("️ Consola de Control de Materias e Inyección de NRCs (Flujo Multi-CSV)")
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
                            if s > mejor_s: 
                                mejor_s, mejor = s, c
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
                    
                    columnas_ordenadas = [
                        "PERIODO", "SEDE", "SUBJ", "COURSE", "PARTEPERIODO", "STATUS",
                        "CAPACIDAD", "GRUPOS", "SECCION", "TIPODEHORARIO",
                        "METODO_EDUCATIVO", "SOCIODEINTEGRACION", "MODODECALIFICAR", "SESION"
                    ]
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
            st.download_button(
                "💥 📥 DESCARGAR TODOS LOS CSVs (.ZIP)", 
                data=st.session_state.zip_file_bytes, 
                file_name="archivos_carga_banner.zip", 
                mime="application/zip", 
                use_container_width=True, 
                type="primary"
            )

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
        if file_errores: 
            input_valido = True
    else:
        txt_errores = st.text_input("✍️ Números de línea separados por comas (Ejemplo: 4, 18, 55):")
        if txt_errores.strip(): 
            input_valido = True
            
    st.markdown("---")
    origen_csv = st.radio("¿Qué CSV vas a recortar para corregir?", ["Utilizar un CSV de la Pestaña 1 (En memoria)", "Subir un CSV desde tu equipo"])
    
    dict_csvs_a_procesar = {}
    if origen_csv == "Utilizar un CSV de la Pestaña 1 (En memoria)":
        if st.session_state.csv_files_to_download:
            options = list(st.session_state.csv_files_to_download.keys())
            sel = st.selectbox("Elige el CSV a depurar:", options)
            if sel: 
                dict_csvs_a_procesar[sel] = pd.read_csv(io.BytesIO(st.session_state.csv_files_to_download[sel]), encoding="utf-8")
        else:
            st.warning("⚠️ No hay CSVs generados en la memoria de la pestaña 1 aún.")
    else:
        file_csv_manual = st.file_uploader("Subir CSV original completo", type=["csv"], key="csv_manual_err")
        if file_csv_manual: 
            dict_csvs_a_procesar[file_csv_manual.name] = pd.read_csv(file_csv_manual, encoding="utf-8")
                
    num_version = st.number_input("🔢 Número de corrección (V):", min_value=1, value=1)
    
    if input_valido and dict_csvs_a_procesar:
        if st.button("🔍 Generar Segmento Corto de Errores", type="primary"):
            st.session_state.delta_files = {}  
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
                    st.session_state.delta_files[out_name] = csv_bytes
            st.rerun()

    if st.session_state.delta_files:
        st.markdown("### 📥 Fragmentos de Error Listos")
        for out_name, csv_bytes in st.session_state.delta_files.items():
            st.download_button(label=f"📥 Descargar fragmento: {out_name}", data=csv_bytes, file_name=out_name, mime="text/csv", use_container_width=True)

# ============================================================
# PESTAÑA 3 (AHORA PASO 2): INYECTAR REPORTE ARGOS Y CLÚSTER
# ============================================================
with tab3:
    st.header("Inyección de NRCs y Generación Estricta de Clúster")
    
    # --- FUNCIONES LOCALES DE RESPALDO PARA EVITAR DEFINED ERRORS ---
    def _local_limpiar_texto(val):
        if pd.isna(val) or val is None:
            return ""
        s = str(val).strip()
        if s.lower() == "nan" or s == "":
            return ""
        if s.endswith(".0"):
            s = s[:-2]
        return s

    def _local_normalizar_mayusculas(val):
        return _local_limpiar_texto(val).upper()

    def _local_seccion_a_dos_digitos(val):
        s = _local_limpiar_texto(val)
        if not s:
            return ""
        if s.isdigit():
            return f"{int(s):02d}"
        return s
    # -----------------------------------------------------------------

    col_a, col_b, col_c = st.columns(3)
    with col_a: file_argos = st.file_uploader("📊 1. Reporte ARGOS (.csv)", type=["csv"])
    with col_b: files_csv_finales = st.file_uploader("📝 2. Archivos CSV finales (Originales + V1, V2...)", type=["csv"], accept_multiple_files=True)
    with col_c: files_xlsx_originales = st.file_uploader("📁 3. Archivos EXCEL originales (.xlsx)", type=["xlsx"], accept_multiple_files=True)
        
    if file_argos and files_csv_finales and files_xlsx_originales:
        if st.button("🚀 Procesar Cruce e Inyectar Pestaña NRC + Clúster", type="primary"):
            try:
                # 1. LEER ARGOS Y LIMPIAR COLUMNAS
                argos_df = pd.read_csv(file_argos, encoding="utf-8", on_bad_lines='skip')
                
                # Normalización de encabezados
                argos_df.columns = [
                    re.sub(r'\.+', '.', str(c).replace('"', '').replace("'", "").strip()) 
                    for c in argos_df.columns
                ]
                
                col_curso = "No.Curso" if "No.Curso" in argos_df.columns else ("No..Curso" if "No..Curso" in argos_df.columns else None)
                if col_curso is None:
                    columnas_candidatas = [c for c in argos_df.columns if "Curso" in c]
                    if columnas_candidatas:
                        col_curso = columnas_candidatas[0]
                    else:
                        raise KeyError("No se encontró la columna del número de curso en el reporte de ARGOS.")

                argos_df["Periodo"] = argos_df["Periodo"].apply(_local_limpiar_texto)
                argos_df["Nivel"] = argos_df["Nivel"].apply(_local_normalizar_mayusculas)
                argos_df["Área"] = argos_df["Área"].apply(_local_normalizar_mayusculas)
                argos_df[col_curso] = argos_df[col_curso].apply(_local_limpiar_texto)
                argos_df["Grupo"] = argos_df["Grupo"].apply(_local_seccion_a_dos_digitos)
                
                argos_df["_llave_maestra"] = (argos_df["Periodo"] + "_" + 
                                              argos_df["Nivel"] + "_" + 
                                              argos_df["Área"] + "_" + 
                                              argos_df[col_curso] + "_" + 
                                              argos_df["Grupo"])
                
                argos_df = argos_df.drop_duplicates(subset=["_llave_maestra"])
                mapa_nrcs = dict(zip(argos_df["_llave_maestra"], argos_df["NRC"]))

                # 2. CONSOLIDAR VERSIONES DE CSV (Última versión gana por PERIODO, SEDE, SECCION)
                dict_csvs_agrupados = {}
                for fc in files_csv_finales:
                    base_name, version = obtener_base_y_version(fc.name)
                    dict_csvs_agrupados.setdefault(base_name, {})[version] = fc
                
                dict_csvs_finalizados = {}  

                for base_name, versiones in dict_csvs_agrupados.items():
                    lista_dfs = []
                    for v in sorted(versiones.keys()):
                        df_v = pd.read_csv(io.BytesIO(versiones[v].getvalue()), encoding="utf-8")
                        df_v.reset_index(drop=True, inplace=True)
                        lista_dfs.append(df_v)
                    
                    if not lista_dfs:
                        continue
                    
                    df_consolidado = pd.concat(lista_dfs, ignore_index=True)
                    df_consolidado.reset_index(drop=True, inplace=True)
                    
                    if "PERIODO" in df_consolidado.columns and "SEDE" in df_consolidado.columns and "SECCION" in df_consolidado.columns:
                        df_consolidado["_temp_per"] = df_consolidado["PERIODO"].apply(_local_limpiar_texto)
                        df_consolidado["_temp_sed"] = df_consolidado["SEDE"].apply(_local_limpiar_texto)
                        df_consolidado["_temp_sec"] = df_consolidado["SECCION"].apply(_local_seccion_a_dos_digitos)
                        
                        df_consolidado = df_consolidado.drop_duplicates(subset=["_temp_per", "_temp_sed", "_temp_sec"], keep='last')
                        df_consolidado.drop(columns=["_temp_per", "_temp_sed", "_temp_sec"], inplace=True)
                    
                    dict_csvs_finalizados[base_name] = df_consolidado.reset_index(drop=True)

                # 3. PROCESAR CADA EXCEL E INYECTAR NRC
                excels_inyectados_zip = io.BytesIO()
                filas_para_cluster_maestro = []
                
                with zipfile.ZipFile(excels_inyectados_zip, "w", zipfile.ZIP_DEFLATED) as zip_out:
                    for fx in files_xlsx_originales:
                        base_x, _ = obtener_base_y_version(fx.name)
                        
                        if base_x in dict_csvs_finalizados:
                            df_csv_perfecto = dict_csvs_finalizados[base_x]
                            wb = openpyxl.load_workbook(io.BytesIO(fx.getvalue()))
                            
                            if HOJA_ALTAS in wb.sheetnames:
                                ws_altas = wb[HOJA_ALTAS]
                                data = list(ws_altas.values)
                                header_excel = [str(c).strip() if c is not None else "" for c in data[0]]
                                df_excel_original = pd.DataFrame(data[1:], columns=header_excel)
                                
                                df_nrc_pestana = df_excel_original.copy()
                                
                                df_excel_original["_k_per"] = df_excel_original["Periodo"].apply(_local_limpiar_texto)
                                df_excel_original["_k_sed"] = df_excel_original["Campus"].apply(_local_limpiar_texto)
                                df_excel_original["_k_sec"] = df_excel_original["Sección"].apply(_local_seccion_a_dos_digitos)
                                
                                df_csv_perfecto["_k_per"] = df_csv_perfecto["PERIODO"].apply(_local_limpiar_texto)
                                df_csv_perfecto["_k_sed"] = df_csv_perfecto["SEDE"].apply(_local_limpiar_texto)
                                df_csv_perfecto["_k_sec"] = df_csv_perfecto["SECCION"].apply(_local_seccion_a_dos_digitos)
                                
                                df_csv_mapping = df_csv_perfecto[["_k_per", "_k_sed", "_k_sec", "SUBJ", "COURSE"]].drop_duplicates(subset=["_k_per", "_k_sed", "_k_sec"])
                                
                                df_excel_original = df_excel_original.merge(df_csv_mapping, on=["_k_per", "_k_sed", "_k_sec"], how="left")
                                
                                df_nrc_pestana["Subject"] = df_excel_original["SUBJ"].combine_first(df_nrc_pestana["Subject"])
                                df_nrc_pestana["Course"] = df_excel_original["COURSE"].combine_first(df_nrc_pestana["Course"])
                                
                                periodo_clean = df_nrc_pestana["Periodo"].apply(_local_limpiar_texto)
                                nivel_clean = df_nrc_pestana["Nivel"].apply(_local_normalizar_mayusculas)
                                subject_clean = df_nrc_pestana["Subject"].apply(_local_normalizar_mayusculas)
                                course_clean = df_nrc_pestana["Course"].apply(_local_limpiar_texto)
                                seccion_clean = df_nrc_pestana["Sección"].apply(_local_seccion_a_dos_digitos)
                                
                                llaves_filas_excel = (periodo_clean + "_" + nivel_clean + "_" + 
                                                      subject_clean + "_" + course_clean + "_" + seccion_clean)
                                
                                vec_nrc = llaves_filas_excel.map(mapa_nrcs)
                                df_nrc_pestana.insert(0, "NRC", vec_nrc)
                                
                                if HOJA_SALIDA_NRC in wb.sheetnames:
                                    del wb[HOJA_SALIDA_NRC]
                                ws_nrc = wb.create_sheet(title=HOJA_SALIDA_NRC)
                                
                                ws_nrc.append(list(df_nrc_pestana.columns))
                                for fila in df_nrc_pestana.values:
                                    ws_nrc.append([None if pd.isna(v) else v for v in fila])
                                
                                excel_buffer = io.BytesIO()
                                wb.save(excel_buffer)
                                zip_out.writestr(fx.name, excel_buffer.getvalue())
                                
                                # RECOLECCIÓN PARA EL CSV DE CLÚSTER
                                for idx_row, row_ex in df_excel_original.iterrows():
                                    filas_para_cluster_maestro.append({
                                        "Periodo": row_ex.get("Periodo"),
                                        "CRN": vec_nrc.iloc[idx_row], 
                                        "datocomplementario": row_ex.get("Clúster")
                                    })
                        else:
                            st.warning(f"⚠️ El archivo Excel `{fx.name}` no encontró su CSV correspondiente.")

                    # 4. ESCRITURA DEL CSV DE CLÚSTER UNIFICADO Y ESTRICTO
                    if filas_para_cluster_maestro:
                        df_cluster_parcial = pd.DataFrame(filas_para_cluster_maestro)
                        df_cluster_final = pd.DataFrame(columns=COLUMNAS_CLUSTER_FINAL)
                        
                        # Inyectar solo las 3 variables de interés
                        df_cluster_final["Periodo"] = df_cluster_parcial["Periodo"].apply(_local_limpiar_texto)
                        df_cluster_final["CRN"] = df_cluster_parcial["CRN"] 
                        df_cluster_final["datocomplementario"] = df_cluster_parcial["datocomplementario"].apply(_local_limpiar_texto)
                        
                        # Rellenar las 21 columnas sobrantes con strings vacíos
                        df_cluster_final = df_cluster_final.fillna("")

                        csv_cluster_bytes = df_cluster_final.to_csv(**CSV_KWARGS_R).encode("utf-8")
                        zip_out.writestr("cluster_unificado.csv", csv_cluster_bytes)

                st.session_state.final_argos_zip = excels_inyectados_zip.getvalue()
                st.success("🎉 ¡Proceso finalizado! Excels actualizados y CSV de Clúster generado con estructura exacta.")
                st.rerun()

            except Exception as e:
                st.error(f"❌ Ocurrió un inconveniente: {str(e)}")

    if st.session_state.final_argos_zip is not None:
        st.markdown("### 📥 Panel de Descarga")
        st.download_button(
            label="📁 📥 DESCARGAR PAQUETE FINAL (.ZIP con Excels modificados + Clúster)",
            data=st.session_state.final_argos_zip,
            file_name="Paquete_Final_ARGOS_y_Cluster.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary"
        )
