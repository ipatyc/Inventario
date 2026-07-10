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
from openpyxl.styles import Font, PatternFill, Alignment

# ================= 1. CONFIGURACIÓN Y FUNCIONES ESTRUCTURALES =================
HOJA_ALTAS = "ALTAS"
HOJA_SALIDA_NRC = "NRC"  
UMBRAL_FUZZY = 0.82  

# Formato puro para Oracle Banner
CSV_KWARGS_R = {
    'index': False,
    'encoding': 'utf-8',
    'sep': ',',
    'lineterminator': '\n'
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

# Inicialización de estados en memoria global
if "original_files_bytes" not in st.session_state: st.session_state.original_files_bytes = {}
if "res_auditoria" not in st.session_state: st.session_state.res_auditoria = None
if "raw_altas" not in st.session_state: st.session_state.raw_altas = None
if "ready_for_download" not in st.session_state: st.session_state.ready_for_download = False
if "zip_file_bytes" not in st.session_state: st.session_state.zip_file_bytes = None
if "csv_files_to_download" not in st.session_state: st.session_state.csv_files_to_download = {}
if "delta_files" not in st.session_state: st.session_state.delta_files = {}
if "final_argos_zip" not in st.session_state: st.session_state.final_argos_zip = None

st.set_page_config(page_title="Consola Iris Cavazos", page_icon="⚙️", layout="wide")
st.title("⚙️ Consola de Control de Materias e Inyección de NRCs")
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
    with col1: file_cat = st.file_uploader("📑 Catálogo de Materias Estatales (Excel)", type=["xlsx"])
    with col2: files_altas = st.file_uploader("📁 Archivos de ALTAS (Puedes subir varios Excel)", accept_multiple_files=True, type=["xlsx"])
    
    # Función de limpieza profunda para celdas antes de exportar
    def limpiar_celda_banner(valor):
        if pd.isna(valor): return ""
        # Quita comillas, saltos de línea, retornos de carro y espacios extra
        s = str(valor).replace('"', '').replace('\n', ' ').replace('\r', '').strip()
        return s

    if files_altas and file_cat:
        if st.button("⚡ Ejecutar Validación Inteligente", type="primary"):
            st.session_state.ready_for_download = False 
            st.toast("Cargando Catálogo de Materias...", icon="📑")
            
            xls_cat = pd.ExcelFile(file_cat)
            indice_cat, indice_cat_claves = {}, {} 
            
            for hoja in xls_cat.sheet_names:
                df_c = xls_cat.parse(hoja)
                if "Nivel" in df_c.columns and "Materia" in df_c.columns:
                    for _, f in df_c.iterrows():
                        niv = normalizar_para_cruce(f.get("Nivel"))
                        mat_o = str(f.get("Materia")).strip()
                        s_val = format_r_string(f.get("Subj"))
                        c_val = format_r_string(f.get("Crse"))
                        indice_cat.setdefault(niv, []).append({
                            "mat_orig": mat_o, "mat_norm": normalizar_para_cruce(f.get("Materia")), 
                            "subj": s_val, "crse": c_val
                        })
                        if pd.notna(s_val) and pd.notna(c_val):
                            indice_cat_claves[(normalizar_para_cruce(s_val), c_val)] = mat_o
            
            piezas = []
            for f in files_altas:
                xls_a = pd.ExcelFile(f)
                hojas_reales = [h for h in xls_a.sheet_names if h.strip().upper() == HOJA_ALTAS]
                if hojas_reales:
                    df_a = xls_a.parse(hojas_reales[0], dtype=str)
                    # 🔥 LIMPIEZA DE TÍTULOS: Quita saltos de línea y espacios en nombres de columna
                    df_a.columns = [str(c).replace('\n', ' ').replace('\r', '').strip() for c in df_a.columns]
                    
                    essential_cols = [c for c in ["Periodo", "Campus", "Subject", "Course"] if c in df_a.columns]
                    if essential_cols: df_a = df_a.dropna(subset=essential_cols, how="all")
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
                        subj_sug, crse_sug, mat_cat_nombre = match_elegido["subj"], match_elegido["crse"], match_elegido["mat_orig"]
                        comentario = "Todo correcto" if subj_orig == subj_sug and crse_orig == crse_sug else "Subj/Crse incorrecto"
                    else:
                        mat_cat_nombre, comentario = mat_excel_orig, "No se encontró en catálogo"
                        subj_sug, crse_sug = subj_orig, crse_orig
                    
                    resultados.append({
                        "Luz Verde": False, "idx": idx, "Archivo": fila.get("ArchivoOrigen"), 
                        "Materia Excel": mat_excel_orig, "Materia Catálogo": mat_cat_nombre, 
                        "Comentario": comentario, "Subj Original": subj_orig, "Crse Original": crse_orig,
                        "Subj Sugerido": subj_sug, "Crse Sugerido": crse_sug,
                        "Llave_Cruce": f"{fila.get('ArchivoOrigen')}|{mat_excel_orig}|{subj_orig}|{crse_orig}"
                    })
                st.session_state.res_auditoria = pd.DataFrame(resultados)
                st.success("¡Revisión finalizada!")

    if st.session_state.res_auditoria is not None:
        st.markdown("### ⚖️ Mesa de Control Interactiva")
        for arch in st.session_state.res_auditoria["Archivo"].unique():
            df_file = st.session_state.res_auditoria[st.session_state.res_auditoria["Archivo"] == arch]
            errores_filas = df_file[df_file["Comentario"] != "Todo correcto"]
            
            if len(errores_filas) == 0:
                st.success(f"✅ **{arch}** — ¡Todo limpio!")
            else:
                with st.expander(f"⚠️ **{arch}** — ({len(errores_filas)} errores)", expanded=True):
                    if st.button("✅ Seleccionar Todo", key=f"sel_all_{arch}"):
                        st.session_state.res_auditoria.loc[st.session_state.res_auditoria["Archivo"] == arch, "Luz Verde"] = True
                        st.rerun()
                    
                    with st.form(key=f"form_{arch}"):
                        df_editado = st.data_editor(errores_filas[["Luz Verde", "Materia Excel", "Materia Catálogo", "Subj Sugerido", "Crse Sugerido"]], use_container_width=True)
                        if st.form_submit_button("💾 Confirmar Selección"):
                            # Actualización eficiente
                            df_editado["Llave_Cruce"] = arch + "|" + df_editado["Materia Excel"] + "|" + df_editado["Subj Sugerido"] + "|" + df_editado["Crse Sugerido"] # Simplificado para demo
                            st.session_state.res_auditoria.update(df_editado)
                            st.rerun()
        
        if st.button("💾 Generar Bloque de Archivos CSV", type="primary"):
            corregido = st.session_state.raw_altas.copy()
            corregido["Subject"] = corregido["Subject"].astype(str)
            corregido["Course"] = corregido["Course"].astype(str)
            for _, row in st.session_state.res_auditoria.iterrows():
                if row["Luz Verde"]:
                    corregido.loc[row["idx"], "Subject"] = str(row["Subj Sugerido"])
                    corregido.loc[row["idx"], "Course"] = str(row["Crse Sugerido"])
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for name, sub in corregido.groupby("ArchivoOrigen"):
                    res = pd.DataFrame()
                    # 🔥 APLICACIÓN DE LIMPIEZA PROFUNDA A CADA CAMPO 🔥
                    res["PERIODO"] = sub["Periodo"].apply(limpiar_celda_banner)
                    res["SEDE"] = sub["Campus"].apply(limpiar_celda_banner)
                    res["SUBJ"] = sub["Subject"].apply(limpiar_celda_banner)
                    res["COURSE"] = sub["Course"].apply(limpiar_celda_banner)
                    res["PARTEPERIODO"] = sub["Parte de Periodo"].apply(limpiar_celda_banner)
                    res["STATUS"] = sub["Estatus"].apply(limpiar_celda_banner)
                    res["CAPACIDAD"] = pd.to_numeric(sub["Capacidad"], errors='coerce').astype('Int64')
                    res["GRUPOS"] = 1
                    res["SECCION"] = pd.to_numeric(sub["Sección"], errors='coerce').astype('Int64')
                    res["TIPODEHORARIO"] = sub["Tipo de Horario"].apply(limpiar_celda_banner)
                    res["METODO_EDUCATIVO"] = sub["Método Educativo"].apply(limpiar_celda_banner)
                    res["SOCIODEINTEGRACION"] = "D2L"
                    res["MODODECALIFICAR"] = sub["Modo de Calificar"].apply(limpiar_celda_banner)
                    res["SESION"] = sub["Sesion"].apply(limpiar_celda_banner)
                    
                    csv_name = f"{name.rsplit('.', 1)[0]}.csv"
                    zip_file.writestr(csv_name, res.to_csv(**CSV_KWARGS_R).encode('utf-8'))
            
            st.session_state.zip_file_bytes = zip_buffer.getvalue()
            st.session_state.ready_for_download = True
            st.rerun()

        if st.session_state.ready_for_download:
            st.download_button("💥 📥 DESCARGAR TODOS LOS CSVs (.ZIP)", data=st.session_state.zip_file_bytes, file_name="archivos_carga_banner.zip", mime="application/zip", use_container_width=True, type="primary")
            
# ============================================================
# PESTAÑA 2: REPORTE DE ERRORES Y ENSAMBLAJE FINAL
# ============================================================
with tab_err:
    st.header("⚠️ Reporte de Errores y Ensamblaje Final")
    st.markdown("Extrae filas con error, corrígelas y genera el archivo para la Pestaña 3.")
    
    st.subheader("✂️ 1. Extraer o corregir el pedacito con errores")
    col_ex1, col_ex2, col_ex3 = st.columns(3)
    with col_ex1: file_base_ext = st.file_uploader("📁 1. Archivo Base (.csv)", type=["csv"], key="ext_base_1")
    with col_ex2: file_err_ext = st.file_uploader("📊 2. Reporte de Errores Banner (.xlsx)", type=["xlsx"], key="ext_err_1")
    with col_ex3: sufijo_version = st.text_input("🔢 Sufijo de versión (Ej: V1, V2):", value="V1", key="suf_v1")
    
    if file_base_ext and file_err_ext:
        df_base = pd.read_csv(file_base_ext, encoding="utf-8", dtype=str)
        df_err = pd.read_excel(file_err_ext, skiprows=2).dropna(subset=["Línea"])
        indices = [r - 2 for r in df_err["Línea"].astype(int).unique().tolist() if 0 <= (r - 2) < len(df_base)]
        
        if indices:
            df_delta = df_base.iloc[indices].copy()
            base_name_ext = file_base_ext.name.rsplit('.', 1)[0].replace("_base", "").replace("_final", "")
            nombre_archivo = f"{base_name_ext}_{sufijo_version}"
            
            modo_delta = st.radio("⚙️ ¿Cómo deseas descargar?", ["Excel (.xlsx)", "CSV (.csv)", "Editar en vivo"], horizontal=True, key="modo_1")
            
            if modo_delta == "Excel (.xlsx)":
                buf = io.BytesIO()
                df_delta.to_excel(buf, index=False)
                st.download_button("📥 Descargar Fragmento", data=buf.getvalue(), file_name=f"{nombre_archivo}.xlsx")
            elif modo_delta == "CSV (.csv)":
                st.download_button("📥 Descargar Fragmento", data=df_delta.to_csv(**CSV_KWARGS_R).encode("utf-8"), file_name=f"{nombre_archivo}.csv")
            else:
                df_editado = st.data_editor(df_delta, key="ed_vivo_1", use_container_width=True)
                st.download_button("📥 Descargar Corregido", data=df_editado.to_csv(**CSV_KWARGS_R).encode("utf-8"), file_name=f"{nombre_archivo}.csv", type="primary")

    st.markdown("---")
    
    st.subheader("💉 2. Inyectar correcciones y generar Archivo Final")
    col_in1, col_in2, col_in3 = st.columns(3)
    with col_in1: file_base_iny = st.file_uploader("📁 1. Archivo Base (.csv)", type=["csv"], key="iny_base_2")
    with col_in2: file_err_iny = st.file_uploader("📊 2. Reporte de Errores (.xlsx)", type=["xlsx"], key="iny_err_2")
    
    with col_in3: 
        file_corr_iny = st.file_uploader("📝 3. Fragmento Corregido", type=["csv", "xlsx"], key="iny_corr_2")
        tipo_final = st.selectbox("Etiqueta del archivo a generar:", ["final", "V1", "V2", "V3", "V4", "V5"], key="suf_v2")
    
    if file_base_iny and file_err_iny and file_corr_iny:
        if st.button("🚀 Ensamblar Archivo Final", type="primary"):
            try:
                df_base = pd.read_csv(file_base_iny, encoding="utf-8", dtype=str)
                df_err = pd.read_excel(file_err_iny, skiprows=2).dropna(subset=["Línea"])
                df_corr = pd.read_excel(file_corr_iny, dtype=str) if file_corr_iny.name.endswith('.xlsx') else pd.read_csv(file_corr_iny, encoding="utf-8", dtype=str)
                
                indices = [r - 2 for r in df_err["Línea"].astype(int).unique().tolist() if 0 <= (r - 2) < len(df_base)]
                
                if len(indices) == len(df_corr):
                    df_final = df_base.copy()
                    for col in df_final.columns:
                        if col in df_corr.columns: df_final.iloc[indices, df_final.columns.get_loc(col)] = df_corr[col].values
                    
                    for col in df_final.columns:
                        df_final[col] = df_final[col].astype(str).str.replace('"', '', regex=False).str.strip().replace(['nan', 'None', '<NA>', 'NaN'], '')
                    
                    base_name_iny = file_base_iny.name.rsplit('.', 1)[0].replace("_base", "").replace("_final", "")
                    out_name = f"{base_name_iny}_{tipo_final}.csv"
                    
                    st.success(f"🎉 ¡Archivo {out_name} listo!")
                    st.download_button(label=f"📁 📥 DESCARGAR {out_name}", data=df_final.to_csv(**CSV_KWARGS_R).encode("utf-8"), file_name=out_name, type="primary", use_container_width=True)
                else:
                    st.error(f"❌ Desajuste: {len(indices)} errores detectados vs {len(df_corr)} filas en el parche.")
            except Exception as e:
                st.error(f"Error: {str(e)}")

# ============================================================
# PESTAÑA 3: INYECCIÓN DE NRCS Y GENERACIÓN DE CLÚSTER
# ============================================================
with tab3:
    st.header("Inyección de NRCs y Generación Estricta de Clúster")
    st.markdown("Procesa todos los archivos, clona ALTAS, sobreescribe correcciones e inyecta el NRC.")
    
    col_a, col_b, col_c = st.columns(3)
    with col_a: file_argos = st.file_uploader("📊 1. Reporte ARGOS (.csv)", type=["csv"])
    with col_b: files_csv_finales = st.file_uploader("📝 2. CSVs Finales Corregidos", type=["csv"], accept_multiple_files=True)
    with col_c: files_xlsx_originales = st.file_uploader("📁 3. Excels Originales", type=["xlsx"], accept_multiple_files=True)
        
    if file_argos and files_csv_finales and files_xlsx_originales:
        if st.button("🚀 PROCESAR Y GENERAR PAQUETE FINAL", type="primary"):
            try:
                argos_df = pd.read_csv(file_argos, encoding="utf-8", on_bad_lines='skip', dtype=str)
                argos_df.columns = [re.sub(r'\.+', '.', str(c).replace('"', '').replace("'", "").strip()) for c in argos_df.columns]
                col_curso = next((c for c in argos_df.columns if "Curso" in c), None)
                if not col_curso: raise KeyError("No se encontró la columna de Curso en ARGOS.")

                argos_df["Periodo"] = argos_df["Periodo"].apply(limpiar_clave_texto)
                argos_df["Nivel"] = argos_df["Nivel"].apply(normalizar_para_cruce)
                argos_df["Área"] = argos_df["Área"].apply(normalizar_para_cruce)
                argos_df[col_curso] = argos_df[col_curso].apply(limpiar_clave_texto)
                argos_df["Grupo"] = argos_df["Grupo"].apply(limpia_seccion_interna)
                
                argos_df["_llave_argos"] = (argos_df["Periodo"] + "_" + argos_df["Nivel"] + "_" + 
                                            argos_df["Área"] + "_" + argos_df[col_curso] + "_" + argos_df["Grupo"])
                argos_df = argos_df.drop_duplicates(subset=["_llave_argos"])
                mapa_nrcs = dict(zip(argos_df["_llave_argos"], argos_df["NRC"]))

                def simplificar_nombre(nombre):
                    n = nombre.lower()
                    for basura in ['.xlsx', '.xls', '.csv', '_final', '_base', '_v1', '_v2', '_v3', '_v4', 'corregidas_', 'errores_']:
                        n = n.replace(basura, '')
                    return n.strip().replace(" ", "")

                excels_inyectados_zip = io.BytesIO()
                filas_para_cluster_maestro = []
                archivos_procesados_con_exito = 0
                alertas_dimensiones, alertas_parejas = [], []
                
                with zipfile.ZipFile(excels_inyectados_zip, "w", zipfile.ZIP_DEFLATED) as zip_out:
                    for fx in files_xlsx_originales:
                        df_csv, fc_usado = None, None
                        base_excel = simplificar_nombre(fx.name)
                        
                        for fc_cand in files_csv_finales:
                            base_csv = simplificar_nombre(fc_cand.name)
                            if base_excel == base_csv or base_excel in base_csv or base_csv in base_excel:
                                df_csv = pd.read_csv(io.BytesIO(fc_cand.getvalue()), encoding="utf-8", dtype=str)
                                fc_usado = fc_cand
                                break
                        
                        if df_csv is not None:
                            wb = openpyxl.load_workbook(io.BytesIO(fx.getvalue()))
                            if HOJA_ALTAS in wb.sheetnames:
                                data = list(wb[HOJA_ALTAS].values)
                                if not data: continue
                                
                                df_excel_original = pd.DataFrame(data[1:], columns=[str(c).strip() if c is not None else "" for c in data[0]])
                                df_excel_original = df_excel_original.dropna(how='all')
                                df_csv = df_csv.dropna(how='all')
                                
                                if "Periodo" in df_excel_original.columns:
                                    df_excel_original = df_excel_original[df_excel_original["Periodo"].astype(str).str.strip() != ""]
                                if "PERIODO" in df_csv.columns:
                                    df_csv = df_csv[df_csv["PERIODO"].astype(str).str.strip() != ""]
                                
                                df_excel_original, df_csv = df_excel_original.reset_index(drop=True), df_csv.reset_index(drop=True)
                                
                                if len(df_excel_original) != len(df_csv):
                                    alertas_dimensiones.append(f"❌ Excel `{fx.name}` tiene **{len(df_excel_original)} filas**, CSV `{fc_usado.name}` tiene **{len(df_csv)} filas**.")
                                    continue
                                
                                df_nrc_pestana = df_excel_original.copy()
                                mapeo_columnas = {
                                    "Periodo": "PERIODO", "Campus": "SEDE", "Subject": "SUBJ", "Course": "COURSE",
                                    "Parte de Periodo": "PARTEPERIODO", "Estatus": "STATUS", "Capacidad": "CAPACIDAD",
                                    "Sección": "SECCION", "Tipo de Horario": "TIPODEHORARIO", "Método Educativo": "METODO_EDUCATIVO",
                                    "Modo de Calificar": "MODODECALIFICAR", "Sesion": "SESION"
                                }
                                
                                for col_ex, col_cs in mapeo_columnas.items():
                                    if col_ex in df_nrc_pestana.columns and col_cs in df_csv.columns:
                                        if col_ex == "Sección": df_nrc_pestana[col_ex] = pd.to_numeric(df_csv[col_cs], errors='coerce').values
                                        else: df_nrc_pestana[col_ex] = df_csv[col_cs].values
                                
                                df_nrc_pestana["Grupos"], df_nrc_pestana["Socio de Integración"] = "1", "D2L"
                                
                                llaves_cruce = (
                                    df_nrc_pestana["Periodo"].apply(limpiar_clave_texto) + "_" + 
                                    df_excel_original["Nivel"].apply(normalizar_para_cruce) + "_" + 
                                    df_nrc_pestana["Subject"].apply(normalizar_para_cruce) + "_" + 
                                    df_nrc_pestana["Course"].apply(limpiar_clave_texto) + "_" + 
                                    df_nrc_pestana["Sección"].apply(str).apply(limpia_seccion_interna)
                                )
                                df_nrc_pestana.insert(0, "NRC", llaves_cruce.map(mapa_nrcs))
                                
                                if HOJA_SALIDA_NRC in wb.sheetnames: del wb[HOJA_SALIDA_NRC]
                                ws_nrc = wb.create_sheet(title=HOJA_SALIDA_NRC)
                                ws_nrc.append(list(df_nrc_pestana.columns))
                                for fila in df_nrc_pestana.values: ws_nrc.append([None if pd.isna(v) else v for v in fila])
                                
                                font_base, font_nrc, font_header = Font(name="Calibri", size=11), Font(name="Calibri", size=11, bold=True), Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                                fill_header, fill_nrc = PatternFill(start_color="1F4E78", fill_type="solid"), PatternFill(start_color="DDEBF7", fill_type="solid")
                                align_header, align_center = Alignment(horizontal="center", vertical="center", wrap_text=True), Alignment(horizontal="center", vertical="center")
                                
                                for row in ws_nrc.iter_rows(min_row=1, max_row=ws_nrc.max_row, min_col=1, max_col=ws_nrc.max_column):
                                    for cell in row: cell.font = font_base
                                for cell in ws_nrc[1]: cell.font, cell.fill, cell.alignment = font_header, fill_header, align_header
                                for cell in ws_nrc['A'][1:]: cell.font, cell.fill, cell.alignment = font_nrc, fill_nrc, align_center
                                
                                for col in ws_nrc.columns:
                                    max_len = 0
                                    for cell in col:
                                        if cell.value: max_len = max(max_len, len(str(cell.value)))
                                    ws_nrc.column_dimensions[col[0].column_letter].width = max(max_len + 3, 11)
                                
                                nombre_salida_excel = fc_usado.name.rsplit('.', 1)[0] + "_con_NRC.xlsx"
                                excel_buffer = io.BytesIO()
                                wb.save(excel_buffer)
                                zip_out.writestr(nombre_salida_excel, excel_buffer.getvalue())
                                archivos_procesados_con_exito += 1
                                
                                for idx_row, row_ex in df_excel_original.iterrows():
                                    filas_para_cluster_maestro.append({
                                        "Periodo": row_ex.get("Periodo"), "CRN": df_nrc_pestana.iloc[idx_row, 0], 
                                        "datocomplementario": row_ex.get("Clúster")
                                    })
                        else:
                            alertas_parejas.append(f"⚠️ `{fx.name}` no encontró ningún CSV compatible.")

                    if filas_para_cluster_maestro:
                        df_parcial = pd.DataFrame(filas_para_cluster_maestro)
                        df_cluster_final = pd.DataFrame(index=df_parcial.index, columns=COLUMNAS_CLUSTER_FINAL)
                        df_cluster_final["Periodo"] = df_parcial["Periodo"].apply(limpiar_clave_texto)
                        df_cluster_final["CRN"] = df_parcial["CRN"] 
                        df_cluster_final["datocomplementario"] = df_parcial["datocomplementario"].apply(limpiar_clave_texto)
                        
                        for col in df_cluster_final.columns:
                            df_cluster_final[col] = df_cluster_final[col].astype(str).str.replace('"', '', regex=False).str.strip().replace(['nan', 'None', '<NA>', 'NaN'], '')
                        
                        zip_out.writestr("cluster_unificado.csv", df_cluster_final.to_csv(**CSV_KWARGS_R).encode("utf-8"))

                if archivos_procesados_con_exito > 0:
                    st.session_state.final_argos_zip = excels_inyectados_zip.getvalue()
                    st.success(f"🎉 ¡Paquete final generado! Se procesaron {archivos_procesados_con_exito} archivos exitosamente.")
                else: st.error("❌ No se pudo procesar ningún archivo.")
                if alertas_dimensiones:
                    st.markdown("### 🚫 Archivos descartados por diferencia de filas:")
                    for alerta in alertas_dimensiones: st.error(alerta)
                if alertas_parejas:
                    st.markdown("### ❓ Archivos sin pareja:")
                    for alerta in alertas_parejas: st.warning(alerta)

            except Exception as e:
                st.error(f"❌ Ocurrió un inconveniente crítico: {str(e)}")

    if st.session_state.final_argos_zip is not None:
        st.markdown("---")
        st.markdown("### 📥 Panel de Descarga")
        st.download_button(
            label="📁 📥 DESCARGAR PAQUETE FINAL (.ZIP)", data=st.session_state.final_argos_zip,
            file_name="Paquete_Final_ARGOS_y_Cluster.zip", mime="application/zip",
            use_container_width=True, type="primary"
        )
