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
HOJA_SALIDA_NRC = "NRC"  # Pestaña en mayúsculas estrictas
UMBRAL_FUZZY = 0.82  

# 🔥 ÚNICO CAMBIO EN CONFIGURACIÓN: Formato puro para Oracle Banner
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

st.set_page_config(page_title="Consola Iris Cavazos", page_icon="⚙️", layout="wide")
st.title("⚙️ Consola de Control de Materias e Inyección de NRCs (Flujo Multi-CSV)")
st.markdown("---")

tab1, tab_err, tab3 = st.tabs([
    "1️⃣ Proceso: Validación y Generar CSVs", 
    "⚠️ Reporte de Errores (Extraer Delta)", 
    "2️⃣ Proceso: Inyección de NRCs Masiva (ARGOS)"
])

# ============================================================
# PESTAÑA 2: REPORTE DE ERRORES Y ENSAMBLAJE FINAL
# ============================================================
with tab_err:
    st.header("⚠️ Reporte de Errores y Ensamblaje Final")
    st.markdown("Extrae filas con error, corrígelas y genera el archivo para la Pestaña 3.")
    
    # --- PASO 1: EXTRAER O EDITAR EL PEDACITO CON ERROR ---
    st.subheader("✂️ 1. Extraer o corregir el pedacito con errores")
    
    col_ex1, col_ex2, col_ex3 = st.columns(3)
    with col_ex1: file_base_ext = st.file_uploader("📁 1. Archivo Base (.csv)", type=["csv"], key="ex_base")
    with col_ex2: file_err_ext = st.file_uploader("📊 2. Reporte de Errores Banner (.xlsx)", type=["xlsx"], key="ex_err")
    with col_ex3: 
        sufijo_version = st.text_input("🔢 Sufijo de versión (Ej: V1, V2, etc.):", value="V1", key="suf_v1")
    
    if file_base_ext and file_err_ext:
        df_base = pd.read_csv(file_base_ext, encoding="utf-8", dtype=str)
        df_err = pd.read_excel(file_err_ext, skiprows=2)
        
        df_err = df_err.dropna(subset=["Línea"])
        renglones = df_err["Línea"].astype(int).unique().tolist()
        indices = [r - 2 for r in renglones if 0 <= (r - 2) < len(df_base)]
        
        if indices:
            df_delta = df_base.iloc[indices].copy()
            # Higiene estricta
            for col in df_delta.columns:
                df_delta[col] = df_delta[col].astype(str).str.replace('"', '', regex=False).str.strip().replace(['nan', 'None', '<NA>', 'NaN'], '')
            
            # Limpiamos el nombre base para no arrastrar basura
            base_name_clean = file_base_ext.name.rsplit('.', 1)[0].replace("_base", "").replace("_final", "").replace("_V1", "").replace("_V2", "")
            
            modo_delta = st.radio("⚙️ ¿Cómo deseas descargar?", ["Excel (.xlsx)", "CSV (.csv)", "Editar en vivo"], horizontal=True)
            
            # Nombre resultante: solo nombre base + sufijo
            nombre_archivo = f"{base_name_clean}_{sufijo_version}"
            
            if modo_delta == "Excel (.xlsx)":
                excel_buffer = io.BytesIO()
                df_delta.to_excel(excel_buffer, index=False)
                st.download_button("📥 Descargar Fragmento", data=excel_buffer.getvalue(), file_name=f"{nombre_archivo}.xlsx")
            elif modo_delta == "CSV (.csv)":
                st.download_button("📥 Descargar Fragmento", data=df_delta.to_csv(**CSV_KWARGS_R).encode("utf-8"), file_name=f"{nombre_archivo}.csv")
            else:
                df_editado = st.data_editor(df_delta, key="ed_vivo", use_container_width=True)
                df_editado_clean = df_editado.copy()
                for col in df_editado_clean.columns:
                    df_editado_clean[col] = df_editado_clean[col].astype(str).str.replace('"', '', regex=False).str.strip().replace(['nan', 'None', '<NA>', 'NaN'], '')
                st.download_button("📥 Descargar Corregido", data=df_editado_clean.to_csv(**CSV_KWARGS_R).encode("utf-8"), file_name=f"{nombre_archivo}.csv", type="primary")

    st.markdown("---")
    
    # --- PASO 2: INYECTAR Y CREAR EL ARCHIVO FINAL ---
    st.subheader("💉 2. Inyectar correcciones y generar Archivo Final")
    
    col_in1, col_in2, col_in3 = st.columns(3)
    with col_in1: file_base_iny = st.file_uploader("📁 1. Archivo Base (.csv)", type=["csv"], key="in_base_2")
    with col_in2: file_err_iny = st.file_uploader("📊 2. Reporte de Errores (.xlsx)", type=["xlsx"], key="in_err_2")
    with col_in3: 
        file_corr_iny = st.file_uploader("📝 3. Fragmento Corregido", type=["csv", "xlsx"], key="in_corr_2")
        tipo_final = st.text_input("Etiqueta final (V1, V2, o final):", value="final", key="suf_v2")
    
    if file_base_iny and file_err_iny and file_corr_iny:
        if st.button("🚀 Ensamblar Archivo Final", type="primary"):
            try:
                df_base = pd.read_csv(file_base_iny, encoding="utf-8", dtype=str)
                df_err = pd.read_excel(file_err_iny, skiprows=2)
                if file_corr_iny.name.lower().endswith('.xlsx'): df_corr = pd.read_excel(file_corr_iny, dtype=str)
                else: df_corr = pd.read_csv(file_corr_iny, encoding="utf-8", dtype=str)
                
                df_err = df_err.dropna(subset=["Línea"])
                renglones = df_err["Línea"].astype(int).unique().tolist()
                indices = [r - 2 for r in renglones if 0 <= (r - 2) < len(df_base)]
                
                if len(indices) == len(df_corr):
                    df_final = df_base.copy()
                    for col in df_corr.columns:
                        df_corr[col] = df_corr[col].astype(str).str.replace('"', '', regex=False).str.strip().replace(['nan', 'None', '<NA>', 'NaN'], '')
                    
                    for col in df_final.columns:
                        if col in df_corr.columns:
                            df_final.iloc[indices, df_final.columns.get_loc(col)] = df_corr[col].values
                    
                    # Limpiamos el nombre base para ensamblar el final
                    base_name_iny = file_base_iny.name.rsplit('.', 1)[0].replace("_base", "").replace("_final", "").replace("_V1", "").replace("_V2", "")
                    out_name = f"{base_name_iny}_{tipo_final}.csv"
                    
                    st.success(f"🎉 ¡Archivo {out_name} listo!")
                    st.download_button(label=f"📁 📥 DESCARGAR {out_name}", data=df_final.to_csv(**CSV_KWARGS_R).encode("utf-8"), file_name=out_name, type="primary", use_container_width=True)
                else:
                    st.error(f"❌ Desajuste: {len(indices)} errores detectados vs {len(df_corr)} filas en el parche.")
            except Exception as e:
                st.error(f"Error: {str(e)}")
            
# ============================================================
# PESTAÑA 2: REPORTE DE ERRORES Y ENSAMBLAJE FINAL
# ============================================================
with tab_err:
    st.header("⚠️ Reporte de Errores y Ensamblaje Final")
    st.markdown("Extrae filas con error, corrígelas y genera el archivo **_final** para la Pestaña 3.")
    
    # --- PASO 1: EXTRAER O EDITAR EL PEDACITO CON ERROR ---
    st.subheader("✂️ 1. Extraer o corregir el pedacito con errores")
    
    col_ex1, col_ex2, col_ex3 = st.columns(3)
    with col_ex1: file_base_ext = st.file_uploader("📁 1. Archivo Base (.csv)", type=["csv"], key="ex_base")
    with col_ex2: file_err_ext = st.file_uploader("📊 2. Reporte de Errores Banner (.xlsx)", type=["xlsx"], key="ex_err")
    with col_ex3: num_v_ext = st.number_input("🔢 Versión de corrección (Ej. 1 para V1):", min_value=1, value=1, key="v_ext")
    
    if file_base_ext and file_err_ext:
        # Leer como texto puro para no alterar tus datos
        df_base = pd.read_csv(file_base_ext, encoding="utf-8", dtype=str)
        df_err = pd.read_excel(file_err_ext, skiprows=2)
        
        # Filtramos exactamente las líneas que dice el Excel
        df_err = df_err.dropna(subset=["Línea"])
        renglones = df_err["Línea"].astype(int).unique().tolist()
        indices = [r - 2 for r in renglones if 0 <= (r - 2) < len(df_base)]
        
        if indices:
            # Copiamos la base y nos quedamos SOLO con los renglones del error
            df_delta = df_base.iloc[indices].copy()
            
            # 🔥 HIGIENE ESTRICTA BANNER 🔥
            for col in df_delta.columns:
                df_delta[col] = df_delta[col].astype(str).str.replace('"', '', regex=False).str.strip().replace(['nan', 'None', '<NA>', 'NaN'], '')
            
            base_name_ext = file_base_ext.name.rsplit('.', 1)[0]
            
            modo_delta = st.radio(
                "⚙️ ¿Cómo deseas corregir las filas con error?", 
                ["Descargar en Excel (.xlsx)", "Descargar en formato CSV (.csv)", "Editar en vivo en la consola"],
                horizontal=True
            )
            
            if modo_delta == "Descargar en Excel (.xlsx)":
                excel_buffer = io.BytesIO()
                df_delta.to_excel(excel_buffer, index=False)
                st.download_button(
                    "📥 Descargar Fragmento en Excel", 
                    data=excel_buffer.getvalue(), 
                    file_name=f"Errores_{base_name_ext}_V{num_v_ext}.xlsx", 
                    type="secondary"
                )
            elif modo_delta == "Descargar en formato CSV (.csv)":
                st.download_button(
                    "📥 Descargar Fragmento en CSV", 
                    data=df_delta.to_csv(**CSV_KWARGS_R).encode("utf-8"), 
                    file_name=f"Errores_{base_name_ext}_V{num_v_ext}.csv", 
                    type="secondary"
                )
            else:
                st.caption("Edita los datos directamente en la tabla y descarga el pedacito ya corregido.")
                df_editado = st.data_editor(df_delta, key="ed_vivo", use_container_width=True)
                
                # 🔥 HIGIENE ESTRICTA BANNER 🔥
                df_editado_clean = df_editado.copy()
                for col in df_editado_clean.columns:
                    df_editado_clean[col] = df_editado_clean[col].astype(str).str.replace('"', '', regex=False).str.strip().replace(['nan', 'None', '<NA>', 'NaN'], '')
                
                st.download_button(
                    "📥 Descargar Fragmento Corregido (.csv)", 
                    data=df_editado_clean.to_csv(**CSV_KWARGS_R).encode("utf-8"), 
                    file_name=f"Corregidas_{base_name_ext}_V{num_v_ext}.csv", 
                    type="primary"
                )
        else:
            st.warning("No se encontraron coincidencias de filas.")

    st.markdown("---")
    
    # --- PASO 2: INYECTAR Y CREAR EL ARCHIVO FINAL ---
    st.subheader("💉 2. Inyectar correcciones y generar Archivo Final")
    st.markdown("Sube los 3 archivos para ensamblar el CSV **_final** limpio.")
    
    col_in1, col_in2, col_in3 = st.columns(3)
    with col_in1: file_base_iny = st.file_uploader("📁 1. Archivo Base (.csv)", type=["csv"], key="in_base_2")
    with col_in2: file_err_iny = st.file_uploader("📊 2. Reporte de Errores (.xlsx)", type=["xlsx"], key="in_err_2")
    with col_in3: file_corr_iny = st.file_uploader("📝 3. Fragmento Corregido", type=["csv", "xlsx"], key="in_corr_2")
    
    if file_base_iny and file_err_iny and file_corr_iny:
        if st.button("🚀 Ensamblar Archivo Final", type="primary"):
            try:
                df_base = pd.read_csv(file_base_iny, encoding="utf-8", dtype=str)
                df_err = pd.read_excel(file_err_iny, skiprows=2)
                
                # Leemos tu parche (ya sea excel o csv)
                if file_corr_iny.name.lower().endswith('.xlsx'):
                    df_corr = pd.read_excel(file_corr_iny, dtype=str)
                else:
                    df_corr = pd.read_csv(file_corr_iny, encoding="utf-8", dtype=str)
                
                df_err = df_err.dropna(subset=["Línea"])
                renglones = df_err["Línea"].astype(int).unique().tolist()
                indices = [r - 2 for r in renglones if 0 <= (r - 2) < len(df_base)]
                
                if len(indices) == len(df_corr):
                    df_final = df_base.copy()
                    
                    # 🔥 HIGIENE ESTRICTA BANNER 🔥
                    for col in df_corr.columns:
                        df_corr[col] = df_corr[col].astype(str).str.replace('"', '', regex=False).str.strip().replace(['nan', 'None', '<NA>', 'NaN'], '')
                    
                    # Inyección exacta de los datos
                    for col in df_final.columns:
                        if col in df_corr.columns:
                            df_final.iloc[indices, df_final.columns.get_loc(col)] = df_corr[col].values
                    
                    # 🔥 HIGIENE ESTRICTA BANNER 🔥
                    for col in df_final.columns:
                        df_final[col] = df_final[col].astype(str).str.replace('"', '', regex=False).str.strip().replace(['nan', 'None', '<NA>', 'NaN'], '')
                    
                    base_name_iny = file_base_iny.name.rsplit('.', 1)[0]
                    out_name = f"{base_name_iny.replace('_base', '').replace('_final', '')}_final.csv"
                    
                    st.success("🎉 ¡Archivo Final listo y limpio!")
                    st.download_button(
                        label=f"📁 📥 DESCARGAR {out_name}", 
                        data=df_final.to_csv(**CSV_KWARGS_R).encode("utf-8"), 
                        file_name=out_name, 
                        type="primary",
                        use_container_width=True
                    )
                else:
                    st.error(f"❌ Desajuste: Tienes {len(indices)} errores, pero el archivo corregido tiene {len(df_corr)} filas.")
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
                # 1. PREPARAR Y NORMALIZAR REPORTE ARGOS
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

                # 2. FUNCIÓN DE LIMPIEZA PARA EMPAREJAMIENTO
                def simplificar_nombre(nombre):
                    n = nombre.lower()
                    for basura in ['.xlsx', '.xls', '.csv', '_final', '_base', '_v1', '_v2', '_v3', '_v4', 'corregidas_', 'errores_']:
                        n = n.replace(basura, '')
                    return n.strip().replace(" ", "")

                # 3. PROCESAMIENTO MATRICIAL
                excels_inyectados_zip = io.BytesIO()
                filas_para_cluster_maestro = []
                archivos_procesados_con_exito = 0
                
                alertas_dimensiones = []
                alertas_parejas = []
                
                with zipfile.ZipFile(excels_inyectados_zip, "w", zipfile.ZIP_DEFLATED) as zip_out:
                    for fx in files_xlsx_originales:
                        df_csv = None
                        fc_usado = None
                        
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
                                ws_altas = wb[HOJA_ALTAS]
                                data = list(ws_altas.values)
                                if not data: continue
                                
                                df_excel_original = pd.DataFrame(data[1:], columns=[str(c).strip() if c is not None else "" for c in data[0]])
                                
                                # Limpieza de filas fantasmas
                                df_excel_original = df_excel_original.dropna(how='all')
                                df_csv = df_csv.dropna(how='all')
                                
                                if "Periodo" in df_excel_original.columns:
                                    df_excel_original = df_excel_original[df_excel_original["Periodo"].astype(str).str.strip() != ""]
                                if "PERIODO" in df_csv.columns:
                                    df_csv = df_csv[df_csv["PERIODO"].astype(str).str.strip() != ""]
                                
                                df_excel_original = df_excel_original.reset_index(drop=True)
                                df_csv = df_csv.reset_index(drop=True)
                                
                                if len(df_excel_original) != len(df_csv):
                                    alertas_dimensiones.append(f"❌ Excel `{fx.name}` tiene **{len(df_excel_original)} filas de datos**, pero el CSV `{fc_usado.name}` tiene **{len(df_csv)} filas de datos**.")
                                    continue
                                
                                # Clonamos la pestaña de Excel original (con todas sus columnas)
                                df_nrc_pestana = df_excel_original.copy()
                                
                                # Diccionario de equivalencias para reescribir solo lo que arreglamos
                                mapeo_columnas = {
                                    "Periodo": "PERIODO", "Campus": "SEDE", "Subject": "SUBJ", "Course": "COURSE",
                                    "Parte de Periodo": "PARTEPERIODO", "Estatus": "STATUS", "Capacidad": "CAPACIDAD",
                                    "Sección": "SECCION", "Tipo de Horario": "TIPODEHORARIO", "Método Educativo": "METODO_EDUCATIVO",
                                    "Modo de Calificar": "MODODECALIFICAR", "Sesion": "SESION"
                                }
                                
                                # Inyección quirúrgica
                                for col_ex, col_cs in mapeo_columnas.items():
                                    if col_ex in df_nrc_pestana.columns and col_cs in df_csv.columns:
                                        if col_ex == "Sección":
                                            df_nrc_pestana[col_ex] = pd.to_numeric(df_csv[col_cs], errors='coerce').values
                                        else:
                                            df_nrc_pestana[col_ex] = df_csv[col_cs].values
                                
                                df_nrc_pestana["Grupos"] = "1"
                                df_nrc_pestana["Socio de Integración"] = "D2L"
                                
                                # Calculamos y traemos el NRC de ARGOS
                                llaves_cruce = (
                                    df_nrc_pestana["Periodo"].apply(limpiar_clave_texto) + "_" + 
                                    df_excel_original["Nivel"].apply(normalizar_para_cruce) + "_" + 
                                    df_nrc_pestana["Subject"].apply(normalizar_para_cruce) + "_" + 
                                    df_nrc_pestana["Course"].apply(limpiar_clave_texto) + "_" + 
                                    df_nrc_pestana["Sección"].apply(str).apply(limpia_seccion_interna)
                                )
                                
                                # Insertamos el NRC en la posición 0 (columna A)
                                df_nrc_pestana.insert(0, "NRC", llaves_cruce.map(mapa_nrcs))
                                
                                # Guardamos la pestaña ensamblada en el Excel
                                if HOJA_SALIDA_NRC in wb.sheetnames: del wb[HOJA_SALIDA_NRC]
                                ws_nrc = wb.create_sheet(title=HOJA_SALIDA_NRC)
                                ws_nrc.append(list(df_nrc_pestana.columns))
                                for fila in df_nrc_pestana.values: 
                                    ws_nrc.append([None if pd.isna(v) else v for v in fila])
                                
                                # ===================================================
                                # ✨ DISEÑO DE FORMATO PROFESIONAL (ESTRICTO) ✨
                                # ===================================================
                                font_base = Font(name="Calibri", size=11, bold=False)
                                font_nrc = Font(name="Calibri", size=11, bold=True)
                                font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                                
                                fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Azul Rey corporativo
                                fill_nrc = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")     # Azul Claro sutil
                                
                                align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)
                                align_center = Alignment(horizontal="center", vertical="center")
                                
                                for row in ws_nrc.iter_rows(min_row=1, max_row=ws_nrc.max_row, min_col=1, max_col=ws_nrc.max_column):
                                    for cell in row:
                                        cell.font = font_base
                                
                                for cell in ws_nrc[1]:
                                    cell.font = font_header
                                    cell.fill = fill_header
                                    cell.alignment = align_header
                                
                                for cell in ws_nrc['A'][1:]:
                                    cell.font = font_nrc
                                    cell.fill = fill_nrc
                                    cell.alignment = align_center
                                
                                for col in ws_nrc.columns:
                                    max_len = 0
                                    col_letter = col[0].column_letter
                                    for cell in col:
                                        if cell.value is not None:
                                            max_len = max(max_len, len(str(cell.value)))
                                    ws_nrc.column_dimensions[col_letter].width = max(max_len + 3, 11)
                                # ===================================================
                                
                                nombre_salida_excel = fc_usado.name.rsplit('.', 1)[0] + "_con_NRC.xlsx"
                                
                                excel_buffer = io.BytesIO()
                                wb.save(excel_buffer)
                                zip_out.writestr(nombre_salida_excel, excel_buffer.getvalue())
                                archivos_procesados_con_exito += 1
                                
                                for idx_row, row_ex in df_excel_original.iterrows():
                                    filas_para_cluster_maestro.append({
                                        "Periodo": row_ex.get("Periodo"),
                                        "CRN": df_nrc_pestana.iloc[idx_row, 0], 
                                        "datocomplementario": row_ex.get("Clúster")
                                    })
                        else:
                            alertas_parejas.append(f"⚠️ `{fx.name}` no encontró ningún CSV compatible.")

                    # 4. CREACIÓN DEL REPORTE DE CLÚSTER UNIFICADO
                    if filas_para_cluster_maestro:
                        df_parcial = pd.DataFrame(filas_para_cluster_maestro)
                        df_cluster_final = pd.DataFrame(index=df_parcial.index, columns=COLUMNAS_CLUSTER_FINAL)
                        
                        df_cluster_final["Periodo"] = df_parcial["Periodo"].apply(limpiar_clave_texto)
                        df_cluster_final["CRN"] = df_parcial["CRN"] 
                        df_cluster_final["datocomplementario"] = df_parcial["datocomplementario"].apply(limpiar_clave_texto)
                        
                        # 🔥 HIGIENE ESTRICTA BANNER 🔥
                        for col in df_cluster_final.columns:
                            df_cluster_final[col] = df_cluster_final[col].astype(str).str.replace('"', '', regex=False).str.strip().replace(['nan', 'None', '<NA>', 'NaN'], '')
                        
                        csv_cluster_bytes = df_cluster_final.to_csv(**CSV_KWARGS_R).encode("utf-8")
                        zip_out.writestr("cluster_unificado.csv", csv_cluster_bytes)

                # --- IMPRESIÓN DE RESULTADOS ---
                if archivos_procesados_con_exito > 0:
                    st.session_state.final_argos_zip = excels_inyectados_zip.getvalue()
                    st.success(f"🎉 ¡Paquete final generado! Se procesaron {archivos_procesados_con_exito} archivos exitosamente.")
                else:
                    st.error("❌ No se pudo procesar ningún archivo.")
                
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
            label="📁 📥 DESCARGAR PAQUETE FINAL (.ZIP)",
            data=st.session_state.final_argos_zip,
            file_name="Paquete_Final_ARGOS_y_Cluster.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary"
        )
