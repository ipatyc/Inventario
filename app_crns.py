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

tab1, tab2 = st.tabs(["1️⃣ Proceso: Validación y Generar CSV", "2️⃣ Proceso: Inyección de NRCs (ARGOS)"])

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
                    
                    # 🚨 FILTRO ANTI-FANTASMAS (Evita el problema de "Muchos" renglones vacíos al estilo R)
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
                    quitar_rep = st.checkbox("🔍 Combinar repetidas (Ver solo
