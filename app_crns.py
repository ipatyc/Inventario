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
            st.session_state.ready_for_download = False  # Resetear descargas anteriores
            st.toast("Cargando Catálogo de Materias...", icon="📑")
            
            xls_cat = pd.ExcelFile(file_cat)
            indice_cat = {}
            for hoja in xls_cat.sheet_names:
                df_c = xls_cat.parse(hoja)
                if "Nivel" in df_c.columns and "Materia" in df_c.columns:
                    for _, f in df_c.iterrows():
                        niv = normalizar_para_cruce(f.get("Nivel"))
                        indice_cat.setdefault(niv, []).append({
                            "mat_orig": str(f.get("Materia")).strip(),
                            "mat_norm": normalizar_para_cruce(f.get("Materia")), 
                            "subj": str(f.get("Subj")).strip(), 
                            "crse": str(f.get("Crse")).strip()
                        })
            
            piezas = []
            for f in files_altas:
                primera_palabra = f.name.split()[0]
                st.info(f"🔍 Checking / Revisando archivo: **{primera_palabra}**")
                st.session_state.original_files_bytes[f.name] = f.getvalue()
                
                xls_a = pd.ExcelFile(f)
                hojas_reales = [h for h in xls_a.sheet_names if h.strip().upper() == HOJA_ALTAS]
                if hojas_reales:
                    df_a = xls_a.parse(hojas_reales[0])
                    df_a["ArchivoOrigen"] = f.name
                    piezas.append(df_a)
            
            if piezas:
                df_total = pd.concat(piezas, ignore_index=True)
                st.session_state.raw_altas = df_total.copy()
                
                resultados = []
                for idx, fila in df_total.iterrows():
                    niv_n = normalizar_para_cruce(fila.get("Nivel"))
                    mat_n = normalizar_para_cruce(fila.get("Nombre de la Materia"))
                    subj_orig = str(fila.get("Subject")).strip()
                    crse_orig = str(fila.get("Course")).strip()
                    
                    candidatos = indice_cat.get(niv_n, [])
                    matches_exactos = [c for c in candidatos if c["mat_norm"] == mat_n]
                    
                    match_elegido = None
                    tipo = "no_encontrado"
                    
                    if matches_exactos:
                        tipo = "exacto"
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
                            tipo = "fuzzy"
                            match_elegido = coincidencia_perf_f if coincidencia_perf_f else mejor
                    
                    subj_sug = match_elegido["subj"] if match_elegido else None
                    crse_sug = match_elegido["crse"] if match_elegido else None
                    mat_cat_nombre = match_elegido["mat_orig"] if match_elegido else "❌ No encontrada en catálogo"
                    
                    if tipo == "no_encontrado":
                        comentario = "No se encontró en catálogo"
                    elif subj_orig == subj_sug and crse_orig == crse_sug:
                        comentario = "Todo correcto"
                    elif subj_orig != subj_sug and crse_orig != crse_sug:
                        comentario = "Subj y Crse incorrectos"
                    elif subj_orig != subj_sug:
                        comentario = "Subject incorrecto"
                    else:
                        comentario = "Course incorrecto"
                    
                    resultados.append({
                        "Luz Verde": False,  # IRIS: Todo desmarcado por defecto
                        "idx": idx, 
                        "Archivo": fila.get("ArchivoOrigen"), 
                        "Materia Excel": fila.get("Nombre de la Materia"), 
                        "Materia Catálogo": mat_cat_nombre, # IRIS: Columna agregada para comparar
                        "Comentario": comentario,
                        "Subj Original": fila.get("Subject"), 
                        "Crse Original": fila.get("Course"),
                        "Subj Sugerido": subj_sug, 
                        "Crse Sugerido": crse_sug
                    })
                
                st.session_state.res_auditoria = pd.DataFrame(resultados)
                st.success("¡Revisión terminada con éxito!")
            else:
                st.error(f"❌ Ninguno de los archivos subidos tiene la pestaña '{HOJA_ALTAS}'")

    if st.session_state.res_auditoria is not None:
        st.markdown("### ⚖️ Mesa de Control Interactiva por Excel")
        
        df_aud = st.session_state.res_auditoria
        archivos_subidos = df_aud["Archivo"].unique()
        
        for arch in archivos_subidos:
            df_file = df_aud[df_aud["Archivo"] == arch]
            errores_filas = df_file[df_file["Comentario"] != "Todo correcto"]
            total_detalles = len(errores_filas)
            
            if total_detalles == 0:
                st.success(f"✅ **{arch}** — ¡Todo perfecto, sin errores estructurales!")
            else:
                with st.expander(f"⚠️ **{arch}** — Observaciones pendientes encontradas", expanded=True):
                    quitar_rep = st.checkbox("🔍 Combinar repetidas (Ver solo 1 renglón por caso)", value=True, key=f"rep_{arch}")
                    
                    if quitar_rep:
                        df_vista = errores_filas.drop_duplicates(subset=["Materia Excel", "Materia Catálogo", "Subj Original", "Crse Original", "Comentario"])
                    else:
                        df_vista = errores_filas
                        
                    columnas_vista = ["Luz Verde", "Materia Excel", "Materia Catálogo", "Comentario", "Subj Original", "Crse Original", "Subj Sugerido", "Crse Sugerido"]
                    
                    df_editado_archivo = st.data_editor(
                        df_vista[columnas_vista],
                        hide_index=True,
                        disabled=["Materia Excel", "Materia Catálogo", "Comentario", "Subj Original", "Crse Original", "Subj Sugerido", "Crse Sugerido"],
                        column_config={
                            "Luz Verde": st.column_config.CheckboxColumn("¿Aplicar?", help="Marca para autorizar este cambio"),
                            "Materia Excel": st.column_config.TextColumn("Materia (Excel)", width="medium"),
                            "Materia Catálogo": st.column_config.TextColumn("Materia (Catálogo)", width="medium"),
                            "Comentario": st.column_config.TextColumn("Diagnóstico", width="medium"),
                            "Subj Original": st.column_config.TextColumn("Subj (Excel)", width="small"),
                            "Crse Original": st.column_config.TextColumn("Crse (Excel)", width="small"),
                            "Subj Sugerido": st.column_config.TextColumn("Subj Sugerido", width="small"),
                            "Crse Sugerido": st.column_config.TextColumn("Crse Sugerido", width="small"),
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
        
        st.markdown("---")
        if st.button("💾 Aplicar Cambios Autorizados y Procesar Todo", type="primary"):
            corregido = st.session_state.raw_altas.copy()
            
            for _, row in st.session_state.res_auditoria.iterrows():
                if row["Luz Verde"] and pd.notna(row["Subj Sugerido"]):
                    corregido.loc[row["idx"], "Subject"] = row["Subj Sugerido"]
                    corregido.loc[row["idx"], "Course"] = row["Crse Sugerido"]
            
            st.session_state.df_corregido = corregido
            
            # Generar Archivo de Resumen para persistencia futura
            summary_str = corregido.to_csv(index=False, encoding="utf-8-sig")
            st.session_state.summary_csv_bytes = summary_str.encode("utf-8-sig")
            
            st.session_state.csv_files_to_download = {}
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for name, sub in corregido.groupby("ArchivoOrigen"):
                    resultado_df = pd.DataFrame()
                    resultado_df["PERIODO"] = sub["Periodo"]
                    resultado_df["SEDE"] = sub["Campus"]
                    resultado_df["SUBJ"] = sub["Subject"]
                    resultado_df["COURSE"] = sub["Course"]
                    resultado_df["PARTEPERIODO"] = sub["Parte de Periodo"]
                    resultado_df["STATUS"]
