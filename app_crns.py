# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import io
import unicodedata
import openpyxl
from difflib import SequenceMatcher

# ================= 1. CONFIGURACIÓN Y FUNCIONES (DEBEN IR AQUÍ ARRIBA) =================
HOJA_ALTAS = "ALTAS"
UMBRAL_FUZZY = 0.72

def quitar_acentos(t):
    if pd.isna(t) or t is None: return ""
    return "".join(c for c in unicodedata.normalize("NFD", str(t)) if unicodedata.category(c) != "Mn")

def normalizar_para_cruce(t):
    return quitar_acentos(str(t).upper().strip())

def similitud(a, b): 
    return SequenceMatcher(None, a, b).ratio()

# Guardar archivos originales en memoria para el Proceso 2
if "original_files_bytes" not in st.session_state: st.session_state.original_files_bytes = {}
if "df_corregido" not in st.session_state: st.session_state.df_corregido = None
if "raw_altas" not in st.session_state: st.session_state.raw_altas = None
if "res_auditoria" not in st.session_state: st.session_state.res_auditoria = None

# Configuración de la página visual
st.set_page_config(page_title="Consola Iris Cavazos", page_icon="🎛️", layout="wide")
st.title("🎛️ Consola de Control de Materias e Inyección de NRCs")
st.markdown("---")

# Pestañas de Navegación
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
            
            st.toast("Cargando Catálogo de Materias...", icon="📑")
            xls_cat = pd.ExcelFile(file_cat)
            indice_cat = {}
            for hoja in xls_cat.sheet_names:
                df_c = xls_cat.parse(hoja)
                if "Nivel" in df_c.columns and "Materia" in df_c.columns:
                    for _, f in df_c.iterrows():
                        niv = normalizar_para_cruce(f.get("Nivel"))
                        indice_cat.setdefault(niv, []).append({
                            "mat_norm": normalizar_para_cruce(f.get("Materia")), 
                            "subj": f.get("Subj"), 
                            "crse": f.get("Crse")
                        })
            
            # Procesando cada archivo de Altas
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
                
                # Auditoría fila por fila
                resultados = []
                for idx, fila in df_total.iterrows():
                    niv_n = normalizar_para_cruce(fila.get("Nivel"))
                    mat_n = normalizar_para_cruce(fila.get("Nombre de la Materia"))
                    candidatos = indice_cat.get(niv_n, [])
                    
                    match = next((c for c in candidatos if c["mat_norm"] == mat_n), None)
                    tipo = "exacto"
                    
                    if not match:
                        mejor, mejor_s = None, -1.0
                        for c in candidatos:
                            s = similitud(mat_n, c["mat_norm"])
                            if s > mejor_s: mejor_s, mejor = s, c
                        if mejor and mejor_s >= UMBRAL_FUZZY:
                            match, tipo = mejor, "fuzzy"
                        else:
                            tipo = "no_encontrado"
                    
                    subj_orig = str(fila.get("Subject")).strip()
                    crse_orig = str(fila.get("Course")).strip()
                    subj_sug = str(match["subj"]).strip() if match else None
                    crse_sug = str(match["crse"]).strip() if match else None
                    
                    if tipo == "no_encontrado":
                        comentario = "No se encontro la materia en el catalogo"
                    elif subj_orig != subj_sug and crse_orig != crse_sug:
                        comentario = "Subject y Course incorrectos"
                    elif subj_orig != subj_sug:
                        comentario = "Subject incorrecto"
                    elif crse_orig != crse_sug:
                        comentario = "Course incorrecto"
                    else:
                        comentario = "Todo correcto"
                    
                    resultados.append({
                        "Luz Verde": True if tipo == "fuzzy" else False,
                        "idx": idx, 
                        "Archivo": fila.get("ArchivoOrigen"), 
                        "Materia": fila.get("Nombre de la Materia"), 
                        "Comentario": comentario,
                        "Subj Original": fila.get("Subject"), 
                        "Crse Original": fila.get("Course"),
                        "Subj Sugerido": match["subj"] if match else None, 
                        "Crse Sugerido": match["crse"] if match else None
                    })
                
                st.session_state.res_auditoria = pd.DataFrame(resultados)
                st.success("¡Revisión terminada con éxito!")
            else:
                st.error(f"❌ Ninguno de los archivos subidos tiene la pestaña '{HOJA_ALTAS}'")

    # Mesa de control interactiva
    if st.session_state.res_auditoria is not None:
        st.markdown("### ⚖️ Mesa de Control de Coincidencias")
        df_aud = st.session_state.res_auditoria
        
        filtro = st.radio("Filtro visual:", ["Solo Errores / Sugerencias", "Ver Todo"], horizontal=True)
        df_vis = df_aud if filtro == "Ver Todo" else df_aud[df_aud["Comentario"] != "Todo correcto"]
        
        df_editado = st.data_editor(
            df_vis,
            hide_index=True,
            disabled=["idx", "Archivo", "Materia", "Comentario", "Subj Original", "Crse Original"],
            column_config={
                "Luz Verde": st.column_config.CheckboxColumn("Luz Verde", help="Aprobar corrección automática"),
                "Subj Sugerido": st.column_config.TextColumn("Subj Sugerido"),
                "Crse Sugerido": st.column_config.TextColumn("Crse Sugerido")
            },
            use_container_width=True
        )
        
        for _, row in df_editado.iterrows():
            st.session_state.res_auditoria.loc[st.session_state.res_auditoria["idx"] == row["idx"], "Luz Verde"] = row["Luz Verde"]
            st.session_state.res_auditoria.loc[st.session_state.res_auditoria["idx"] == row["idx"], "Subj Sugerido"] = row["Subj Sugerido"]
            st.session_state.res_auditoria.loc[st.session_state.res_auditoria["idx"] == row["idx"], "Crse Sugerido"] = row["Crse Sugerido"]
            
        if st.button("💾 Aplicar Luz Verde y Generar CSVs", type="primary"):
            corregido = st.session_state.raw_altas.copy()
            
            for _, row in st.session_state.res_auditoria.iterrows():
                if row["Luz Verde"] and pd.notna(row["Subj Sugerido"]):
                    corregido.loc[row["idx"], "Subject"] = row["Subj Sugerido"]
                    corregido.loc[row["idx"], "Course"] = row["Crse Sugerido"]
            
            st.session_state.df_corregido = corregido
            st.success("¡Cambios guardados! Descarga tus archivos .csv estructurados aquí abajo:")
            
            for name, sub in corregido.groupby("ArchivoOrigen"):
                resultado_df = pd.DataFrame()
                
                resultado_df["PERIODO"] = sub["Periodo"]
                resultado_df["SEDE"] = sub["Campus"]
                resultado_df["SUBJ"] = sub["Subject"]
                resultado_df["COURSE"] = sub["Course"]
                resultado_df["PARTEPERIODO"] = sub["Parte de Periodo"]
                resultado_df["STATUS"] = sub["Estatus"]
                resultado_df["CAPACIDAD"] = sub["Capacidad"]
                resultado_df["GRUPOS"] = 1
                resultado_df["SECCION"] = pd.to_numeric(sub["Sección"], errors="coerce").fillna(0).astype(int)
                resultado_df["TIPODEHORARIO"] = sub["Tipo de Horario"]
                resultado_df["METODO_EDUCATIVO"] = sub["Método Educativo"]
                resultado_df["SOCIODEINTEGRACION"] = "D2L"
                resultado_df["MODODECALIFICAR"] = sub["Modo de Calificar"]
                resultado_df["SESION"] = sub["Sesion"]
                
                columnas_ordenadas = ["PERIODO", "SEDE", "SUBJ", "COURSE", "PARTEPERIODO", "STATUS",
                                      "CAPACIDAD", "GRUPOS", "SECCION", "TIPODEHORARIO",
                                      "METODO_EDUCATIVO", "SOCIODEINTEGRACION", "MODODECALIFICAR", "SESION"]
                resultado_df = resultado_df[columnas_ordenadas]
                
                nombre_base = name.rsplit('.', 1)[0] if '.' in name else name
                csv_filename = f"{nombre_base}.csv"
                
                csv_buffer = io.StringIO()
                resultado_df.to_csv(csv_buffer, index=False, encoding="utf-8-sig")
                
                st.download_button(
                    label=f"📥 Descargar {csv_filename}", 
                    data=csv_buffer.getvalue(), 
                    file_name=csv_filename, 
                    mime="text/csv"
                )

# ============================================================
# PESTAÑA 2: INYECTAR EN REPORTE ARGOS (CONSTRUIR HOJA "CRNs")
# ============================================================
with tab2:
    st.header("Inyección de NRCs desde Reporte de ARGOS")
    
    if st.session_state.df_corregido is None:
        st.warning("⚠️ Primero debes completar el Proceso 1 (Validar, dar Luz Verde y Procesar) para poder usar esta pestaña.")
    else:
        file_argos = st.file_uploader("📊 Cargar Reporte de ARGOS (.csv)", type=["csv"])
        
        if file_argos:
            if st.button("🚀 Cruzar Datos y Modificar Excels", type="primary"):
                try:
                    argos_df = pd.read_csv(file_argos, encoding="utf-8")
                except:
                    argos_df = pd.read_csv(file_argos, encoding="latin-1")
                
                st.info("Estandarizando llaves y realizando left_join de R...")
                solicitud_p2 = st.session_state.df_corregido.copy()
                
                solicitud_p2["_k_per"] = solicitud_p2["Periodo"].astype(str).str.strip()
                solicitud_p2["_k_niv"] = solicitud_p2["Nivel"].apply(normalizer_para_cruce)
                solicitud_p2["_k_sub"] = solicitud_p2["Subject"].apply(normalizer_para_cruce)
                solicitud_p2["_k_crs"] = solicitud_p2["Course"].astype(str).str.strip()
                
                def pad_seccion(v):
                    try: return f"{int(float(str(v).strip())):02d}"
                    except: return str(v).strip()
                solicitud_p2["_k_sec"] = solicitud_p2["Sección"].apply(pad_seccion)
                
                argos_df["_k_per"] = argos_df["Periodo"].astype(str).str.strip()
                argos_df["_k_niv"] = argos_df["Nivel"].apply(normalizer_para_cruce)
                argos_df["_k_sub"] = argos_df["Área"].apply(normalizer_para_cruce)
                argos_df["_k_crs"] = argos_df["No..Curso"].astype(str).str.strip()
                argos_df["_k_sec"] = argos_df["Grupo"].apply(pad_seccion)
                
                llaves_cruce = ["_k_per", "_k_niv", "_k_sub", "_k_crs", "_k_sec"]
                
                argos_subset = argos_df[llaves_cruce + ["NRC"]]
                fusion = solicitud_p2.merge(argos_subset, on=llaves_cruce, how="left")
                fusion.drop(columns=llaves_cruce, inplace=True)
                
                fusion = fusion.drop_duplicates(subset=["NRC"], keep="first")
                columnas_finales = ["NRC"] + [c for c in fusion.columns if c != "NRC" and c != "ArchivoOrigen"]
                
                st.success("¡Cruce completado con éxito!")
                st.markdown("#### 📥 Descarga tus Excels modificados con la pestaña 'CRNs':")
                
                for name, sub in fusion.groupby(st.session_state.df_corregido["ArchivoOrigen"]):
                    df_escribir = sub[columnas_finales].copy()
                    
                    original_bytes = st.session_state.original_files_bytes[name]
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
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
