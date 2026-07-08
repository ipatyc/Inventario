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
UMBRAL_FUZZY = 0.82  # Precisión para búsqueda de nombres

# 🚨 PARÁMETROS EXACTOS DE R: Esto clona el comportamiento del write.csv() de tu código original
CSV_KWARGS_R = {
    'index': False,
    'encoding': 'utf-8',
    'quoting': csv.QUOTE_NONNUMERIC, # Pone comillas solo a textos, no a números
    'lineterminator': '\r\n',        # Salto de línea estricto de Windows (evita ORA-01400)
    'na_rep': 'NA'                   # Las celdas vacías se escriben como "NA" (como en R)
}

def quitar_acentos(t):
    if pd.isna(t) or t is None: return ""
    return "".join(c for c in unicodedata.normalize("NFD", str(t)) if unicodedata.category(c) != "Mn")

def normalizar_para_cruce(t):
    return quitar_acentos(str(t).upper().strip())

def similitud(a, b): 
    return SequenceMatcher(None, a, b).ratio()

# 🧼 Limpiador adaptado para el CSV: Deja np.nan en los vacíos para que 'na_rep' escriba "NA"
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

# Inicialización de estados en memoria de Streamlit
if "original_files_bytes" not in st.session_state: st.session_state.original_files_bytes = {}
if "df_corregido" not in st.session_state: st.session_state.df_corregido = None
if "raw_altas" not in st.session_state: st.session_state.raw_altas = None
if "res_auditoria" not in st.session_state: st.session_state.res_auditoria = None
if "ready_for_download" not in st.session_state: st.session_state.ready_for_download = False
if "zip_file_bytes" not in st.session_state: st.session_state.zip_file_bytes = None
if "summary_csv_bytes" not in st.session_state: st.session_state.summary_csv_bytes = None
if "csv_files_to_download" not in st.session_state: st.session_state.csv_files_to_download = {}

st.set_page_config(page_title="Consola Iris Cavazos", page_icon="🎛️", layout="wide")
st.title("🎛️ Consola de Control de Materias e Inyección de NRCs")
st.markdown("---")

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
                st.info(f"🔍 Checking / Revisando archivo: **{f.name.split()[0]}**")
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
                    quitar_rep = st.checkbox("🔍 Combinar repetidas", value=True, key=f"rep_{arch}")
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
        if st.button("💾 Aplicar Cambios Autorizados y Procesar Todo", type="primary"):
            corregido = st.session_state.raw_altas.copy()
            for _, row in st.session_state.res_auditoria.iterrows():
                if row["Luz Verde"] and pd.notna(row["Subj Sugerido"]):
                    corregido.loc[row["idx"], "Subject"] = row["Subj Sugerido"]
                    corregido.loc[row["idx"], "Course"] = row["Crse Sugerido"]
            
            st.session_state.df_corregido = corregido
            st.session_state.summary_csv_bytes = corregido.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.session_state.csv_files_to_download = {}
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for name, sub in corregido.groupby("ArchivoOrigen"):
                    resultado_df = pd.DataFrame()
                    
                    # Estructuración exacta a tipos de R (Textos a String, Numéricos a Int64)
                    resultado_df["PERIODO"] = sub["Periodo"].apply(format_r_string)
                    resultado_df["SEDE"] = sub["Campus"].apply(format_r_string)
                    resultado_df["SUBJ"] = sub["Subject"].apply(format_r_string)
                    resultado_df["COURSE"] = sub["Course"].apply(format_r_string)
                    resultado_df["PARTEPERIODO"] = sub["Parte de Periodo"].apply(format_r_string)
                    resultado_df["STATUS"] = sub["Estatus"].apply(format_r_string)
                    
                    # Conversión a números reales para evitar comillas en la carga de Banner
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
                    # Aplicamos los Kwargs de R
                    csv_string = resultado_df.to_csv(**CSV_KWARGS_R)
                    zip_file.writestr(csv_filename, csv_string.encode('utf-8'))
                    st.session_state.csv_files_to_download[csv_filename] = csv_string.encode('utf-8')
            
            st.session_state.zip_file_bytes = zip_buffer.getvalue()
            st.session_state.ready_for_download = True
            st.rerun()

        if st.session_state.ready_for_download:
            st.markdown("### 📥 Panel de Descarga de Resultados")
            st.download_button("📝 📥 DESCARGAR ARCHIVO DE RESUMEN", data=st.session_state.summary_csv_bytes, file_name="resumen.csv", mime="text/csv", use_container_width=True)
            st.download_button("💥 📥 DESCARGAR TODOS LOS CSVs (.ZIP)", data=st.session_state.zip_file_bytes, file_name="archivos_carga.zip", mime="application/zip", use_container_width=True, type="primary")

# ============================================================
# PESTAÑA: FILTRADO POR REPORTE DE ERRORES (BANNER)
# ============================================================
with tab_err:
    st.header("⚠️ Extracción Delta por Reporte de Errores (Banner)")
    metodo_input = st.radio("🛠️ Método para ingresar las líneas con error:", ["Subir archivo Excel de Banner", "Escribir números de línea manualmente"], horizontal=True)
    
    file_errores = None
    txt_errores = ""
    input_valido = False
    
    if metodo_input == "Subir archivo Excel de Banner":
        file_errores = st.file_uploader("📥 Cargar Reporte de Errores (.xlsx)", type=["xlsx"])
        if file_errores: input_valido = True
    else:
        txt_errores = st.text_input("✍️ Escribe las líneas separadas por comas (Ejemplo: 4, 18, 25):")
        if txt_errores.strip(): input_valido = True
            
    st.markdown("---")
    origen_csv = st.radio("¿Base de datos original?", ["Utilizar los CSVs generados en la Pestaña 1 (En memoria)", "Subir un CSV manualmente"])
    
    dict_csvs_a_procesar = {}
    
    if origen_csv == "Utilizar los CSVs generados en la Pestaña 1 (En memoria)":
        if st.session_state.csv_files_to_download:
            opciones = list(st.session_state.csv_files_to_download.keys())
            sel = st.selectbox("Elige el CSV a depurar:", opciones)
            if sel: dict_csvs_a_procesar[sel] = pd.read_csv(io.BytesIO(st.session_state.csv_files_to_download[sel]), encoding="utf-8")
    else:
        file_csv_manual = st.file_uploader("Subir CSV original", type=["csv"])
        if file_csv_manual: dict_csvs_a_procesar[file_csv_manual.name] = pd.read_csv(file_csv_manual, encoding="utf-8")
                
    num_version = st.number_input("🔢 Indica el número de corrección (V):", min_value=1, value=1)
    
    if input_valido and dict_csvs_a_procesar:
        if st.button("🔍 Generar Delta", type="primary"):
            renglones_errores = []
            if metodo_input == "Subir archivo Excel de Banner":
                df_err_excel = pd.read_excel(file_errores, skiprows=2)
                renglones_errores = df_err_excel["Línea"].dropna().astype(int).unique().tolist()
            else:
                renglones_errores = [int(p.strip()) for p in txt_errores.split(",") if p.strip().isdigit()]
                
            for nombre_archivo, df_datos in dict_csvs_a_procesar.items():
                indices = [r - 2 for r in renglones_errores if 0 <= (r - 2) < len(df_datos)]
                df_delta = df_datos.iloc[indices]
                if not df_delta.empty:
                    # Exportar aplicando formato estricto de R
                    csv_bytes = df_delta.to_csv(**CSV_KWARGS_R).encode("utf-8")
                    out_name = f"{nombre_archivo.rsplit('.', 1)[0]}_Correccion_V{num_version}.csv"
                    st.download_button(label=f"📥 Descargar {out_name}", data=csv_bytes, file_name=out_name, mime="text/csv", use_container_width=True)

# ============================================================
# PESTAÑA 3: INYECTAR EN REPORTE ARGOS (CONSTRUIR HOJA "CRNs")
# ============================================================
with tab3:
    st.header("Inyección de NRCs desde Reporte de ARGOS")
    mismo_momento = st.session_state.df_corregido is not None
    procesar_cruce = False
    df_base_cruce, file_argos = None, None
    dict_bytes_altas = {}
    
    if mismo_momento:
        file_argos = st.file_uploader("📊 Cargar Reporte ARGOS (.csv)", type=["csv"], key="a1")
        if file_argos and st.button("🚀 Cruzar Datos", type="primary"):
            df_base_cruce = st.session_state.df_corregido.copy()
            dict_bytes_altas = st.session_state.original_files_bytes
            procesar_cruce = True
    else:
        c1, c2, c3 = st.columns(3)
        with c1: file_argos = st.file_uploader("📊 ARGOS (.csv)", type=["csv"], key="a2")
        with c2: file_resumen = st.file_uploader("📝 Resumen (.csv)", type=["csv"])
        with c3: files_altas_p2 = st.file_uploader("📁 ALTAS (.xlsx)", accept_multiple_files=True, type=["xlsx"])
        if file_argos and file_resumen and files_altas_p2 and st.button("🚀 Cruzar Datos", type="primary"):
            df_base_cruce = pd.read_csv(file_resumen, encoding="utf-8")
            for f in files_altas_p2: dict_bytes_altas[f.name] = f.getvalue()
            procesar_cruce = True

    if procesar_cruce and df_base_cruce is not None:
        argos_df = pd.read_csv(file_argos, encoding="utf-8", on_bad_lines='skip')
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
        
        llaves = ["_k_per", "_k_niv", "_k_sub", "_k_crs", "_k_sec"]
        fusion = solicitud_p2.merge(argos_df[llaves + ["NRC"]], on=llaves, how="left").drop(columns=llaves).drop_duplicates(subset=["NRC"], keep="first")
        columnas_finales = ["NRC"] + [c for c in fusion.columns if c not in ["NRC", "ArchivoOrigen"]]
        
        for name, sub in fusion.groupby("ArchivoOrigen"):
            if name in dict_bytes_altas:
                wb = openpyxl.load_workbook(io.BytesIO(dict_bytes_altas[name]))
                if "CRNs" in wb.sheetnames: del wb["CRNs"]
                ws = wb.create_sheet(title="CRNs")
                df_e = sub[columnas_finales].copy()
                ws.append(list(df_e.columns))
                for r in df_e.values: ws.append(list(r))
                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                st.download_button(f"⬇️ Descargar {name.rsplit('.', 1)[0]} CRNs.xlsx", data=excel_buffer.getvalue(), file_name=f"{name.rsplit('.', 1)[0]} CRNs.xlsx")

        st.markdown("---")
        columnas_cluster = ["Periodo", "CRN", "Tipo.de.Reunión", "Fecha.Inicio", "Fecha.Fin", "Dom", "Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "horarioIni", "horarioFin", "Inicio.de.sesión", "edificio", "salon", "Tipo.de.horario", "indCategoria", "idInstructor", "responsabilidad", "Ind.principal", "ind.sobre.paso", "datocomplementario"]
        df_cluster = pd.DataFrame(columns=columnas_cluster)
        df_cluster["Periodo"] = fusion["Periodo"].apply(format_r_string)
        df_cluster["CRN"] = pd.to_numeric(fusion["NRC"], errors='coerce').astype('Int64')
        df_cluster["datocomplementario"] = fusion.get("Clúster", fusion.get("Cluster", np.nan))
        
        # Exportar clúster con formato estricto de R
        csv_cluster_bytes = df_cluster.to_csv(**CSV_KWARGS_R).encode("utf-8")
        st.download_button("🧩 📥 DESCARGAR CLÚSTER UNIFICADO", data=csv_cluster_bytes, file_name="cluster_unificado.csv", mime="text/csv", type="primary")
