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

def normalizar_llave(val):
    if pd.isna(val) or val is None: return ""
    s = str(val).strip()
    if s.endswith(".0"): s = s[:-2]
    return s.upper()

def normalizar_seccion(val):
    if pd.isna(val) or val is None: return ""
    s = str(val).strip()
    if s.endswith(".0"): s = s[:-2]
    try:
        return str(int(float(s)))
    except:
        return s.upper()

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
st.title("🎛️ Consola de Control de Materias e Inyección de NRCs (Edición Consolidador)")
st.markdown("---")

tab1, tab_err, tab3 = st.tabs([
    "1️⃣ Proceso: Validación y Generar CSV", 
    "⚠️ Reporte de Errores y Consolidador V1, V2...", 
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
        
        st.markdown("#### 📊 Resumen Estadístico de Observations")
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
            st.markdown("### 📥 Panel de Descarga de Resultados")
            st.download_button("📝 📥 DESCARGAR ARCHIVO DE RESUMEN", data=st.session_state.summary_csv_bytes, file_name="resumen.csv", mime="text/csv", use_container_width=True)
            st.download_button("💥 📥 DESCARGAR TODOS LOS CSVs (.ZIP)", data=st.session_state.zip_file_bytes, file_name="archivos_carga.zip", mime="application/zip", use_container_width=True, type="primary")

# ============================================================
# PESTAÑA 2: REPORTE DE ERRORES Y CONSOLIDADOR INTELIGENTE V1, V2
# ============================================================
with tab_err:
    st.header("⚠️ Sección A: Extracción Delta por Reporte de Errores")
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
        file_csv_manual = st.file_uploader("Subir CSV original completo", type=["csv"], key="csv_manual_err")
        if file_csv_manual: 
            dict_csvs_a_procesar[file_csv_manual.name] = pd.read_csv(file_csv_manual, encoding="utf-8")
                
    num_version = st.number_input("🔢 Indica el número de corrección (V):", min_value=1, value=1)
    
    if input_valido and dict_csvs_a_procesar:
        if st.button("🔍 Generar Delta (Segmento de Errores)", type="primary"):
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
                    out_name = f"{nombre_archivo.rsplit('.', 1)[0]}_Correccion_V{num_version}.csv"
                    st.download_button(label=f"📥 Descargar {out_name}", data=csv_bytes, file_name=out_name, mime="text/csv", use_container_width=True)

    # 🔄 SECCIÓN B: RE-GENERADOR Y CONSOLIDADOR DE RESUMEN MAESTRO V1, V2
    st.markdown("---")
    st.header("🔄 Sección B: Re-generador de Resumen Maestro con Cambios Manuales")
    st.markdown("Esta sección toma tu **Resumen completo desactualizado**, se alinea con el **CSV Original**, e inyecta todas tus versiones corregidas a mano (`V1`, `V2`, `V3`...). Esto generará un **Resumen Nuevo y Corregido** listo para cruzarse con ARGOS en la pestaña 3.")
    
    col_c1, col_c2, col_c3 = st.columns(3)
    with col_c1:
        base_resumen = st.file_uploader("📝 1. Sube tu archivo de RESUMEN original (.csv de Pestaña 1)", type=["csv"], key="summary_consolidador")
    with col_c2:
        csv_original = st.file_uploader("🗂️ 2. Sube el CSV Original completo del Campus/Materia", type=["csv"], key="csv_org_consolidador")
    with col_c3:
        archivos_corregidos = st.file_uploader("✏️ 3. Sube tus archivos corregidos a mano (V1, V2, V3...)", type=["csv"], accept_multiple_files=True, key="manuales_consolidador")
        
    if base_resumen and archivos_corregidos:
        if st.button("🧱 Fusionar Correcciones y Re-generar Resumen Maestro", type="primary"):
            try:
                df_resumen_master = pd.read_csv(base_resumen, encoding="utf-8")
                
                # Detectar dinámicamente nombres de columnas en el archivo resumen de entrada
                col_per_res = "Periodo" if "Periodo" in df_resumen_master.columns else "PERIODO"
                col_sed_res = "Campus" if "Campus" in df_resumen_master.columns else "SEDE"
                col_sec_res = "Sección" if "Sección" in df_resumen_master.columns else ("SECCION" if "SECCION" in df_resumen_master.columns else "Seccion")
                col_sub_res = "Subject" if "Subject" in df_resumen_master.columns else "SUBJ"
                col_crs_res = "Course" if "Course" in df_resumen_master.columns else "COURSE"
                
                # Normalizar columnas llave del resumen maestro para el copiado exacto
                df_resumen_master["_k_per"] = df_resumen_master[col_per_res].apply(normalizar_llave)
                df_resumen_master["_k_sed"] = df_resumen_master[col_sed_res].apply(normalizar_llave)
                df_resumen_master["_k_sec"] = df_resumen_master[col_sec_res].apply(normalizar_seccion)
                
                total_reemplazos = 0
                
                # Iterar e inyectar cada archivo V1, V2... subido
                for f_corr in archivos_corregidos:
                    df_c = pd.read_csv(io.BytesIO(f_corr.getvalue()), encoding="utf-8")
                    
                    df_c["_k_per"] = df_c["PERIODO"].apply(normalizar_llave)
                    df_c["_k_sed"] = df_c["SEDE"].apply(normalizar_llave)
                    df_c["_k_sec"] = df_c["SECCION"].apply(normalizar_seccion)
                    
                    for _, row in df_c.iterrows():
                        mask = (
                            (df_resumen_master["_k_per"] == row["_k_per"]) &
                            (df_resumen_master["_k_sed"] == row["_k_sed"]) &
                            (df_resumen_master["_k_sec"] == row["_k_sec"])
                        )
                        if mask.any():
                            idx_target = df_resumen_master[mask].index
                            df_resumen_master.loc[idx_target, col_sub_res] = row["SUBJ"]
                            df_resumen_master.loc[idx_target, col_crs_res] = row["COURSE"]
                            total_reemplazos += 1
                
                # Limpiar columnas temporales de control
                df_resumen_master = df_resumen_master.drop(columns=["_k_per", "_k_sed", "_k_sec"])
                
                # Guardar el nuevo resumen maestro 100% corregido en la memoria global de la app
                st.session_state.df_corregido = df_resumen_master.copy()
                
                # Generar los bytes del nuevo resumen aplicando UTF-8 con BOM para conservar compatibilidad con Excel
                csv_fusionado_bytes = df_resumen_master.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                st.session_state.summary_csv_bytes = csv_fusionado_bytes
                
                st.success(f"🎉 ¡Resumen Re-generado con éxito! Se inyectaron {total_reemplazos} correcciones de tus archivos V1/V2 en el maestro.")
                st.markdown("#### 📥 Descarga tu nuevo Resumen Maestro Corregido:")
                st.download_button("📝 Descargar NUEVO_RESUMEN_CORREGIDO.csv", data=csv_fusionado_bytes, file_name="NUEVO_RESUMEN_CORREGIDO.csv", mime="text/csv", use_container_width=True)
                st.info("💡 **¡Excelente!** La aplicación guardó este nuevo resumen en su memoria temporal. Ya puedes pasar directo a la pestaña **2️⃣ Proceso: Inyección de NRCs (ARGOS)**.")
            except Exception as e:
                st.error(f"❌ Ocurrió un error al procesar el pegado en el resumen: {str(e)}")

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
        st.info("✅ Detectamos que tienes el Resumen Nuevo Corregido en memoria (venga del Consolidador de la Pestaña 2). Se utilizará de forma automática.")
        file_argos = st.file_uploader("📊 Cargar Reporte ARGOS (.csv)", type=["csv"], key="a1")
        if file_argos and st.button("🚀 Cruzar Datos", type="primary"):
            df_base_cruce = st.session_state.df_corregido.copy()
            dict_bytes_altas = st.session_state.original_files_bytes
            procesar_cruce = True
    else:
        c1, c2, c3 = st.columns(3)
        with c1: file_argos = st.file_uploader("📊 ARGOS (.csv)", type=["csv"], key="a2")
        with c2: file_resumen = st.file_uploader("📝 Resumen o CSV Completo Corregido (.csv)", type=["csv"])
        with c3: files_altas_p2 = st.file_uploader("📁 ALTAS (.xlsx)", accept_multiple_files=True, type=["xlsx"])
        if file_argos and file_resumen and files_altas_p2 and st.button("🚀 Cruzar Datos", type="primary"):
            df_base_cruce = pd.read_csv(file_resumen, encoding="utf-8")
            for f in files_altas_p2: dict_bytes_altas[f.name] = f.getvalue()
            procesar_cruce = True

    if procesar_cruce and df_base_cruce is not None:
        try:
            argos_df = pd.read_csv(file_argos, encoding="utf-8", on_bad_lines='skip')
            solicitud_p2 = df_base_cruce.copy()
            
            # Detectar nombres de columnas si viene de formato R o de resumen Excel
            col_per = "PERIODO" if "PERIODO" in solicitud_p2.columns else "Periodo"
            col_sed = "SEDE" if "SEDE" in solicitud_p2.columns else "Campus"
            col_sub = "SUBJ" if "SUBJ" in solicitud_p2.columns else "Subject"
            col_crs = "COURSE" if "COURSE" in solicitud_p2.columns else "Course"
            col_sec = "SECCION" if "SECCION" in solicitud_p2.columns else "Sección"
            col_niv = "Nivel" if "Nivel" in solicitud_p2.columns else None
            
            solicitud_p2["_k_per"] = solicitud_p2[col_per].astype(str).str.strip().apply(lambda x: x[:-2] if x.endswith(".0") else x)
            solicitud_p2["_k_sub"] = solicitud_p2[col_sub].apply(normalizar_para_cruce)
            solicitud_p2["_k_crs"] = solicitud_p2[col_crs].astype(str).str.strip().apply(lambda x: x[:-2] if x.endswith(".0") else x)
            solicitud_p2["_k_sec"] = solicitud_p2[col_sec].apply(limpia_seccion_interna)
            
            argos_df["_k_per"] = argos_df["Periodo"].astype(str).str.strip().apply(lambda x: x[:-2] if x.endswith(".0") else x)
            argos_df["_k_sub"] = argos_df["Área"].apply(normalizar_para_cruce)
            argos_df["_k_crs"] = argos_df["No..Curso"].astype(str).str.strip().apply(lambda x: x[:-2] if x.endswith(".0") else x)
            argos_df["_k_sec"] = argos_df["Grupo"].apply(limpia_seccion_interna)
            
            llaves = ["_k_per", "_k_sub", "_k_crs", "_k_sec"]
            
            if col_niv and "Nivel" in argos_df.columns:
                solicitud_p2["_k_niv"] = solicitud_p2[col_niv].apply(normalizar_para_cruce)
                argos_df["_k_niv"] = argos_df["Nivel"].apply(normalizar_para_cruce)
                llaves.append("_k_niv")
            
            fusion = solicitud_p2.merge(argos_df[llaves + ["NRC"]], on=llaves, how="left").drop(columns=[c for c in llaves if c in solicitud_p2.columns or c.startswith("_k_")])
            fusion = fusion.drop_duplicates(subset=["NRC"], keep="first") if "NRC" in fusion.columns else fusion
            
            if "NRC" in fusion.columns:
                columnas_finales = ["NRC"] + [c for c in fusion.columns if c not in ["NRC", "ArchivoOrigen"]]
            else:
                columnas_finales = list(fusion.columns)
                
            if dict_bytes_altas:
                for name, sub in fusion.groupby(fusion.get("ArchivoOrigen", "Archivo")):
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
            else:
                st.markdown("#### 📊 Vista previa del cruce realizado:")
                st.dataframe(fusion, use_container_width=True)
    
            st.markdown("---")
            columnas_cluster = ["Periodo", "CRN", "Tipo.de.Reunión", "Fecha.Inicio", "Fecha.Fin", "Dom", "Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "horarioIni", "horarioFin", "Inicio.de.sesión", "edificio", "salon", "Tipo.de.horario", "indCategoria", "idInstructor", "responsabilidad", "Ind.principal", "ind.sobre.paso", "datocomplementario"]
            df_cluster = pd.DataFrame(columns=columnas_cluster)
            df_cluster["Periodo"] = fusion[col_per].apply(format_r_string) if col_per in fusion.columns else ""
            df_cluster["CRN"] = pd.to_numeric(fusion["NRC"], errors='coerce').astype('Int64') if "NRC" in fusion.columns else np.nan
            df_cluster["datocomplementario"] = fusion.get("Clúster", fusion.get("Cluster", np.nan))
            
            csv_cluster_bytes = df_cluster.to_csv(**CSV_KWARGS_R).encode("utf-8")
            st.download_button("🧩 📥 DESCARGAR CLÚSTER UNIFICADO", data=csv_cluster_bytes, file_name="cluster_unificado.csv", mime="text/csv", type="primary")
        except Exception as e:
            st.error(f"❌ Error durante el cruce con ARGOS: {str(e)}")
