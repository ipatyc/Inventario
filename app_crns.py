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
import difflib
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

def sin_espacios(x):
    v = format_r_string(x)
    if pd.isna(v): return ""
    return str(v).replace(" ", "").upper()

def limpiar_espacios_y_mayusculas(x):
    if pd.isna(x): return ""
    s = re.sub(r'\s+', ' ', str(x))
    return s.strip().upper()

def normalizar_para_busqueda(texto):
    s = str(texto).lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return re.sub(r'[^a-z0-9]', '', s)

# Inicialización de estados en memoria global
if "original_files_bytes" not in st.session_state: st.session_state.original_files_bytes = {}
if "res_auditoria" not in st.session_state: st.session_state.res_auditoria = None
if "raw_altas" not in st.session_state: st.session_state.raw_altas = None
if "ready_for_download" not in st.session_state: st.session_state.ready_for_download = False
if "zip_file_bytes" not in st.session_state: st.session_state.zip_file_bytes = None
if "csv_files_to_download" not in st.session_state: st.session_state.csv_files_to_download = {}
if "delta_files" not in st.session_state: st.session_state.delta_files = {}
if "final_argos_zip" not in st.session_state: st.session_state.final_argos_zip = None
if "df_cruce_rapido" not in st.session_state: st.session_state.df_cruce_rapido = None

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
    col_tit, col_btn = st.columns([4, 1])
    with col_tit:
        st.header("Validación de Claves, Horarios y Generación de CSV")
    with col_btn:
        if st.button("🔄 Limpiar / Recomenzar", type="secondary", use_container_width=True):
            claves_a_borrar = ["file_cat_uploader", "file_cat_ext_uploader", "files_altas_uploader", "res_auditoria", "raw_altas", "ready_for_download", "cat_avanzado_cache", "df_manual_fijo"]
            for k in claves_a_borrar:
                if k in st.session_state:
                    del st.session_state[k]
            st.rerun()

    col1, col2, col3 = st.columns(3)
    with col1: file_cat = st.file_uploader("📑 Catálogo Básico (Niveles)", type=["xlsx"], key="file_cat_uploader")
    with col2: file_cat_ext = st.file_uploader("📚 Catálogo Avanzado (SCBCRSE)", type=["csv", "xlsx"], key="file_cat_ext_uploader")
    with col3: files_altas = st.file_uploader("📁 Archivos ALTAS", accept_multiple_files=True, type=["xlsx"], key="files_altas_uploader")

    columnas_esperadas = [
        "Periodo", "Campus", "Subject", "Course", "Nivel", "Nombre de la Materia",
        "Parte de Periodo", "Estatus", "Capacidad", "Sección", 
        "Tipo de Horario", "Método Educativo", "Modo de Calificar", "Sesion", "Clúster"
    ]
    mapa_huellas = {normalizar_para_busqueda(col): col for col in columnas_esperadas}

    def cargar_catalogo_avanzado():
        if "cat_avanzado_cache" in st.session_state:
            return st.session_state.cat_avanzado_cache, st.session_state.indice_nombres_avanzado
            
        cat_avanzado = {}
        indice_nombres_avanzado = {}
        
        if file_cat_ext is not None:
            try:
                if file_cat_ext.name.lower().endswith('.csv'):
                    df_ext = pd.read_csv(file_cat_ext, dtype=str, encoding='utf-8', on_bad_lines='skip')
                else:
                    df_ext = pd.read_excel(file_cat_ext, dtype=str)
                
                df_ext.columns = [str(c).strip().upper() for c in df_ext.columns]

                for _, row in df_ext.iterrows():
                    subj = sin_espacios(row.get("SCBCRSE_SUBJ_CODE"))
                    crse = sin_espacios(row.get("SCBCRSE_CRSE_NUMB"))
                    if not subj or not crse: continue
                    
                    t_short = limpiar_espacios_y_mayusculas(row.get("SCBCRSE_TITLE"))
                    t_long = limpiar_espacios_y_mayusculas(row.get("SCRSYLN_LONG_COURSE_TITLE"))
                    schd = sin_espacios(row.get("SCRSCHD_SCHD_CODE"))
                    insm = sin_espacios(row.get("SCRSCHD_INSM_CODE"))
                    
                    llave = (subj, crse)
                    if llave not in cat_avanzado:
                        cat_avanzado[llave] = {"titles": set(), "schd": set(), "insm": set()}
                    
                    if t_short and t_short != "NAN": 
                        cat_avanzado[llave]["titles"].add(t_short)
                        indice_nombres_avanzado[normalizar_para_cruce(t_short)] = llave
                    if t_long and t_long != "NAN": 
                        cat_avanzado[llave]["titles"].add(t_long)
                        indice_nombres_avanzado[normalizar_para_cruce(t_long)] = llave
                    if schd and schd != "NAN": cat_avanzado[llave]["schd"].add(schd)
                    if insm and insm != "NAN": cat_avanzado[llave]["insm"].add(insm)
                
                st.session_state.cat_avanzado_cache = cat_avanzado
                st.session_state.indice_nombres_avanzado = indice_nombres_avanzado
            except Exception as e:
                st.error(f"Error al leer el Catálogo Avanzado: {e}")
                
        return cat_avanzado, indice_nombres_avanzado

    if files_altas and file_cat:
        if st.button("⚡ Ejecutar Validación Inteligente", type="primary"):
            st.session_state.ready_for_download = False 
            st.toast("Cargando Catálogos y validando...", icon="📑")

            cat_avanzado, indice_nombres_avanzado = cargar_catalogo_avanzado()
            
            xls_cat = pd.ExcelFile(file_cat)
            indice_cat, indice_cat_claves = {}, {} 
            for hoja in xls_cat.sheet_names:
                df_c = xls_cat.parse(hoja)
                if "Nivel" in df_c.columns and "Materia" in df_c.columns:
                    for _, f in df_c.iterrows():
                        niv = normalizar_para_cruce(f.get("Nivel"))
                        mat_o = limpiar_espacios_y_mayusculas(f.get("Materia"))
                        s_val = sin_espacios(f.get("Subj"))
                        c_val = sin_espacios(f.get("Crse"))
                        indice_cat.setdefault(niv, []).append({
                            "mat_orig": mat_o, "mat_norm": normalizar_para_cruce(f.get("Materia")), 
                            "subj": s_val, "crse": c_val
                        })
                        if pd.notna(s_val) and pd.notna(c_val):
                            indice_cat_claves[(normalizar_para_cruce(s_val), c_val)] = mat_o

            piezas = []
            for f in files_altas:
                st.session_state.original_files_bytes[f.name] = f.getvalue()
                xls_a = pd.ExcelFile(f)
                hojas_reales = [h for h in xls_a.sheet_names if h.strip().upper() == HOJA_ALTAS]
                if hojas_reales:
                    df_a = xls_a.parse(hojas_reales[0], dtype=str)
                    nuevas_columnas = []
                    for col in df_a.columns:
                        huella = normalizar_para_busqueda(col)
                        if huella in mapa_huellas: nuevas_columnas.append(mapa_huellas[huella])
                        else: nuevas_columnas.append(col)
                    df_a.columns = nuevas_columnas

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
                    mat_excel_orig = limpiar_espacios_y_mayusculas(fila.get("Nombre de la Materia"))
                    mat_n = normalizar_para_cruce(mat_excel_orig)
                    subj_orig = sin_espacios(fila.get("Subject"))
                    crse_orig = sin_espacios(fila.get("Course"))
                    
                    horario_orig = sin_espacios(fila.get("Tipo de Horario"))
                    metodo_orig = sin_espacios(fila.get("Método Educativo"))

                    subj_sug, crse_sug = subj_orig, crse_orig
                    comentario_nombres = ""
                    mat_cat_nombre = mat_excel_orig

                    if cat_avanzado: 
                        if (subj_orig, crse_orig) in cat_avanzado:
                            titulos_permitidos = [normalizar_para_cruce(t) for t in cat_avanzado[(subj_orig, crse_orig)]["titles"]]
                            if mat_n in titulos_permitidos:
                                comentario_nombres = "Nombre y Claves OK"
                            else:
                                lista_tits = list(cat_avanzado[(subj_orig, crse_orig)]["titles"])
                                mat_cat_nombre = lista_tits[0] if lista_tits else mat_excel_orig
                                comentario_nombres = "Clave OK, pero Nombre difiere del Catálogo"
                        elif mat_n in indice_nombres_avanzado:
                            subj_sug, crse_sug = indice_nombres_avanzado[mat_n]
                            comentario_nombres = "Claves incorrectas (Match por Nombre en Cat. Avanzado)"
                        else:
                            comentario_nombres = "No hallado en Avanzado. Buscando en Básico..."
                    
                    if not comentario_nombres or "Buscando en Básico" in comentario_nombres:
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
                            if mejor and mejor_s >= UMBRAL_FUZZY: match_elegido = mejor

                        if match_elegido:
                            subj_sug, crse_sug, mat_cat_nombre = match_elegido["subj"], match_elegido["crse"], match_elegido["mat_orig"]
                            if subj_orig == subj_sug and crse_orig == crse_sug: comentario_nombres = "Todo correcto (Cat. Básico)"
                            else: comentario_nombres = "Claves sugeridas (Cat. Básico)"
                        else:
                            s_excel_norm, c_excel_norm = normalizar_para_cruce(subj_orig), crse_orig
                            if (s_excel_norm, c_excel_norm) in indice_cat_claves:
                                mat_cat_nombre = indice_cat_claves[(s_excel_norm, c_excel_norm)]
                                comentario_nombres = "Nombre sugerido por Claves (Cat. Básico)"
                            else:
                                comentario_nombres = "No se encontró en ningún catálogo"

                    horario_sug, metodo_sug = horario_orig, metodo_orig
                    comentario_horario = "Sin catálogo para validar"
                    comentario_metodo = "Sin catálogo para validar"

                    if cat_avanzado and (subj_sug, crse_sug) in cat_avanzado:
                        permitidos_h = cat_avanzado[(subj_sug, crse_sug)]["schd"]
                        permitidos_m = cat_avanzado[(subj_sug, crse_sug)]["insm"]
                        
                        if permitidos_h:
                            if horario_orig in permitidos_h: comentario_horario = "Horario OK"
                            else:
                                comentario_horario = f"Error. Valores permitidos: {', '.join(permitidos_h)}"
                                if len(permitidos_h) == 1: horario_sug = list(permitidos_h)[0]
                                else: horario_sug = "" 
                        else: comentario_horario = "Sin restricciones en catálogo"

                        if permitidos_m:
                            if metodo_orig in permitidos_m: comentario_metodo = "Método OK"
                            else:
                                comentario_metodo = f"Error. Valores permitidos: {', '.join(permitidos_m)}"
                                if len(permitidos_m) == 1: metodo_sug = list(permitidos_m)[0]
                                else: metodo_sug = ""
                        else: comentario_metodo = "Sin restricciones en catálogo"

                    resultados.append({
                        "Luz Verde": False, "idx": idx, "Archivo": fila.get("ArchivoOrigen"), 
                        "Materia Excel": mat_excel_orig, "Materia Catálogo": mat_cat_nombre, 
                        "Comentario Nombres": comentario_nombres, "Subj Original": subj_orig, "Crse Original": crse_orig,
                        "Subj Sugerido": subj_sug, "Crse Sugerido": crse_sug,
                        "Horario Original": horario_orig, "Horario Sugerido": horario_sug, "Comentario Horario": comentario_horario,
                        "Método Original": metodo_orig, "Método Sugerido": metodo_sug, "Comentario Método": comentario_metodo,
                        "Llave_Cruce": f"{fila.get('ArchivoOrigen')}|{mat_excel_orig}|{subj_orig}|{crse_orig}|{idx}"
                    })

                st.session_state.res_auditoria = pd.DataFrame(resultados)
                st.success("¡Revisión de catálogos finalizada!")
            else:
                st.error(f"❌ Ninguno de los archivos subidos tiene filas válidas en la pestaña '{HOJA_ALTAS}'")

    if st.session_state.res_auditoria is not None:
        st.markdown("### ⚖️ Mesa de Control (Dividida en 2 Partes)")
        df_aud = st.session_state.res_auditoria
        archivos_subidos = df_aud["Archivo"].unique()

        for arch in archivos_subidos:
            df_file = df_aud[df_aud["Archivo"] == arch]
            
            cond_nombres = ~df_file["Comentario Nombres"].isin(["Nombre y Claves OK", "Todo correcto (Cat. Básico)"])
            cond_horario = ~df_file["Comentario Horario"].isin(["Horario OK", "Sin restricciones en catálogo", "Sin catálogo para validar"])
            cond_metodo = ~df_file["Comentario Método"].isin(["Método OK", "Sin restricciones en catálogo", "Sin catálogo para validar"])
            
            errores_filas = df_file[cond_nombres | cond_horario | cond_metodo]
            total_detalles = len(errores_filas)

            if total_detalles == 0:
                st.success(f"✅ **{arch}** — ¡Todo limpio, Claves, Horarios y Métodos validados!")
            else:
                with st.expander(f"⚠️ **{arch}** — ({total_detalles} advertencias detectadas)", expanded=True):
                    df_vista = errores_filas.copy()

                    with st.form(key=f"form_{arch}"):
                        st.markdown("Revisa las dos tablas a continuación. Activa la casilla **'¿Aplicar?'** para aceptar las sugerencias.")
                        
                        col_nombres, col_metodos = st.tabs(["🏷️ PARTE 1: Nombres y Claves", "🕒 PARTE 2: Métodos y Horarios"])
                        
                        with col_nombres:
                            columnas_nombres = ["Luz Verde", "Materia Excel", "Materia Catálogo", "Comentario Nombres", "Subj Original", "Crse Original", "Subj Sugerido", "Crse Sugerido"]
                            df_editado_nombres = st.data_editor(
                                df_vista[columnas_nombres], hide_index=True,
                                disabled=["Materia Excel", "Materia Catálogo", "Comentario Nombres", "Subj Original", "Crse Original"],
                                column_config={"Luz Verde": st.column_config.CheckboxColumn("¿Aplicar?")},
                                key=f"edit_nom_{arch}", use_container_width=True
                            )
                            
                        with col_metodos:
                            columnas_metodos = ["Luz Verde", "Materia Excel", "Horario Original", "Horario Sugerido", "Comentario Horario", "Método Original", "Método Sugerido", "Comentario Método"]
                            df_editado_metodos = st.data_editor(
                                df_vista[columnas_metodos], hide_index=True,
                                disabled=["Materia Excel", "Horario Original", "Comentario Horario", "Método Original", "Comentario Método"],
                                column_config={"Luz Verde": st.column_config.CheckboxColumn("¿Aplicar?")},
                                key=f"edit_met_{arch}", use_container_width=True
                            )

                        btn_guardar = st.form_submit_button("💾 Confirmar Selección de Ambas Pestañas")

                        if btn_guardar:
                            df_final_edits = df_vista.copy()
                            df_final_edits["Luz Verde"] = df_editado_nombres["Luz Verde"] | df_editado_metodos["Luz Verde"]
                            df_final_edits["Subj Sugerido"] = df_editado_nombres["Subj Sugerido"]
                            df_final_edits["Crse Sugerido"] = df_editado_nombres["Crse Sugerido"]
                            df_final_edits["Horario Sugerido"] = df_editado_metodos["Horario Sugerido"]
                            df_final_edits["Método Sugerido"] = df_editado_metodos["Método Sugerido"]

                            df_master = st.session_state.res_auditoria.copy()
                            df_master.set_index("Llave_Cruce", inplace=True)
                            df_final_edits.set_index("Llave_Cruce", inplace=True)

                            df_master.update(df_final_edits[["Luz Verde", "Subj Sugerido", "Crse Sugerido", "Horario Sugerido", "Método Sugerido"]])
                            st.session_state.res_auditoria = df_master.reset_index()
                            st.rerun()

        st.markdown("---")
        if st.button("💾 Generar Bloque de Archivos CSV", type="primary"):
            corregido = st.session_state.raw_altas.copy()

            for col in ["Subject", "Course", "Tipo de Horario", "Método Educativo"]:
                if col in corregido.columns: corregido[col] = corregido[col].astype(str)

            for _, row in st.session_state.res_auditoria.iterrows():
                if row["Luz Verde"]:
                    if pd.notna(row["Subj Sugerido"]): corregido.loc[row["idx"], "Subject"] = str(row["Subj Sugerido"])
                    if pd.notna(row["Crse Sugerido"]): corregido.loc[row["idx"], "Course"] = str(row["Crse Sugerido"])
                    if pd.notna(row["Horario Sugerido"]) and row["Horario Sugerido"] != "": corregido.loc[row["idx"], "Tipo de Horario"] = str(row["Horario Sugerido"])
                    if pd.notna(row["Método Sugerido"]) and row["Método Sugerido"] != "": corregido.loc[row["idx"], "Método Educativo"] = str(row["Método Sugerido"])

            st.session_state.csv_files_to_download = {}
            zip_buffer = io.BytesIO()
            errores_encontrados = False

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for name, sub in corregido.groupby("ArchivoOrigen"):
                    columnas_requeridas_csv = [
                        "Periodo", "Campus", "Subject", "Course", "Nivel", 
                        "Parte de Periodo", "Estatus", "Capacidad", 
                        "Sección", "Tipo de Horario", "Método Educativo", 
                        "Modo de Calificar", "Sesion", "Clúster"
                    ]

                    faltantes = [c for c in columnas_requeridas_csv if c not in sub.columns]
                    if faltantes:
                        st.error(f"❌ **Error en `{name}`**: Faltan columnas: **{', '.join(faltantes)}**.")
                        errores_encontrados = True
                        continue

                    resultado_df = pd.DataFrame()
                    resultado_df["PERIODO"] = sub["Periodo"].apply(format_r_string)
                    resultado_df["SEDE"] = sub["Campus"].apply(format_r_string)
                    resultado_df["SUBJ"] = sub["Subject"].apply(lambda x: sin_espacios(format_r_string(x)) if pd.notna(format_r_string(x)) else np.nan)
                    resultado_df["COURSE"] = sub["Course"].apply(lambda x: sin_espacios(format_r_string(x)) if pd.notna(format_r_string(x)) else np.nan)
                    resultado_df["PARTEPERIODO"] = sub["Parte de Periodo"].apply(format_r_string)
                    resultado_df["STATUS"] = sub["Estatus"].apply(format_r_string)
                    resultado_df["CAPACIDAD"] = pd.to_numeric(sub["Capacidad"], errors='coerce').astype('Int64')
                    resultado_df["GRUPOS"] = pd.Series(1, index=resultado_df.index, dtype="Int64")
                    resultado_df["SECCION"] = pd.to_numeric(sub["Sección"], errors='coerce').astype('Int64')
                    resultado_df["TIPODEHORARIO"] = sub["Tipo de Horario"].apply(lambda x: sin_espacios(format_r_string(x)))
                    resultado_df["METODO_EDUCATIVO"] = sub["Método Educativo"].apply(lambda x: sin_espacios(format_r_string(x)))
                    resultado_df["SOCIODEINTEGRACION"] = "D2L"
                    resultado_df["MODODECALIFICAR"] = sub["Modo de Calificar"].apply(format_r_string)
                    resultado_df["SESION"] = sub["Sesion"].apply(format_r_string)

                    def aplicar_reglas_cluster(fila):
                        nivel_actual = str(fila.get("Nivel", "")).strip().upper()
                        cluster_excel = format_r_string(fila.get("Clúster"))
                        if "BACHILLERATO" in nivel_actual: return "BACHILLERATO"
                        return cluster_excel

                    resultado_df["datocomplementario"] = sub.apply(aplicar_reglas_cluster, axis=1)

                    columnas_ordenadas = [
                        "PERIODO", "SEDE", "SUBJ", "COURSE", "PARTEPERIODO", "STATUS",
                        "CAPACIDAD", "GRUPOS", "SECCION", "TIPODEHORARIO",
                        "METODO_EDUCATIVO", "SOCIODEINTEGRACION", "MODODECALIFICAR", "SESION",
                        "datocomplementario"
                    ]
                    resultado_df = resultado_df[columnas_ordenadas]

                    for col in resultado_df.columns:
                        resultado_df[col] = resultado_df[col].astype(str).str.replace('"', '', regex=False).str.strip().replace(['nan', 'None', '<NA>', 'NaN'], '')

                    csv_filename = f"{name.rsplit('.', 1)[0] if '.' in name else name}.csv"
                    csv_string = resultado_df.to_csv(**CSV_KWARGS_R)
                    zip_file.writestr(csv_filename, csv_string.encode('utf-8'))
                    st.session_state.csv_files_to_download[csv_filename] = csv_string.encode('utf-8')

            if not errores_encontrados:
                st.session_state.zip_file_bytes = zip_buffer.getvalue()
                st.session_state.ready_for_download = True
                st.rerun()

        if st.session_state.ready_for_download:
            st.markdown("### 📥 Panel de Descarga")
            st.download_button(
                "💥 📥 DESCARGAR TODOS LOS CSVs (.ZIP)", 
                data=st.session_state.zip_file_bytes, 
                file_name="archivos_carga_banner.zip", 
                mime="application/zip", use_container_width=True, type="primary"
            )

        # ============================================================
        # MÓDULO: CREACIÓN MANUAL Y BUSCADOR MÁGICO FLEXIBLE
        # ============================================================
        st.markdown("---")
        st.subheader("📝 Creación de CSV Manual (Con Autocompletado Flexible)")
        
        st.markdown("#### 🪄 Buscador y Agregador Automático")
        st.info("Puedes buscar escribiendo **SUBJ + COURSE**, **Nombre + COURSE**, **Nombre + SUBJ**, o **solo el Nombre de la Materia**.")
        
        c1, c2, c3 = st.columns(3)
        with c1: input_nombre_busq = st.text_input("Nombre / Título (Opcional)", key="input_nom_busq").strip()
        with c2: input_subj_busq = st.text_input("SUBJ (Opcional)", key="input_subj_busq").strip().upper()
        with c3: input_crse_busq = st.text_input("COURSE (Opcional)", key="input_crse_busq").strip().upper()
        
        if st.button("🪄 Buscar y Agregar a la Tabla", type="secondary", use_container_width=True):
            if not file_cat_ext:
                st.warning("⚠️ Primero sube tu Catálogo Avanzado arriba para poder buscar.")
            elif not input_nombre_busq and not input_subj_busq and not input_crse_busq:
                st.warning("⚠️ Ingresa al menos un criterio (Nombre, SUBJ o COURSE).")
            else:
                cat_av_cache, indice_nom_av = cargar_catalogo_avanzado()
                
                encontrado_subj, encontrado_crse = None, None
                subj_clean = sin_espacios(input_subj_busq)
                crse_clean = sin_espacios(input_crse_busq)
                nom_clean_norm = normalizar_para_cruce(input_nombre_busq) if input_nombre_busq else ""

                if subj_clean and crse_clean and (subj_clean, crse_clean) in cat_av_cache:
                    encontrado_subj, encontrado_crse = subj_clean, crse_clean
                elif nom_clean_norm:
                    if nom_clean_norm in indice_nom_av:
                        s_m, c_m = indice_nom_av[nom_clean_norm]
                        if (not subj_clean or s_m == subj_clean) and (not crse_clean or c_m == crse_clean):
                            encontrado_subj, encontrado_crse = s_m, c_m
                    
                    if not encontrado_subj:
                        for (s, c), data in cat_av_cache.items():
                            match_parcial = any(nom_clean_norm in normalizar_para_cruce(t) or normalizar_para_cruce(t) in nom_clean_norm for t in data["titles"])
                            if match_parcial:
                                if subj_clean and s != subj_clean: continue
                                if crse_clean and c != crse_clean: continue
                                encontrado_subj, encontrado_crse = s, c
                                break
                elif subj_clean and not crse_clean:
                    for (s, c) in cat_av_cache.keys():
                        if s == subj_clean:
                            encontrado_subj, encontrado_crse = s, c
                            break

                if encontrado_subj and encontrado_crse:
                    info = cat_av_cache[(encontrado_subj, encontrado_crse)]
                    horario_magico = list(info["schd"])[0] if info["schd"] else ""
                    metodo_magico = list(info["insm"])[0] if info["insm"] else ""
                    
                    nuevo_renglon = pd.DataFrame([{
                        "PERIODO": "", "SEDE": "", "SUBJ": encontrado_subj, "COURSE": encontrado_crse,
                        "PARTEPERIODO": "", "STATUS": "", "CAPACIDAD": "", "GRUPOS": "1", 
                        "SECCION": "", "TIPODEHORARIO": horario_magico, 
                        "METODO_EDUCATIVO": metodo_magico, "SOCIODEINTEGRACION": "D2L", 
                        "MODODECALIFICAR": "", "SESION": "", "datocomplementario": ""
                    }])
                    
                    st.session_state.df_manual_fijo = pd.concat([nuevo_renglon, st.session_state.df_manual_fijo], ignore_index=True)
                    st.success(f"✅ ¡Materia encontrada y agregada: {encontrado_subj} {encontrado_crse}!")
                    st.rerun()
                else:
                    st.error("❌ No se encontró ninguna materia que coincida con esos datos en el Catálogo Avanzado.")

        columnas_manual = [
            "PERIODO", "SEDE", "SUBJ", "COURSE", "PARTEPERIODO", "STATUS",
            "CAPACIDAD", "GRUPOS", "SECCION", "TIPODEHORARIO",
            "METODO_EDUCATIVO", "SOCIODEINTEGRACION", "MODODECALIFICAR", "SESION",
            "datocomplementario"
        ]

        if "df_manual_fijo" not in st.session_state:
            df_ini = pd.DataFrame([[""] * len(columnas_manual)], columns=columnas_manual)
            df_ini["GRUPOS"] = "1"
            df_ini["SOCIODEINTEGRACION"] = "D2L"
            st.session_state.df_manual_fijo = df_ini

        st.session_state.df_manual_fijo = st.data_editor(
            st.session_state.df_manual_fijo,
            num_rows="dynamic",
            use_container_width=True,
            key="editor_manual_seguro"
        )

        col_man1, col_man2 = st.columns([3, 1])
        with col_man1:
            nombre_csv_manual = st.text_input("Nombre del archivo:", value="carga_manual.csv", key="nom_manual_seguro")
        
        with col_man2:
            st.write("") 
            st.write("") 
            
            df_out_manual = st.session_state.df_manual_fijo.copy()
            for col in df_out_manual.columns:
                df_out_manual[col] = df_out_manual[col].astype(str).str.replace('"', '', regex=False).str.strip().replace(['nan', 'None', '<NA>', 'NaN'], '')
            
            if "SUBJ" in df_out_manual.columns: df_out_manual["SUBJ"] = df_out_manual["SUBJ"].str.upper()
            if "COURSE" in df_out_manual.columns: df_out_manual["COURSE"] = df_out_manual["COURSE"].str.upper()
                
            csv_manual_string = df_out_manual.to_csv(**CSV_KWARGS_R)
            
            st.download_button(
                label="📥 Descargar CSV Manual",
                data=csv_manual_string.encode('utf-8'),
                file_name=nombre_csv_manual if nombre_csv_manual.endswith(".csv") else f"{nombre_csv_manual}.csv",
                mime="text/csv",
                type="primary",
                use_container_width=True
            )

# ============================================================
# PESTAÑA 2 (ANTERIORMENTE TAB_ERR): REPORTE DE ERRORES (EXTRAER DELTA)
# ============================================================
with tab_err:
    st.header("Reporte de Errores y Generación de Deltas")
    st.markdown("Aquí puedes analizar diferencias o extraer deltas si lo requieres.")
    # Espacio reservado funcional de respaldo

# ============================================================
# PESTAÑA 3: INYECCIÓN DE NRCS Y CRUCES CON ARGOS
# ============================================================
with tab3:
    st.header("Inyección de NRCs y Cruces con ARGOS")
    
    modo_inyeccion = st.radio(
        "🛠️ **Elige tu escenario de archivos disponibles:**",
        ["📦 Completo (Tengo ARGOS, CSV Final y Excel Original)", 
         "⚡ Rápido (Tengo ARGOS y CSV Final)"],
        horizontal=True
    )
    st.markdown("---")

    columnas_esperadas_t3 = [
        "Periodo", "Campus", "Subject", "Course", "Nivel", "Nombre de la Materia",
        "Parte de Periodo", "Estatus", "Capacidad", "Sección", 
        "Tipo de Horario", "Método Educativo", "Modo de Calificar", "Sesion", "Clúster"
    ]
    mapa_huellas_t3 = {normalizar_para_busqueda(col): col for col in columnas_esperadas_t3}

    # ====================================================================
    # MODO A: COMPLETO (3 ARCHIVOS)
    # ====================================================================
    if modo_inyeccion == "📦 Completo (Tengo ARGOS, CSV Final y Excel Original)":
        st.markdown("Procesa los Excels originales, inyecta los NRCs de ARGOS cruzando con el CSV final y genera un ZIP con los Excels actualizados.")
        col_a, col_b, col_c = st.columns(3)
        with col_a: file_argos = st.file_uploader("📊 1. Reporte ARGOS (.csv)", type=["csv"], key="arg_c")
        with col_b: files_csv_finales = st.file_uploader("📝 2. CSVs Finales", type=["csv"], accept_multiple_files=True, key="csv_c")
        with col_c: files_xlsx_originales = st.file_uploader("📁 3. Excels Originales", type=["xlsx"], accept_multiple_files=True, key="xls_c")
            
        if file_argos and files_csv_finales and files_xlsx_originales:
            if st.button("🚀 PROCESAR Y GENERAR EXCELS CON NRC", type="primary"):
                try:
                    argos_df = pd.read_csv(file_argos, encoding="utf-8", on_bad_lines='skip', dtype=str)
                    argos_df.columns = [re.sub(r'\.+', '.', str(c).replace('"', '').replace("'", "").strip()) for c in argos_df.columns]
                    
                    col_argos_cluster = next((c for c in argos_df.columns if "cluster" in normalizar_para_busqueda(c)), None)
                    col_argos_area = next((c for c in argos_df.columns if "area" in normalizar_para_busqueda(c)), None)
                    col_argos_curso = next((c for c in argos_df.columns if "curso" in normalizar_para_busqueda(c)), None)

                    if not col_argos_cluster: raise KeyError("No se encontró la columna de Cluster en ARGOS.")
                    if not col_argos_area: raise KeyError("No se encontró la columna de Área en ARGOS.")
                    if not col_argos_curso: raise KeyError("No se encontró la columna de Curso en ARGOS.")

                    argos_df["Periodo"] = argos_df["Periodo"].apply(format_r_string).apply(ultra_limpiar if 'ultra_limpiar' in globals() else str)
                    # Definición local segura de ultra_limpiar para t3 si no existe
                    def t3_ul(x):
                        if pd.isna(x): return ""
                        s = str(x).strip().upper().replace(" ", "")
                        if s.endswith(".0"): s = s[:-2]
                        return s

                    def t3_uls(x):
                        if pd.isna(x): return ""
                        s = str(x).strip().upper().replace(" ", "")
                        if s.endswith(".0"): s = s[:-2]
                        if s.isdigit(): return f"{int(s):02d}"
                        return s

                    argos_df["Periodo"] = argos_df["Periodo"].apply(t3_ul)
                    argos_df["Nivel"] = argos_df["Nivel"].apply(t3_ul)
                    argos_df[col_argos_cluster] = argos_df[col_argos_cluster].apply(t3_ul)
                    argos_df[col_argos_area] = argos_df[col_argos_area].apply(t3_ul)
                    argos_df[col_argos_curso] = argos_df[col_argos_curso].apply(t3_ul)
                    argos_df["Grupo"] = argos_df["Grupo"].apply(t3_uls)
                    
                    argos_df["_llave_argos"] = (
                        argos_df["Periodo"] + "_" + 
                        argos_df["Nivel"] + "_" + 
                        argos_df[col_argos_cluster] + "_" + 
                        argos_df[col_argos_area] + "_" + 
                        argos_df[col_argos_curso] + "_" + 
                        argos_df["Grupo"]
                    )
                    argos_df = argos_df.drop_duplicates(subset=["_llave_argos"])
                    mapa_nrcs = dict(zip(argos_df["_llave_argos"], argos_df["NRC"]))
                    llaves_argos_disponibles = list(mapa_nrcs.keys())

                    excels_inyectados_zip = io.BytesIO()
                    archivos_procesados_con_exito = 0
                    alertas_dimensiones, alertas_parejas = [], []
                    alertas_nrc_faltantes = []
                    
                    with zipfile.ZipFile(excels_inyectados_zip, "w", zipfile.ZIP_DEFLATED) as zip_out:
                        for fx in files_xlsx_originales:
                            df_csv, fc_usado = None, None
                            base_excel = simplificar_nombre(fx.name) if 'simplificar_nombre' in globals() else fx.name.lower()
                            
                            for fc_cand in files_csv_finales:
                                base_csv = simplificar_nombre(fc_cand.name) if 'simplificar_nombre' in globals() else fc_cand.name.lower()
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
                                    
                                    nuevas_columnas = []
                                    for col in df_excel_original.columns:
                                        huella = normalizar_para_busqueda(col)
                                        if huella in mapa_huellas_t3:
                                            nuevas_columnas.append(mapa_huellas_t3[huella])
                                        else:
                                            nuevas_columnas.append(col)
                                    df_excel_original.columns = nuevas_columnas
                                    
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
                                    
                                    def cor_niv(cluster_val):
                                        cluster_str = str(cluster_val).strip().lower()
                                        if "posgrado" in cluster_str: return "POSGRADO"
                                        elif "bachillerato" in cluster_str: return "BACHILLERATO"
                                        else: return "LICENCIATURA"

                                    cluster_csv_series = df_csv["datocomplementario"] if "datocomplementario" in df_csv.columns else pd.Series([""] * len(df_csv))
                                    nivel_corregido = cluster_csv_series.apply(cor_niv).apply(t3_ul)
                                    cluster_limpio_csv = cluster_csv_series.apply(t3_ul)
                                    
                                    llaves_cruce = (
                                        df_nrc_pestana["Periodo"].apply(t3_ul) + "_" + 
                                        nivel_corregido + "_" + 
                                        cluster_limpio_csv + "_" + 
                                        df_nrc_pestana["Subject"].apply(t3_ul) + "_" + 
                                        df_nrc_pestana["Course"].apply(t3_ul) + "_" + 
                                        df_nrc_pestana["Sección"].apply(t3_uls)
                                    )
                                    
                                    nrc_mapeados = llaves_cruce.map(mapa_nrcs)
                                    
                                    faltantes = llaves_cruce[nrc_mapeados.isna()]
                                    if not faltantes.empty:
                                        for llave_rota in faltantes.unique():
                                            sugerencias = difflib.get_close_matches(str(llave_rota), llaves_argos_disponibles, n=1, cutoff=0.5)
                                            sugerencia_txt = f"👉 En ARGOS lo más parecido es: **{sugerencias[0]}**" if sugerencias else "👉 (No se encontró nada parecido en ARGOS)"
                                            alertas_nrc_faltantes.append(f"❌ `{fx.name}` buscó: **{llave_rota}** \n{sugerencia_txt}")

                                    df_nrc_pestana.insert(0, "NRC", nrc_mapeados)
                                    
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
                                    
                            else:
                                alertas_parejas.append(f"⚠️ `{fx.name}` no encontró ningún CSV compatible.")

                    if archivos_procesados_con_exito > 0:
                        st.session_state.final_argos_zip = excels_inyectados_zip.getvalue()
                        st.success(f"🎉 ¡Proceso finalizado! Se procesaron {archivos_procesados_con_exito} archivos exitosamente.")
                    else: 
                        st.error("❌ No se pudo procesar ningún archivo.")
                    
                    if alertas_nrc_faltantes:
                        st.markdown("### 🔍 Radar de Llaves Rotas (Comparación Frente a Frente):")
                        for alerta in alertas_nrc_faltantes: st.error(alerta)
                            
                    if alertas_dimensiones:
                        st.markdown("### 🚫 Archivos descartados por diferencia de filas:")
                        for alerta in alertas_dimensiones: st.error(alerta)
                    if alertas_parejas:
                        st.markdown("### ❓ Archivos sin pareja:")
                        for alerta in alertas_parejas: st.warning(alerta)

                except Exception as e:
                    st.error(f"❌ Ocurrió un inconveniente crítico: {str(e)}")

        if st.session_state.final_argos_zip is not None:
            st.markdown("### 📥 Panel de Descarga (Excels Inyectados)")
            st.download_button(
                label="📁 📥 DESCARGAR EXCELS CON NRC (.ZIP)", data=st.session_state.final_argos_zip,
                file_name="Excels_Finales_con_NRC.zip", mime="application/zip",
                use_container_width=True, type="primary"
            )

    # ====================================================================
    # MODO B: RÁPIDO (SOLO CSV FINAL + ARGOS)
    # ====================================================================
    elif modo_inyeccion == "⚡ Rápido (Tengo ARGOS y CSV Final)":
        st.markdown("Extrae los NRC de ARGOS cruzando con el CSV final y genera una tabla limpia que puedes copiar o descargar en Excel.")
        
        col_r1, col_r2 = st.columns(2)
        with col_r1: file_argos_rap = st.file_uploader("📊 1. Reporte ARGOS (.csv)", type=["csv"], key="arg_r")
        with col_r2: files_csv_rap = st.file_uploader("📝 2. CSVs Finales", type=["csv"], accept_multiple_files=True, key="csv_r")
        
        if file_argos_rap and files_csv_rap:
            if st.button("⚡ Cruzar NRC y Generar Tabla", type="primary"):
                try:
                    argos_df = pd.read_csv(file_argos_rap, encoding="utf-8", on_bad_lines='skip', dtype=str)
                    argos_df.columns = [re.sub(r'\.+', '.', str(c).replace('"', '').replace("'", "").strip()) for c in argos_df.columns]
                    
                    col_argos_cluster = next((c for c in argos_df.columns if "cluster" in normalizar_para_busqueda(c)), None)
                    col_argos_area = next((c for c in argos_df.columns if "area" in normalizar_para_busqueda(c)), None)
                    col_argos_curso = next((c for c in argos_df.columns if "curso" in normalizar_para_busqueda(c)), None)

                    if not col_argos_cluster: raise KeyError("No se encontró la columna de Cluster en ARGOS.")
                    if not col_argos_area: raise KeyError("No se encontró la columna de Área en ARGOS.")
                    if not col_argos_curso: raise KeyError("No se encontró la columna de Curso en ARGOS.")

                    def t3_ul(x):
                        if pd.isna(x): return ""
                        s = str(x).strip().upper().replace(" ", "")
                        if s.endswith(".0"): s = s[:-2]
                        return s

                    def t3_uls(x):
                        if pd.isna(x): return ""
                        s = str(x).strip().upper().replace(" ", "")
                        if s.endswith(".0"): s = s[:-2]
                        if s.isdigit(): return f"{int(s):02d}"
                        return s

                    argos_df["Periodo"] = argos_df["Periodo"].apply(t3_ul)
                    argos_df["Nivel"] = argos_df["Nivel"].apply(t3_ul)
                    argos_df[col_argos_cluster] = argos_df[col_argos_cluster].apply(t3_ul)
                    argos_df[col_argos_area] = argos_df[col_argos_area].apply(t3_ul)
                    argos_df[col_argos_curso] = argos_df[col_argos_curso].apply(t3_ul)
                    argos_df["Grupo"] = argos_df["Grupo"].apply(t3_uls)
                    
                    argos_df["_llave_argos"] = (
                        argos_df["Periodo"] + "_" + 
                        argos_df["Nivel"] + "_" + 
                        argos_df[col_argos_cluster] + "_" + 
                        argos_df[col_argos_area] + "_" + 
                        argos_df[col_argos_curso] + "_" + 
                        argos_df["Grupo"]
                    )
                    argos_df = argos_df.drop_duplicates(subset=["_llave_argos"])
                    mapa_nrcs = dict(zip(argos_df["_llave_argos"], argos_df["NRC"]))
                    llaves_argos_disponibles = list(mapa_nrcs.keys())
                    
                    dfs_combinados = []
                    alertas_nrc_rapido = []
                    
                    for fc in files_csv_rap:
                        df_c = pd.read_csv(io.BytesIO(fc.getvalue()), encoding="utf-8", dtype=str)
                        df_c = df_c.dropna(how='all')
                        
                        cluster_csv_series = df_c["datocomplementario"] if "datocomplementario" in df_c.columns else pd.Series([""] * len(df_c))
                        
                        def cor_niv(cluster_val):
                            cluster_str = str(cluster_val).strip().lower()
                            if "posgrado" in cluster_str: return "POSGRADO"
                            elif "bachillerato" in cluster_str: return "BACHILLERATO"
                            else: return "LICENCIATURA"

                        nivel_csv = cluster_csv_series.apply(cor_niv).apply(t3_ul)
                        cluster_limpio_csv = cluster_csv_series.apply(t3_ul)
                        
                        llaves_csv = (
                            df_c.get("PERIODO", pd.Series(dtype=str)).apply(t3_ul) + "_" + 
                            nivel_csv + "_" + 
                            cluster_limpio_csv + "_" + 
                            df_c.get("SUBJ", pd.Series(dtype=str)).apply(t3_ul) + "_" + 
                            df_c.get("COURSE", pd.Series(dtype=str)).apply(t3_ul) + "_" + 
                            df_c.get("SECCION", pd.Series(dtype=str)).apply(t3_uls)
                        )
                        
                        nrc_asignados = llaves_csv.map(mapa_nrcs)
                        
                        faltantes = llaves_csv[nrc_asignados.isna()]
                        if not faltantes.empty:
                            for llave_rota in faltantes.unique():
                                sugerencias = difflib.get_close_matches(str(llave_rota), llaves_argos_disponibles, n=1, cutoff=0.5)
                                sugerencia_txt = f"👉 En ARGOS lo más parecido es: **{sugerencias[0]}**" if sugerencias else "👉 (No se encontró nada parecido)"
                                alertas_nrc_rapido.append(f"❌ Buscó: **{llave_rota}** \n{sugerencia_txt}")
                            
                        df_c.insert(0, "NRC", nrc_asignados)
                        
                        df_c = df_c.rename(columns={
                            "TIPODEHORARIO": "TIPO DE HORARIO",
                            "METODO_EDUCATIVO": "METODO_ED"
                        })
                        
                        columnas_deseadas = ["NRC", "PERIODO", "SUBJ", "COURSE", "CAPACIDAD", "SECCION", "TIPO DE HORARIO", "METODO_ED", "datocomplementario"]
                        columnas_finales = [c for c in columnas_deseadas if c in df_c.columns]
                        
                        dfs_combinados.append(df_c[columnas_finales])
                        
                    if dfs_combinados:
                        df_resultado_rapido = pd.concat(dfs_combinados, ignore_index=True)
                        st.session_state.df_cruce_rapido = df_resultado_rapido
                        st.success("✅ ¡Tabla cruzada generada exitosamente!")
                        
                        if alertas_nrc_rapido:
                            st.warning("⚠️ Ojo: Algunas filas no encontraron su NRC. Revisa las discrepancias abajo:")
                            with st.expander("🔍 Ver llaves que no cruzaron (Frente a Frente)"):
                                for a in alertas_nrc_rapido: st.write(a)
                    else:
                        st.error("No se pudo procesar la información de los CSV.")
                        
                except Exception as e:
                    st.error(f"❌ Ocurrió un error en el cruce rápido: {str(e)}")

        if st.session_state.df_cruce_rapido is not None:
            st.markdown("### 📋 Resultados del Cruce (NRC inyectados)")
            st.dataframe(st.session_state.df_cruce_rapido, use_container_width=True)
            
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                st.markdown("#### 📝 ¿Quieres copiar la tabla directamente?")
                st.info("Haz clic en el botón de **'Copiar'** en la esquina superior derecha del cuadro de abajo y pégalo (Ctrl+V) en tu Excel.")
                tsv_rapido = st.session_state.df_cruce_rapido.to_csv(index=False, sep='\t')
                st.code(tsv_rapido, language="text")
            
            with col_b2:
                st.markdown("#### 📥 O descárgala en formato Excel")
                st.info("Si prefieres descargar el archivo listo para usar, usa este botón:")
                excel_rapido_buffer = io.BytesIO()
                st.session_state.df_cruce_rapido.to_excel(excel_rapido_buffer, index=False)
                
                st.download_button(
                    label="📥 Descargar esta tabla (.xlsx)",
                    data=excel_rapido_buffer.getvalue(),
                    file_name="Cruce_Rapido_NRC.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
# ============================================================
# PESTAÑA 2: REPORTE DE ERRORES Y ENSAMBLAJE FINAL
# ============================================================
with tab_err:
    st.header("⚠️ Reporte de Errores y Ensamblaje Final")
    st.markdown("Extrae filas con error, corrígelas y genera el archivo para la Pestaña 3.")
    
    # --- PASO 1: EXTRAER O EDITAR EL PEDACITO CON ERROR ---
    st.subheader("✂️ 1. Extraer o corregir el pedacito con errores")
    
    # 🔥 LA FUNCIÓN ESTÁ AQUÍ MISMO PARA QUE NUNCA MARQUE "NOT DEFINED" 🔥
    def limpiar_nombre_columna(col):
        if pd.isna(col): return ""
        return " ".join(str(col).split())
    
    col_ex1, col_ex2, col_ex3 = st.columns(3)
    with col_ex1: file_base_ext = st.file_uploader("📁 1. Archivo Base (.csv)", type=["csv"], key="ext_base_1")
    with col_ex2: file_err_ext = st.file_uploader("📊 2. Reporte de Errores Banner (.xlsx)", type=["xlsx"], key="ext_err_1")
    with col_ex3: sufijo_version = st.text_input("🔢 Sufijo de versión (Ej: V1, V2):", value="V1", key="suf_v1")
    
    # Memoria caché interna para proteger la conexión del data_editor
    if "df_delta_cache" not in st.session_state: st.session_state.df_delta_cache = None
    if "nombre_delta_cache" not in st.session_state: st.session_state.nombre_delta_cache = None
    if "llave_control_archivos" not in st.session_state: st.session_state.llave_control_archivos = ""

    if file_base_ext and file_err_ext:
        # Monitoreamos si cambiaste de archivos subidos para resetear la memoria limpia
        llave_actual = f"{file_base_ext.name}_{file_err_ext.name}_{sufijo_version}"
        if st.session_state.llave_control_archivos != llave_actual:
            st.session_state.df_delta_cache = None
            st.session_state.nombre_delta_cache = None
            st.session_state.llave_control_archivos = llave_actual

        # Botón disparador: Lee y procesa los archivos una sola vez
        if st.button("🔍 Cargar y Procesar Reporte de Errores", use_container_width=True, type="secondary"):
            try:
                df_base = pd.read_csv(file_base_ext, encoding="utf-8", dtype=str)
                df_err = pd.read_excel(file_err_ext, skiprows=2)
                
                # Destruimos saltos de línea y espacios en los títulos de la tabla de errores usando la función local
                df_err.columns = [limpiar_nombre_columna(c) for c in df_err.columns]
                
                # Búsqueda higiénica de la columna Línea (ignora acentos y mayúsculas/minúsculas)
                col_linea = [c for c in df_err.columns if "linea" in str(c).strip().lower().replace("í", "i")]
                
                if not col_linea:
                    st.error("❌ No se encontró la columna 'Línea' en el reporte de errores. Verifica la estructura de tu archivo.")
                else:
                    nombre_col = col_linea[0]
                    df_err = df_err.dropna(subset=[nombre_col])
                    
                    # Extracción segura mapeando floats a enteros e ignorando celdas vacías (NaN)
                    indices = [int(float(r)) - 2 for r in df_err[nombre_col].unique().tolist() if pd.notna(r) and 0 <= (int(float(r)) - 2) < len(df_base)]
                    
                    if indices:
                        st.session_state.df_delta_cache = df_base.iloc[indices].copy()
                        base_name_ext = file_base_ext.name.rsplit('.', 1)[0].replace("_base", "").replace("_final", "")
                        st.session_state.nombre_delta_cache = f"{base_name_ext}_{sufijo_version}"
                        st.success(f"🎉 Éxito: Se aislaron {len(indices)} filas con anomalías. Configura tu descarga abajo.")
                    else:
                        st.warning("⚠️ No se identificaron números de línea válidos dentro de las dimensiones del archivo base.")
            except Exception as e:
                st.error(f"❌ Error crítico de lectura física: {str(e)}")

    # Interfaz de Descarga/Edición: Lee directo de la memoria RAM (Caché), evitando desconexiones por clics continuos
    if st.session_state.df_delta_cache is not None:
        st.markdown("---")
        modo_delta = st.radio("⚙️ ¿Cómo deseas descargar o corregir el fragmento?", ["Excel (.xlsx)", "CSV (.csv)", "Editar en vivo"], horizontal=True, key="modo_1")
        
        nombre_archivo = st.session_state.nombre_delta_cache
        df_delta = st.session_state.df_delta_cache
        
        if modo_delta == "Excel (.xlsx)":
            buf = io.BytesIO()
            df_delta.to_excel(buf, index=False)
            st.download_button("📥 Descargar Fragmento (.xlsx)", data=buf.getvalue(), file_name=f"{nombre_archivo}.xlsx", type="primary", use_container_width=True)
        
        elif modo_delta == "CSV (.csv)":
            st.download_button("📥 Descargar Fragmento (.csv)", data=df_delta.to_csv(**CSV_KWARGS_R).encode("utf-8"), file_name=f"{nombre_archivo}.csv", type="primary", use_container_width=True)
        
        else:
            st.info("✏️ **Modo Edición Interactiva:** Escribe tus ajustes directamente en las celdas de la tabla. Al finalizar, haz clic en el botón inferior para exportarlo.")
            df_editado = st.data_editor(df_delta, key="ed_vivo_1", use_container_width=True)
            st.download_button("📥 Descargar Parche Corregido (.csv)", data=df_editado.to_csv(**CSV_KWARGS_R).encode("utf-8"), file_name=f"{nombre_archivo}.csv", type="primary", use_container_width=True)

    
# --- PASO 2: INYECTAR Y CREAR EL ARCHIVO FINAL ---
    st.subheader("💉 2. Inyectar correcciones y generar Archivo Final")
    
    col_in1, col_in2, col_in3 = st.columns(3)
    with col_in1: file_base_iny = st.file_uploader("📁 1. Archivo Base (.csv)", type=["csv"], key="iny_base_2")
    with col_in2: file_err_iny = st.file_uploader("📊 2. Reporte de Errores (.xlsx)", type=["xlsx"], key="iny_err_2")
    with col_in3: 
        file_corr_iny = st.file_uploader("📝 3. Fragmento Corregido", type=["csv", "xlsx"], key="iny_corr_2")
        tipo_final = st.text_input("Etiqueta final (V1, V2, final):", value="final", key="suf_v2")
    
    # 🔥 INICIALIZAR MEMORIA: Para que la app no "truene" al querer descargar
    if "archivo_final_bytes" not in st.session_state: st.session_state.archivo_final_bytes = None
    if "archivo_final_nombre" not in st.session_state: st.session_state.archivo_final_nombre = None
    
    if file_base_iny and file_err_iny and file_corr_iny:
        if st.button("🚀 Ensamblar Archivo Final", type="primary"):
            try: # 🛡️ ESCUDO GLOBAL ANTI-CAÍDAS
                df_base = pd.read_csv(file_base_iny, encoding="utf-8", dtype=str)
                df_err = pd.read_excel(file_err_iny, skiprows=2)
                
                # 🔥 Búsqueda inteligente de la columna Línea 
                col_linea_iny = [c for c in df_err.columns if "linea" in str(c).strip().lower().replace("í", "i")]
                
                if not col_linea_iny:
                    st.error("❌ No se encontró la columna 'Línea' en el reporte.")
                else:
                    nombre_col_iny = col_linea_iny[0]
                    df_err = df_err.dropna(subset=[nombre_col_iny])
                    df_corr = pd.read_excel(file_corr_iny, dtype=str) if file_corr_iny.name.endswith('.xlsx') else pd.read_csv(file_corr_iny, encoding="utf-8", dtype=str)
                    
                    # 🔥 Conversión segura de floats/nans a int
                    indices = [int(float(r)) - 2 for r in df_err[nombre_col_iny].unique().tolist() if pd.notna(r) and 0 <= (int(float(r)) - 2) < len(df_base)]
                    
                    if len(indices) == len(df_corr):
                        df_final = df_base.copy()
                        for col in df_final.columns:
                            if col in df_corr.columns: df_final.iloc[indices, df_final.columns.get_loc(col)] = df_corr[col].values
                        
                        # Limpieza final para Banner
                        for col in df_final.columns:
                            df_final[col] = df_final[col].astype(str).str.replace('"', '', regex=False).str.strip().replace(['nan', 'None', '<NA>', 'NaN'], '')
                        
                        base_name_iny = file_base_iny.name.rsplit('.', 1)[0].replace("_base", "").replace("_final", "")
                        out_name = f"{base_name_iny}_{tipo_final}.csv"
                        
                        # 🔥 GUARDAMOS EN LA MEMORIA PARA QUE EL BOTÓN SOBREVIVA
                        st.session_state.archivo_final_bytes = df_final.to_csv(**CSV_KWARGS_R).encode("utf-8")
                        st.session_state.archivo_final_nombre = out_name
                        st.success(f"🎉 ¡Archivo {out_name} listo! Da clic en el botón debajo para descargar.")
                    else:
                        st.error(f"❌ Desajuste de filas: {len(indices)} errores en Banner vs {len(df_corr)} filas corregidas en tu archivo.")
            except Exception as e:
                st.error(f"❌ Error interno al ensamblar: {str(e)}")

    # 🔥 BOTÓN DE DESCARGA AFUERA DEL "IF BUTTON": Esto evita la recarga mortal
    if st.session_state.archivo_final_bytes is not None:
        st.download_button(
            label=f"📁 📥 DESCARGAR {st.session_state.archivo_final_nombre}", 
            data=st.session_state.archivo_final_bytes, 
            file_name=st.session_state.archivo_final_nombre, 
            type="primary", 
            use_container_width=True
        )


# ============================================================
# PESTAÑA 3: INYECCIÓN DE NRCS Y CRUCES CON ARGOS
# ============================================================
with tab3:
    import datetime
    import difflib
    
    # --- INICIALIZAR MEMORIA ---
    if "df_cruce_rapido" not in st.session_state: st.session_state.df_cruce_rapido = None

    st.header("Inyección de NRCs y Cruces con ARGOS")
    
    # 👇 Selector de Modo de Trabajo
    modo_inyeccion = st.radio(
        "🛠️ **Elige tu escenario de archivos disponibles:**",
        ["📦 Completo (Tengo ARGOS, CSV Final y Excel Original)", 
         "⚡ Rápido (Tengo ARGOS y CSV Final)"],
        horizontal=True
    )
    st.markdown("---")
    
    # 🔥 FUNCIONES LOCALES DE BÚSQUEDA Y LIMPIEZA
    def normalizar_para_busqueda_t3(texto):
        s = str(texto).lower()
        s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
        return re.sub(r'[^a-z0-9]', '', s)

    def ultra_limpiar(x):
        if pd.isna(x): return ""
        s = str(x).strip().upper().replace(" ", "")
        if s.endswith(".0"): s = s[:-2]
        return s

    def ultra_limpiar_seccion(x):
        if pd.isna(x): return ""
        s = str(x).strip().upper().replace(" ", "")
        if s.endswith(".0"): s = s[:-2]
        if s.isdigit(): return f"{int(s):02d}"
        return s

    def corregir_nivel_por_cluster_csv(cluster_val):
        cluster_str = str(cluster_val).strip().lower()
        if "posgrado" in cluster_str: return "POSGRADO"
        elif "bachillerato" in cluster_str: return "BACHILLERATO"
        else: return "LICENCIATURA"
        
    def simplificar_nombre(nombre):
        n = nombre.lower()
        for basura in ['.xlsx', '.xls', '.csv', '_final', '_base', '_v1', '_v2', '_v3', '_v4', 'corregidas_', 'errores_']:
            n = n.replace(basura, '')
        return n.strip().replace(" ", "")

    columnas_esperadas_t3 = [
        "Periodo", "Campus", "Subject", "Course", "Nivel", "Nombre de la Materia",
        "Parte de Periodo", "Estatus", "Capacidad", "Sección", 
        "Tipo de Horario", "Método Educativo", "Modo de Calificar", "Sesion", "Clúster"
    ]
    mapa_huellas_t3 = {normalizar_para_busqueda_t3(col): col for col in columnas_esperadas_t3}

    # ====================================================================
    # MODO A: COMPLETO (3 ARCHIVOS)
    # ====================================================================
    if modo_inyeccion == "📦 Completo (Tengo ARGOS, CSV Final y Excel Original)":
        st.markdown("Procesa los Excels originales, inyecta los NRCs de ARGOS cruzando con el CSV final y genera un ZIP con los Excels actualizados.")
        col_a, col_b, col_c = st.columns(3)
        with col_a: file_argos = st.file_uploader("📊 1. Reporte ARGOS (.csv)", type=["csv"], key="arg_c")
        with col_b: files_csv_finales = st.file_uploader("📝 2. CSVs Finales", type=["csv"], accept_multiple_files=True, key="csv_c")
        with col_c: files_xlsx_originales = st.file_uploader("📁 3. Excels Originales", type=["xlsx"], accept_multiple_files=True, key="xls_c")
            
        if file_argos and files_csv_finales and files_xlsx_originales:
            if st.button("🚀 PROCESAR Y GENERAR EXCELS CON NRC", type="primary"):
                try:
                    argos_df = pd.read_csv(file_argos, encoding="utf-8", on_bad_lines='skip', dtype=str)
                    argos_df.columns = [re.sub(r'\.+', '.', str(c).replace('"', '').replace("'", "").strip()) for c in argos_df.columns]
                    
                    # 🔥 Búsqueda exacta de las 3 columnas en ARGOS
                    col_argos_cluster = next((c for c in argos_df.columns if "cluster" in normalizar_para_busqueda_t3(c)), None)
                    col_argos_area = next((c for c in argos_df.columns if "area" in normalizar_para_busqueda_t3(c)), None)
                    col_argos_curso = next((c for c in argos_df.columns if "curso" in normalizar_para_busqueda_t3(c)), None)

                    if not col_argos_cluster: raise KeyError("No se encontró la columna de Cluster en ARGOS.")
                    if not col_argos_area: raise KeyError("No se encontró la columna de Área en ARGOS.")
                    if not col_argos_curso: raise KeyError("No se encontró la columna de Curso en ARGOS.")

                    # 🔥 APLICAMOS LA ULTRA-LIMPIEZA A ARGOS
                    argos_df["Periodo"] = argos_df["Periodo"].apply(ultra_limpiar)
                    argos_df["Nivel"] = argos_df["Nivel"].apply(ultra_limpiar)
                    argos_df[col_argos_cluster] = argos_df[col_argos_cluster].apply(ultra_limpiar)
                    argos_df[col_argos_area] = argos_df[col_argos_area].apply(ultra_limpiar)
                    argos_df[col_argos_curso] = argos_df[col_argos_curso].apply(ultra_limpiar)
                    argos_df["Grupo"] = argos_df["Grupo"].apply(ultra_limpiar_seccion)
                    
                    # LLAVE DE ARGOS INCLUYENDO CLÚSTER Y ÁREA
                    argos_df["_llave_argos"] = (
                        argos_df["Periodo"] + "_" + 
                        argos_df["Nivel"] + "_" + 
                        argos_df[col_argos_cluster] + "_" + 
                        argos_df[col_argos_area] + "_" + 
                        argos_df[col_argos_curso] + "_" + 
                        argos_df["Grupo"]
                    )
                    argos_df = argos_df.drop_duplicates(subset=["_llave_argos"])
                    mapa_nrcs = dict(zip(argos_df["_llave_argos"], argos_df["NRC"]))
                    llaves_argos_disponibles = list(mapa_nrcs.keys())

                    excels_inyectados_zip = io.BytesIO()
                    archivos_procesados_con_exito = 0
                    alertas_dimensiones, alertas_parejas = [], []
                    alertas_nrc_faltantes = []
                    
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
                                    
                                    nuevas_columnas = []
                                    for col in df_excel_original.columns:
                                        huella = normalizar_para_busqueda_t3(col)
                                        if huella in mapa_huellas_t3:
                                            nuevas_columnas.append(mapa_huellas_t3[huella])
                                        else:
                                            nuevas_columnas.append(col)
                                    df_excel_original.columns = nuevas_columnas
                                    
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
                                    
                                    # Extraer Clúster y Nivel corregido
                                    cluster_csv_series = df_csv["datocomplementario"] if "datocomplementario" in df_csv.columns else pd.Series([""] * len(df_csv))
                                    nivel_corregido = cluster_csv_series.apply(corregir_nivel_por_cluster_csv).apply(ultra_limpiar)
                                    cluster_limpio_csv = cluster_csv_series.apply(ultra_limpiar)
                                    
                                    # LLAVE DE CRUCE DESDE EL CSV INCLUYENDO CLÚSTER Y ÁREA (Subject)
                                    llaves_cruce = (
                                        df_nrc_pestana["Periodo"].apply(ultra_limpiar) + "_" + 
                                        nivel_corregido + "_" + 
                                        cluster_limpio_csv + "_" + 
                                        df_nrc_pestana["Subject"].apply(ultra_limpiar) + "_" + 
                                        df_nrc_pestana["Course"].apply(ultra_limpiar) + "_" + 
                                        df_nrc_pestana["Sección"].apply(ultra_limpiar_seccion)
                                    )
                                    
                                    nrc_mapeados = llaves_cruce.map(mapa_nrcs)
                                    
                                    faltantes = llaves_cruce[nrc_mapeados.isna()]
                                    if not faltantes.empty:
                                        for llave_rota in faltantes.unique():
                                            sugerencias = difflib.get_close_matches(str(llave_rota), llaves_argos_disponibles, n=1, cutoff=0.5)
                                            sugerencia_txt = f"👉 En ARGOS lo más parecido es: **{sugerencias[0]}**" if sugerencias else "👉 (No se encontró nada parecido en ARGOS)"
                                            alertas_nrc_faltantes.append(f"❌ `{fx.name}` buscó: **{llave_rota}** \n{sugerencia_txt}")

                                    df_nrc_pestana.insert(0, "NRC", nrc_mapeados)
                                    
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
                                    
                            else:
                                alertas_parejas.append(f"⚠️ `{fx.name}` no encontró ningún CSV compatible.")

                    if archivos_procesados_con_exito > 0:
                        st.session_state.final_argos_zip = excels_inyectados_zip.getvalue()
                        st.success(f"🎉 ¡Proceso finalizado! Se procesaron {archivos_procesados_con_exito} archivos exitosamente.")
                    else: 
                        st.error("❌ No se pudo procesar ningún archivo.")
                    
                    if alertas_nrc_faltantes:
                        st.markdown("### 🔍 Radar de Llaves Rotas (Comparación Frente a Frente):")
                        for alerta in alertas_nrc_faltantes: st.error(alerta)
                            
                    if alertas_dimensiones:
                        st.markdown("### 🚫 Archivos descartados por diferencia de filas:")
                        for alerta in alertas_dimensiones: st.error(alerta)
                    if alertas_parejas:
                        st.markdown("### ❓ Archivos sin pareja:")
                        for alerta in alertas_parejas: st.warning(alerta)

                except Exception as e:
                    st.error(f"❌ Ocurrió un inconveniente crítico: {str(e)}")

        if st.session_state.final_argos_zip is not None:
            st.markdown("### 📥 Panel de Descarga (Excels Inyectados)")
            st.download_button(
                label="📁 📥 DESCARGAR EXCELS CON NRC (.ZIP)", data=st.session_state.final_argos_zip,
                file_name="Excels_Finales_con_NRC.zip", mime="application/zip",
                use_container_width=True, type="primary"
            )

    # ====================================================================
    # MODO B: RÁPIDO (SOLO CSV FINAL + ARGOS)
    # ====================================================================
    elif modo_inyeccion == "⚡ Rápido (Tengo ARGOS y CSV Final)":
        st.markdown("Extrae los NRC de ARGOS cruzando con el CSV final y genera una tabla limpia que puedes copiar o descargar en Excel.")
        
        col_r1, col_r2 = st.columns(2)
        with col_r1: file_argos_rap = st.file_uploader("📊 1. Reporte ARGOS (.csv)", type=["csv"], key="arg_r")
        with col_r2: files_csv_rap = st.file_uploader("📝 2. CSVs Finales", type=["csv"], accept_multiple_files=True, key="csv_r")
        
        if file_argos_rap and files_csv_rap:
            if st.button("⚡ Cruzar NRC y Generar Tabla", type="primary"):
                try:
                    # 1. Preparar Diccionario de ARGOS
                    argos_df = pd.read_csv(file_argos_rap, encoding="utf-8", on_bad_lines='skip', dtype=str)
                    argos_df.columns = [re.sub(r'\.+', '.', str(c).replace('"', '').replace("'", "").strip()) for c in argos_df.columns]
                    
                    # 🔥 Búsqueda exacta de las 3 columnas en ARGOS
                    col_argos_cluster = next((c for c in argos_df.columns if "cluster" in normalizar_para_busqueda_t3(c)), None)
                    col_argos_area = next((c for c in argos_df.columns if "area" in normalizar_para_busqueda_t3(c)), None)
                    col_argos_curso = next((c for c in argos_df.columns if "curso" in normalizar_para_busqueda_t3(c)), None)

                    if not col_argos_cluster: raise KeyError("No se encontró la columna de Cluster en ARGOS.")
                    if not col_argos_area: raise KeyError("No se encontró la columna de Área en ARGOS.")
                    if not col_argos_curso: raise KeyError("No se encontró la columna de Curso en ARGOS.")

                    argos_df["Periodo"] = argos_df["Periodo"].apply(ultra_limpiar)
                    argos_df["Nivel"] = argos_df["Nivel"].apply(ultra_limpiar)
                    argos_df[col_argos_cluster] = argos_df[col_argos_cluster].apply(ultra_limpiar)
                    argos_df[col_argos_area] = argos_df[col_argos_area].apply(ultra_limpiar)
                    argos_df[col_argos_curso] = argos_df[col_argos_curso].apply(ultra_limpiar)
                    argos_df["Grupo"] = argos_df["Grupo"].apply(ultra_limpiar_seccion)
                    
                    # LLAVE DE ARGOS INCLUYENDO CLÚSTER Y ÁREA
                    argos_df["_llave_argos"] = (
                        argos_df["Periodo"] + "_" + 
                        argos_df["Nivel"] + "_" + 
                        argos_df[col_argos_cluster] + "_" + 
                        argos_df[col_argos_area] + "_" + 
                        argos_df[col_argos_curso] + "_" + 
                        argos_df["Grupo"]
                    )
                    argos_df = argos_df.drop_duplicates(subset=["_llave_argos"])
                    mapa_nrcs = dict(zip(argos_df["_llave_argos"], argos_df["NRC"]))
                    llaves_argos_disponibles = list(mapa_nrcs.keys())
                    
                    # 2. Leer CSVs y armar tabla final
                    dfs_combinados = []
                    alertas_nrc_rapido = []
                    
                    for fc in files_csv_rap:
                        df_c = pd.read_csv(io.BytesIO(fc.getvalue()), encoding="utf-8", dtype=str)
                        df_c = df_c.dropna(how='all')
                        
                        cluster_csv_series = df_c["datocomplementario"] if "datocomplementario" in df_c.columns else pd.Series([""] * len(df_c))
                        nivel_csv = cluster_csv_series.apply(corregir_nivel_por_cluster_csv).apply(ultra_limpiar)
                        cluster_limpio_csv = cluster_csv_series.apply(ultra_limpiar)
                        
                        # LLAVE DE CRUCE DESDE EL CSV INCLUYENDO CLÚSTER Y ÁREA (SUBJ)
                        llaves_csv = (
                            df_c.get("PERIODO", pd.Series(dtype=str)).apply(ultra_limpiar) + "_" + 
                            nivel_csv + "_" + 
                            cluster_limpio_csv + "_" + 
                            df_c.get("SUBJ", pd.Series(dtype=str)).apply(ultra_limpiar) + "_" + 
                            df_c.get("COURSE", pd.Series(dtype=str)).apply(ultra_limpiar) + "_" + 
                            df_c.get("SECCION", pd.Series(dtype=str)).apply(ultra_limpiar_seccion)
                        )
                        
                        nrc_asignados = llaves_csv.map(mapa_nrcs)
                        
                        faltantes = llaves_csv[nrc_asignados.isna()]
                        if not faltantes.empty:
                            for llave_rota in faltantes.unique():
                                sugerencias = difflib.get_close_matches(str(llave_rota), llaves_argos_disponibles, n=1, cutoff=0.5)
                                sugerencia_txt = f"👉 En ARGOS lo más parecido es: **{sugerencias[0]}**" if sugerencias else "👉 (No se encontró nada parecido)"
                                alertas_nrc_rapido.append(f"❌ Buscó: **{llave_rota}** \n{sugerencia_txt}")
                            
                        df_c.insert(0, "NRC", nrc_asignados)
                        
                        df_c = df_c.rename(columns={
                            "TIPODEHORARIO": "TIPO DE HORARIO",
                            "METODO_EDUCATIVO": "METODO_ED"
                        })
                        
                        columnas_deseadas = ["NRC", "PERIODO", "SUBJ", "COURSE", "CAPACIDAD", "SECCION", "TIPO DE HORARIO", "METODO_ED", "datocomplementario"]
                        columnas_finales = [c for c in columnas_deseadas if c in df_c.columns]
                        
                        dfs_combinados.append(df_c[columnas_finales])
                        
                    if dfs_combinados:
                        df_resultado_rapido = pd.concat(dfs_combinados, ignore_index=True)
                        st.session_state.df_cruce_rapido = df_resultado_rapido
                        st.success("✅ ¡Tabla cruzada generada exitosamente!")
                        
                        if alertas_nrc_rapido:
                            st.warning("⚠️ Ojo: Algunas filas no encontraron su NRC. Revisa las discrepancias abajo:")
                            with st.expander("🔍 Ver llaves que no cruzaron (Frente a Frente)"):
                                for a in alertas_nrc_rapido: st.write(a)
                    else:
                        st.error("No se pudo procesar la información de los CSV.")
                        
                except Exception as e:
                    st.error(f"❌ Ocurrió un error en el cruce rápido: {str(e)}")

        if st.session_state.df_cruce_rapido is not None:
            # 👇 NUEVO: SECCIÓN DE RESULTADOS CON OPCIÓN DE COPIADO FÁCIL
            st.markdown("### 📋 Resultados del Cruce (NRC inyectados)")
            
            # Vista interactiva visual
            st.dataframe(st.session_state.df_cruce_rapido, use_container_width=True)
            
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                st.markdown("#### 📝 ¿Quieres copiar la tabla directamente?")
                st.info("Haz clic en el botón de **'Copiar'** en la esquina superior derecha del cuadro de abajo y pégalo (Ctrl+V) en tu Excel. Los datos se acomodarán solos.")
                
                # Convertimos la tabla a formato de texto separado por tabulaciones (TSV)
                # Al pegarlo en Excel, las tabulaciones separan el texto en columnas automáticamente.
                tsv_rapido = st.session_state.df_cruce_rapido.to_csv(index=False, sep='\t')
                st.code(tsv_rapido, language="text")
            
            with col_b2:
                st.markdown("#### 📥 O descárgala en formato Excel")
                st.info("Si prefieres descargar el archivo listo para usar, usa este botón:")
                excel_rapido_buffer = io.BytesIO()
                st.session_state.df_cruce_rapido.to_excel(excel_rapido_buffer, index=False)
                
                st.download_button(
                    label="📥 Descargar esta tabla (.xlsx)",
                    data=excel_rapido_buffer.getvalue(),
                    file_name="Cruce_Rapido_NRC.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
