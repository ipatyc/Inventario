import streamlit as st
import pandas as pd
import io
import re
import unicodedata
from difflib import SequenceMatcher

# ================= CONFIGURACIÓN Y FUNCIONES (¡ESTO DEBE IR ARRIBA!) =================
HOJA_ALTAS = "ALTAS"
UMBRAL_FUZZY = 0.72

def normalizer(t):
    if pd.isna(t) or t is None: return ""
    return "".join(c for c in unicodedata.normalize("NFD", str(t).upper().strip()) if unicodedata.category(c) != "Mn")

def similitud(a, b): 
    return SequenceMatcher(None, a, b).ratio()

def _sec_pad(v):
    s = str(v).strip()
    if s == "" or str(s).lower() == "nan": return ""
    try: return str(int(float(s))).zfill(2)
    except: return s
# ===================================================================================

# Después de este bloque ya puede seguir el resto de tu código (st.set_page_config, st.title, etc.)
# Configuración de la página de Streamlit
st.set_page_config(page_title="Consola Iris Cavazos", page_icon="🎛️", layout="wide")

st.title("🎛️ Consola de Control de Materias e Inyección de NRCs")
st.markdown("---")

# ================= CONFIGURACIÓN Y FUNCIONES =================
HOJA_ALTAS = "ALTAS"  # Pestaña configurada en plural
UMBRAL_FUZZY = 0.72

def normalizar(t):
    if pd.isna(t) or t is None: return ""
    return "".join(c for c in unicodedata.normalize("NFD", str(t).upper().strip()) if unicodedata.category(c) != "Mn")

def similitud(a, b): 
    return SequenceMatcher(None, a, b).ratio()

# Pad para las secciones (ej. 1 -> '01')
def _sec_pad(v):
    s = str(v).strip()
    if s == "" or str(s).lower() == "nan": return ""
    try: return str(int(float(s))).zfill(2)
    except: return s

# ================= ESTADOS EN MEMORIA (SESSION STATE) =================
if "df_corregido" not in st.session_state: st.session_state.df_corregido = None
if "raw_altas" not in st.session_state: st.session_state.raw_altas = None
if "res_auditoria" not in st.session_state: st.session_state.res_auditoria = None

# Crear pestañas para separar los dos procesos
tab1, tab2 = st.tabs(["1️⃣ Proceso: Validación e Interactividad", "2️⃣ Proceso: Inyección de NRCs (ARGOS)"])

# ============================================================
# PESTAÑA 1: VALIDACIÓN E INTERACTIVIDAD
# ============================================================
with tab1:
    st.header("Validación de Claves de Materias")
    
    # Subida de archivos
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        file_cat = st.file_uploader("📑 Inventario de Materias (Catálogo - Excel)", type=["xlsx"])
    with col_up2:
        files_altas = st.file_uploader("📁 Archivos Excel de ALTAS (Puedes subir varios juntos)", accept_multiple_files=True, type=["xlsx"])
    
    if files_altas and file_cat:
        if st.button("⚡ Ejecutar Validación Inteligente", type="primary"):
            with st.spinner("Procesando catálogos y analizando archivos de Altas..."):
                # 1. Leer e indexar Catálogo
                xls_cat = pd.ExcelFile(file_cat)
                indice_cat = {}
                for hoja in xls_cat.sheet_names:
                    df_c = xls_cat.parse(hoja)
                    if "Nivel" in df_c.columns and "Materia" in df_c.columns:
                        for _, f in df_c.iterrows():
                            niv = normalizer(f.get("Nivel"))
                            indice_cat.setdefault(niv, []).append({
                                "mat_norm": normalizer(f.get("Materia")), "mat_orig": f.get("Materia"),
                                "subj": f.get("Subj"), "crse": f.get("Crse")
                            })
                
                # 2. Leer archivos de Altas buscando la pestaña "ALTAS"
                piezas = []
                for f in files_altas:
                    xls_a = pd.ExcelFile(f)
                    hojas_reales = [h for h in xls_a.sheet_names if h.strip().upper() == HOJA_ALTAS.upper()]
                    if hojas_reales:
                        df_a = xls_a.parse(hojas_reales[0])
                        df_a["ArchivoOrigen"] = f.name
                        piezas.append(df_a)
                
                if piezas:
                    df_total = pd.concat(piezas, ignore_index=True)
                    st.session_state.raw_altas = df_total.copy()
                    
                    # 3. Evaluar Fila por Fila (Auditoría)
                    resultados = []
                    for idx, fila in df_total.iterrows():
                        niv_n = normalizer(fila.get("Nivel"))
                        mat_n = normalizer(fila.get("Nombre de la Materia"))
                        candidatos = indice_cat.get(niv_n, [])
                        
                        match = next((c for c in candidatos if c["mat_norm"] == mat_n), None)
                        tipo, score = "exacto", 1.0
                        
                        if not match:
                            mejor, mejor_s = None, -1.0
                            for c in candidatos:
                                s = similitud(mat_n, c["mat_norm"])
                                if s > mejor_s: mejor_s, mejor = s, c
                            if mejor and mejor_s >= UMBRAL_FUZZY:
                                match, tipo, score = mejor, "fuzzy", mejor_s
                            else:
                                tipo, score = "no_encontrado", max(mejor_s, 0.0)
                        
                        resultados.append({
                            "Luz Verde": True if tipo == "fuzzy" else False, # Auto-marcar sugerencias fuzzy
                            "idx": idx, 
                            "Archivo Origen": fila.get("ArchivoOrigen"), 
                            "Nivel": fila.get("Nivel"),
                            "Materia Original": fila.get("Nombre de la Materia"), 
                            "Dictamen": "Todo correcto" if tipo == "exacto" else "Revisar/Fuzzy Match" if tipo == "fuzzy" else "No encontrado",
                            "Subj Original": fila.get("Subject"), 
                            "Crse Original": fila.get("Course"),
                            "Subj Sugerido": match["subj"] if match else None, 
                            "Crse Sugerido": match["crse"] if match else None
                        })
                    
                    st.session_state.res_auditoria = pd.DataFrame(resultados)
                    st.success("¡Análisis completado con éxito!")
                else:
                    st.error(f"❌ Ninguno de los archivos subidos contiene la pestaña '{HOJA_ALTAS}'.")

    # Mostrar Mesa de Control Interactiva si ya hay auditoría
    if st.session_state.res_auditoria is not None:
        st.markdown("### ⚖️ Mesa de Control Interactiva")
        st.info("A continuación se muestran las incidencias. Las filas con 'Fuzzy Match' tienen la casilla 'Luz Verde' activa por defecto. Puedes desmarcar o modificar lo que consideres.")
        
        df_aud = st.session_state.res_auditoria
        
        # Filtro rápido en pantalla
        filtro = st.radio("Filtrar vista:", ["Solo Incidencias (Fuzzy / No Encontradas)", "Mostrar Todo (Incluyendo Perfectas)"], horizontal=True)
        if filtro == "Solo Incidencias (Fuzzy / No Encontradas)":
            df_visible = df_aud[df_aud["Dictamen"] != "Todo correcto"]
        else:
            df_visible = df_aud

        # Tabla Interactiva Avanzada de Streamlit
        df_editado = st.data_editor(
            df_visible,
            hide_index=True,
            disabled=["idx", "Archivo Origen", "Nivel", "Materia Original", "Dictamen", "Subj Original", "Crse Original"],
            column_config={
                "Luz Verde": st.column_config.CheckboxColumn("Luz Verde", help="Dale luz verde para aplicar el Subj/Crse sugerido"),
                "Subj Sugerido": st.column_config.TextColumn("Subj Sugerido"),
                "Crse Sugerido": st.column_config.TextColumn("Crse Sugerido")
            },
            use_container_width=True
        )
        
        # Sincronizar cambios del editor al estado global
        for _, row in df_editado.iterrows():
            st.session_state.res_auditoria.loc[st.session_state.res_auditoria["idx"] == row["idx"], "Luz Verde"] = row["Luz Verde"]
            st.session_state.res_auditoria.loc[st.session_state.res_auditoria["idx"] == row["idx"], "Subj Sugerido"] = row["Subj Sugerido"]
            st.session_state.res_auditoria.loc[st.session_state.res_auditoria["idx"] == row["idx"], "Crse Sugerido"] = row["Crse Sugerido"]

        if st.button("💾 Aplicar Luz Verde y Generar CSVs SQL", type="primary"):
            corregido = st.session_state.raw_altas.copy()
            df_final_aud = st.session_state.res_auditoria
            
            # Aplicar correcciones aprobadas
            for _, row in df_final_aud.iterrows():
                if row["Luz Verde"] and pd.notna(row["Subj Sugerido"]):
                    corregido.loc[row["idx"], "Subject"] = row["Subj Sugerido"]
                    corregido.loc[row["idx"], "Course"] = row["Crse Sugerido"]
            
            st.session_state.df_corregido = corregido
            st.success("¡Cambios consolidados en memoria! Ya puedes descargar los archivos individuales abajo:")
            
            # Generar botones de descarga por archivo original estructurando las columnas SQL
            st.markdown("#### 📥 Archivos Listos para Descargar:")
            for name, sub in corregido.groupby("ArchivoOrigen"):
                columnas = ["Periodo","Campus","Parte de Periodo","Estatus","Capacidad",
                            "Sección","Tipo de Horario","Método Educativo","Modo de Calificar","Sesion"]
                df_csv = pd.DataFrame()
                for c in columnas:
                    df_csv[c] = sub[c] if c in sub.columns else ""
                
                df_csv.rename(columns={"Campus":"SEDE", "Parte de Periodo":"PARTEPERIODO", 
                                       "Estatus":"STATUS", "Sección":"SECCION", 
                                       "Tipo de Horario":"TIPODEHORARIO", "Método Educativo":"METODO_EDUCATIVO",
                                       "Modo de Calificar":"MODODECALIFICAR", "Sesion":"SESION"}, inplace=True)
                df_csv["PERIODO"] = df_csv["Periodo"]
                df_csv.drop(columns=["Periodo"], inplace=True)
                df_csv["SUBJ"] = sub["Subject"]
                df_csv["COURSE"] = sub["Course"]
                df_csv["GRUPOS"] = 1
                df_csv["SOCIODEINTEGRACION"] = "D2L"
                
                csv_buffer = io.StringIO()
                df_csv.to_csv(csv_buffer, index=False, encoding="utf-8-sig")
                
                fname_out = f"{name.split('.')[0]}_PROCESADO_SQL.csv"
                st.download_button(label=f"⬇️ Descargar {fname_out}", data=csv_buffer.getvalue(), file_name=fname_out, mime="text/csv")

# ============================================================
# PESTAÑA 2: INYECTAR ARGOS
# ============================================================
with tab2:
    st.header("Inyección de NRCs con Reporte de ARGOS")
    
    if st.session_state.df_corregido is None:
        st.warning("⚠️ Primero debes completar el Proceso 1 (Ejecutar validación, dar Luz Verde y Procesar) antes de poder hacer el cruce con ARGOS aquí.")
    else:
        file_argos = st.file_uploader("📊 Cargar Reporte de ARGOS (.csv)", type=["csv"])
        
        if file_argos:
            if st.button("🚀 Cruzar Información y Generar Excels Finales", type="primary"):
                try:
                    argos_df = pd.read_csv(file_argos, encoding="utf-8")
                except:
                    argos_df = pd.read_csv(file_argos, encoding="latin-1")
                
                with st.spinner("Realizando cruce inteligente de llaves primarias..."):
                    df_m = st.session_state.df_corregido.copy()
                    
                    # Crear llaves homologadas para el cruce
                    df_m["_k_p"] = df_m["Periodo"].astype(str).str.strip()
                    df_m["_k_n"] = df_m["Nivel"].apply(normalizer)
                    df_m["_k_s"] = df_m["Subject"].apply(normalizer)
                    df_m["_k_c"] = df_m["Course"].astype(str).str.strip()
                    df_m["_k_sec"] = df_m["Sección"].apply(_sec_pad)
                    
                    argos_df["_k_p"] = argos_df["Periodo"].astype(str).str.strip()
                    argos_df["_k_n"] = argos_df["Nivel"].apply(normalizer)
                    argos_df["_k_s"] = argos_df["Área"].apply(normalizer)
                    argos_df["_k_c"] = argos_df["No. Curso"].astype(str).str.strip()
                    argos_df["_k_sec"] = argos_df["Grupo"].apply(_sec_pad)
                    
                    llaves = ["_k_p", "_k_n", "_k_s", "_k_c", "_k_sec"]
                    
                    # Fusión y limpieza
                    fusion = df_m.merge(argos_df[llaves + ["NRC"]], on=llaves, how="left")
                    fusion = fusion.drop_duplicates(subset=["NRC"], keep="first")
                    fusion.drop(columns=llaves, inplace=True)
                    
                    st.success("¡Cruce completado exitosamente!")
                    st.markdown("#### 📥 Descargar Excels con pestaña NRC:")
                    
                    for name, sub in fusion.groupby("ArchivoOrigen"):
                        fname_out = f"{name.split('.')[0]}_CON_NRC.xlsx"
                        
                        out_excel = io.BytesIO()
                        with pd.ExcelWriter(out_excel, engine='openpyxl') as writer:
                            sub.drop(columns=["ArchivoOrigen"]).to_excel(writer, sheet_name="DATOS_SOLICITUD", index=False)
                            if "NRC" in sub.columns:
                                nrc_only = sub[["NRC", "Nivel", "Subject", "Course"]]
                                nrc_only.to_excel(writer, sheet_name="NRC_ASIGNADOS", index=False)
                        
                        st.download_button(label=f"⬇️ Descargar {fname_out}", data=out_excel.getvalue(), file_name=fname_out, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
