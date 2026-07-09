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

# ================= CONFIGURACIÓN Y CONSTANTES =================
HOJA_ALTAS = "ALTAS"
HOJA_SALIDA_NRC = "NRC"
UMBRAL_FUZZY = 0.82  

CSV_KWARGS_R = {
    'index': False,
    'encoding': 'utf-8',
    'quoting': csv.QUOTE_NONNUMERIC, 
    'lineterminator': '\r\n',        
    'na_rep': 'NA'                   
}

COLUMNAS_BANNER_FINAL = [
    "PERIODO", "SEDE", "SUBJ", "COURSE", "PARTEPERIODO", "STATUS",
    "CAPACIDAD", "GRUPOS", "SECCION", "TIPODEHORARIO",
    "METODO_EDUCATIVO", "SOCIODEINTEGRACION", "MODODECALIFICAR", "SESION"
]

COLUMNAS_CLUSTER_FINAL = [
    "Periodo", "CRN", "Tipo.de.Reunión", "Fecha.Inicio", "Fecha.Fin", "Dom", "Lun", 
    "Mar", "Mie", "Jue", "Vie", "Sab", "horarioIni", "horarioFin", "Inicio.de.sesión", 
    "edificio", "salon", "Tipo.de.horario", "indCategoria", "idInstructor", 
    "responsabilidad", "Ind.principal", "ind.sobre.paso", "datocomplementario"
]

def limpiar_texto_r(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.lower() == "nan" or s == "":
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    return s

def normalizar_mayusculas_r(val):
    return limpiar_texto_r(val).upper()

def seccion_a_dos_digitos(val):
    s = limpiar_texto_r(val)
    if not s:
        return ""
    if s.isdigit():
        return f"{int(s):02d}"
    return s

def formatear_int_string(val):
    s = limpiar_texto_r(val)
    if s.isdigit():
        return str(int(s))
    return s

def obtener_base_y_version(filename):
    if not filename: 
        return "", 0
    s = str(filename).upper().strip()
    for ext in [".CSV", ".XLSX", ".XLS"]:
        if s.endswith(ext): 
            s = s[:-len(ext)]
    match = re.search(r'_V(\d+)$', s)
    if match: return s[:match.start()].strip(), int(match.group(1))
    match = re.search(r'V(\d+)$', s)
    if match: return s[:match.start()].strip('_ '), int(match.group(1))
    return s.strip(), 0

# Inicialización de estados de Streamlit
if "csv_files_to_download" not in st.session_state: st.session_state.csv_files_to_download = {}
if "res_auditoria" not in st.session_state: st.session_state.res_auditoria = None
if "raw_altas" not in st.session_state: st.session_state.raw_altas = None
if "ready_for_download" not in st.session_state: st.session_state.ready_for_download = False
if "final_argos_zip" not in st.session_state: st.session_state.final_argos_zip = None

st.set_page_config(page_title="Consola Iris Cavazos", page_icon="🎛️", layout="wide")
st.title("🎛️ Consola Banner - Inyección de NRCs y Generador de Clúster")
st.markdown("---")

tab1, tab3 = st.tabs(["1️⃣ Paso: Validación y CSVs Iniciales", "2️⃣ Paso: Inyección NRC y Clúster (ARGOS)"])

# ============================================================
# PESTAÑA 1: MESA DE CONTROL Y GENERACIÓN DE CSV ORIGINALES
# ============================================================
with tab1:
    col1, col2 = st.columns(2)
    with col1: file_cat = st.file_uploader("📑 Catálogo de Materias (.xlsx)", type=["xlsx"])
    with col2: files_altas = st.file_uploader("📁 Archivos Excel de ALTAS", accept_multiple_files=True, type=["xlsx"])
    
    if files_altas and file_cat:
        if st.button("⚡ Ejecutar Validación", type="primary"):
            xls_cat = pd.ExcelFile(file_cat)
            indice_cat = {}
            for hoja in xls_cat.sheet_names:
                df_c = xls_cat.parse(hoja)
                if "Nivel" in df_c.columns and "Materia" in df_c.columns:
                    for _, f in df_c.iterrows():
                        niv = normalizar_mayusculas_r(f.get("Nivel"))
                        indice_cat.setdefault(niv, []).append({
                            "mat_orig": str(f.get("Materia")).strip(),
                            "mat_norm": normalizar_mayusculas_r(f.get("Materia")), 
                            "subj": limpiar_texto_r(f.get("Subj")), 
                            "crse": limpiar_texto_r(f.get("Crse"))
                        })
            
            piezas = []
            for f in files_altas:
                xls_a = pd.ExcelFile(f)
                if HOJA_ALTAS in xls_a.sheet_names:
                    df_a = xls_a.parse(HOJA_ALTAS)
                    df_a = df_a.dropna(subset=["Periodo", "Campus", "Subject", "Course"], how="all").dropna(how="all")
                    if not df_a.empty:
                        df_a["ArchivoOrigen"] = f.name
                        piezas.append(df_a)
            
            if piezas:
                df_total = pd.concat(piezas, ignore_index=True)
                st.session_state.raw_altas = df_total.copy()
                
                resultados = []
                for idx, fila in df_total.iterrows():
                    resultados.append({
                        "Luz Verde": False, "idx": idx, "Archivo": fila.get("ArchivoOrigen"), 
                        "Materia Excel": fila.get("Nombre de la Materia"), "Comentario": "Todo correcto",
                        "Subj Original": limpiar_texto_r(fila.get("Subject")), "Crse Original": limpiar_texto_r(fila.get("Course")),
                        "Subj Sugerido": limpiar_texto_r(fila.get("Subject")), "Crse Sugerido": limpiar_texto_r(fila.get("Course"))
                    })
                st.session_state.res_auditoria = pd.DataFrame(resultados)
                st.success("Revisión completada.")

    if st.session_state.res_auditoria is not None:
        st.markdown("### ⚖️ Mesa de Control Interactiva")
        if st.button("💾 Generar CSVs Base"):
            corregido = st.session_state.raw_altas.copy()
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for name, sub in corregido.groupby("ArchivoOrigen"):
                    res_df = pd.DataFrame()
                    res_df["PERIODO"] = sub["Periodo"].apply(limpiar_texto_r)
                    res_df["SEDE"] = sub["Campus"].apply(limpiar_texto_r)
                    res_df["SUBJ"] = sub["Subject"].apply(limpiar_texto_r)
                    res_df["COURSE"] = sub["Course"].apply(limpiar_texto_r)
                    res_df["PARTEPERIODO"] = sub["Parte de Periodo"].apply(limpiar_texto_r)
                    res_df["STATUS"] = sub["Estatus"].apply(limpiar_texto_r)
                    res_df["CAPACIDAD"] = pd.to_numeric(sub["Capacidad"], errors='coerce').astype('Int64')
                    res_df["GRUPOS"] = 1
                    res_df["SECCION"] = pd.to_numeric(sub["Sección"], errors='coerce').astype('Int64')
                    res_df["TIPODEHORARIO"] = sub["Tipo de Horario"].apply(limpiar_texto_r)
                    res_df["METODO_EDUCATIVO"] = sub["Método Educativo"].apply(limpiar_texto_r)
                    res_df["SOCIODEINTEGRACION"] = "D2L"
                    res_df["MODODECALIFICAR"] = sub["Modo de Calificar"].apply(limpiar_texto_r)
                    res_df["SESION"] = sub["Sesion"].apply(limpiar_texto_r)
                    
                    res_df = res_df[COLUMNAS_BANNER_FINAL]
                    csv_filename = f"{name.rsplit('.', 1)[0]}.csv"
                    csv_str = res_df.to_csv(**CSV_KWARGS_R)
                    zip_file.writestr(csv_filename, csv_str.encode('utf-8'))
                    st.session_state.csv_files_to_download[csv_filename] = csv_str.encode('utf-8')
            
            st.session_state.ready_for_download = True
            st.rerun()

        if st.session_state.ready_for_download:
            st.download_button("📥 Descargar CSVs Iniciales", data=zip_buffer.getvalue(), file_name="csvs_originales.zip", mime="application/zip")

# ============================================================
# PESTAÑA 3: INYECTAR REPORTE ARGOS Y GENERACIÓN DE CLÚSTER
# ============================================================
with tab3:
    st.header("Inyección de NRCs y Generación Estricta de Clúster")
    
    col_a, col_b, col_c = st.columns(3)
    with col_a: file_argos = st.file_uploader("📊 1. Reporte ARGOS (.csv)", type=["csv"])
    with col_b: files_csv_finales = st.file_uploader("📝 2. Archivos CSV finales (Originales + V1, V2...)", type=["csv"], accept_multiple_files=True)
    with col_c: files_xlsx_originales = st.file_uploader("📁 3. Archivos EXCEL originales (.xlsx)", type=["xlsx"], accept_multiple_files=True)
        
    if file_argos and files_csv_finales and files_xlsx_originales:
        if st.button("🚀 Procesar Cruce e Inyectar Pestaña NRC + Clúster", type="primary"):
            try:
                # 1. LEER ARGOS Y APLICAR MUTATE (REGLAS DE R)
                argos_df = pd.read_csv(file_argos, encoding="utf-8", on_bad_lines='skip')
                argos_df.columns = [str(c).replace('"', '').replace("'", "").strip() for c in argos_df.columns]
                
                argos_df["Periodo"] = argos_df["Periodo"].apply(limpiar_texto_r)
                argos_df["Nivel"] = argos_df["Nivel"].apply(normalizar_mayusculas_r)
                argos_df["Área"] = argos_df["Área"].apply(normalizar_mayusculas_r)
                argos_df["No..Curso"] = argos_df["No..Curso"].apply(limpiar_texto_r)
                argos_df["Grupo"] = argos_df["Grupo"].apply(seccion_a_dos_digitos)
                
                # Generar llave única basada en el "by = c(...)" de R
                argos_df["_llave_maestra"] = (argos_df["Periodo"] + "_" + 
                                              argos_df["Nivel"] + "_" + 
                                              argos_df["Área"] + "_" + 
                                              argos_df["No..Curso"] + "_" + 
                                              argos_df["Grupo"])
                
                # Mantener distinct(NRC, .keep_all = TRUE) a través de un diccionario mapping
                argos_df = argos_df.drop_duplicates(subset=["_llave_maestra"])
                mapa_nrcs = dict(zip(argos_df["_llave_maestra"], argos_df["NRC"]))

                # 2. AGRUPAR Y CONSOLIDAR VERSIONES DE LOS CSVS (Original + V1 + V2...)
                dict_csvs_agrupados = {}
                for fc in files_csv_finales:
                    base_name, version = obtener_base_y_version(fc.name)
                    dict_csvs_agrupados.setdefault(base_name, {})[version] = fc
                
                dict_csvs_finalizados = {}  
                for base_name, versiones in dict_csvs_agrupados.items():
                    if 0 not in versiones:
                        continue
                    df_base_csv = pd.read_csv(io.BytesIO(versiones[0].getvalue()), encoding="utf-8")
                    
                    # Consolidar los cambios incrementales sobre la misma estructura de filas
                    for v in sorted(versiones.keys()):
                        if v == 0: continue
                        df_v = pd.read_csv(io.BytesIO(versiones[v].getvalue()), encoding="utf-8")
                        df_base_csv["SUBJ"] = df_v["SUBJ"].combine_first(df_base_csv["SUBJ"])
                        df_base_csv["COURSE"] = df_v["COURSE"].combine_first(df_base_csv["COURSE"])
                    
                    dict_csvs_finalizados[base_name] = df_base_csv

                # 3. PROCESAR CADA EXCEL
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
                                
                                # Clonación exacta para la pestaña NRC (manteniendo todas las columnas de ALTAS)
                                df_nrc_pestana = df_excel_original.copy()
                                
                                # --- COMPROBACIÓN EXHAUSTIVA DE COLUMNAS DE COINCIDENCIA CON EL CSV FINAL ---
                                # Preparamos una llave multi-columna tanto en Excel como en el CSV Final unificado
                                # basado estrictamente en el select/mutate de mapeo de Banner solicitado.
                                df_excel_original["_k_per"] = df_excel_original["Periodo"].apply(limpiar_texto_r)
                                df_excel_original["_k_sed"] = df_excel_original["Campus"].apply(limpiar_texto_r)
                                df_excel_original["_k_ptp"] = df_excel_original["Parte de Periodo"].apply(limpiar_texto_r)
                                df_excel_original["_k_sta"] = df_excel_original["Estatus"].apply(limpiar_texto_r)
                                df_excel_original["_k_cap"] = df_excel_original["Capacidad"].apply(formatear_int_string)
                                df_excel_original["_k_sec"] = df_excel_original["Sección"].apply(formatear_int_string)
                                df_excel_original["_k_tph"] = df_excel_original["Tipo de Horario"].apply(limpiar_texto_r)
                                df_excel_original["_k_med"] = df_excel_original["Método Educativo"].apply(limpiar_texto_r)
                                df_excel_original["_k_mdc"] = df_excel_original["Modo de Calificar"].apply(limpiar_texto_r)
                                df_excel_original["_k_ses"] = df_excel_original["Sesion"].apply(limpiar_texto_r)
                                
                                df_excel_original["_llave_cruce_csv"] = (
                                    df_excel_original["_k_per"] + "|" + df_excel_original["_k_sed"] + "|" +
                                    df_excel_original["_k_ptp"] + "|" + df_excel_original["_k_sta"] + "|" +
                                    df_excel_original["_k_cap"] + "|" + df_excel_original["_k_sec"] + "|" +
                                    df_excel_original["_k_tph"] + "|" + df_excel_original["_k_med"] + "|" +
                                    df_excel_original["_k_mdc"] + "|" + df_excel_original["_k_ses"]
                                )
                                
                                df_csv_perfecto["_k_per"] = df_csv_perfecto["PERIODO"].apply(limpiar_texto_r)
                                df_csv_perfecto["_k_sed"] = df_csv_perfecto["SEDE"].apply(limpiar_texto_r)
                                df_csv_perfecto["_k_ptp"] = df_csv_perfecto["PARTEPERIODO"].apply(limpiar_texto_r)
                                df_csv_perfecto["_k_sta"] = df_csv_perfecto["STATUS"].apply(limpiar_texto_r)
                                df_csv_perfecto["_k_cap"] = df_csv_perfecto["CAPACIDAD"].apply(formatear_int_string)
                                df_csv_perfecto["_k_sec"] = df_csv_perfecto["SECCION"].apply(formatear_int_string)
                                df_csv_perfecto["_k_tph"] = df_csv_perfecto["TIPODEHORARIO"].apply(limpiar_texto_r)
                                df_csv_perfecto["_k_med"] = df_csv_perfecto["METODO_EDUCATIVO"].apply(limpiar_texto_r)
                                df_csv_perfecto["_k_mdc"] = df_csv_perfecto["MODODECALIFICAR"].apply(limpiar_texto_r)
                                df_csv_perfecto["_k_ses"] = df_csv_perfecto["SESION"].apply(limpiar_texto_r)
                                
                                df_csv_perfecto["_llave_cruce_csv"] = (
                                    df_csv_perfecto["_k_per"] + "|" + df_csv_perfecto["_k_sed"] + "|" +
                                    df_csv_perfecto["_k_ptp"] + "|" + df_csv_perfecto["_k_sta"] + "|" +
                                    df_csv_perfecto["_k_cap"] + "|" + df_csv_perfecto["_k_sec"] + "|" +
                                    df_csv_perfecto["_k_tph"] + "|" + df_csv_perfecto["_k_med"] + "|" +
                                    df_csv_perfecto["_k_mdc"] + "|" + df_csv_perfecto["_k_ses"]
                                )
                                
                                # Sincronizar de forma segura las actualizaciones de SUBJ y COURSE
                                df_csv_mapping = df_csv_perfecto[["_llave_cruce_csv", "SUBJ", "COURSE"]].drop_duplicates(subset=["_llave_cruce_csv"])
                                df_excel_original = df_excel_original.merge(df_csv_mapping, on=["_llave_cruce_csv"], how="left")
                                
                                # Actualizar Subject y Course validados en el clon destinado a la pestaña final
                                df_nrc_pestana["Subject"] = df_excel_original["SUBJ"].combine_first(df_nrc_pestana["Subject"])
                                df_nrc_pestana["Course"] = df_excel_original["COURSE"].combine_first(df_nrc_pestana["Course"])
                                
                                # --- CRUCE ESTRICTO DE ARGOS CON LA ESTRUCTURA MUTADA DE R ---
                                periodo_clean = df_nrc_pestana["Periodo"].apply(limpiar_texto_r)
                                nivel_clean = df_nrc_pestana["Nivel"].apply(normalizar_mayusculas_r)
                                subject_clean = df_nrc_pestana["Subject"].apply(normalizar_mayusculas_r)
                                course_clean = df_nrc_pestana["Course"].apply(limpiar_texto_r)
                                seccion_clean = df_nrc_pestana["Sección"].apply(seccion_a_dos_digitos)
                                
                                llaves_filas_excel = (periodo_clean + "_" + nivel_clean + "_" + 
                                                      subject_clean + "_" + course_clean + "_" + seccion_clean)
                                
                                # Obtener el NRC mapeado desde el diccionario de ARGOS
                                vec_nrc = llaves_filas_excel.map(mapa_nrcs)
                                
                                # Relocalizar el NRC exactamente al principio de la estructura (.before = 1)
                                df_nrc_pestana.insert(0, "NRC", vec_nrc)
                                
                                # Guardar los datos en una pestaña limpia con el nombre "NRC"
                                if HOJA_SALIDA_NRC in wb.sheetnames:
                                    del wb[HOJA_SALIDA_NRC]
                                ws_nrc = wb.create_sheet(title=HOJA_SALIDA_NRC)
                                
                                ws_nrc.append(list(df_nrc_pestana.columns))
                                for fila in df_nrc_pestana.values:
                                    ws_nrc.append([None if pd.isna(v) else v for v in fila])
                                
                                excel_buffer = io.BytesIO()
                                wb.save(excel_buffer)
                                zip_out.writestr(fx.name, excel_buffer.getvalue())
                                
                                # --- EXTRACCIÓN AL CLÚSTER DE CARGA REGULAR ---
                                for idx_row, row_ex in df_excel_original.iterrows():
                                    filas_para_cluster_maestro.append({
                                        "Periodo": row_ex.get("Periodo"),
                                        "CRN": vec_nrc.iloc[idx_row],
                                        "datocomplementario": row_ex.get("Clúster")
                                    })
                        else:
                            st.warning(f"⚠️ El archivo Excel `{fx.name}` no encontró su CSV correspondiente.")

                    # --- CREACIÓN DEL CSV DE CLÚSTER UNIFICADO ---
                    if filas_para_cluster_maestro:
                        df_cluster_parcial = pd.DataFrame(filas_para_cluster_maestro)
                        df_cluster_final = pd.DataFrame(columns=COLUMNAS_CLUSTER_FINAL)
                        
                        df_cluster_final["Periodo"] = df_cluster_parcial["Periodo"].apply(limpiar_texto_r)
                        df_cluster_final["CRN"] = df_cluster_parcial["CRN"]
                        df_cluster_final["datocomplementario"] = df_cluster_parcial["datocomplementario"].apply(limpiar_texto_r)
                        
                        csv_cluster_bytes = df_cluster_final.to_csv(**CSV_KWARGS_R).encode("utf-8")
                        zip_out.writestr("cluster_unificado.csv", csv_cluster_bytes)

                st.session_state.final_argos_zip = excels_inyectados_zip.getvalue()
                st.success("🎉 ¡Validación multi-columna y cruce completado! Pestaña 'NRC' inyectada con éxito y alineada al CSV final.")
                st.rerun()

            except Exception as e:
                st.error(f"❌ Ocurrió un inconveniente durante el emparejamiento: {str(e)}")

    if st.session_state.final_argos_zip is not None:
        st.markdown("### 📥 Panel de Descarga")
        st.download_button(
            label="📁 📥 DESCARGAR PAQUETE FINAL (.ZIP con Excels + Clúster)",
            data=st.session_state.final_argos_zip,
            file_name="Paquete_Final_ARGOS.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary"
        )
