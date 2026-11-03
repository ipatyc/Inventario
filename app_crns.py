

# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import io, re, csv, zipfile, datetime, difflib, unicodedata, hashlib
import openpyxl
from collections import defaultdict
from difflib import SequenceMatcher
from openpyxl.styles import Font, PatternFill, Alignment

# ============================================================
# 1. CONFIGURACIÓN GENERAL
# ============================================================
HOJA_ALTAS, HOJA_SALIDA_NRC, UMBRAL_FUZZY = "ALTAS", "NRC", 0.82
CSV_KWARGS_R = {"index": False, "encoding": "utf-8", "sep": ",", "lineterminator": "\n"}

COLUMNAS_CSV = ["PERIODO", "SEDE", "SUBJ", "COURSE", "PARTEPERIODO", "STATUS", "CAPACIDAD",
                "GRUPOS", "SECCION", "TIPODEHORARIO", "METODO_EDUCATIVO",
                "SOCIODEINTEGRACION", "MODODECALIFICAR", "SESION", "datocomplementario"]

COLUMNAS_ALTAS = ["Periodo", "Campus", "Subject", "Course", "Nivel", "Nombre de la Materia",
                  "Parte de Periodo", "Estatus", "Capacidad", "Sección", "Tipo de Horario",
                  "Método Educativo", "Modo de Calificar", "Sesion", "Clúster", "Responsable"]

CLUSTERS_PERMITIDOS = ["Ingenieria", "Bachillerato", "Negocios", "Ciencias Exactas",
                       "Posgrado Online", "Humanidades", "Idiomas y ADN", "TJYG",
                       "Smart Cities", "Ejecutivas", "EGEL", "Intercambio",
                       "Consejeria", "Posgrado"]

COLUMNAS_CLUSTER_FINAL = [
    "Periodo", "CRN", "Tipo.de.Reunión", "Fecha.Inicio", "Fecha.Fin", "Dom", "Lun",
    "Mar", "Mie", "Jue", "Vie", "Sab", "horarioIni", "horarioFin", "Inicio.de.sesión",
    "edificio", "salon", "Tipo.de.horario", "indCategoria", "idInstructor",
    "responsabilidad", "Ind.principal", "ind.sobre.paso", "datocomplementario"
]

# ============================================================
# 2. LIMPIEZA Y NORMALIZACIÓN
# ============================================================
def limpiar_clave_texto(x):
    if x is None or pd.isna(x): return ""
    s = str(x).strip()
    if s.lower() in ("nan", "none", "nat", "<na>"): return ""
    return s[:-2] if re.fullmatch(r"\d+\.0", s) else s

def quitar_acentos(x):
    return "".join(c for c in unicodedata.normalize("NFD", limpiar_clave_texto(x))
                   if unicodedata.category(c) != "Mn")

def normalizar_para_cruce(x):
    return quitar_acentos(x).upper().strip()

def normalizar_para_busqueda(x):
    return re.sub(r"[^a-z0-9]", "", quitar_acentos(x).lower())

def normalizar_para_busqueda_t3(x):
    return normalizar_para_busqueda(x)

def limpiar_nombre_columna(x):
    return " ".join(limpiar_clave_texto(x).split())

def limpiar_espacios_y_mayusculas(x):
    return re.sub(r"\s+", " ", limpiar_clave_texto(x)).strip().upper()

def sin_espacios(x):
    return re.sub(r"\s+", "", limpiar_clave_texto(x)).upper()

def ultra_limpiar(x):
    return sin_espacios(x)

def clave_seccion(x):
    s = limpiar_clave_texto(x).upper()
    if s.isdigit() and int(s) <= 99: return f"{int(s):02d}"
    return s

def limpia_seccion_interna(x):
    return clave_seccion(x)

def ultra_limpiar_seccion(x):
    return clave_seccion(x).replace(" ", "")

def format_r_string(x):
    s = limpiar_clave_texto(x)
    return s if s else np.nan

def similitud(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def corregir_nivel_por_cluster_csv(x):
    c = normalizar_para_cruce(x)
    if "POSGRADO" in c: return "POSGRADO"
    if "BACHILLERATO" in c: return "BACHILLERATO"
    return "LICENCIATURA"

def simplificar_nombre(nombre):
    n = nombre.lower()
    for parte in [".xlsx", ".xlsm", ".xls", ".csv", "_final", "_base", "_v1", "_v2",
                  "_v3", "_v4", "corregidas_", "errores_"]:
        n = n.replace(parte, "")
    return re.sub(r"\s+", "", n)

def cluster_oficial(x):
    original = limpiar_clave_texto(x)
    return next((c for c in CLUSTERS_PERMITIDOS
                 if normalizar_para_busqueda(c) == normalizar_para_busqueda(original)), original)


# ============================================================
# 3. ENCABEZADOS REALES DEL NUEVO EXCEL ALTAS
# ============================================================

COLUMNAS_ALTAS_REALES = [
    "AÑO", "CICLO", "PERIODO", "CLUSTER", "NIVEL", "REQUERIMIENTO",
    "RESPONSABLE", "SUBJ", "COURSE", "NOMBRE_MATERIA", "SEDE",
    "PARTE_PERIODO", "STATUS", "CAPACIDAD", "SECCION", "TIPO_HORARIO",
    "METODO_EDUCATIVO", "SOCIO_INTEGRACION", "MODO_CALIFICAR",
    "SESION", "COMENTARIOS", "FECHA_REGISTRO"
]

ALIAS_COLUMNAS = {
    "ano": "Año", "ciclo": "Ciclo", "periodo": "Periodo",
    "cluster": "Clúster", "nivel": "Nivel",
    "requerimiento": "Requerimiento",
    "responsable": "Responsable", "responsables": "Responsable",
    "coordinador": "Responsable", "nombrederesponsable": "Responsable",
    "subj": "Subject", "subject": "Subject", "area": "Subject",
    "course": "Course", "crse": "Course", "nocurso": "Course",
    "nombremateria": "Nombre de la Materia",
    "nombredelamateria": "Nombre de la Materia",
    "materia": "Nombre de la Materia",
    "sede": "Campus", "campus": "Campus",
    "parteperiodo": "Parte de Periodo",
    "status": "Estatus", "estatus": "Estatus",
    "capacidad": "Capacidad", "cupo": "Capacidad",
    "seccion": "Sección", "grupo": "Sección",
    "tipohorario": "Tipo de Horario",
    "metodoeducativo": "Método Educativo",
    "sociointegracion": "Socio Integración",
    "sociodeintegracion": "Socio Integración",
    "modocalificar": "Modo de Calificar",
    "mododecalificar": "Modo de Calificar",
    "sesion": "Sesion", "comentarios": "Comentarios",
    "fecharegistro": "Fecha Registro"
}

MAPA_COLUMNAS = {normalizar_para_busqueda(c): c for c in COLUMNAS_ALTAS}
MAPA_COLUMNAS.update(ALIAS_COLUMNAS)
MAPA = MAPA_COLUMNAS

def normalizar_columnas_altas(df):
    df = df.copy()
    df.columns = [
        MAPA_COLUMNAS.get(normalizar_para_busqueda(c), limpiar_nombre_columna(c))
        for c in df.columns
    ]
    duplicadas = df.columns[df.columns.duplicated()].tolist()
    if duplicadas:
        raise ValueError(f"Columnas duplicadas después de normalizar: {duplicadas}")
    return df


# ============================================================
# 4. RESPONSABLES Y NOMBRES DE ARCHIVOS
# ============================================================
def nombre_seguro(x):
    s = re.sub(r'[\\/:*?"<>|]', "", limpiar_clave_texto(x))
    return re.sub(r"\s+", " ", s).strip()

def clave_responsable(x):
    return normalizar_para_busqueda(x)

def responsable_normalizado(x):
    s = nombre_seguro(x)
    if not s: return ""
    return " ".join(p.capitalize() for p in quitar_acentos(s).lower().split())

def responsables_unicos(valores):
    encontrados = {}
    for valor in valores:
        nombre = responsable_normalizado(valor)
        if nombre: encontrados.setdefault(clave_responsable(nombre), nombre)
    return sorted(encontrados.values(), key=lambda x: normalizar_para_busqueda(x))

def nombre_solicitud(df, extension=".xlsm"):
    # El nombre se obtiene de los registros ALTAS, nunca del nombre del archivo cargado.
    periodos = sorted({nombre_seguro(x) for x in df["Periodo"] if nombre_seguro(x)})
    responsables = responsables_unicos(df["Responsable"])
    clusters = {}
    for valor in df["Clúster"]:
        nombre = nombre_seguro(cluster_oficial(valor))
        if nombre: clusters.setdefault(normalizar_para_busqueda(nombre), nombre)
    if not periodos or not responsables or not clusters:
        raise ValueError("Falta Periodo, Responsable o Clúster para nombrar SOLICITUDES.")
    nombres_cluster = sorted(clusters.values(), key=normalizar_para_busqueda)
    unir = lambda nombres: " y ".join(nombres) if len(nombres) == 2 else ", ".join(nombres)
    return f"{'-'.join(periodos)} {unir(responsables)} {unir(nombres_cluster)} CRNs{extension}"

def nombre_csv_altas(periodo, responsable):
    periodo, responsable = nombre_seguro(periodo), responsable_normalizado(responsable)
    if not periodo or not responsable: raise ValueError("Falta Periodo o Responsable para nombrar CSV_ALTAS.")
    return f"{periodo} {responsable} CSV_ALTAS.csv"

# ============================================================
# 5. CATÁLOGO PA OPCIONAL
# ============================================================
def leer_csv_pa(archivo):
    datos = archivo.getvalue() if hasattr(archivo, "getvalue") else archivo
    df = None
    for enc in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            df = pd.read_csv(io.BytesIO(datos), dtype=str, encoding=enc, sep=None,
                             engine="python", keep_default_na=False)
            break
        except (UnicodeError, pd.errors.ParserError):
            continue
    if df is None: raise ValueError("No se pudo leer el catálogo PA.")

    df.columns = [normalizar_para_busqueda(str(c).replace("Ã", "A").replace("Ã¡", "a")) for c in df.columns]
    requeridas = ["periodo", "area", "nocurso", "grupo", "nrc"]
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes: raise ValueError(f"PA: faltan columnas {faltantes}")

    pa = df.copy()
    for c in ["periodo", "area", "nocurso", "nrc"]: pa[c] = pa[c].apply(normalizar_para_cruce)
    pa["grupo"] = pa["grupo"].apply(clave_seccion)
    pa = pa[pa["nrc"].ne("")].copy()

    llaves = ["periodo", "area", "nocurso", "grupo"]
    conflictos = pa.groupby("nrc")[llaves].nunique(dropna=False).gt(1).any(axis=1)
    nrc_conflictivos = set(conflictos[conflictos].index)

    pa_unica = pa.drop_duplicates(subset=["nrc"])
    pa_reserva = pd.concat([
        pa_unica[~pa_unica["nrc"].isin(nrc_conflictivos)],
        pa[pa["nrc"].isin(nrc_conflictivos)]
    ]).drop_duplicates(subset=["nrc"] + llaves)

    advertencias = [{"Tipo": "PA", "Detalle": f"NRC {nrc} presenta registros contradictorios."}
                    for nrc in sorted(nrc_conflictivos)]
    return pa_reserva, advertencias

# ============================================================
# 6. ASIGNACIÓN GLOBAL DE SECCIONES
# ============================================================
def asignar_secciones_globales(df_altas, pa=None):
    requeridas = ["Periodo", "Subject", "Course", "Nivel", "Sección"]
    faltantes = [c for c in requeridas if c not in df_altas.columns]
    if faltantes: raise ValueError(f"ALTAS: faltan columnas {faltantes}")

    salida = df_altas.copy().reset_index(drop=True)
    salida["Sección Original"] = salida["Sección"].apply(clave_seccion)
    salida["Sección"] = salida["Sección Original"]
    ocupadas, errores = {}, []

    if pa is not None:
        for _, fila in pa.iterrows():
            llave = tuple(normalizar_para_cruce(fila[c]) for c in ["periodo", "area", "nocurso"])
            grupo = clave_seccion(fila["grupo"])
            if all(llave) and grupo: ocupadas.setdefault(llave, set()).add(grupo)

    for i, fila in salida.iterrows():
        nivel = normalizar_para_cruce(fila["Nivel"])
        if nivel not in ("LICENCIATURA", "POSGRADO"): continue
        llave = tuple(normalizar_para_cruce(fila[c]) for c in ["Periodo", "Subject", "Course"])

        if not all(llave):
            errores.append({"Fila": i + 1, "Tipo": "Sección", "Detalle": "Periodo, SUBJ o CRSE vacío."})
            salida.at[i, "Sección"] = ""
            continue

        usadas = ocupadas.setdefault(llave, set())
        nueva = next((f"{n:02d}" for n in range(1, 100) if f"{n:02d}" not in usadas), None)
        if nueva is None:
            errores.append({"Fila": i + 1, "Tipo": "Sección", "Detalle": f"Sin secciones 01-99 para {llave}."})
            salida.at[i, "Sección"] = ""
        else:
            salida.at[i, "Sección"] = nueva
            usadas.add(nueva)

    salida["Sección Modificada"] = salida["Sección Original"] != salida["Sección"]
    return salida, pd.DataFrame(errores, columns=["Fila", "Tipo", "Detalle"])

# ============================================================
# 8. RESTRICCIONES INTERNAS
# ============================================================
# Las restricciones se validarán dentro de Streamlit.
# No se solicitará cargar el Excel RESTRICCIONES CRNs.
# Las reglas específicas pendientes deberán incorporarse aquí.

def auditar_restricciones_internas(df):
    incidencias = []
    for i, fila in df.reset_index(drop=True).iterrows():
        periodo = limpiar_clave_texto(fila.get("Periodo"))
        nivel = normalizar_para_cruce(fila.get("Nivel"))
        estatus = normalizar_para_cruce(fila.get("Estatus"))

        if not periodo: incidencias.append({"Fila": i + 1, "Tipo": "Periodo", "Detalle": "Periodo vacío."})
        if nivel not in ("LICENCIATURA", "POSGRADO", "BACHILLERATO"):
            incidencias.append({"Fila": i + 1, "Tipo": "Nivel", "Detalle": f"Nivel no reconocido: {nivel}"})
        if estatus and estatus not in ("A", "R"):
            incidencias.append({"Fila": i + 1, "Tipo": "Estatus", "Detalle": f"Estatus no reconocido: {estatus}"})

    return pd.DataFrame(incidencias, columns=["Fila", "Tipo", "Detalle"])

# ============================================================
# 9. CONSERVAR EXCEL ORIGINALES Y MACROS
# ============================================================
def crear_solicitud_excel(df_grupo, archivos_originales):
    origenes = df_grupo["ArchivoOrigen"].dropna().unique().tolist()
    if len(origenes) != 1: raise ValueError("No se pueden fusionar libros con macros de distintos orígenes.")
    origen = origenes[0]
    if origen not in archivos_originales: raise ValueError(f"No se encontró el archivo original: {origen}")

    extension = ".xlsm" if origen.lower().endswith(".xlsm") else ".xlsx"
    wb = openpyxl.load_workbook(io.BytesIO(archivos_originales[origen]), keep_vba=(extension == ".xlsm"))
    hoja = next((h for h in wb.sheetnames if h.strip().upper() == HOJA_ALTAS), None)
    if hoja is None: raise ValueError(f"{origen}: falta la hoja ALTAS.")
    ws = wb[hoja]

    encabezados = {normalizar_para_busqueda(c.value): c.column for c in ws[1] if c.value is not None}

    campos = {
        "Subject": ["subj", "subject", "area"],
        "Course": ["course", "crse", "nocurso"],
        "Sección": ["seccion", "grupo"],
        "Tipo de Horario": ["tipohorario"],
        "Método Educativo": ["metodoeducativo"],
        "Modo de Calificar": ["modocalificar", "mododecalificar"],
        "Parte de Periodo": ["parteperiodo"],
        "Estatus": ["status", "estatus"],
        "Capacidad": ["capacidad", "cupo"],
        "Sesion": ["sesion"],
        "Campus": ["sede", "campus"]
    }


    for _, fila in df_grupo.iterrows():
        numero = fila.get("_FilaExcel")
        if pd.isna(numero): raise ValueError(f"{origen}: falta _FilaExcel.")
        for campo, alternativas in campos.items():
            columna = next((encabezados[a] for a in alternativas if a in encabezados), None)
            if columna is not None and campo in fila.index:
                valor = fila[campo]
                ws.cell(int(numero), columna).value = None if pd.isna(valor) else str(valor)

    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue(), extension

# ============================================================
# 10. GENERAR ZIP SOLICITUDES + CSV_ALTAS
# ============================================================
def generar_zip_altas(df_final, archivos_originales, funcion_csv):
    requeridas = ["Periodo", "Responsable", "Clúster", "ArchivoOrigen", "_FilaExcel"]
    faltantes = [c for c in requeridas if c not in df_final.columns]
    if faltantes: raise ValueError(f"Faltan columnas para exportar: {faltantes}")

    df = df_final.copy()
    df["Periodo"] = df["Periodo"].apply(limpiar_clave_texto)
    df["Responsable"] = df["Responsable"].apply(responsable_normalizado)
    if df["Periodo"].eq("").any() or df["Responsable"].eq("").any() or df["Clúster"].apply(limpiar_clave_texto).eq("").any():
        raise ValueError("Hay ALTAS sin Periodo, Responsable o Clúster.")

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        # Un Excel por archivo original. Evitar colisiones con consecutivos.
        nombres = set()
        for origen, grupo in df.groupby("ArchivoOrigen", sort=False):
            contenido, extension = crear_solicitud_excel(grupo, archivos_originales)
            nombre = nombre_solicitud(grupo, extension)
            ruta, consecutivo = f"SOLICITUDES/{nombre}", 2
            while ruta.casefold() in nombres:
                ruta = f"SOLICITUDES/{nombre[:-len(extension)]} ({consecutivo}){extension}"
                consecutivo += 1
            z.writestr(ruta, contenido)
            nombres.add(ruta.casefold())

        # CSV_ALTAS: un CSV por periodo y responsable, aunque provenga de varios Excel.
        for (periodo, responsable), grupo in df.groupby(["Periodo", "Responsable"], sort=False):
            nombre = nombre_csv_altas(periodo, responsable)
            contenido = funcion_csv(grupo)
            if isinstance(contenido, str): contenido = contenido.encode("utf-8")
            z.writestr(f"CSV_ALTAS/{nombre}", contenido)

    return buffer.getvalue()

# ============================================================
# 11. ESTADOS DE STREAMLIT
# ============================================================
ESTADOS_INICIALES = {
    "original_files_bytes": {}, "res_auditoria": None, "raw_altas": None,
    "ready_for_download": False, "zip_file_bytes": None, "csv_files_to_download": {},
    "csv_consolidado_bytes": None, "delta_files": {}, "final_argos_zip": None,
    "df_cruce_rapido": None, "df_delta_cache": None, "nombre_delta_cache": None,
    "llave_control_archivos": "", "archivo_final_bytes": None, "archivo_final_nombre": None,
    "cat_avanzado_cache": None, "cat_avanzado_firma": None, "indice_nombres_avanzado": None,
    "cat_pa_cache": None, "advertencias_pa": [], "errores_restricciones": None,
    "df_secciones_corregidas": None, "errores_secciones": None, "secciones_validadas": False,
    "zip_solicitudes_csv_altas": None,
    "df_manual_fijo": None, "manual_candidatos": [], "manual_busqueda_realizada": False
}
for clave, valor in ESTADOS_INICIALES.items():
    if clave not in st.session_state: st.session_state[clave] = valor

# ============================================================
# 12. CONFIGURACIÓN VISUAL
# ============================================================
st.set_page_config(page_title="Consola Iris Cavazos", page_icon="⚙️", layout="wide")
st.title("⚙️ Consola de Control de Materias e Inyección de NRCs")
st.markdown("---")

tab1, tab_err, tab3 = st.tabs([
    "1️⃣ Proceso: Validación y Generar CSVs",
    "⚠️ Reporte de Errores (Extraer Delta)",
    "2️⃣ Proceso: Inyección de NRCs Masiva (ARGOS)"
])

# ============================================================
# PESTAÑA 1: VALIDACIÓN, CONFIGURACIONES Y GENERACIÓN
# ============================================================
with tab1:
    st.header("⚙️ Validación de ALTAS y Generación de Archivos")

    # 1. REINICIAR
    if st.button("🔄 Limpiar / Recomenzar", key="limpiar_t1"):
        claves = ["cat_ext_t1", "pa_t1", "altas_t1", "raw_altas", "res_auditoria",
                  "cat_pa_cache", "advertencias_pa", "errores_restricciones",
                  "df_secciones_corregidas", "errores_secciones", "zip_solicitudes_csv_altas",
                  "original_files_bytes", "cat_avanzado_cache", "cat_avanzado_firma",
                  "df_manual_fijo", "manual_candidatos", "manual_busqueda_realizada",
                  "confirmar_errores_t1", "t1_pendientes_alta", "t1_confirmar_altas",
                  "t1_firma_archivos", "t1_catalogo", "t1_version_editor"]
        claves += [k for k in st.session_state if k.startswith(("nom_t1_", "met_t1_", "manual_t1_", "t1_editor_"))]
        for k in claves: st.session_state.pop(k, None)
        st.rerun()

    # 2. CARGA DE ARCHIVOS
    st.subheader("📁 Archivos de entrada")
    c1, c2, c3 = st.columns(3)
    file_cat_ext = c1.file_uploader("📚 Catálogo Avanzado (SCBCRSE)", type=["csv", "xlsx"], key="cat_ext_t1")
    file_pa = c2.file_uploader("📊 Catálogo PA (Opcional)", type=["csv"], key="pa_t1")
    files_altas = c3.file_uploader("📁 Archivos ALTAS", type=["xlsx", "xlsm"], accept_multiple_files=True, key="altas_t1")
    st.caption("Las restricciones son internas. La PA es opcional y los Excel pueden contener macros.")

    # Invalidar resultados cuando cambien los archivos.
    firma_actual = tuple((f.name, hashlib.sha256(f.getvalue()).hexdigest())
                         for f in ([file_cat_ext, file_pa] + list(files_altas or [])) if f is not None)
    if st.session_state.get("t1_firma_archivos") != firma_actual:
        for k in ["raw_altas", "res_auditoria", "cat_pa_cache", "advertencias_pa",
                  "errores_restricciones", "df_secciones_corregidas", "errores_secciones",
                  "zip_solicitudes_csv_altas", "original_files_bytes", "t1_pendientes_alta", "t1_catalogo",
                  "confirmar_errores_t1", "t1_confirmar_altas"]:
            st.session_state.pop(k, None)
        st.session_state.t1_firma_archivos = firma_actual
        st.session_state.t1_version_editor = st.session_state.get("t1_version_editor", 0) + 1

    # 3. CATÁLOGO AVANZADO
    def cargar_catalogo_avanzado():
        if file_cat_ext is None: return {}
        firma = hashlib.sha256(file_cat_ext.getvalue()).hexdigest()
        if st.session_state.get("cat_avanzado_firma") == firma:
            return st.session_state.get("cat_avanzado_cache") or {}

        datos = io.BytesIO(file_cat_ext.getvalue())
        df = pd.read_csv(datos, dtype=str, encoding="utf-8-sig", keep_default_na=False) \
            if file_cat_ext.name.lower().endswith(".csv") else pd.read_excel(datos, dtype=str).fillna("")
        df.columns = [str(c).strip().upper() for c in df.columns]
        faltantes = {"SCBCRSE_SUBJ_CODE", "SCBCRSE_CRSE_NUMB"} - set(df.columns)
        if faltantes: raise ValueError(f"Catálogo Avanzado: faltan columnas {sorted(faltantes)}")

        catalogo = {}
        for f in df.to_dict("records"):
            llave = (sin_espacios(f.get("SCBCRSE_SUBJ_CODE")), sin_espacios(f.get("SCBCRSE_CRSE_NUMB")))
            if not all(llave): continue
            info = catalogo.setdefault(llave, {"titles": set(), "schd": set(), "insm": set(),
                                                "gmod": set(), "pares": set()})
            for campo in ["SCBCRSE_TITLE", "SCRSYLN_LONG_COURSE_TITLE"]:
                titulo = limpiar_espacios_y_mayusculas(f.get(campo))
                if titulo: info["titles"].add(titulo)
            h, m, g = (sin_espacios(f.get(c)) for c in
                       ["SCRSCHD_SCHD_CODE", "SCRSCHD_INSM_CODE", "SCRGMOD_GMOD_CODE"])
            if h: info["schd"].add(h)
            if m: info["insm"].add(m)
            if g: info["gmod"].add(g)
            if h and m: info["pares"].add((h, m))

        st.session_state.cat_avanzado_cache = catalogo
        st.session_state.cat_avanzado_firma = firma
        return catalogo

    # 4. VALIDAR MATERIA Y CONFIGURACIONES
    def validar_materia(fila, catalogo):
        subj, crse = sin_espacios(fila.get("Subject")), sin_espacios(fila.get("Course"))
        nombre = limpiar_espacios_y_mayusculas(fila.get("Nombre de la Materia"))
        horario, metodo = sin_espacios(fila.get("Tipo de Horario")), sin_espacios(fila.get("Método Educativo"))
        modo = sin_espacios(fila.get("Modo de Calificar"))
        llave = (subj, crse)

        r = {"Luz Verde": False, "Materia Excel": nombre, "Materia Catálogo": "",
             "Subj Original": subj, "Crse Original": crse, "Subj Sugerido": subj, "Crse Sugerido": crse,
             "Horario Original": horario, "Horario Sugerido": horario,
             "Método Original": metodo, "Método Sugerido": metodo,
             "Modo de Calificar Original": modo, "Modo de Calificar Sugerido": modo,
             "Comentario Nombres": "", "Comentario Horario": "", "Comentario Método": "",
             "Comentario Modo de Calificar": "", "Pendientes Alta": [], "Requiere Corrección": False}

        if llave not in catalogo:
            candidatos = []
            for (s, c), info in catalogo.items():
                p_nombre = max([similitud(normalizar_para_cruce(nombre), normalizar_para_cruce(t))
                                for t in info["titles"]] or [0])
                puntaje = p_nombre * 0.55 + similitud(subj, s) * 0.25 + similitud(crse, c) * 0.20
                if puntaje >= 0.63: candidatos.append((puntaje, s, c))
            if candidatos:
                p, s, c = max(candidatos)
                r["Subj Sugerido"], r["Crse Sugerido"] = s, c
                r["Comentario Nombres"] = f"SUBJ/CRSE no existen. Posible coincidencia: {s} {c} ({p:.0%})"
            else:
                r["Comentario Nombres"] = "Materia no encontrada: revisar si requiere alta en catálogo."
            r["Requiere Corrección"] = True
            return r

        info = catalogo[llave]
        titulos = sorted(info["titles"])
        r["Materia Catálogo"] = titulos[0] if titulos else ""
        nombres = {normalizar_para_cruce(t) for t in titulos}

        if not titulos:
            r["Comentario Nombres"] = "Materia encontrada por claves; catálogo sin título verificable."
        elif normalizar_para_cruce(nombre) in nombres:
            r["Comentario Nombres"] = "Todo correcto"
        else:
            r["Comentario Nombres"] = "SUBJ/CRSE correctos, pero el nombre difiere."
            r["Requiere Corrección"] = True

        campos = [
            ("schd", horario, "Tipo de Horario", "Comentario Horario"),
            ("insm", metodo, "Método Educativo", "Comentario Método"),
            ("gmod", modo, "Modo de Calificar", "Comentario Modo de Calificar")
        ]
        for campo, actual, etiqueta, comentario in campos:
            if not actual:
                r[comentario] = f"{etiqueta} vacío: requiere revisión."
                r["Requiere Corrección"] = True
            elif actual in info[campo]:
                r[comentario] = "Correcto"
            else:
                r[comentario] = f"{etiqueta} {actual} no existe para esta materia; requiere alta."
                r["Pendientes Alta"].append((etiqueta, actual))

        if horario in info["schd"] and metodo in info["insm"] and info["pares"]:
            if (horario, metodo) not in info["pares"]:
                r["Comentario Método"] = "Horario y método existen, pero la combinación requiere revisión."
                r["Requiere Corrección"] = True
        return r

    # 5. EXTRAER CONFIGURACIONES PENDIENTES DE ALTA
    def obtener_pendientes_altas(df, auditoria):
        pendientes = []
        for _, r in auditoria.iterrows():
            fila = df.iloc[int(r["idx"])]
            for tipo, valor in r["Pendientes Alta"]:
                pendientes.append({
                    "Periodo": limpiar_clave_texto(fila.get("Periodo")),
                    "SUBJ": sin_espacios(fila.get("Subject")),
                    "COURSE": sin_espacios(fila.get("Course")),
                    "Materia": limpiar_clave_texto(fila.get("Nombre de la Materia")),
                    "Configuración": tipo, "Valor": valor,
                    "Responsable": responsable_normalizado(fila.get("Responsable")),
                    "Archivo": fila.get("ArchivoOrigen", "")
                })
        columnas = ["Periodo", "SUBJ", "COURSE", "Materia", "Configuración", "Valor", "Responsable", "Archivo"]
        return pd.DataFrame(pendientes, columns=columnas)

    # 6. VALIDACIÓN INTELIGENTE
    if files_altas and file_cat_ext and st.button("⚡ Ejecutar Validación Inteligente", type="primary",
                                                   use_container_width=True, key="validar_t1"):
        try:
            with st.spinner("Revisando materias, configuraciones y secciones..."):
                catalogo = cargar_catalogo_avanzado()
                if not catalogo: raise ValueError("El Catálogo Avanzado no contiene materias válidas.")
                pa, avisos = leer_csv_pa(file_pa) if file_pa else (None, [])
                piezas, errores_archivo, originales = [], [], {}

                for archivo in files_altas:
                    if archivo.name in originales:
                        errores_archivo.append(f"Nombre de archivo repetido: {archivo.name}")
                        continue
                    originales[archivo.name] = archivo.getvalue()
                    libro = pd.ExcelFile(io.BytesIO(archivo.getvalue()))
                    hojas = [h for h in libro.sheet_names if h.strip().upper() == HOJA_ALTAS]
                    if not hojas:
                        errores_archivo.append(f"{archivo.name}: falta hoja ALTAS")
                        continue

                    df = libro.parse(hojas[0], dtype=str)
                    df["_FilaExcel"] = range(2, len(df) + 2)
                    df = normalizar_columnas_altas(df)
                    faltantes = set(COLUMNAS_ALTAS) - set(df.columns)
                    if faltantes:
                        errores_archivo.append(f"{archivo.name}: faltan columnas {sorted(faltantes)}")
                        continue
                    df = df.dropna(subset=["Periodo", "Subject", "Course"], how="all").copy()
                    df["ArchivoOrigen"] = archivo.name
                    piezas.append(df)

                if errores_archivo: raise ValueError("\n".join(errores_archivo))
                if not piezas: raise ValueError("No se encontraron registros ALTAS.")

                total = pd.concat(piezas, ignore_index=True)
                total["Responsable"] = total["Responsable"].apply(responsable_normalizado)
                auditoria = []
                for i, fila in total.iterrows():
                    r = validar_materia(fila, catalogo)
                    r.update({"idx": i, "Archivo": fila["ArchivoOrigen"]})
                    auditoria.append(r)

                auditoria = pd.DataFrame(auditoria)
                pendientes = obtener_pendientes_altas(total, auditoria)
                secciones, errores_sec = asignar_secciones_globales(total, pa)

                st.session_state.original_files_bytes = originales
                st.session_state.raw_altas = total
                st.session_state.res_auditoria = auditoria
                st.session_state.cat_pa_cache = pa
                st.session_state.t1_pendientes_alta = pendientes
                st.session_state.advertencias_pa = avisos
                st.session_state.errores_restricciones = auditar_restricciones_internas(total)
                st.session_state.df_secciones_corregidas = secciones
                st.session_state.errores_secciones = errores_sec
                st.session_state.zip_solicitudes_csv_altas = None
                st.session_state.confirmar_errores_t1 = False
                st.session_state.t1_confirmar_altas = False
                st.session_state.t1_version_editor = st.session_state.get("t1_version_editor", 0) + 1

            st.success(f"Validación completada: {len(total)} registros de {len(piezas)} archivo(s).")
        except Exception as e:
            st.error(f"Error de validación: {e}")

    # 7. RESULTADOS
    if isinstance(st.session_state.get("res_auditoria"), pd.DataFrame):
        auditoria = st.session_state.res_auditoria
        pendientes = st.session_state.get("t1_pendientes_alta", pd.DataFrame())
        avisos_pa = st.session_state.get("advertencias_pa") or []
        errores_reglas = st.session_state.get("errores_restricciones")

        # 7.1 CONFIGURACIONES PENDIENTES DE ALTA
        st.divider()
        if not pendientes.empty:
            st.subheader("📚 Configuraciones pendientes de alta")
            resumen_altas = pendientes.drop_duplicates(subset=["SUBJ", "COURSE", "Configuración", "Valor"])
            resumen_altas = resumen_altas.sort_values(["SUBJ", "COURSE", "Configuración", "Valor"])
            st.warning(f"Se encontraron {len(resumen_altas)} configuraciones por dar de alta.")

            with st.expander("👀 Ver configuraciones pendientes", expanded=True):
                st.dataframe(resumen_altas, hide_index=True, use_container_width=True)

            buffer_altas = io.BytesIO()
            with pd.ExcelWriter(buffer_altas, engine="openpyxl") as writer:
                resumen_altas.to_excel(writer, index=False, sheet_name="Altas_Pendientes")
                pendientes.to_excel(writer, index=False, sheet_name="Detalle")
                for ws in writer.sheets.values():
                    ws.freeze_panes, ws.auto_filter.ref = "A2", ws.dimensions
                    for celda in ws[1]:
                        celda.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                        celda.fill = PatternFill(start_color="7030A0", fill_type="solid")
                        celda.alignment = Alignment(horizontal="center", vertical="center")
                    for columna in ws.columns:
                        ancho = max((len(str(c.value)) for c in columna if c.value is not None), default=8)
                        ws.column_dimensions[columna[0].column_letter].width = min(max(ancho + 3, 12), 45)

            st.download_button("📥 Descargar configuraciones pendientes (.xlsx)", buffer_altas.getvalue(),
                               file_name="Configuraciones_Pendientes_Alta.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True, key="descargar_altas_t1")
            st.checkbox("Ya revisé las configuraciones pendientes y deseo continuar con advertencia",
                        key="t1_confirmar_altas")
            st.caption("Si actualizas el Catálogo Avanzado y vuelves a validar, las configuraciones existentes desaparecerán.")
        else:
            st.success("No hay configuraciones pendientes de alta.")

        # 7.2 MESA DE CONTROL: SOLO SI HAY CORRECCIONES
        corregibles = auditoria[auditoria["Requiere Corrección"] == True].copy()
        if not corregibles.empty:
            st.subheader("⚖️ Mesa de Control — Correcciones necesarias")
            for archivo in corregibles["Archivo"].unique():
                sub = corregibles[corregibles["Archivo"] == archivo]
                with st.expander(f"⚠️ {archivo} — {len(sub)} registros", expanded=True):
                    with st.form(f"revision_{hashlib.md5(archivo.encode()).hexdigest()[:10]}"):
                        t1, t2 = st.tabs(["Nombres y Claves", "Horarios y Métodos"])
                        cols_nom = ["Luz Verde", "Materia Excel", "Materia Catálogo", "Comentario Nombres",
                                    "Subj Original", "Crse Original", "Subj Sugerido", "Crse Sugerido"]
                        cols_met = ["Luz Verde", "Materia Excel", "Horario Original", "Horario Sugerido",
                                    "Comentario Horario", "Método Original", "Método Sugerido",
                                    "Comentario Método", "Modo de Calificar Original",
                                    "Modo de Calificar Sugerido", "Comentario Modo de Calificar"]

                        with t1:
                            nom = st.data_editor(sub[cols_nom], hide_index=True, use_container_width=True,
                                                 disabled=["Materia Excel", "Materia Catálogo", "Comentario Nombres",
                                                           "Subj Original", "Crse Original"],
                                                 key=f"nom_t1_{archivo}_{st.session_state.t1_version_editor}")
                        with t2:
                            met = st.data_editor(sub[cols_met], hide_index=True, use_container_width=True,
                                                 disabled=["Materia Excel", "Horario Original", "Comentario Horario",
                                                           "Método Original", "Comentario Método",
                                                           "Modo de Calificar Original", "Comentario Modo de Calificar"],
                                                 key=f"met_t1_{archivo}_{st.session_state.t1_version_editor}")

                        if st.form_submit_button("💾 Guardar correcciones", use_container_width=True):
                            for i in sub.index:
                                auditoria.at[i, "Luz Verde"] = bool(nom.at[i, "Luz Verde"] or met.at[i, "Luz Verde"])
                                for c in ["Subj Sugerido", "Crse Sugerido"]: auditoria.at[i, c] = nom.at[i, c]
                                for c in ["Horario Sugerido", "Método Sugerido", "Modo de Calificar Sugerido"]:
                                    auditoria.at[i, c] = met.at[i, c]
                            st.session_state.res_auditoria = auditoria
                            st.session_state.zip_solicitudes_csv_altas = None
                            st.session_state.t1_version_editor += 1
                            st.success("Correcciones guardadas.")
                            st.rerun()

        # 7.3 ADVERTENCIAS
        if avisos_pa:
            st.subheader("⚠️ Advertencias PA")
            st.dataframe(pd.DataFrame(avisos_pa), hide_index=True, use_container_width=True)

        if isinstance(errores_reglas, pd.DataFrame) and not errores_reglas.empty:
            st.subheader("📋 Restricciones internas")
            st.dataframe(errores_reglas, hide_index=True, use_container_width=True)

        hay_otras_advertencias = bool(
            avisos_pa or (isinstance(errores_reglas, pd.DataFrame) and not errores_reglas.empty)
            or not corregibles.empty
        )
        if hay_otras_advertencias:
            st.checkbox("Revisé las advertencias y correcciones; deseo continuar",
                        key="confirmar_errores_t1")

        # 8. CORRECCIONES Y SECCIONES AUTOMÁTICAS
        def construir_altas_corregidas():
            df = st.session_state.raw_altas.copy()
            aprobadas = st.session_state.res_auditoria
            for _, f in aprobadas[aprobadas["Luz Verde"] == True].iterrows():
                i = int(f["idx"])
                for destino, origen in {
                    "Subject": "Subj Sugerido", "Course": "Crse Sugerido",
                    "Tipo de Horario": "Horario Sugerido", "Método Educativo": "Método Sugerido",
                    "Modo de Calificar": "Modo de Calificar Sugerido"
                }.items():
                    valor = limpiar_clave_texto(f.get(origen))
                    if valor: df.at[i, destino] = valor
            return asignar_secciones_globales(df, st.session_state.get("cat_pa_cache"))

        st.divider()
        st.subheader("🔢 Secciones asignadas")
        try:
            df_sec, errores_sec = construir_altas_corregidas()
            st.session_state.df_secciones_corregidas = df_sec
            st.session_state.errores_secciones = errores_sec
            columnas = ["ArchivoOrigen", "Periodo", "Responsable", "Subject", "Course",
                        "Sección Original", "Sección", "Sección Modificada"]
            with st.expander("Ver secciones asignadas"):
                st.dataframe(df_sec[[c for c in columnas if c in df_sec.columns]],
                             hide_index=True, use_container_width=True)
            st.metric("Secciones modificadas", int(df_sec["Sección Modificada"].sum()))
            if not errores_sec.empty:
                st.error("Hay secciones que no pudieron asignarse.")
                st.dataframe(errores_sec, hide_index=True, use_container_width=True)
        except Exception as e:
            st.error(f"Error al calcular secciones: {e}")

        # 9. CSV BANNER: 15 COLUMNAS ORIGINALES
        def preparar_csv_banner(df):
            r = pd.DataFrame(index=df.index)
            equivalencias = {
                "PERIODO": "Periodo", "SEDE": "Campus", "SUBJ": "Subject", "COURSE": "Course",
                "PARTEPERIODO": "Parte de Periodo", "STATUS": "Estatus",
                "CAPACIDAD": "Capacidad", "SECCION": "Sección",
                "TIPODEHORARIO": "Tipo de Horario", "METODO_EDUCATIVO": "Método Educativo",
                "MODODECALIFICAR": "Modo de Calificar", "SESION": "Sesion"
            }
            faltantes = set(equivalencias.values()) - set(df.columns)
            if faltantes: raise ValueError(f"Faltan columnas para Banner: {sorted(faltantes)}")
            for destino, origen in equivalencias.items(): r[destino] = df[origen].apply(format_r_string)
            for c in ["SUBJ", "COURSE", "TIPODEHORARIO", "METODO_EDUCATIVO"]: r[c] = r[c].apply(sin_espacios)
            r["SECCION"] = df["Sección"].apply(clave_seccion)
            r["CAPACIDAD"] = pd.to_numeric(r["CAPACIDAD"], errors="coerce").astype("Int64")
            r["GRUPOS"], r["SOCIODEINTEGRACION"] = 1, "D2L"
            r["datocomplementario"] = df.apply(
                lambda f: "Bachillerato" if normalizar_para_cruce(f.get("Nivel")) == "BACHILLERATO"
                else cluster_oficial(f.get("Clúster")), axis=1)
            return r[COLUMNAS_CSV].fillna("").to_csv(**CSV_KWARGS_R)

        # 10. GENERAR SOLICITUDES + CSV_ALTAS
        st.divider()
        st.subheader("📁 Nombres de SOLICITUDES (según contenido de ALTAS)")
        try:
            vista, _ = construir_altas_corregidas()
            previstos, usados = [], set()
            for origen, grupo in vista.groupby("ArchivoOrigen", sort=False):
                extension = ".xlsm" if origen.lower().endswith(".xlsm") else ".xlsx"
                nombre = nombre_solicitud(grupo, extension)
                final, n = nombre, 2
                while final.casefold() in usados:
                    final = f"{nombre[:-len(extension)]} ({n}){extension}"
                    n += 1
                usados.add(final.casefold())
                previstos.append({"Archivo recibido": origen, "Nombre SOLICITUDES": final})
            st.dataframe(pd.DataFrame(previstos), hide_index=True, use_container_width=True)
        except Exception as e:
            st.warning(f"No se pudieron calcular los nombres: {e}")
        if st.button("💾 Generar SOLICITUDES y CSV_ALTAS", type="primary",
                     use_container_width=True, key="generar_salidas_t1"):
            try:
                if not pendientes.empty and not st.session_state.get("t1_confirmar_altas"):
                    raise ValueError("Revisa y confirma las configuraciones pendientes de alta.")
                if hay_otras_advertencias and not st.session_state.get("confirmar_errores_t1"):
                    raise ValueError("Confirma las advertencias antes de continuar.")

                corregido, errores = construir_altas_corregidas()
                if not errores.empty: raise ValueError("Hay secciones sin asignar.")
                if corregido["Periodo"].apply(limpiar_clave_texto).eq("").any():
                    raise ValueError("Hay registros sin periodo.")
                if corregido["Responsable"].apply(limpiar_clave_texto).eq("").any():
                    raise ValueError("Hay registros sin responsable.")
                if corregido["Clúster"].apply(limpiar_clave_texto).eq("").any():
                    raise ValueError("Hay registros sin clúster; no se puede nombrar SOLICITUDES.")

                with st.spinner("Generando SOLICITUDES y CSV_ALTAS..."):
                    zip_final = generar_zip_altas(corregido, st.session_state.original_files_bytes,
                                                  preparar_csv_banner)
                st.session_state.zip_solicitudes_csv_altas = zip_final
                st.success("SOLICITUDES y CSV_ALTAS generados correctamente.")
            except Exception as e:
                st.session_state.zip_solicitudes_csv_altas = None
                st.error(f"No se pudieron generar los archivos: {e}")

        if st.session_state.get("zip_solicitudes_csv_altas"):
            st.download_button("📥 Descargar SOLICITUDES + CSV_ALTAS (.ZIP)",
                               data=st.session_state.zip_solicitudes_csv_altas,
                               file_name="ALTAS_PROCESADAS.zip", mime="application/zip",
                               type="primary", use_container_width=True)

    # ========================================================
    # 11. BUSCADOR Y CREACIÓN MANUAL
    # ========================================================
    st.divider()
    st.subheader("🔍 Buscador y creación manual")

    if not isinstance(st.session_state.get("df_manual_fijo"), pd.DataFrame):
        st.session_state.df_manual_fijo = pd.DataFrame(columns=COLUMNAS_CSV + ["Responsable", "Nivel"])
    if "manual_candidatos" not in st.session_state: st.session_state.manual_candidatos = []

    with st.form("buscar_manual_t1"):
        b1, b2, b3 = st.columns(3)
        nombre = b1.text_input("Nombre de materia", key="manual_t1_nombre")
        subj = b2.text_input("SUBJ", key="manual_t1_subj")
        crse = b3.text_input("COURSE", key="manual_t1_crse")
        buscar = st.form_submit_button("🪄 Buscar opciones")

    if buscar:
        if not file_cat_ext: st.warning("Carga el Catálogo Avanzado.")
        elif not any([nombre, subj, crse]): st.warning("Ingresa al menos un criterio.")
        else:
            cat = cargar_catalogo_avanzado()
            candidatos = []
            for (s, c), info in cat.items():
                titulos = sorted(info["titles"])
                pn = max([similitud(normalizar_para_busqueda(nombre), normalizar_para_busqueda(t))
                          for t in titulos] or [0]) if nombre else 0
                ps = similitud(sin_espacios(subj), s) if subj else 0
                pc = similitud(sin_espacios(crse), c) if crse else 0
                puntaje = (pn + ps + pc) / sum(bool(x) for x in [nombre, subj, crse])
                if puntaje >= 0.45: candidatos.append((puntaje, s, c, titulos[0] if titulos else ""))
            st.session_state.manual_candidatos = sorted(candidatos, reverse=True)[:20]

    candidatos = st.session_state.manual_candidatos
    if candidatos:
        opciones = [(s, c) for _, s, c, _ in candidatos]
        etiquetas = {(s, c): f"{s} {c} - {t} ({p:.0%})" for p, s, c, t in candidatos}
        llave = st.selectbox("Materia correcta", opciones, format_func=lambda x: etiquetas[x], key="manual_t1_materia")
        confirmar = st.checkbox(f"Confirmo: {etiquetas[llave]}", key=f"manual_t1_confirmar_{llave}")
        info = cargar_catalogo_avanzado()[llave]

        d1, d2 = st.columns(2)
        with d1:
            horario = st.selectbox("Tipo de horario", [""] + sorted(info["schd"]), key="manual_t1_horario")
            periodo = st.text_input("Periodo", key="manual_t1_periodo")
            responsable = st.text_input("Responsable", key="manual_t1_responsable")
            parte = st.text_input("Parte de periodo", key="manual_t1_parte")
            capacidad = st.text_input("Capacidad", key="manual_t1_capacidad")
            seccion = st.text_input("Sección (se asignará automáticamente)", key="manual_t1_seccion")
        with d2:
            metodos = sorted({m for h, m in info["pares"] if h == horario and m}) if horario else sorted(info["insm"])
            opciones_metodo = [""] + metodos
            if st.session_state.get("manual_t1_metodo", "") not in opciones_metodo:
                st.session_state.manual_t1_metodo = ""
            metodo = st.selectbox("Método educativo", opciones_metodo, key="manual_t1_metodo")
            sede = st.text_input("Sede", key="manual_t1_sede")
            estatus = st.text_input("Estatus", key="manual_t1_estatus")
            modo = st.selectbox("Modo de calificar", [""] + sorted(info["gmod"]), key="manual_t1_modo")
            sesion = st.text_input("Sesión", key="manual_t1_sesion")

        e1, e2 = st.columns(2)
        nivel = e1.selectbox("Nivel", ["LICENCIATURA", "BACHILLERATO", "POSGRADO"], key="manual_t1_nivel")
        cluster = e2.selectbox("Clúster", CLUSTERS_PERMITIDOS, index=None, key="manual_t1_cluster")

        if st.button("➕ Agregar materia", key="manual_t1_agregar"):
            if not confirmar or not cluster or not periodo or not responsable:
                st.warning("Confirma la materia, periodo, responsable y clúster.")
            else:
                nuevo = dict(zip(COLUMNAS_CSV, [
                    periodo, sede, llave[0], llave[1], parte, estatus, capacidad, "1",
                    seccion, horario, metodo, "D2L", modo, sesion,
                    "Bachillerato" if nivel == "BACHILLERATO" else cluster
                ]))
                nuevo.update({"Responsable": responsable_normalizado(responsable), "Nivel": nivel})
                st.session_state.df_manual_fijo = pd.concat(
                    [st.session_state.df_manual_fijo, pd.DataFrame([nuevo])], ignore_index=True)
                st.session_state.manual_t1_editor_version = st.session_state.get("manual_t1_editor_version", 0) + 1
                st.success("Materia agregada.")

    # 12. TABLA MANUAL EDITABLE
    st.subheader("📋 Tabla manual")
    if "manual_t1_editor_version" not in st.session_state: st.session_state.manual_t1_editor_version = 0
    df_manual = st.session_state.df_manual_fijo.copy()
    df_manual.insert(0, "Seleccionar", False)

    with st.form("form_tabla_manual_t1"):
        editado = st.data_editor(df_manual, hide_index=True, num_rows="fixed", use_container_width=True,
                                 key=f"manual_editor_t1_{st.session_state.manual_t1_editor_version}")
        guardar_manual = st.form_submit_button("💾 Guardar cambios de la tabla", use_container_width=True)

    if guardar_manual:
        st.session_state.df_manual_fijo = editado.drop(columns="Seleccionar").copy()
        st.session_state.manual_t1_seleccion = editado.index[editado["Seleccionar"]].tolist()
        st.session_state.manual_t1_editor_version += 1
        st.success("Cambios guardados.")
        st.rerun()

    seleccion = st.session_state.get("manual_t1_seleccion", [])
    a1, a2, a3 = st.columns(3)

    if a1.button("Nuevo renglón", key="manual_t1_nuevo"):
        nuevo = {c: "" for c in st.session_state.df_manual_fijo.columns}
        nuevo.update({"GRUPOS": "1", "SOCIODEINTEGRACION": "D2L"})
        st.session_state.df_manual_fijo = pd.concat(
            [st.session_state.df_manual_fijo, pd.DataFrame([nuevo])], ignore_index=True)
        st.session_state.manual_t1_editor_version += 1
        st.rerun()

    if a2.button("Copiar seleccionados", key="manual_t1_copiar"):
        if seleccion:
            copias = st.session_state.df_manual_fijo.iloc[seleccion].copy()
            st.session_state.df_manual_fijo = pd.concat(
                [st.session_state.df_manual_fijo, copias], ignore_index=True)
            st.session_state.manual_t1_editor_version += 1
            st.session_state.manual_t1_seleccion = []
            st.rerun()
        else: st.warning("Primero selecciona y guarda los renglones.")

    if a3.button("Eliminar seleccionados", key="manual_t1_eliminar"):
        if seleccion:
            st.session_state.df_manual_fijo = st.session_state.df_manual_fijo.drop(
                index=seleccion).reset_index(drop=True)
            st.session_state.manual_t1_editor_version += 1
            st.session_state.manual_t1_seleccion = []
            st.rerun()
        else: st.warning("Primero selecciona y guarda los renglones.")

    # 13. CSV MANUAL
    df_out = st.session_state.df_manual_fijo.copy()
    if not df_out.empty:
        if df_out["PERIODO"].apply(limpiar_clave_texto).eq("").any() or df_out["Responsable"].apply(limpiar_clave_texto).eq("").any():
            st.warning("Hay registros manuales sin Periodo o Responsable.")
        else:
            temporal = pd.DataFrame({
                "Periodo": df_out["PERIODO"], "Subject": df_out["SUBJ"], "Course": df_out["COURSE"],
                "Nivel": df_out["Nivel"], "Sección": df_out["SECCION"]
            })
            temporal, errores_manual = asignar_secciones_globales(temporal, st.session_state.get("cat_pa_cache"))
            df_out["SECCION"] = temporal["Sección"].values

            if not errores_manual.empty:
                st.error("Hay errores de secciones en los registros manuales.")
                st.dataframe(errores_manual, hide_index=True)
            else:
                df_out["Responsable"] = df_out["Responsable"].apply(responsable_normalizado)
                df_out["GRUPOS"], df_out["SOCIODEINTEGRACION"] = "1", "D2L"
                for c in ["SUBJ", "COURSE", "TIPODEHORARIO", "METODO_EDUCATIVO"]:
                    df_out[c] = df_out[c].apply(sin_espacios)

                buffer = io.BytesIO()
                with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
                    for (periodo, resp), grupo in df_out.groupby(["PERIODO", "Responsable"], sort=False):
                        nombre = nombre_csv_altas(periodo, resp)
                        z.writestr(f"CSV_ALTAS/{nombre}",
                                   grupo[COLUMNAS_CSV].to_csv(**CSV_KWARGS_R).encode("utf-8"))

                st.download_button("📥 Descargar CSV_ALTAS Manual (.ZIP)", buffer.getvalue(),
                                   file_name="CSV_ALTAS_MANUAL.zip", mime="application/zip",
                                   type="primary", use_container_width=True)

# ============================================================
# FIN DE PESTAÑA 1
# ============================================================


# ============================================================
# PESTAÑA 2: REPORTE DE ERRORES Y ENSAMBLAJE FINAL
# ============================================================
with tab_err:
    st.header("⚠️ Reporte de Errores y Ensamblaje Final")
    st.caption("Extrae las líneas rechazadas por Banner, corrígelas y reintegra cada registro en su posición original.")

    if st.button("🔄 Limpiar / Recomenzar", key="limpiar_t2"):
        claves = ["t2_base_ext", "t2_errores_ext", "t2_modo", "t2_version", "t2_base_iny", "t2_errores_iny",
                  "t2_corregido_iny", "t2_etiqueta", "t2_delta", "t2_delta_nombre", "t2_delta_firma",
                  "t2_delta_editado", "t2_delta_indices", "t2_editor_version", "t2_final_bytes",
                  "t2_final_nombre", "t2_final_firma", "t2_editor_vista"]
        claves += [k for k in st.session_state if k.startswith("t2_editor_datos_")]
        for k in claves: st.session_state.pop(k, None)
        st.rerun()

    # FUNCIONES AUXILIARES
    COL_FILA_T2 = "_FILA_ORIGINAL_BANNER"
    def leer_csv_t2(archivo):
        df = pd.read_csv(io.BytesIO(archivo.getvalue()), dtype=str, encoding="utf-8-sig", keep_default_na=False, skip_blank_lines=True)
        if df.columns.duplicated().any(): raise ValueError("El CSV tiene encabezados duplicados.")
        if COL_FILA_T2 in df.columns: raise ValueError(f"El CSV original ya contiene la columna reservada {COL_FILA_T2}.")
        return df.reset_index(drop=True)

    def leer_reporte_t2(archivo, cantidad):
        reporte = pd.read_excel(io.BytesIO(archivo.getvalue()), skiprows=2, dtype=str)
        reporte.columns = [limpiar_nombre_columna(c) for c in reporte.columns]
        columna = next((c for c in reporte.columns if "linea" in normalizar_para_busqueda(c)), None)
        if columna is None: raise ValueError("No se encontró la columna Línea en el reporte de Banner.")
        indices, invalidas = [], []
        for valor in reporte[columna].dropna().unique():
            try:
                numero = float(str(valor).strip())
                if not numero.is_integer(): raise ValueError()
                indice = int(numero) - 2  # Línea 2 del CSV = primer registro, después del encabezado.
                if 0 <= indice < cantidad: indices.append(indice)
                else: invalidas.append(str(valor))
            except (ValueError, TypeError): invalidas.append(str(valor))
        if invalidas: raise ValueError(f"Líneas inválidas o fuera del CSV: {invalidas[:15]}")
        if not indices: raise ValueError("No se encontraron líneas válidas con errores.")
        return list(dict.fromkeys(indices))

    def nombre_base_t2(nombre):
        return re.sub(r"(?i)_(base|final|v\d+)$", "", nombre.rsplit(".", 1)[0])

    def firma_archivos_t2(*archivos):
        return tuple((a.name, hashlib.sha256(a.getvalue()).hexdigest()) if a else None for a in archivos)

    def exportar_delta_t2(df, formato):
        if formato == "csv": return df.to_csv(**CSV_KWARGS_R).encode("utf-8")
        buffer = io.BytesIO()
        df.to_excel(buffer, index=False, engine="openpyxl")
        return buffer.getvalue()

    # 1. EXTRAER REGISTROS
    st.subheader("✂️ 1. Extraer registros con errores")
    c1, c2, c3 = st.columns(3)
    base_ext = c1.file_uploader("📁 CSV original de Banner", type=["csv"], key="t2_base_ext")
    errores_ext = c2.file_uploader("📊 Reporte de errores (.xlsx)", type=["xlsx"], key="t2_errores_ext")
    version = c3.text_input("Versión del fragmento", value="V1", key="t2_version")
    firma = (firma_archivos_t2(base_ext, errores_ext), version)
    if st.session_state.get("t2_delta_firma") != firma:
        for k in ["t2_delta", "t2_delta_nombre", "t2_delta_editado", "t2_delta_indices", "t2_editor_version"]:
            st.session_state.pop(k, None)
        st.session_state.t2_delta_firma = firma

    if base_ext and errores_ext and st.button("🔍 Extraer líneas con errores", key="t2_extraer", use_container_width=True):
        try:
            base = leer_csv_t2(base_ext)
            indices = leer_reporte_t2(errores_ext, len(base))
            delta = base.iloc[indices].copy().reset_index(drop=True)
            delta.insert(0, COL_FILA_T2, [i + 2 for i in indices])
            st.session_state.t2_delta = delta.copy()
            st.session_state.t2_delta_editado = delta.copy()
            st.session_state.t2_delta_indices = indices
            st.session_state.t2_delta_nombre = f"{nombre_base_t2(base_ext.name)}_{version}"
            st.session_state.t2_editor_version = st.session_state.get("t2_editor_version", 0) + 1
            st.success(f"Se extrajeron {len(delta)} registros con errores de {len(base)} registros.")
        except Exception as e:
            st.session_state.t2_delta = None
            st.error(f"Error al extraer: {e}")

    delta = st.session_state.get("t2_delta")
    if isinstance(delta, pd.DataFrame):
        st.markdown("#### 📋 Fragmento de errores")
        st.caption(f"La columna {COL_FILA_T2} identifica la fila original. No la cambies ni la elimines al corregir fuera de Streamlit.")
        modo = st.radio("¿Cómo deseas corregirlo?", ["Excel (.xlsx)", "CSV (.csv)", "Editar en Streamlit"], horizontal=True, key="t2_modo")
        nombre_delta = st.session_state.t2_delta_nombre
        if modo == "Excel (.xlsx)":
            st.download_button("📥 Descargar fragmento Excel", exportar_delta_t2(st.session_state.t2_delta_editado, "xlsx"),
                               file_name=f"{nombre_delta}.xlsx", use_container_width=True)
        elif modo == "CSV (.csv)":
            st.download_button("📥 Descargar fragmento CSV", exportar_delta_t2(st.session_state.t2_delta_editado, "csv"),
                               file_name=f"{nombre_delta}.csv", mime="text/csv", use_container_width=True)
        else:
            with st.form("t2_form_edicion"):
                editado = st.data_editor(st.session_state.t2_delta_editado, hide_index=True, num_rows="fixed",
                                         use_container_width=True, disabled=[COL_FILA_T2],
                                         key=f"t2_editor_datos_{st.session_state.t2_editor_version}")
                guardar = st.form_submit_button("💾 Guardar correcciones", use_container_width=True)
            if guardar:
                st.session_state.t2_delta_editado = editado.copy()
                st.session_state.t2_editor_version += 1
                st.success("Cambios guardados correctamente.")
                st.rerun()
            st.download_button("📥 Descargar fragmento corregido", exportar_delta_t2(st.session_state.t2_delta_editado, "csv"),
                               file_name=f"{nombre_delta}.csv", mime="text/csv", use_container_width=True)

    # 2. ENSAMBLAR ARCHIVO FINAL
    st.divider()
    st.subheader("💉 2. Inyectar correcciones y generar CSV final")
    c1, c2, c3 = st.columns(3)
    base_iny = c1.file_uploader("📁 CSV original", type=["csv"], key="t2_base_iny")
    errores_iny = c2.file_uploader("📊 Reporte de errores", type=["xlsx"], key="t2_errores_iny")
    corregido_iny = c3.file_uploader("📝 Fragmento corregido", type=["csv", "xlsx"], key="t2_corregido_iny")
    etiqueta = st.text_input("Etiqueta del archivo final", value="final", key="t2_etiqueta")
    firma_final = (firma_archivos_t2(base_iny, errores_iny, corregido_iny), etiqueta)
    if st.session_state.get("t2_final_firma") != firma_final:
        st.session_state.t2_final_bytes = None
        st.session_state.t2_final_nombre = None
        st.session_state.t2_final_firma = firma_final

    if base_iny and errores_iny and corregido_iny and st.button("🚀 Ensamblar Archivo Final", type="primary", key="t2_ensamblar", use_container_width=True):
        try:
            base = leer_csv_t2(base_iny)
            indices = leer_reporte_t2(errores_iny, len(base))
            if corregido_iny.name.lower().endswith(".xlsx"):
                corregidas = pd.read_excel(io.BytesIO(corregido_iny.getvalue()), dtype=str, keep_default_na=False)
            else:
                corregidas = pd.read_csv(io.BytesIO(corregido_iny.getvalue()), dtype=str,
                                         encoding="utf-8-sig", keep_default_na=False)
            corregidas = corregidas.fillna("").reset_index(drop=True)
            if corregidas.columns.duplicated().any(): raise ValueError("El fragmento contiene encabezados duplicados.")
            if COL_FILA_T2 not in corregidas.columns:
                raise ValueError(f"El fragmento debe conservar {COL_FILA_T2}. Vuelve a extraerlo con esta versión de la pestaña 2.")
            if len(corregidas) != len(indices):
                raise ValueError(f"El reporte tiene {len(indices)} líneas, pero el fragmento tiene {len(corregidas)}.")
            if set(corregidas.columns) - {COL_FILA_T2} != set(base.columns):
                raise ValueError("Las columnas del fragmento no coinciden con las del CSV original.")
            numeros = pd.to_numeric(corregidas[COL_FILA_T2], errors="coerce")
            if numeros.isna().any() or (numeros % 1 != 0).any(): raise ValueError("Hay identificadores de fila inválidos.")
            numeros = numeros.astype(int)
            if numeros.duplicated().any(): raise ValueError("Hay filas originales repetidas en el fragmento.")
            esperadas = {i + 2 for i in indices}
            if set(numeros) != esperadas: raise ValueError("Las filas identificadas no coinciden con las líneas del reporte de errores.")
            final = base.copy()
            for j, numero in enumerate(numeros):
                final.iloc[numero - 2] = corregidas.loc[j, list(base.columns)].to_numpy()
            if len(final) != len(base): raise ValueError("Cambió la cantidad de registros originales.")
            st.session_state.t2_final_bytes = final.to_csv(**CSV_KWARGS_R).encode("utf-8")
            etiqueta_segura = re.sub(r'[^a-zA-Z0-9_-]', '_', etiqueta.strip()) or "final"
            st.session_state.t2_final_nombre = f"{nombre_base_t2(base_iny.name)}_{etiqueta_segura}.csv"
            st.success(f"Archivo final generado: {len(indices)} correcciones, {len(final)} registros. Se respetó la posición original de cada fila.")
        except Exception as e:
            st.session_state.t2_final_bytes = None
            st.error(f"Error al ensamblar: {e}")

    if st.session_state.get("t2_final_bytes"):
        st.download_button("📥 DESCARGAR CSV FINAL", data=st.session_state.t2_final_bytes,
                           file_name=st.session_state.t2_final_nombre, mime="text/csv", type="primary",
                           use_container_width=True, key="t2_descargar_final")
# ============================================================
# FIN DE PESTAÑA 2
# ============================================================
# ============================================================
# PESTAÑA 3: ARGOS, CRUCE RAPIDO E INYECCION SEGURA DE NRC
# ============================================================
with tab3:
    st.header("📊 Inyección de NRCs y Cruces con ARGOS")
    if st.button("🔄 Limpiar / Recomenzar", key="limpiar_t3"):
        claves = ["t3_modo", "t3_argos_completo", "t3_csv_completo", "t3_excel_completo", "t3_argos_rapido", "t3_csv_rapido", "final_argos_zip", "df_cruce_rapido", "t3_alertas", "t3_resultado_editado", "t3_editor_version", "t3_columnas", "t3_firma_completo", "t3_firma_rapido", "t3_tabla_revision", "t3_informe_zip"]
        claves += [k for k in st.session_state if k.startswith("t3_editor_")]
        for k in claves: st.session_state.pop(k, None)
        st.rerun()

    modo = st.radio("Selecciona el proceso:", ["📦 Completo (ARGOS + CSV Final + Excel Original)", "⚡ Rápido (ARGOS + CSV Final)"], horizontal=True, key="t3_modo")
    st.divider()

    # Normalizar los encabezados de ARGOS, incluidos AREA / Ãrea.
    def leer_argos_t3(archivo):
        datos = archivo.getvalue()
        df = None
        for enc in ("utf-8-sig", "cp1252"):
            try:
                df = pd.read_csv(io.BytesIO(datos), dtype=str, encoding=enc, keep_default_na=False, on_bad_lines="error")
                break
            except UnicodeError: continue
        if df is None: raise ValueError("No se pudo leer ARGOS.")
        huellas = {normalizar_para_busqueda_t3(c): c for c in df.columns}
        aliases = {"periodo": ["periodo"], "nivel": ["nivel"], "cluster": ["cluster"], "subj": ["area", "rea", "subj"], "crse": ["nocurso", "course", "crse"], "grupo": ["grupo", "seccion"], "nrc": ["nrc", "crn"], "sede": ["campus", "sede"]}
        columnas = {k: next((huellas[a] for a in opciones if a in huellas), None) for k, opciones in aliases.items()}
        faltantes = [k for k in ("periodo", "nivel", "cluster", "subj", "crse", "grupo", "nrc") if columnas[k] is None]
        if faltantes: raise ValueError(f"ARGOS: faltan columnas {faltantes}.")
        if df.columns.duplicated().any(): raise ValueError("ARGOS tiene encabezados duplicados.")
        for campo, col in columnas.items():
            df[f"_{campo}"] = df[col].apply(ultra_limpiar_seccion if campo == "grupo" else ultra_limpiar) if col else ""
        df = df[df[["_periodo", "_subj", "_crse", "_grupo", "_nrc"]].ne("").all(axis=1)].copy()
        # Un NRC repetido en el reporte no crea un segundo grupo; un NRC contradictorio se informa.
        llave_nrc = ["_periodo", "_subj", "_crse", "_grupo", "_nivel", "_cluster", "_sede"]
        conflictos = df.groupby("_nrc")[llave_nrc].nunique().gt(1).any(axis=1)
        avisos = [{"Tipo": "NRC contradictorio en ARGOS", "NRC": n} for n in conflictos[conflictos].index]
        df = df.drop_duplicates(subset=["_nrc"] + llave_nrc).copy()
        df["_llave"] = df[["_periodo", "_nivel", "_cluster", "_subj", "_crse", "_grupo"]].agg("_".join, axis=1)
        mapa = df.groupby("_llave")["_nrc"].agg(lambda x: sorted(set(x))).to_dict()
        return df, mapa, avisos

    def leer_csv_t3(archivo):
        df = pd.read_csv(io.BytesIO(archivo.getvalue()), dtype=str, encoding="utf-8-sig", keep_default_na=False)
        if df.columns.duplicated().any(): raise ValueError(f"{archivo.name}: encabezados duplicados.")
        faltantes = [c for c in COLUMNAS_CSV if c not in df.columns]
        if faltantes: raise ValueError(f"{archivo.name}: faltan columnas Banner {faltantes}.")
        return df.dropna(how="all").reset_index(drop=True)

    def nivel_csv_t3(cluster):
        c = normalizar_para_cruce(cluster)
        if "POSGRADO" in c: return "POSGRADO"
        if "BACHILLERATO" in c: return "BACHILLERATO"
        return "LICENCIATURA"

    def llaves_csv_t3(df):
        cluster = df["datocomplementario"].apply(ultra_limpiar)
        nivel = df["datocomplementario"].apply(nivel_csv_t3)
        return (df["PERIODO"].apply(ultra_limpiar) + "_" + nivel + "_" + cluster + "_" + df["SUBJ"].apply(ultra_limpiar) + "_" + df["COURSE"].apply(ultra_limpiar) + "_" + df["SECCION"].apply(ultra_limpiar_seccion))

    def cruzar_nrc_t3(df_csv, argos_df, mapa, nombre="CSV"):
        llaves = llaves_csv_t3(df_csv)
        resultados, alertas = [], []
        # Busqueda secundaria solo si la llave corta identifica un NRC unico.
        cortas = argos_df.groupby(["_periodo", "_subj", "_crse", "_grupo"])["_nrc"].agg(lambda x: sorted(set(x))).to_dict()
        for i, (llave, (_, fila)) in enumerate(zip(llaves, df_csv.iterrows())):
            candidatos = mapa.get(llave, [])
            tipo = "Llave completa"
            if not candidatos:
                corta = (ultra_limpiar(fila["PERIODO"]), ultra_limpiar(fila["SUBJ"]), ultra_limpiar(fila["COURSE"]), ultra_limpiar_seccion(fila["SECCION"]))
                candidatos = cortas.get(corta, [])
                tipo = "Llave corta unica" if len(candidatos) == 1 else "Llave corta ambigua o ausente"
            if len(candidatos) == 1: resultados.append(candidatos[0])
            else:
                resultados.append("")
                alertas.append({"Archivo": nombre, "Fila CSV": i + 2, "PERIODO": fila["PERIODO"], "SUBJ": fila["SUBJ"], "COURSE": fila["COURSE"], "SECCION": fila["SECCION"], "Estado": "Varios NRC" if candidatos else "No encontrado", "NRC candidatos": ", ".join(candidatos), "Busqueda": tipo})
        return pd.Series(resultados, index=df_csv.index, dtype="str"), alertas

    def firma_t3(*archivos):
        return tuple((a.name, hashlib.sha256(a.getvalue()).hexdigest()) if a else None for a in archivos)

    def formatear_excel_t3(ws):
        ws.freeze_panes = "A2"
        if not ws.tables: ws.auto_filter.ref = ws.dimensions
        encabezados = {str(c.value): c.column for c in ws[1]}
        idx_nrc = encabezados.get("NRC")
        fondo, nrc_fondo = PatternFill(start_color="7030A0", fill_type="solid"), PatternFill(start_color="E4D5F4", fill_type="solid")
        for celda in ws[1]:
            celda.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            celda.fill = fondo
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for fila in ws.iter_rows(min_row=2):
            for celda in fila:
                celda.alignment = Alignment(horizontal="center", vertical="center")
                if celda.column == idx_nrc:
                    celda.fill = nrc_fondo
                    celda.font = Font(name="Calibri", size=11, bold=True)
        for columna in ws.columns:
            ancho = max((len(str(c.value)) for c in columna if c.value is not None), default=8)
            ws.column_dimensions[columna[0].column_letter].width = min(max(ancho + 3, 11), 45)

    # 3. MODO COMPLETO: emparejar registros por contenido, no por nombre ni posicion.
    if modo.startswith("📦"):
        st.subheader("📦 Inyección de NRC en SOLICITUDES")
        c1, c2, c3 = st.columns(3)
        argos = c1.file_uploader("📊 ARGOS actualizado (.csv)", type=["csv"], key="t3_argos_completo")
        archivos_csv = c2.file_uploader("📝 CSV_ALTAS finales", type=["csv"], accept_multiple_files=True, key="t3_csv_completo")
        archivos_excel = c3.file_uploader("📁 SOLICITUDES originales", type=["xlsx", "xlsm"], accept_multiple_files=True, key="t3_excel_completo")
        firma = firma_t3(argos, *(archivos_csv or []), *(archivos_excel or []))
        if st.session_state.get("t3_firma_completo") != firma:
            st.session_state.final_argos_zip, st.session_state.t3_alertas = None, []
            st.session_state.t3_firma_completo = firma
        if argos and archivos_csv and archivos_excel and st.button("🚀 Procesar y generar Excel con NRC", type="primary", use_container_width=True, key="t3_procesar_completo"):
            try:
                with st.spinner("Buscando NRC y verificando correspondencias..."):
                    argos_df, mapa, avisos = leer_argos_t3(argos)
                    csv_partes, alertas = [], list(avisos)
                    for fc in archivos_csv:
                        parte = leer_csv_t3(fc)
                        parte["_ArchivoCSV"] = fc.name
                        csv_partes.append(parte)
                    csv_total = pd.concat(csv_partes, ignore_index=True)
                    nrc, avisos_cruce = cruzar_nrc_t3(csv_total, argos_df, mapa, "CSV_ALTAS")
                    alertas.extend(avisos_cruce)
                    csv_total["_NRC"] = nrc
                    # Una misma llave CSV no debe tener NRC diferentes.
                    columnas_clave = ["PERIODO", "SUBJ", "COURSE", "SECCION", "SEDE"]
                    for c in columnas_clave: csv_total[f"_K_{c}"] = csv_total[c].apply(ultra_limpiar_seccion if c == "SECCION" else ultra_limpiar)
                    claves = [f"_K_{c}" for c in columnas_clave]
                    repetidas = csv_total.duplicated(subset=claves, keep=False)
                    if repetidas.any():
                        alertas.append({"Tipo": "CSV repetido o ambiguo", "Detalle": f"{int(repetidas.sum())} filas con la misma combinación de periodo, materia, sección y sede. Revisar antes de inyectar."})
                    csv_validos = csv_total[~repetidas & csv_total["_NRC"].ne("")].copy()
                    mapa_csv = {tuple(f[c] for c in claves): f for _, f in csv_validos.iterrows()}
                    buffer, procesados, nombres_zip = io.BytesIO(), 0, set()
                    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
                        for fx in archivos_excel:
                            ext = ".xlsm" if fx.name.lower().endswith(".xlsm") else ".xlsx"
                            wb = openpyxl.load_workbook(io.BytesIO(fx.getvalue()), keep_vba=(ext == ".xlsm"))
                            hoja = next((h for h in wb.sheetnames if h.strip().upper() == HOJA_ALTAS), None)
                            if not hoja:
                                alertas.append({"Archivo": fx.name, "Estado": "Falta hoja ALTAS"})
                                continue
                            ws_altas = wb[hoja]
                            encabezados = {normalizar_para_busqueda(c.value): c.column for c in ws_altas[1] if c.value is not None}
                            campos = {"PERIODO": ["periodo"], "SUBJ": ["subj", "subject"], "COURSE": ["course", "crse"], "SECCION": ["seccion", "grupo"], "SEDE": ["sede", "campus"]}
                            indices = {k: next((encabezados[a] for a in alias if a in encabezados), None) for k, alias in campos.items()}
                            if any(v is None for v in indices.values()):
                                alertas.append({"Archivo": fx.name, "Estado": "Faltan encabezados ALTAS", "Detalle": str([k for k, v in indices.items() if v is None])})
                                continue
                            filas, encontrados = [], 0
                            for fila_excel in range(2, ws_altas.max_row + 1):
                                valores = {k: ws_altas.cell(fila_excel, col).value for k, col in indices.items()}
                                if all(not limpiar_clave_texto(v) for v in valores.values()): continue
                                llave = tuple((ultra_limpiar_seccion(valores[k]) if k == "SECCION" else ultra_limpiar(valores[k])) for k in columnas_clave)
                                coincidencia = mapa_csv.get(llave)
                                if coincidencia is None:
                                    alertas.append({"Archivo": fx.name, "Fila Excel": fila_excel, "Estado": "Sin correspondencia unica en CSV", "Detalle": str(llave)})
                                    filas.append((fila_excel, valores, ""))
                                else:
                                    encontrados += 1
                                    filas.append((fila_excel, valores, coincidencia["_NRC"]))
                            if not filas: continue
                            # Bloquear este libro si tiene coincidencias faltantes o ambiguas.
                            if encontrados != len(filas):
                                alertas.append({"Archivo": fx.name, "Estado": "Libro no exportado", "Detalle": f"{encontrados}/{len(filas)} filas con NRC unico"})
                                continue
                            if HOJA_SALIDA_NRC in wb.sheetnames: del wb[HOJA_SALIDA_NRC]
                            ws = wb.create_sheet(HOJA_SALIDA_NRC)
                            ws.append(["Fila ALTAS", "NRC"] + list(COLUMNAS_ALTAS_REALES))
                            for fila_excel, valores, codigo in filas:
                                valores_originales = {normalizar_para_busqueda(c.value): c.column for c in ws_altas[1] if c.value is not None}
                                ws.append([fila_excel, codigo] + [ws_altas.cell(fila_excel, valores_originales[normalizar_para_busqueda(c)]).value if normalizar_para_busqueda(c) in valores_originales else "" for c in COLUMNAS_ALTAS_REALES])
                            formatear_excel_t3(ws)
                            salida = io.BytesIO()
                            wb.save(salida)
                            nombre = fx.name.rsplit(".", 1)[0] + "_con_NRC" + ext
                            original_nombre, n = nombre, 2
                            while nombre in nombres_zip:
                                nombre = original_nombre[:-len(ext)] + f" ({n})" + ext
                                n += 1
                            nombres_zip.add(nombre)
                            z.writestr(nombre, salida.getvalue())
                            procesados += 1
                    st.session_state.final_argos_zip = buffer.getvalue() if procesados else None
                    st.session_state.t3_alertas = alertas
                    if procesados: st.success(f"Se generaron {procesados} Excel con NRC.")
                    else: st.error("No se generó ningún Excel. Revisa las incidencias de correspondencia.")
            except Exception as e:
                st.session_state.final_argos_zip = None
                st.error(f"Error de inyección: {e}")
        alertas = st.session_state.get("t3_alertas") or []
        if alertas:
            with st.expander(f"⚠️ Incidencias de cruce ({len(alertas)})", expanded=True):
                st.dataframe(pd.DataFrame(alertas), hide_index=True, use_container_width=True)
        if st.session_state.get("final_argos_zip"):
            st.download_button("📥 Descargar Excel con NRC (.ZIP)", st.session_state.final_argos_zip, file_name="Excels_Finales_con_NRC.zip", mime="application/zip", type="primary", use_container_width=True)

    # 4. MODO RAPIDO: tabla editable y descarga.
    else:
        st.subheader("⚡ Cruce rápido de NRC")
        c1, c2 = st.columns(2)
        argos = c1.file_uploader("📊 ARGOS actualizado (.csv)", type=["csv"], key="t3_argos_rapido")
        archivos_csv = c2.file_uploader("📝 CSV finales", type=["csv"], accept_multiple_files=True, key="t3_csv_rapido")
        firma = firma_t3(argos, *(archivos_csv or []))
        if st.session_state.get("t3_firma_rapido") != firma:
            st.session_state.df_cruce_rapido = None
            st.session_state.t3_resultado_editado = None
            st.session_state.t3_editor_version = 0
            st.session_state.t3_firma_rapido = firma
            st.session_state.pop("t3_columnas", None)
        if argos and archivos_csv and st.button("⚡ Cruzar NRC y generar tabla", type="primary", use_container_width=True, key="t3_procesar_rapido"):
            try:
                argos_df, mapa, avisos = leer_argos_t3(argos)
                resultados, alertas = [], list(avisos)
                for archivo in archivos_csv:
                    df = leer_csv_t3(archivo)
                    nrc, faltantes = cruzar_nrc_t3(df, argos_df, mapa, archivo.name)
                    df.insert(0, "NRC", nrc.values)
                    df = df.rename(columns={"TIPODEHORARIO": "TIPO DE HORARIO", "METODO_EDUCATIVO": "METODO_ED", "datocomplementario": "Cluster"})
                    columnas = ["NRC", "PERIODO", "SUBJ", "COURSE", "CAPACIDAD", "SECCION", "TIPO DE HORARIO", "METODO_ED", "Cluster"]
                    resultados.append(df[columnas])
                    alertas.extend(faltantes)
                resultado = pd.concat(resultados, ignore_index=True)
                st.session_state.df_cruce_rapido = resultado.copy()
                st.session_state.t3_resultado_editado = resultado.copy()
                st.session_state.t3_editor_version = st.session_state.get("t3_editor_version", 0) + 1
                st.session_state.t3_alertas = alertas
                st.success(f"Cruce completado: {len(resultado)} registros; {int(resultado['NRC'].ne('').sum())} con NRC.")
            except Exception as e:
                st.session_state.t3_resultado_editado = None
                st.error(f"Error en cruce rápido: {e}")
        alertas = st.session_state.get("t3_alertas") or []
        if alertas:
            with st.expander(f"⚠️ Incidencias ({len(alertas)})"):
                st.dataframe(pd.DataFrame(alertas), hide_index=True, use_container_width=True)
        resultado = st.session_state.get("t3_resultado_editado")
        if isinstance(resultado, pd.DataFrame):
            st.subheader("📋 Resultados del cruce")
            with st.form("t3_form_editor"):
                editado = st.data_editor(resultado, hide_index=True, num_rows="fixed", use_container_width=True, key=f"t3_editor_{st.session_state.t3_editor_version}")
                guardar = st.form_submit_button("💾 Guardar cambios", use_container_width=True)
            if guardar:
                st.session_state.t3_resultado_editado = editado.copy()
                st.session_state.t3_editor_version += 1
                st.rerun()
            columnas = st.multiselect("Columnas para copiar o descargar", options=list(resultado.columns), default=list(resultado.columns), key="t3_columnas")
            if columnas:
                mostrar = st.session_state.t3_resultado_editado[columnas].copy()
                st.dataframe(mostrar, hide_index=True, use_container_width=True)
                st.code(mostrar.to_csv(index=False, sep="\t"), language="text")
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    mostrar.to_excel(writer, index=False, sheet_name="Cruce_NRC")
                    formatear_excel_t3(writer.sheets["Cruce_NRC"])
                st.download_button("📥 Descargar Excel (.xlsx)", buffer.getvalue(), file_name="Cruce_Rapido_NRC.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
            else: st.warning("Selecciona al menos una columna.")
# ============================================================
# FIN DE PESTAÑA 3
# ============================================================



# ============================================================
# PESTAÑA 4: GESTIÓN DE CRUCES
# ============================================================
with tab4:
    st.header("🔗 Gestión de Cruces")
    st.caption("Generación de claves AA-ZZ por periodo, continuación de cruces y consolidación.")

    COL_COMPLETO = ["Clave", "NRC", "Tipo", "Responsable", "Periodo", "SUB", "COURSE", "Comentarios"]
    COL_TI = ["Clave", "NRC", "Periodo"]
    COLUMNAS_ORIGEN = ["NO_CRUCE", "PERIODO", "RESPONSABLE", "NRC", "TIPO", "CLUSTER_CRN", "SUBJ", "COURSE", "MATERIA"]

    # ========================================================
    # 1. FUNCIONES DE LIMPIEZA
    # ========================================================
    def t4_texto(x):
        if x is None or pd.isna(x): return ""
        s = str(x).strip()
        return "" if s.lower() in ("nan", "none", "nat") else re.sub(r"\.0$", "", s)

    def t4_normal(x):
        return re.sub(r"\s+", " ", quitar_acentos(t4_texto(x)).upper()).strip()

    def t4_col(x):
        return re.sub(r"[^A-Z0-9]", "", t4_normal(x))

    def t4_clave_num(clave):
        c = t4_normal(clave)
        if not re.fullmatch(r"[A-Z]{2}", c) or c < "AA" or c > "ZZ":
            raise ValueError(f"Clave inválida: {clave}. Se admiten AA hasta ZZ.")
        return (ord(c[0]) - 65) * 26 + ord(c[1]) - 65

    def t4_num_clave(n):
        if not 0 <= n < 676: raise ValueError("Se agotaron las claves AA-ZZ para un periodo.")
        return chr(65 + n // 26) + chr(65 + n % 26)

    # ========================================================
    # 2. LECTURA DE ARCHIVOS
    # ========================================================
    def t4_leer_tabla(archivo, hoja=None):
        datos = archivo.getvalue()
        if archivo.name.lower().endswith(".csv"):
            ultimo = None
            for cod in ("utf-8-sig", "cp1252"):
                try:
                    return pd.read_csv(io.BytesIO(datos), dtype=str, keep_default_na=False,
                                       encoding=cod, sep=None, engine="python")
                except (UnicodeError, pd.errors.ParserError) as e:
                    ultimo = e
            raise ValueError(f"No se pudo leer {archivo.name}: {ultimo}")

        libro = pd.ExcelFile(io.BytesIO(datos))
        nombre = next((h for h in libro.sheet_names if t4_col(h) == t4_col(hoja)), None) if hoja else libro.sheet_names[0]
        if nombre is None: raise ValueError(f"{archivo.name}: falta la hoja '{hoja}'.")
        return libro.parse(nombre, dtype=str, keep_default_na=False)

    def t4_estandarizar(df, origen=False):
        d = df.copy()
        mapa = {t4_col(c): c for c in (COLUMNAS_ORIGEN if origen else COL_COMPLETO)}
        if origen:
            mapa.update({"SUBJECT": "SUBJ", "AREA": "SUBJ", "CRSE": "COURSE", "NOCURSO": "COURSE"})
        else:
            mapa.update({"SUBJ": "SUB", "SUBJECT": "SUB"})

        d.columns = [mapa.get(t4_col(c), t4_texto(c)) for c in d.columns]
        if d.columns.duplicated().any(): raise ValueError("Hay encabezados duplicados después de normalizarlos.")

        oblig = COLUMNAS_ORIGEN if origen else COL_TI
        faltan = [c for c in oblig if c not in d.columns]
        if faltan: raise ValueError(f"Faltan columnas: {faltan}")

        if not origen:
            for c in COL_COMPLETO:
                if c not in d.columns: d[c] = ""
        for c in d.columns: d[c] = d[c].map(t4_texto)
        return d

    def t4_leer_nuevos(archivos):
        piezas, nombres = [], set()
        for pos, f in enumerate(archivos):
            if f.name in nombres:
                raise ValueError(f"Hay archivos con el mismo nombre: {f.name}. Renómbralos para diferenciarlos.")
            nombres.add(f.name)
            d = t4_estandarizar(t4_leer_tabla(f, "Cruces"), origen=True)
            d = d[d[["NO_CRUCE", "PERIODO", "NRC", "TIPO"]].ne("").any(axis=1)].copy()
            d["_origen"], d["_fila"], d["_orden_archivo"] = f.name, d.index + 2, pos
            piezas.append(d)

        if not piezas: raise ValueError("No hay archivos de Cruces.")
        return pd.concat(piezas, ignore_index=True)

    # ========================================================
    # 3. VALIDACIONES DE MANDANTES E HIJOS
    # ========================================================
    def t4_validar_nuevos(df):
        problemas = []
        for _, r in df.iterrows():
            faltan = [c for c in ("NO_CRUCE", "PERIODO", "NRC", "TIPO", "SUBJ", "COURSE") if not t4_texto(r[c])]
            if faltan: problemas.append(f"{r['_origen']}, fila {r['_fila']}: faltan {', '.join(faltan)}.")
            if t4_normal(r["TIPO"]) not in ("MANDANTE", "CRUCE"):
                problemas.append(f"{r['_origen']}, fila {r['_fila']}: TIPO debe ser Mandante o Cruce.")
            if t4_texto(r["NRC"]) and not re.fullmatch(r"\d+", t4_texto(r["NRC"])):
                problemas.append(f"{r['_origen']}, fila {r['_fila']}: NRC no numérico.")

        for llave, g in df.groupby(["PERIODO", "_origen", "NO_CRUCE"], sort=False):
            mandantes = sum(g["TIPO"].map(t4_normal) == "MANDANTE")
            hijos = sum(g["TIPO"].map(t4_normal) == "CRUCE")
            if mandantes != 1 or hijos < 1:
                problemas.append(f"Periodo {llave[0]}, {llave[1]}, cruce {llave[2]}: {mandantes} mandante(s), {hijos} hijo(s).")
            if g["NRC"].duplicated().any(): problemas.append(f"Cruce {llave}: NRC repetido dentro del cruce.")

        repetidos = df.groupby(["PERIODO", "NRC"])["NO_CRUCE"].size()
        for periodo, nrc in repetidos[repetidos > 1].index:
            problemas.append(f"Periodo {periodo}: NRC {nrc} aparece más de una vez en las plantillas.")
        return problemas

    # ========================================================
    # 4. LEER CRUCES ANTERIORES
    # ========================================================
    def t4_leer_anteriores(archivos):
        completos, tis = [], []
        for f in archivos:
            original = t4_leer_tabla(f)
            completo = set(t4_col(c) for c in COL_COMPLETO).issubset(set(t4_col(c) for c in original.columns))
            d = t4_estandarizar(original)
            d = d[d[COL_TI].ne("").any(axis=1)].copy()
            if d[COL_TI].eq("").any().any():
                raise ValueError(f"{f.name}: hay filas anteriores incompletas.")

            for _, r in d.iterrows(): t4_clave_num(r["Clave"])
            (completos if completo else tis).append(d[COL_COMPLETO if completo else COL_TI])

        full = pd.concat(completos, ignore_index=True) if completos else pd.DataFrame(columns=COL_COMPLETO)
        ti = pd.concat(tis + [full[COL_TI]], ignore_index=True) if tis or not full.empty else pd.DataFrame(columns=COL_TI)

        if not ti.empty:
            ti = ti.drop_duplicates()
            por_nrc = ti.groupby(["Periodo", "NRC"])["Clave"].nunique()
            if (por_nrc > 1).any():
                raise ValueError("Los cruces anteriores contienen NRC asociados a claves distintas en el mismo periodo.")
            por_clave = ti.groupby(["Periodo", "Clave"])["NRC"].nunique()
            if (por_clave < 2).any():
                raise ValueError("Hay claves anteriores con menos de dos NRC.")

        return full.drop_duplicates(), ti

    # ========================================================
    # 5. ASIGNACIÓN DE CLAVES AA-ZZ
    # ========================================================
    def t4_generar(nuevos, anteriores_full, anteriores_ti, ultimas):
        problemas = t4_validar_nuevos(nuevos)
        if problemas: raise ValueError("\n".join(problemas[:35]))

        ocupados = {}
        for _, r in anteriores_ti.iterrows():
            ocupados.setdefault(r["Periodo"], set()).add(t4_clave_num(r["Clave"]))

        # Continuación manual: reservar todas las claves hasta la última.
        for p, ultima in ultimas.items():
            if ultima: ocupados.setdefault(p, set()).update(range(t4_clave_num(ultima) + 1))

        prev_nrc = set(zip(anteriores_ti["Periodo"], anteriores_ti["NRC"]))
        nuevo_nrc = set(zip(nuevos["PERIODO"], nuevos["NRC"]))
        repetidos = prev_nrc & nuevo_nrc
        if repetidos:
            raise ValueError(f"Estos NRC ya están en cruces anteriores: {sorted(repetidos)[:15]}.")

        filas = []
        for (periodo, archivo, no), g in nuevos.groupby(["PERIODO", "_origen", "NO_CRUCE"], sort=False):
            usadas = ocupados.setdefault(periodo, set())
            siguiente = next((n for n in range(676) if n not in usadas), None)
            if siguiente is None: raise ValueError(f"Periodo {periodo}: no quedan claves AA-ZZ.")

            usadas.add(siguiente)
            clave = t4_num_clave(siguiente)
            g = g.assign(_tipo_orden=g["TIPO"].map(lambda x: 0 if t4_normal(x) == "MANDANTE" else 1))
            g = g.sort_values(["_tipo_orden", "_fila"], kind="stable")

            for _, r in g.iterrows():
                filas.append({
                    "Clave": clave, "NRC": r["NRC"],
                    "Tipo": "Mandante" if t4_normal(r["TIPO"]) == "MANDANTE" else "Cruce",
                    "Responsable": r["RESPONSABLE"], "Periodo": periodo,
                    "SUB": r["SUBJ"], "COURSE": r["COURSE"], "Comentarios": ""
                })

        full = pd.DataFrame(filas, columns=COL_COMPLETO)
        ti = full[COL_TI].copy()
        cons_ti = pd.concat([anteriores_ti, ti], ignore_index=True).drop_duplicates() if not anteriores_ti.empty else ti.copy()

        cons_full = None
        if anteriores_ti.empty:
            cons_full = full.copy()
        elif not anteriores_full.empty:
            cubiertos = set(map(tuple, anteriores_full[COL_TI].to_numpy()))
            todos = set(map(tuple, anteriores_ti[COL_TI].to_numpy()))
            if cubiertos == todos:
                cons_full = pd.concat([anteriores_full, full], ignore_index=True).drop_duplicates()

        return full, ti, cons_full, cons_ti

    # ========================================================
    # 6. CREAR EXCEL DE SALIDA
    # ========================================================
    def t4_excel(df, hoja):
        salida = io.BytesIO()
        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=hoja, index=False)
            ws = writer.sheets[hoja]
            ws.freeze_panes, ws.auto_filter.ref = "A2", ws.dimensions

            for celda in ws[1]:
                celda.font = Font(bold=True, color="FFFFFF")
                celda.fill = PatternFill(fill_type="solid", fgColor="7030A0")
                celda.alignment = Alignment(horizontal="center", vertical="center")

            for col in ws.columns:
                ancho = min(40, max(12, max(len(str(c.value or "")) for c in col) + 2))
                ws.column_dimensions[col[0].column_letter].width = ancho
        return salida.getvalue()

    # ========================================================
    # 7. LIMPIAR Y CARGAR ARCHIVOS
    # ========================================================
    if st.button("🔄 Limpiar Cruces", key="t4_limpiar"):
        for k in list(st.session_state):
            if k.startswith("t4_"): st.session_state.pop(k, None)
        st.rerun()

    st.subheader("1. Cargar plantillas")
    archivos_nuevos = st.file_uploader("Plantillas con hoja Cruces", type=["xlsx", "xlsm"],
                                       accept_multiple_files=True, key="t4_nuevos")
    modo_t4 = st.radio("Continuación de claves", [
        "Empezar en AA", "Cargar cruces anteriores", "Indicar última clave manualmente"
    ], key="t4_modo")

    archivos_anteriores = []
    if modo_t4 == "Cargar cruces anteriores":
        archivos_anteriores = st.file_uploader(
            "Cruces anteriores: completo o TI (.xlsx, .xlsm, .csv)",
            type=["xlsx", "xlsm", "csv"], accept_multiple_files=True,
            key="t4_anteriores"
        ) or []

    if modo_t4 == "Indicar última clave manualmente":
        st.caption("Escribe la última clave utilizada por periodo. Si está vacía, comienza en AA.")

    # ========================================================
    # 8. VALIDAR Y GENERAR
    # ========================================================
    st.subheader("2. Revisar y asignar claves")
    firma = tuple((f.name, hashlib.sha256(f.getvalue()).hexdigest())
                  for f in list(archivos_nuevos or []) + list(archivos_anteriores)) + (modo_t4,)

    if st.session_state.get("t4_firma") != firma:
        st.session_state.t4_firma = firma
        st.session_state.t4_resultados = None
        st.session_state.t4_editor_version = st.session_state.get("t4_editor_version", 0) + 1

    if archivos_nuevos:
        try:
            nuevos = t4_leer_nuevos(archivos_nuevos)
            periodos = sorted(p for p in nuevos["PERIODO"].unique() if p)
            ultimas = {}

            if modo_t4 == "Indicar última clave manualmente":
                cols = st.columns(min(3, max(1, len(periodos))))
                for i, p in enumerate(periodos):
                    ultimas[p] = cols[i % len(cols)].text_input(
                        f"Última clave del periodo {p}", key=f"t4_ultima_{p}"
                    ).strip().upper()

            if modo_t4 == "Cargar cruces anteriores" and not archivos_anteriores:
                st.info("Carga los archivos anteriores para continuar la numeración.")
            else:
                anteriores_full, anteriores_ti = t4_leer_anteriores(archivos_anteriores) if archivos_anteriores else (
                    pd.DataFrame(columns=COL_COMPLETO), pd.DataFrame(columns=COL_TI)
                )

                if st.button("🔍 Validar y generar claves", type="primary", key="t4_generar"):
                    full, ti, cons_full, cons_ti = t4_generar(nuevos, anteriores_full, anteriores_ti, ultimas)
                    st.session_state.t4_resultados = (full, ti, cons_full, cons_ti)
                    st.session_state.t4_editor_version += 1
                    cantidad = full.groupby(["Periodo", "Clave"]).ngroups
                    st.success(f"Se asignaron {cantidad} cruces nuevos.")

        except Exception as e:
            st.error(f"Revisión de cruces: {e}")

    # ========================================================
    # 9. TABLA EDITABLE Y PERSISTENTE
    # ========================================================
    resultado = st.session_state.get("t4_resultados")
    if resultado is not None:
        full, ti, cons_full, cons_ti = resultado
        st.subheader("3. Revisar resultados")
        st.caption("Puedes editar los datos descriptivos. Clave, NRC, Tipo y Periodo permanecen protegidos.")

        with st.form("t4_edicion"):
            editado = st.data_editor(full, hide_index=True, num_rows="fixed",
                                     use_container_width=True,
                                     disabled=["Clave", "NRC", "Tipo", "Periodo"],
                                     key=f"t4_editor_{st.session_state.t4_editor_version}")
            guardar = st.form_submit_button("💾 Guardar cambios", use_container_width=True)

        if guardar:
            full = editado.copy()
            ti = full[COL_TI].copy()

            if cons_full is not None:
                anteriores = cons_full.iloc[:len(cons_full) - len(resultado[0])].copy()
                cons_full = pd.concat([anteriores, full], ignore_index=True)

            anteriores_ti = cons_ti.iloc[:len(cons_ti) - len(resultado[1])].copy()
            cons_ti = pd.concat([anteriores_ti, ti], ignore_index=True)

            st.session_state.t4_resultados = (full, ti, cons_full, cons_ti)
            st.session_state.t4_editor_version += 1
            st.success("Cambios guardados.")
            st.rerun()

        # ====================================================
        # 10. DESCARGAS
        # ====================================================
        salida = st.radio("¿Qué deseas descargar?", [
            "Solo nuevos", "Consolidado completo", "Ambas opciones"
        ], index=2, horizontal=True, key="t4_salida")

        archivos = {}
        if salida in ("Solo nuevos", "Ambas opciones"):
            archivos["Cruces_Nuevos_Completo.xlsx"] = t4_excel(full[COL_COMPLETO], "Cruces")
            archivos["Cruces_Nuevos_TI.xlsx"] = t4_excel(ti[COL_TI], "TI")

        if salida in ("Consolidado completo", "Ambas opciones"):
            archivos["Cruces_Consolidado_TI.xlsx"] = t4_excel(cons_ti[COL_TI], "TI")
            if cons_full is not None:
                archivos["Cruces_Consolidado_Completo.xlsx"] = t4_excel(cons_full[COL_COMPLETO], "Cruces")
            else:
                st.warning("No se puede generar el consolidado completo de 8 columnas porque los anteriores solo contienen Clave, NRC y Periodo.")

        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
            for nombre, datos in archivos.items(): z.writestr(nombre, datos)

        st.download_button("📥 Descargar cruces (.ZIP)", buffer.getvalue(),
                           file_name="CRUCES_PROCESADOS.zip", mime="application/zip",
                           type="primary", use_container_width=True, key="t4_descargar")

# ============================================================
# FIN DE PESTAÑA 4
# ============================================================
