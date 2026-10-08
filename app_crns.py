

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
    periodos = sorted({limpiar_clave_texto(x) for x in df["Periodo"] if limpiar_clave_texto(x)})
    responsables = responsables_unicos(df["Responsable"])
    if not periodos or not responsables: raise ValueError("Falta Periodo o Responsable para nombrar SOLICITUDES.")
    union = " y " if len(responsables) == 2 else ", "
    return f"{'-'.join(periodos)} {union.join(responsables)} CRNs U-ERRE{extension}"

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

    df.columns = [normalizar_para_busqueda(str(c).replace("Ã", "A")) for c in df.columns]
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
# 7. DETECCIÓN DE POSIBLES ALTAS DUPLICADAS
# ============================================================
def detectar_altas_duplicadas(df):
    columnas = ["Periodo", "Subject", "Course", "Nivel", "Responsable",
                "Parte de Periodo", "Tipo de Horario", "Método Educativo"]
    disponibles = [c for c in columnas if c in df.columns]
    if not disponibles: return pd.DataFrame()

    temporal = df.copy()
    for c in disponibles: temporal[f"_dup_{c}"] = temporal[c].apply(normalizar_para_cruce)
    claves = [f"_dup_{c}" for c in disponibles]
    repetidas = temporal.duplicated(subset=claves, keep=False)
    resultado = df.loc[repetidas].copy()
    if not resultado.empty: resultado["Observación"] = "Posible ALTA duplicada; requiere revisión."
    return resultado

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
        "Subject": ["subject", "subj", "area"], "Course": ["course", "crse", "nocurso"],
        "Sección": ["seccion", "grupo"], "Tipo de Horario": ["tipodehorario"],
        "Método Educativo": ["metodoeducativo"], "Modo de Calificar": ["mododecalificar"]
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
    requeridas = ["Periodo", "Responsable", "ArchivoOrigen", "_FilaExcel"]
    faltantes = [c for c in requeridas if c not in df_final.columns]
    if faltantes: raise ValueError(f"Faltan columnas para exportar: {faltantes}")

    df = df_final.copy()
    df["Periodo"] = df["Periodo"].apply(limpiar_clave_texto)
    df["Responsable"] = df["Responsable"].apply(responsable_normalizado)
    if df["Periodo"].eq("").any() or df["Responsable"].eq("").any():
        raise ValueError("Hay ALTAS sin Periodo o Responsable.")

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        # SOLICITUDES: un Excel por archivo original, con todos sus periodos y responsables.
        nombres = set()
        for origen, grupo in df.groupby("ArchivoOrigen", sort=False):
            contenido, extension = crear_solicitud_excel(grupo, archivos_originales)
            nombre = nombre_solicitud(grupo, extension)
            ruta = f"SOLICITUDES/{nombre}"
            if ruta in nombres: raise ValueError(f"Nombre de solicitud repetido: {nombre}")
            z.writestr(ruta, contenido)
            nombres.add(ruta)

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
    "duplicados_altas": None, "zip_solicitudes_csv_altas": None,
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
# PESTAÑA 1: VALIDACIÓN DE ALTAS, PA Y GENERACIÓN DE ARCHIVOS
# ============================================================

with tab1:
    st.header("⚙️ Validación de ALTAS y Generación de Archivos")

    # 1. REINICIAR
    if st.button("🔄 Limpiar / Recomenzar", key="limpiar_t1"):
        claves = ["cat_ext_t1", "pa_t1", "altas_t1", "raw_altas", "res_auditoria",
                  "cat_pa_cache", "advertencias_pa", "errores_restricciones", "duplicados_altas",
                  "df_secciones_corregidas", "errores_secciones", "zip_solicitudes_csv_altas",
                  "original_files_bytes", "cat_avanzado_cache", "cat_avanzado_firma",
                  "df_manual_fijo", "manual_candidatos", "manual_busqueda_realizada",
                  "confirmar_errores_t1", "confirmar_duplicados_t1", "manual_editor_t1"]
        claves += [k for k in st.session_state if k.startswith(("nom_t1_", "met_t1_", "manual_t1_"))]
        for k in claves: st.session_state.pop(k, None)
        st.rerun()

    # 2. CARGA DE ARCHIVOS
    st.subheader("📁 Archivos de entrada")
    c1, c2, c3 = st.columns(3)
    file_cat_ext = c1.file_uploader("📚 Catálogo Avanzado (SCBCRSE)", type=["csv", "xlsx"], key="cat_ext_t1")
    file_pa = c2.file_uploader("📊 Catálogo PA (Opcional)", type=["csv"], key="pa_t1")
    files_altas = c3.file_uploader("📁 Archivos ALTAS", type=["xlsx", "xlsm"], accept_multiple_files=True, key="altas_t1")
    st.caption("Las restricciones son internas. La PA es opcional y los Excel pueden contener macros.")

    # 3. COLUMNAS Y CATÁLOGO AVANZADO
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
        if faltantes: raise ValueError(f"Catálogo Avanzado: faltan {sorted(faltantes)}")

        catalogo = {}
        for _, f in df.iterrows():
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
            if h or m: info["pares"].add((h, m))

        st.session_state.cat_avanzado_cache = catalogo
        st.session_state.cat_avanzado_firma = firma
        return catalogo

    # 4. VALIDAR MATERIA
    def validar_materia(fila, catalogo):
        subj, crse = sin_espacios(fila.get("Subject")), sin_espacios(fila.get("Course"))
        nombre = limpiar_espacios_y_mayusculas(fila.get("Nombre de la Materia"))
        horario, metodo = sin_espacios(fila.get("Tipo de Horario")), sin_espacios(fila.get("Método Educativo"))
        modo = sin_espacios(fila.get("Modo de Calificar"))

        r = {"Luz Verde": False, "Materia Excel": nombre, "Materia Catálogo": "",
             "Subj Original": subj, "Crse Original": crse, "Subj Sugerido": subj, "Crse Sugerido": crse,
             "Horario Original": horario, "Horario Sugerido": horario, "Método Original": metodo,
             "Método Sugerido": metodo, "Modo de Calificar Original": modo,
             "Modo de Calificar Sugerido": modo, "Comentario Nombres": "",
             "Comentario Horario": "", "Comentario Método": "", "Comentario Modo de Calificar": ""}

        llave = (subj, crse)
        if llave not in catalogo:
            candidatos = []
            for (s, c), info in catalogo.items():
                p_nombre = max([similitud(normalizar_para_cruce(nombre), normalizar_para_cruce(t))
                                for t in info["titles"]] or [0])
                p = p_nombre * 0.55 + similitud(subj, s) * 0.25 + similitud(crse, c) * 0.20
                candidatos.append((p, s, c))
            if not candidatos or max(candidatos)[0] < 0.63:
                r["Comentario Nombres"] = "Materia no encontrada en catálogo"
                return r
            p, s, c = max(candidatos)
            llave = (s, c)
            r["Subj Sugerido"], r["Crse Sugerido"] = s, c
            r["Comentario Nombres"] = f"Claves sugeridas ({p:.0%})"

        info = catalogo[llave]
        titulos = sorted(info["titles"])
        r["Materia Catálogo"] = titulos[0] if titulos else ""
        if not r["Comentario Nombres"]:
            nombres = {normalizar_para_cruce(t) for t in titulos}
            r["Comentario Nombres"] = "Todo correcto" if normalizar_para_cruce(nombre) in nombres \
                else "Clave OK, pero Nombre difiere"

        campos = [
            ("schd", horario, "Horario", "Comentario Horario", "Horario Sugerido"),
            ("insm", metodo, "Método", "Comentario Método", "Método Sugerido"),
            ("gmod", modo, "Modo de calificar", "Comentario Modo de Calificar", "Modo de Calificar Sugerido")
        ]
        for campo, actual, etiqueta, comentario, sugerido in campos:
            permitidos = info[campo]
            if not permitidos: r[comentario] = "Sin restricciones en catálogo"
            elif actual in permitidos: r[comentario] = f"{etiqueta} OK"
            else:
                r[comentario] = "Error. Permitidos: " + ", ".join(sorted(permitidos))
                r[sugerido] = sorted(permitidos)[0] if len(permitidos) == 1 else ""

        if horario and metodo and info["pares"] and (horario, metodo) not in info["pares"]:
            r["Comentario Método"] += " | Combinación horario-método no encontrada"
        return r

    # 5. VALIDACIÓN MASIVA
    if files_altas and file_cat_ext and st.button("⚡ Ejecutar Validación Inteligente", type="primary", key="validar_t1"):
        try:
            catalogo = cargar_catalogo_avanzado()
            if not catalogo: raise ValueError("El Catálogo Avanzado no contiene materias válidas.")
            pa, avisos = leer_csv_pa(file_pa) if file_pa else (None, [])
            piezas, errores_archivo, originales = [], [], {}

            for archivo in files_altas:
                originales[archivo.name] = archivo.getvalue()
                libro = pd.ExcelFile(io.BytesIO(archivo.getvalue()))
                hojas = [h for h in libro.sheet_names if h.strip().upper() == HOJA_ALTAS]
                if not hojas:
                    errores_archivo.append(f"{archivo.name}: falta hoja ALTAS")
                    continue

                df = libro.parse(hojas[0], dtype=str)
                df["_FilaExcel"] = range(2, len(df) + 2)
                df = normalizar_columnas_altas(df).dropna(how="all")
                faltantes = {"Periodo", "Subject", "Course", "Nivel", "Sección"} - set(df.columns)
                if faltantes:
                    errores_archivo.append(f"{archivo.name}: faltan columnas {sorted(faltantes)}")
                    continue
                if "Responsable" not in df.columns:
                    errores_archivo.append(f"{archivo.name}: no se encontró Responsable. Encabezados: {list(df.columns)}")
                    continue

                df = df.dropna(subset=["Periodo", "Subject", "Course"], how="all").copy()
                df["ArchivoOrigen"] = archivo.name
                piezas.append(df)

            if errores_archivo: raise ValueError("\n".join(errores_archivo))
            if not piezas: raise ValueError("No hay registros ALTAS válidos.")

            total = pd.concat(piezas, ignore_index=True)
            total["Responsable"] = total["Responsable"].apply(responsable_normalizado)
            auditoria = []
            for i, fila in total.iterrows():
                r = validar_materia(fila, catalogo)
                r.update({"idx": i, "Archivo": fila["ArchivoOrigen"]})
                auditoria.append(r)

            st.session_state.original_files_bytes = originales
            st.session_state.raw_altas = total
            st.session_state.res_auditoria = pd.DataFrame(auditoria)
            st.session_state.cat_pa_cache = pa
            st.session_state.advertencias_pa = avisos
            st.session_state.errores_restricciones = auditar_restricciones_internas(total)
            st.session_state.duplicados_altas = detectar_altas_duplicadas(total)
            st.session_state.zip_solicitudes_csv_altas = None
            st.session_state.confirmar_errores_t1 = False
            st.session_state.confirmar_duplicados_t1 = False
            st.success(f"Validación completada: {len(total)} registros de {len(piezas)} archivo(s).")
        except Exception as e:
            st.error(f"Error de validación: {e}")

    # 6. MESA DE CONTROL
    if isinstance(st.session_state.get("res_auditoria"), pd.DataFrame):
        auditoria = st.session_state.res_auditoria
        st.divider()
        st.subheader("⚖️ Mesa de Control")

        for archivo in auditoria["Archivo"].unique():
            sub = auditoria[auditoria["Archivo"] == archivo]
            pendientes = sub[
                ~sub["Comentario Nombres"].eq("Todo correcto") |
                ~sub["Comentario Horario"].isin(["Horario OK", "Sin restricciones en catálogo"]) |
                ~sub["Comentario Método"].isin(["Método OK", "Sin restricciones en catálogo"]) |
                ~sub["Comentario Modo de Calificar"].isin(["Modo de calificar OK", "Sin restricciones en catálogo"])
            ]
            if pendientes.empty:
                st.success(f"{archivo}: sin advertencias de catálogo.")
                continue

            with st.expander(f"⚠️ {archivo} — {len(pendientes)} advertencias", expanded=True):
                with st.form(f"revision_{hashlib.md5(archivo.encode()).hexdigest()[:10]}"):
                    t1, t2 = st.tabs(["Nombres y Claves", "Horarios y Métodos"])
                    cols_nom = ["Luz Verde", "Materia Excel", "Materia Catálogo", "Comentario Nombres",
                                "Subj Original", "Crse Original", "Subj Sugerido", "Crse Sugerido"]
                    cols_met = ["Luz Verde", "Materia Excel", "Horario Original", "Horario Sugerido",
                                "Comentario Horario", "Método Original", "Método Sugerido",
                                "Comentario Método", "Modo de Calificar Original",
                                "Modo de Calificar Sugerido", "Comentario Modo de Calificar"]

                    with t1:
                        nom = st.data_editor(pendientes[cols_nom], hide_index=True, use_container_width=True,
                                             disabled=["Materia Excel", "Materia Catálogo", "Comentario Nombres",
                                                       "Subj Original", "Crse Original"],
                                             key=f"nom_t1_{archivo}")
                    with t2:
                        met = st.data_editor(pendientes[cols_met], hide_index=True, use_container_width=True,
                                             disabled=["Materia Excel", "Horario Original", "Comentario Horario",
                                                       "Método Original", "Comentario Método",
                                                       "Modo de Calificar Original", "Comentario Modo de Calificar"],
                                             key=f"met_t1_{archivo}")

                    if st.form_submit_button("💾 Confirmar correcciones", use_container_width=True):
                        for i in pendientes.index:
                            auditoria.at[i, "Luz Verde"] = bool(nom.at[i, "Luz Verde"] or met.at[i, "Luz Verde"])
                            for c in ["Subj Sugerido", "Crse Sugerido"]: auditoria.at[i, c] = nom.at[i, c]
                            for c in ["Horario Sugerido", "Método Sugerido", "Modo de Calificar Sugerido"]:
                                auditoria.at[i, c] = met.at[i, c]
                        st.session_state.res_auditoria = auditoria
                        st.session_state.zip_solicitudes_csv_altas = None
                        st.rerun()

        # 7. ADVERTENCIAS, DUPLICADOS Y CONFIRMACIONES
        avisos_pa = st.session_state.get("advertencias_pa") or []
        errores_reglas = st.session_state.get("errores_restricciones")
        duplicados = st.session_state.get("duplicados_altas")
        errores_cat = auditoria[
            ~auditoria["Comentario Nombres"].eq("Todo correcto") |
            ~auditoria["Comentario Horario"].isin(["Horario OK", "Sin restricciones en catálogo"]) |
            ~auditoria["Comentario Método"].isin(["Método OK", "Sin restricciones en catálogo"]) |
            ~auditoria["Comentario Modo de Calificar"].isin(["Modo de calificar OK", "Sin restricciones en catálogo"])
        ]

        if avisos_pa:
            st.subheader("⚠️ Advertencias PA")
            st.dataframe(pd.DataFrame(avisos_pa), hide_index=True, use_container_width=True)
        if isinstance(errores_reglas, pd.DataFrame) and not errores_reglas.empty:
            st.subheader("📋 Restricciones internas")
            st.dataframe(errores_reglas, hide_index=True, use_container_width=True)
        if isinstance(duplicados, pd.DataFrame) and not duplicados.empty:
            st.subheader("🔁 Posibles ALTAS duplicadas")
            st.dataframe(duplicados, hide_index=True, use_container_width=True)

        hay_advertencias = bool(len(errores_cat) or avisos_pa or
                               (isinstance(errores_reglas, pd.DataFrame) and not errores_reglas.empty))
        hay_duplicados = isinstance(duplicados, pd.DataFrame) and not duplicados.empty

        if hay_advertencias:
            st.checkbox("Confirmo que revisé las advertencias y deseo continuar",
                        key="confirmar_errores_t1")
        if hay_duplicados:
            st.checkbox("Confirmo que revisé los posibles duplicados y deseo conservarlos",
                        key="confirmar_duplicados_t1")

        # 8. APLICAR CORRECCIONES Y SECCIONES
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
        st.subheader("🔢 Corrección global de secciones")
        if st.button("🔄 Recalcular secciones", key="recalcular_t1"):
            try:
                corregido, errores = construir_altas_corregidas()
                st.session_state.df_secciones_corregidas = corregido
                st.session_state.errores_secciones = errores
                st.session_state.zip_solicitudes_csv_altas = None
            except Exception as e: st.error(str(e))

        df_sec = st.session_state.get("df_secciones_corregidas")
        if isinstance(df_sec, pd.DataFrame):
            columnas = ["ArchivoOrigen", "Periodo", "Responsable", "Subject", "Course",
                        "Sección Original", "Sección", "Sección Modificada"]
            st.dataframe(df_sec[[c for c in columnas if c in df_sec.columns]],
                         hide_index=True, use_container_width=True)
            st.metric("Secciones modificadas", int(df_sec["Sección Modificada"].sum()))

        errores_sec = st.session_state.get("errores_secciones")
        if isinstance(errores_sec, pd.DataFrame) and not errores_sec.empty:
            st.error("Hay secciones sin asignar.")
            st.dataframe(errores_sec, hide_index=True, use_container_width=True)

        # 9. PREPARAR CSV BANNER
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
        if st.button("💾 Generar SOLICITUDES y CSV_ALTAS", type="primary",
                     use_container_width=True, key="generar_salidas_t1"):
            try:
                if hay_advertencias and not st.session_state.get("confirmar_errores_t1"):
                    raise ValueError("Confirma primero las advertencias.")
                if hay_duplicados and not st.session_state.get("confirmar_duplicados_t1"):
                    raise ValueError("Confirma primero los posibles duplicados.")

                corregido, errores = construir_altas_corregidas()
                if not errores.empty: raise ValueError("Hay secciones sin asignar. Revisa los errores.")
                st.session_state.df_secciones_corregidas = corregido
                st.session_state.errores_secciones = errores

                zip_final = generar_zip_altas(corregido, st.session_state.original_files_bytes,
                                              preparar_csv_banner)
                st.session_state.zip_solicitudes_csv_altas = zip_final
                st.success("SOLICITUDES y CSV_ALTAS generados correctamente.")
            except Exception as e:
                st.session_state.zip_solicitudes_csv_altas = None
                st.error(f"No se pudieron generar los archivos: {e}")

        if st.session_state.get("zip_solicitudes_csv_altas"):
            st.download_button("📥 Descargar SOLICITUDES + CSV_ALTAS (.ZIP)",
                               st.session_state.zip_solicitudes_csv_altas,
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
                st.success("Materia agregada.")

    # 12. TABLA MANUAL EDITABLE Y PERSISTENTE
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
        st.success("Cambios guardados.")

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
            st.session_state.df_manual_fijo = pd.concat([st.session_state.df_manual_fijo, copias], ignore_index=True)
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

    # 13. DESCARGA CSV MANUAL POR PERIODO Y RESPONSABLE
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
    st.caption("Extrae las líneas con errores de Banner, corrígelas y vuelve a integrarlas al CSV original.")

    # 1. REINICIAR PESTAÑA
    if st.button("🔄 Limpiar / Recomenzar", key="limpiar_t2"):
        claves = ["t2_base_ext", "t2_errores_ext", "t2_modo", "t2_version",
                  "t2_base_iny", "t2_errores_iny", "t2_corregido_iny", "t2_etiqueta",
                  "t2_delta", "t2_delta_nombre", "t2_delta_firma", "t2_delta_editado",
                  "t2_editor_version", "t2_final_bytes", "t2_final_nombre", "t2_final_firma"]
        for k in claves: st.session_state.pop(k, None)
        st.rerun()

    # 2. FUNCIONES AUXILIARES
    def leer_csv_t2(archivo):
        return pd.read_csv(io.BytesIO(archivo.getvalue()), dtype=str, encoding="utf-8-sig",
                           keep_default_na=False, skip_blank_lines=True)

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
                indice = int(numero) - 2
                if 0 <= indice < cantidad: indices.append(indice)
                else: invalidas.append(str(valor))
            except (ValueError, TypeError):
                invalidas.append(str(valor))

        if invalidas: raise ValueError(f"Líneas inválidas o fuera del CSV: {invalidas[:15]}")
        if not indices: raise ValueError("No se encontraron líneas válidas con errores.")
        return list(dict.fromkeys(indices))

    def nombre_base_t2(nombre):
        return re.sub(r"(?i)_(base|final|v\d+)$", "", nombre.rsplit(".", 1)[0])

    def firma_archivos_t2(*archivos):
        return tuple((a.name, hashlib.sha256(a.getvalue()).hexdigest()) if a else None for a in archivos)

    # ========================================================
    # 3. EXTRAER REGISTROS CON ERRORES
    # ========================================================
    st.subheader("✂️ 1. Extraer registros con errores")
    c1, c2, c3 = st.columns(3)
    base_ext = c1.file_uploader("📁 CSV original de Banner", type=["csv"], key="t2_base_ext")
    errores_ext = c2.file_uploader("📊 Reporte de errores (.xlsx)", type=["xlsx"], key="t2_errores_ext")
    version = c3.text_input("Versión del fragmento", value="V1", key="t2_version")

    firma = (firma_archivos_t2(base_ext, errores_ext), version)
    if st.session_state.get("t2_delta_firma") != firma:
        for k in ["t2_delta", "t2_delta_nombre", "t2_delta_editado", "t2_editor_version"]:
            st.session_state.pop(k, None)
        st.session_state.t2_delta_firma = firma

    if base_ext and errores_ext and st.button("🔍 Extraer líneas con errores", key="t2_extraer", use_container_width=True):
        try:
            base = leer_csv_t2(base_ext)
            indices = leer_reporte_t2(errores_ext, len(base))
            delta = base.iloc[indices].copy().reset_index(drop=True)
            st.session_state.t2_delta = delta
            st.session_state.t2_delta_editado = delta.copy()
            st.session_state.t2_delta_nombre = f"{nombre_base_t2(base_ext.name)}_{version}"
            st.session_state.t2_editor_version = 0
            st.success(f"Se extrajeron {len(delta)} registros con errores de {len(base)} registros.")
        except Exception as e:
            st.session_state.t2_delta = None
            st.error(f"Error al extraer: {e}")

    delta = st.session_state.get("t2_delta")
    if isinstance(delta, pd.DataFrame):
        st.markdown("#### 📋 Fragmento de errores")
        modo = st.radio("¿Cómo deseas corregirlo?", ["Excel (.xlsx)", "CSV (.csv)", "Editar en Streamlit"],
                        horizontal=True, key="t2_modo")
        nombre_delta = st.session_state.t2_delta_nombre

        if modo == "Excel (.xlsx)":
            buffer = io.BytesIO()
            st.session_state.t2_delta_editado.to_excel(buffer, index=False, engine="openpyxl")
            st.download_button("📥 Descargar fragmento Excel", buffer.getvalue(),
                               file_name=f"{nombre_delta}.xlsx", use_container_width=True)

        elif modo == "CSV (.csv)":
            st.download_button("📥 Descargar fragmento CSV",
                               st.session_state.t2_delta_editado.to_csv(**CSV_KWARGS_R).encode("utf-8"),
                               file_name=f"{nombre_delta}.csv", use_container_width=True)

        else:
            if "t2_editor_version" not in st.session_state: st.session_state.t2_editor_version = 0
            with st.form("t2_form_edicion"):
                editado = st.data_editor(st.session_state.t2_delta_editado, hide_index=True,
                                         num_rows="fixed", use_container_width=True,
                                         key=f"t2_editor_{st.session_state.t2_editor_version}")
                guardar = st.form_submit_button("💾 Guardar correcciones", use_container_width=True)

            if guardar:
                st.session_state.t2_delta_editado = editado.copy()
                st.session_state.t2_editor_version += 1
                st.success("Cambios guardados correctamente.")
                st.rerun()

            st.download_button("📥 Descargar fragmento corregido",
                               st.session_state.t2_delta_editado.to_csv(**CSV_KWARGS_R).encode("utf-8"),
                               file_name=f"{nombre_delta}.csv", use_container_width=True)

    # ========================================================
    # 4. ENSAMBLAR ARCHIVO FINAL
    # ========================================================
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

    if base_iny and errores_iny and corregido_iny:
        if st.button("🚀 Ensamblar Archivo Final", type="primary", key="t2_ensamblar", use_container_width=True):
            try:
                base = leer_csv_t2(base_iny)
                indices = leer_reporte_t2(errores_iny, len(base))
                if corregido_iny.name.lower().endswith(".xlsx"):
                    corregidas = pd.read_excel(io.BytesIO(corregido_iny.getvalue()),
                                               dtype=str, keep_default_na=False)
                else:
                    corregidas = leer_csv_t2(corregido_iny)

                corregidas = corregidas.fillna("").reset_index(drop=True)
                if len(corregidas) != len(indices):
                    raise ValueError(f"El reporte tiene {len(indices)} líneas, pero el fragmento tiene {len(corregidas)}.")
                if set(base.columns) != set(corregidas.columns):
                    raise ValueError("Las columnas del fragmento no coinciden con las del CSV original.")
                if corregidas.columns.duplicated().any():
                    raise ValueError("El fragmento contiene columnas duplicadas.")

                final = base.copy()
                final.iloc[indices, :] = corregidas[base.columns].to_numpy()
                if len(final) != len(base): raise ValueError("Cambió la cantidad de registros originales.")

                st.session_state.t2_final_bytes = final.to_csv(**CSV_KWARGS_R).encode("utf-8")
                st.session_state.t2_final_nombre = f"{nombre_base_t2(base_iny.name)}_{etiqueta}.csv"
                st.success(f"Archivo final generado: {len(indices)} correcciones, {len(final)} registros.")
            except Exception as e:
                st.session_state.t2_final_bytes = None
                st.error(f"Error al ensamblar: {e}")

    # 5. DESCARGA FINAL
    if st.session_state.get("t2_final_bytes"):
        st.download_button("📥 DESCARGAR CSV FINAL", data=st.session_state.t2_final_bytes,
                           file_name=st.session_state.t2_final_nombre, mime="text/csv",
                           type="primary", use_container_width=True, key="t2_descargar_final")

# ============================================================
# FIN DE PESTAÑA 2
# ============================================================


# ============================================================
# PESTAÑA 3: INYECCIÓN DE NRCs Y CRUCES CON ARGOS
# ============================================================

with tab3:
    st.header("📊 Inyección de NRCs y Cruces con ARGOS")

    # 1. REINICIAR PESTAÑA
    if st.button("🔄 Limpiar / Recomenzar", key="limpiar_t3"):
        claves = ["t3_modo", "t3_argos_completo", "t3_csv_completo", "t3_excel_completo",
                  "t3_argos_rapido", "t3_csv_rapido", "final_argos_zip", "df_cruce_rapido",
                  "t3_alertas", "t3_resultado_editado", "t3_editor_version", "t3_columnas",
                  "t3_firma_completo", "t3_firma_rapido"]
        for k in claves: st.session_state.pop(k, None)
        st.rerun()

    modo = st.radio("Selecciona el proceso:", [
        "📦 Completo (ARGOS + CSV Final + Excel Original)",
        "⚡ Rápido (ARGOS + CSV Final)"
    ], horizontal=True, key="t3_modo")
    st.divider()

    # ========================================================
    # 2. FUNCIONES AUXILIARES
    # ========================================================
    def leer_argos_t3(archivo):
        df = pd.read_csv(io.BytesIO(archivo.getvalue()), dtype=str, encoding="utf-8-sig",
                         keep_default_na=False, on_bad_lines="skip")
        df.columns = [limpiar_nombre_columna(c) for c in df.columns]
        huellas = {normalizar_para_busqueda_t3(c): c for c in df.columns}

        columnas = {
            "periodo": huellas.get("periodo"),
            "nivel": huellas.get("nivel"),
            "cluster": huellas.get("cluster"),
            "subj": huellas.get("area"),
            "crse": huellas.get("nocurso"),
            "grupo": huellas.get("grupo"),
            "nrc": huellas.get("nrc")
        }
        faltantes = [k for k, v in columnas.items() if v is None]
        if faltantes: raise ValueError(f"ARGOS: faltan columnas {faltantes}.")
        if df.columns.duplicated().any(): raise ValueError("ARGOS contiene encabezados duplicados.")

        for campo, col in columnas.items():
            df[f"_{campo}"] = df[col].apply(ultra_limpiar_seccion if campo == "grupo" else ultra_limpiar)

        df = df[df["_nrc"].ne("") & df["_periodo"].ne("") & df["_subj"].ne("") &
                df["_crse"].ne("") & df["_grupo"].ne("")].copy()

        df["_llave"] = df[["_periodo", "_nivel", "_cluster", "_subj", "_crse", "_grupo"]].agg("_".join, axis=1)
        conflictos = df.groupby("_llave")["_nrc"].nunique()
        conflictos = conflictos[conflictos > 1]
        if not conflictos.empty:
            raise ValueError(f"ARGOS tiene {len(conflictos)} combinaciones con NRC distintos. "
                             f"Ejemplos: {conflictos.index.tolist()[:5]}")

        df = df.drop_duplicates(subset=["_llave"])
        return df, dict(zip(df["_llave"], df[columnas["nrc"]]))

    def nivel_csv_t3(cluster):
        c = normalizar_para_cruce(cluster)
        if "POSGRADO" in c: return "POSGRADO"
        if "BACHILLERATO" in c: return "BACHILLERATO"
        return "LICENCIATURA"

    def llaves_csv_t3(df):
        requeridas = ["PERIODO", "SUBJ", "COURSE", "SECCION", "datocomplementario"]
        faltantes = [c for c in requeridas if c not in df.columns]
        if faltantes: raise ValueError(f"CSV: faltan columnas {faltantes}")

        cluster = df["datocomplementario"].apply(ultra_limpiar)
        nivel = df["datocomplementario"].apply(nivel_csv_t3)
        return (df["PERIODO"].apply(ultra_limpiar) + "_" + nivel + "_" + cluster + "_" +
                df["SUBJ"].apply(ultra_limpiar) + "_" + df["COURSE"].apply(ultra_limpiar) + "_" +
                df["SECCION"].apply(ultra_limpiar_seccion))

    def cruzar_nrc_t3(df_csv, mapa):
        llaves = llaves_csv_t3(df_csv)
        nrc = llaves.map(mapa)
        alertas = []
        for llave in llaves[nrc.isna()].unique():
            similares = difflib.get_close_matches(str(llave), list(mapa), n=1, cutoff=0.5)
            alertas.append({"Llave sin NRC": llave,
                            "Coincidencia ARGOS": similares[0] if similares else "Sin coincidencia"})
        return nrc, alertas

    def leer_csv_t3(archivo):
        return pd.read_csv(io.BytesIO(archivo.getvalue()), dtype=str, encoding="utf-8-sig",
                           keep_default_na=False).dropna(how="all").reset_index(drop=True)

    def formatear_excel_t3(ws):
        ws.freeze_panes, ws.auto_filter.ref = "A2", ws.dimensions
        fondo = PatternFill(start_color="1F4E78", fill_type="solid")
        fondo_nrc = PatternFill(start_color="DDEBF7", fill_type="solid")
        encabezados = {str(c.value): c.column for c in ws[1]}
        idx_nrc = encabezados.get("NRC")

        for celda in ws[1]:
            celda.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            celda.fill = fondo
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for fila in ws.iter_rows(min_row=2):
            for celda in fila:
                celda.font = Font(name="Calibri", size=11, bold=(celda.column == idx_nrc))
                celda.alignment = Alignment(horizontal="center", vertical="center")
                if idx_nrc and celda.column == idx_nrc: celda.fill = fondo_nrc

        for columna in ws.columns:
            ancho = max((len(str(c.value)) for c in columna if c.value is not None), default=8)
            ws.column_dimensions[columna[0].column_letter].width = min(max(ancho + 3, 11), 45)

    def firma_t3(*archivos):
        return tuple((a.name, hashlib.sha256(a.getvalue()).hexdigest()) if a else None for a in archivos)

    # ========================================================
    # 3. MODO COMPLETO: ARGOS + CSV + EXCEL ORIGINAL
    # ========================================================
    if modo.startswith("📦"):
        st.subheader("📦 Inyección masiva de NRC en Excel")
        c1, c2, c3 = st.columns(3)
        argos = c1.file_uploader("📊 Reporte ARGOS (.csv)", type=["csv"], key="t3_argos_completo")
        archivos_csv = c2.file_uploader("📝 CSV Finales", type=["csv"], accept_multiple_files=True, key="t3_csv_completo")
        archivos_excel = c3.file_uploader("📁 Excel Originales", type=["xlsx", "xlsm"],
                                         accept_multiple_files=True, key="t3_excel_completo")

        firma = firma_t3(argos, *(archivos_csv or []), *(archivos_excel or []))
        if st.session_state.get("t3_firma_completo") != firma:
            st.session_state.final_argos_zip = None
            st.session_state.t3_alertas = []
            st.session_state.t3_firma_completo = firma

        if argos and archivos_csv and archivos_excel:
            if st.button("🚀 Procesar y generar Excel con NRC", type="primary",
                         use_container_width=True, key="t3_procesar_completo"):
                try:
                    df_argos, mapa = leer_argos_t3(argos)
                    buffer, alertas, procesados, usados = io.BytesIO(), [], 0, set()

                    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
                        for fx in archivos_excel:
                            parejas = [fc for fc in archivos_csv if simplificar_nombre(fc.name) == simplificar_nombre(fx.name)]
                            if len(parejas) != 1:
                                alertas.append(f"{fx.name}: se encontraron {len(parejas)} CSV compatibles.")
                                continue

                            fc = parejas[0]
                            if fc.name in usados:
                                alertas.append(f"{fx.name}: CSV {fc.name} utilizado previamente.")
                                continue

                            df_csv = leer_csv_t3(fc)
                            extension = ".xlsm" if fx.name.lower().endswith(".xlsm") else ".xlsx"
                            wb = openpyxl.load_workbook(io.BytesIO(fx.getvalue()), keep_vba=(extension == ".xlsm"))
                            hoja = next((h for h in wb.sheetnames if h.strip().upper() == HOJA_ALTAS), None)
                            if hoja is None:
                                alertas.append(f"{fx.name}: falta hoja ALTAS.")
                                continue

                            df_excel = pd.read_excel(io.BytesIO(fx.getvalue()), sheet_name=hoja, dtype=str).dropna(how="all")
                            df_excel = df_excel.reset_index(drop=True)
                            if len(df_excel) != len(df_csv):
                                alertas.append(f"{fx.name}: Excel {len(df_excel)} filas y CSV {len(df_csv)} filas.")
                                continue

                            nrc, faltantes = cruzar_nrc_t3(df_csv, mapa)
                            alertas.extend([{"Archivo": fx.name, **a} for a in faltantes])

                            df_nrc = df_excel.copy()
                            equivalencias = {
                                "Periodo": "PERIODO", "Campus": "SEDE", "Subject": "SUBJ", "Course": "COURSE",
                                "Parte de Periodo": "PARTEPERIODO", "Estatus": "STATUS",
                                "Capacidad": "CAPACIDAD", "Sección": "SECCION",
                                "Tipo de Horario": "TIPODEHORARIO", "Método Educativo": "METODO_EDUCATIVO",
                                "Modo de Calificar": "MODODECALIFICAR", "Sesion": "SESION"
                            }
                            for origen, destino in equivalencias.items():
                                if origen in df_nrc and destino in df_csv: df_nrc[origen] = df_csv[destino].values

                            df_nrc.insert(0, "NRC", nrc.values)
                            df_nrc["Grupos"], df_nrc["Socio de Integración"] = "1", "D2L"

                            if HOJA_SALIDA_NRC in wb.sheetnames: del wb[HOJA_SALIDA_NRC]
                            ws = wb.create_sheet(HOJA_SALIDA_NRC)
                            ws.append(list(df_nrc.columns))
                            for fila in df_nrc.itertuples(index=False, name=None):
                                ws.append([None if pd.isna(v) else v for v in fila])

                            formatear_excel_t3(ws)
                            salida = io.BytesIO()
                            wb.save(salida)
                            nombre = fx.name.rsplit(".", 1)[0] + "_con_NRC" + extension
                            z.writestr(nombre, salida.getvalue())
                            usados.add(fc.name)
                            procesados += 1

                    st.session_state.final_argos_zip = buffer.getvalue() if procesados else None
                    st.session_state.t3_alertas = alertas
                    if procesados: st.success(f"Se generaron {procesados} Excel con NRC.")
                    else: st.error("No se pudo procesar ningún Excel.")

                except Exception as e:
                    st.session_state.final_argos_zip = None
                    st.error(f"Error de inyección: {e}")

        alertas = st.session_state.get("t3_alertas") or []
        if alertas:
            with st.expander(f"⚠️ Advertencias ({len(alertas)})"):
                for a in alertas: st.warning(str(a))

        if st.session_state.get("final_argos_zip"):
            st.download_button("📥 Descargar Excel con NRC (.ZIP)", st.session_state.final_argos_zip,
                               file_name="Excels_Finales_con_NRC.zip", mime="application/zip",
                               type="primary", use_container_width=True)

    # ========================================================
    # 4. MODO RÁPIDO: ARGOS + CSV FINAL
    # ========================================================
    else:
        st.subheader("⚡ Cruce rápido de NRC")
        c1, c2 = st.columns(2)
        argos = c1.file_uploader("📊 Reporte ARGOS (.csv)", type=["csv"], key="t3_argos_rapido")
        archivos_csv = c2.file_uploader("📝 CSV Finales", type=["csv"],
                                       accept_multiple_files=True, key="t3_csv_rapido")

        firma = firma_t3(argos, *(archivos_csv or []))
        if st.session_state.get("t3_firma_rapido") != firma:
            st.session_state.df_cruce_rapido = None
            st.session_state.t3_resultado_editado = None
            st.session_state.t3_editor_version = 0
            st.session_state.t3_firma_rapido = firma

        if argos and archivos_csv:
            if st.button("⚡ Cruzar NRC y generar tabla", type="primary",
                         use_container_width=True, key="t3_procesar_rapido"):
                try:
                    df_argos, mapa = leer_argos_t3(argos)
                    resultados, alertas = [], []

                    for archivo in archivos_csv:
                        df = leer_csv_t3(archivo)
                        nrc, faltantes = cruzar_nrc_t3(df, mapa)
                        df.insert(0, "NRC", nrc.values)
                        df = df.rename(columns={"TIPODEHORARIO": "TIPO DE HORARIO",
                                                "METODO_EDUCATIVO": "METODO_ED",
                                                "datocomplementario": "Cluster"})
                        columnas = ["NRC", "PERIODO", "SUBJ", "COURSE", "CAPACIDAD",
                                    "SECCION", "TIPO DE HORARIO", "METODO_ED", "Cluster"]
                        resultados.append(df[[c for c in columnas if c in df.columns]])
                        alertas.extend([{"Archivo": archivo.name, **a} for a in faltantes])

                    resultado = pd.concat(resultados, ignore_index=True)
                    st.session_state.df_cruce_rapido = resultado.copy()
                    st.session_state.t3_resultado_editado = resultado.copy()
                    st.session_state.t3_editor_version = 0
                    st.session_state.t3_alertas = alertas
                    st.success(f"Cruce completado: {len(resultado)} registros.")
                except Exception as e:
                    st.session_state.df_cruce_rapido = None
                    st.error(f"Error en cruce rápido: {e}")

        alertas = st.session_state.get("t3_alertas") or []
        if alertas:
            with st.expander(f"⚠️ NRC no encontrados ({len(alertas)})"):
                st.dataframe(pd.DataFrame(alertas), hide_index=True, use_container_width=True)

        resultado = st.session_state.get("t3_resultado_editado")
        if isinstance(resultado, pd.DataFrame):
            st.subheader("📋 Resultados del cruce")

            with st.form("t3_form_editor"):
                editado = st.data_editor(resultado, hide_index=True, num_rows="fixed",
                                         use_container_width=True,
                                         key=f"t3_editor_{st.session_state.t3_editor_version}")
                guardar = st.form_submit_button("💾 Guardar cambios", use_container_width=True)

            if guardar:
                st.session_state.t3_resultado_editado = editado.copy()
                st.session_state.t3_editor_version += 1
                st.success("Cambios guardados.")
                st.rerun()

            columnas = st.multiselect("Columnas para copiar o descargar",
                                      options=list(resultado.columns), default=list(resultado.columns),
                                      key="t3_columnas")
            if columnas:
                mostrar = st.session_state.t3_resultado_editado[columnas].copy()
                st.dataframe(mostrar, hide_index=True, use_container_width=True)

                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("#### 📋 Copiar a Excel")
                    st.code(mostrar.to_csv(index=False, sep="\t"), language="text")
                with c2:
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        mostrar.to_excel(writer, index=False, sheet_name="Cruce_NRC")
                        formatear_excel_t3(writer.sheets["Cruce_NRC"])

                    st.download_button("📥 Descargar Excel (.xlsx)", buffer.getvalue(),
                                       file_name="Cruce_Rapido_NRC.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       type="primary", use_container_width=True)
            else:
                st.warning("Selecciona al menos una columna.")

# ============================================================
# FIN DE PESTAÑA 3
# ============================================================
