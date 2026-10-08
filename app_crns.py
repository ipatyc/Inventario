
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
import datetime
import difflib
from difflib import SequenceMatcher
from openpyxl.styles import Font, PatternFill, Alignment

# ============================================================
# 1. CONFIGURACIÓN GENERAL
# ============================================================

HOJA_ALTAS = "ALTAS"
HOJA_SALIDA_NRC = "NRC"
UMBRAL_FUZZY = 0.82

CSV_KWARGS_R = {
    "index": False,
    "encoding": "utf-8",
    "sep": ",",
    "lineterminator": "\n"
}

COLUMNAS_CLUSTER_FINAL = [
    "Periodo", "CRN", "Tipo.de.Reunión", "Fecha.Inicio", "Fecha.Fin",
    "Dom", "Lun", "Mar", "Mie", "Jue", "Vie", "Sab",
    "horarioIni", "horarioFin", "Inicio.de.sesión", "edificio",
    "salon", "Tipo.de.horario", "indCategoria", "idInstructor",
    "responsabilidad", "Ind.principal", "ind.sobre.paso",
    "datocomplementario"
]

# ============================================================
# 2. FUNCIONES DE LIMPIEZA
# ============================================================

def quitar_acentos(t):
    if t is None or pd.isna(t): return ""
    return "".join(c for c in unicodedata.normalize("NFD", str(t))
                   if unicodedata.category(c) != "Mn")

def normalizar_para_cruce(t):
    if t is None or pd.isna(t): return ""
    s = str(t).strip()
    if s.endswith(".0"): s = s[:-2]
    return quitar_acentos(s).upper().strip()

def similitud(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def limpiar_clave_texto(val):
    if val is None or pd.isna(val): return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none", ""): return ""
    return s[:-2] if s.endswith(".0") else s

def format_r_string(val):
    s = limpiar_clave_texto(val)
    return s if s else np.nan

def limpia_seccion_interna(x):
    s = limpiar_clave_texto(x)
    if s.isdigit(): return f"{int(s):02d}"
    return s

def sin_espacios(x):
    return limpiar_clave_texto(x).replace(" ", "").upper()

def limpiar_espacios_y_mayusculas(x):
    if x is None or pd.isna(x): return ""
    return re.sub(r"\s+", " ", str(x)).strip().upper()

def normalizar_para_busqueda(texto):
    return re.sub(r"[^a-z0-9]", "", quitar_acentos(texto).lower())

def limpiar_nombre_columna(col):
    return " ".join(str(col).split()) if not pd.isna(col) else ""

def normalizar_para_busqueda_t3(texto):
    return normalizar_para_busqueda(texto)

def ultra_limpiar(x):
    return sin_espacios(x)

def ultra_limpiar_seccion(x):
    return limpia_seccion_interna(x).upper().replace(" ", "")

def corregir_nivel_por_cluster_csv(cluster_val):
    c = normalizar_para_cruce(cluster_val)
    if "POSGRADO" in c: return "POSGRADO"
    if "BACHILLERATO" in c: return "BACHILLERATO"
    return "LICENCIATURA"

def simplificar_nombre(nombre):
    n = nombre.lower()
    for basura in [".xlsx", ".xls", ".csv", "_final", "_base", "_v1",
                   "_v2", "_v3", "_v4", "corregidas_", "errores_"]:
        n = n.replace(basura, "")
    return n.strip().replace(" ", "")

def clave_seccion(valor):
    s = limpiar_clave_texto(valor)
    if s.isdigit() and int(s) <= 99: return f"{int(s):02d}"
    return s.upper()

# ============================================================
# 3. LECTURA DEL CATÁLOGO PA (OPCIONAL)
# ============================================================

def leer_csv_pa(archivo):
    """
    PA:
    Periodo = Periodo
    Área = SUBJ
    No. Curso = CRSE
    Grupo = Sección
    NRC = Identificador único

    No modifica el CSV original.
    """

    datos = archivo.getvalue() if hasattr(archivo, "getvalue") else archivo
    df = None

    for encoding in ["utf-8-sig", "utf-8", "cp1252"]:
        try:
            df = pd.read_csv(io.BytesIO(datos), dtype=str, encoding=encoding,
                             keep_default_na=False, sep=None, engine="python")
            break
        except (UnicodeError, pd.errors.ParserError):
            continue

    if df is None: raise ValueError("No se pudo leer el archivo PA.")

    # Normalizar encabezados, incluyendo acentos
    df.columns = [normalizar_para_busqueda(str(c).replace("Ã", "A")) for c in df.columns]
    requeridas = ["periodo", "area", "nocurso", "grupo", "nrc"]
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes: raise ValueError(f"PA: faltan columnas {faltantes}")

    pa = df.copy()
    for col in ["periodo", "area", "nocurso", "nrc"]:
        pa[col] = pa[col].apply(normalizar_para_cruce)
    pa["grupo"] = pa["grupo"].apply(clave_seccion)

    pa = pa[pa["nrc"].ne("")].copy()
    llaves = ["periodo", "area", "nocurso", "grupo"]

    conflictos = pa.groupby("nrc")[llaves].nunique(dropna=False).gt(1).any(axis=1)
    nrc_conflictivos = set(conflictos[conflictos].index)

    # Un NRC repetido idéntico se toma una vez; uno contradictorio se conserva para revisión
    pa_unica = pa.drop_duplicates(subset=["nrc"])
    pa_reserva = pd.concat([
        pa_unica[~pa_unica["nrc"].isin(nrc_conflictivos)],
        pa[pa["nrc"].isin(nrc_conflictivos)]
    ]).drop_duplicates(subset=["nrc"] + llaves)

    advertencias = [
        {"Tipo": "PA", "Detalle": f"NRC {nrc} tiene registros contradictorios; revisar."}
        for nrc in sorted(nrc_conflictivos)
    ]
    return pa_reserva, advertencias

# ============================================================
# 4. ASIGNACIÓN GLOBAL DE SECCIONES
# ============================================================

def asignar_secciones_globales(df_altas, pa=None):
    """
    Revisa TODAS las ALTAS juntas, aunque sean varios Excel.
    Llave: Periodo + SUBJ + CRSE.
    PA opcional: reserva grupos existentes.
    Licenciatura y Posgrado: 01 a 99.
    Bachillerato: conserva sección original.
    """

    requeridas = ["Periodo", "Subject", "Course", "Nivel", "Sección"]
    faltantes = [c for c in requeridas if c not in df_altas.columns]
    if faltantes: raise ValueError(f"ALTAS: faltan columnas {faltantes}")

    salida = df_altas.copy().reset_index(drop=True)
    salida["Sección Original"] = salida["Sección"].apply(clave_seccion)
    salida["Sección"] = salida["Sección Original"]
    ocupadas, errores = {}, []

    # Reservar secciones de PA
    if pa is not None:
        for _, fila in pa.iterrows():
            llave = tuple(normalizar_para_cruce(fila[c]) for c in ["periodo", "area", "nocurso"])
            grupo = clave_seccion(fila["grupo"])
            if all(llave) and grupo: ocupadas.setdefault(llave, set()).add(grupo)

    # Numerar todas las ALTAS de manera global
    for i, fila in salida.iterrows():
        nivel = normalizar_para_cruce(fila["Nivel"])
        if nivel not in ["LICENCIATURA", "POSGRADO"]: continue

        llave = tuple(normalizar_para_cruce(fila[c]) for c in ["Periodo", "Subject", "Course"])
        if not all(llave):
            errores.append({"Fila": i + 1, "Tipo": "Sección",
                            "Detalle": "Periodo, SUBJ o CRSE vacío; no se puede asignar sección."})
            salida.at[i, "Sección"] = ""
            continue

        usadas = ocupadas.setdefault(llave, set())
        nueva = next((f"{n:02d}" for n in range(1, 100) if f"{n:02d}" not in usadas), None)

        if nueva is None:
            errores.append({"Fila": i + 1, "Tipo": "Sección",
                            "Detalle": f"Sin secciones disponibles 01-99 para {llave}."})
            salida.at[i, "Sección"] = ""
        else:
            salida.at[i, "Sección"] = nueva
            usadas.add(nueva)

    salida["Sección Modificada"] = salida["Sección Original"] != salida["Sección"]
    return salida, pd.DataFrame(errores, columns=["Fila", "Tipo", "Detalle"])

# ============================================================
# 5. RESTRICCIONES DE ALTAS
# ============================================================

def leer_restricciones_excel(archivo):
    """Carga todas las hojas del Excel de restricciones sin modificar sus datos."""
    libro = pd.ExcelFile(archivo)
    reglas = {}
    for hoja in libro.sheet_names:
        df = libro.parse(hoja, dtype=str).dropna(how="all")
        df.columns = [limpiar_nombre_columna(c) for c in df.columns]
        reglas[hoja] = df
    return reglas

def auditar_restricciones(df_altas, reglas):
    """
    Validación preliminar de periodo, campus, requerimiento y estatus.
    Las reglas complejas por nivel, clúster y parte de periodo
    requieren integración adicional.
    """
    columnas = ["Fila", "Tipo", "Detalle"]
    if not reglas: return pd.DataFrame(columns=columnas)

    tabla = next((t for t in reglas.values()
                  if "Periodo" in t.columns and "Campus" in t.columns), None)
    if tabla is None:
        return pd.DataFrame([{"Fila": "", "Tipo": "Restricciones",
                              "Detalle": "No se encontró una tabla con Periodo y Campus."}])

    incidencias = []
    for i, fila in df_altas.reset_index(drop=True).iterrows():
        periodo = normalizar_para_cruce(fila.get("Periodo"))[-2:]
        posibles = tabla[tabla["Periodo"].apply(normalizar_para_cruce) == periodo]

        if posibles.empty:
            incidencias.append({"Fila": i + 1, "Tipo": "Periodo",
                                "Detalle": f"Periodo {periodo}: sin regla definida."})
            continue

        for col in ["Campus", "Requerimiento", "Estatus"]:
            if col not in posibles.columns or col not in fila.index: continue
            opciones = set()
            for valor in posibles[col].dropna():
                opciones.update(normalizar_para_cruce(v) for v in str(valor).split("/"))
            opciones.discard("")
            actual = normalizar_para_cruce(fila.get(col))
            if actual and opciones and actual not in opciones:
                incidencias.append({"Fila": i + 1, "Tipo": col,
                                    "Detalle": f"{actual} no permitido para periodo {periodo}. "
                                               f"Opciones: {', '.join(sorted(opciones))}"})

    return pd.DataFrame(incidencias, columns=columnas)

# ============================================================
# 6. ESTADOS GLOBALES DE STREAMLIT
# ============================================================

ESTADOS_INICIALES = {
    "original_files_bytes": {},
    "res_auditoria": None,
    "raw_altas": None,
    "ready_for_download": False,
    "zip_file_bytes": None,
    "csv_files_to_download": {},
    "csv_consolidado_bytes": None,
    "delta_files": {},
    "final_argos_zip": None,
    "df_cruce_rapido": None,
    "df_delta_cache": None,
    "nombre_delta_cache": None,
    "llave_control_archivos": "",
    "archivo_final_bytes": None,
    "archivo_final_nombre": None,
    "cat_avanzado_cache": None,
    "cat_avanzado_firma": None,
    "indice_nombres_avanzado": None,
    "cat_pa_cache": None,
    "cat_pa_firma": None,
    "advertencias_pa": [],
    "reglas_restricciones": {},
    "errores_restricciones": None,
    "df_secciones_corregidas": None,
    "errores_secciones": None,
    "resumen_secciones": None,
    "secciones_validadas": False
}

for clave, valor in ESTADOS_INICIALES.items():
    if clave not in st.session_state:
        st.session_state[clave] = valor

# ============================================================
# 7. CONFIGURACIÓN VISUAL
# ============================================================

st.set_page_config(page_title="Consola Iris Cavazos", page_icon="⚙️", layout="wide")
st.title("⚙️ Consola de Control de Materias e Inyección de NRCs")
st.markdown("---")

# ============================================================
# 8. CREACIÓN DE PESTAÑAS
# ============================================================

tab1, tab_err, tab3 = st.tabs([
    "1️⃣ Proceso: Validación y Generar CSVs",
    "⚠️ Reporte de Errores (Extraer Delta)",
    "2️⃣ Proceso: Inyección de NRCs Masiva (ARGOS)"
])

# ============================================================
# NOMBRES Y CARPETAS: SOLICITUDES + CSV_ALTAS
# ============================================================

def nombre_seguro(valor):
    texto = limpiar_clave_texto(valor)
    texto = re.sub(r'[\\/:*?"<>|]', "", texto)
    return re.sub(r"\s+", " ", texto).strip()

def nombre_archivo_altas(periodo, responsable, tipo, extension):
    periodo, responsable = nombre_seguro(periodo), nombre_seguro(responsable)
    if not periodo or not responsable:
        raise ValueError("Falta Periodo o Responsable para generar el nombre.")
    if tipo == "SOLICITUDES": return f"{periodo} {responsable} CRNs U-ERRE{extension}"
    if tipo == "CSV_ALTAS": return f"{periodo} {responsable} CSV_ALTAS.csv"
    raise ValueError(f"Tipo de archivo desconocido: {tipo}")

def crear_solicitud_excel(df_grupo, archivos_originales):
    """
    Conserva el libro original y actualiza su hoja ALTAS.
    Cada grupo debe proceder de un solo archivo original.
    """
    origenes = df_grupo["ArchivoOrigen"].dropna().unique().tolist()
    if len(origenes) != 1:
        raise ValueError("La solicitud reúne varios Excel originales; no se pueden combinar sus macros automáticamente.")

    origen = origenes[0]
    if origen not in archivos_originales:
        raise ValueError(f"No se encontró el Excel original: {origen}")

    extension = ".xlsm" if origen.lower().endswith(".xlsm") else ".xlsx"
    wb = openpyxl.load_workbook(io.BytesIO(archivos_originales[origen]), keep_vba=(extension == ".xlsm"))
    hoja = next((h for h in wb.sheetnames if h.strip().upper() == "ALTAS"), None)
    if hoja is None: raise ValueError(f"{origen} no contiene la hoja ALTAS.")

    ws = wb[hoja]
    encabezados = {normalizar_para_busqueda(c.value): c.column for c in ws[1] if c.value is not None}
    equivalencias = {
        "Periodo": ["periodo"], "Responsable": ["responsable"],
        "Subject": ["subject", "subj", "area"], "Course": ["course", "crse", "nocurso"],
        "Sección": ["seccion", "grupo"], "Tipo de Horario": ["tipodehorario"],
        "Método Educativo": ["metodoeducativo"], "Modo de Calificar": ["mododecalificar"]
    }

    # Solo actualizar filas que realmente pertenecen al grupo.
    for _, fila in df_grupo.iterrows():
        numero_excel = fila.get("_FilaExcel")
        if pd.isna(numero_excel): raise ValueError(f"Falta _FilaExcel para {origen}.")
        numero_excel = int(numero_excel)

        for campo, alternativas in equivalencias.items():
            if campo not in df_grupo.columns: continue
            columna = next((encabezados[a] for a in alternativas if a in encabezados), None)
            if columna is not None:
                valor = fila.get(campo)
                ws.cell(numero_excel, columna).value = None if pd.isna(valor) else str(valor)

    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue(), extension

def generar_zip_altas(df_final, archivos_originales, funcion_csv):
    """
    Genera SOLICITUDES y CSV_ALTAS agrupados por Periodo + Responsable.
    df_final ya debe contener las secciones corregidas globalmente.
    """
    obligatorias = ["Periodo", "Responsable", "ArchivoOrigen", "_FilaExcel", "Sección"]
    faltantes = [c for c in obligatorias if c not in df_final.columns]
    if faltantes: raise ValueError(f"Faltan columnas: {faltantes}")

    df = df_final.copy()
    df["Periodo"] = df["Periodo"].apply(limpiar_clave_texto)
    df["Responsable"] = df["Responsable"].apply(nombre_seguro)

    if df["Periodo"].eq("").any() or df["Responsable"].eq("").any():
        raise ValueError("Hay registros sin Periodo o Responsable.")

    # No permitir dos archivos con el mismo nombre dentro del ZIP.
    nombres_usados = set()
    buffer = io.BytesIO()

    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zip_salida:
        for (periodo, responsable), grupo in df.groupby(["Periodo", "Responsable"], sort=False):
            nombre_csv = nombre_archivo_altas(periodo, responsable, "CSV_ALTAS", ".csv")
            ruta_csv = f"CSV_ALTAS/{nombre_csv}"
            if ruta_csv in nombres_usados: raise ValueError(f"Nombre CSV duplicado: {ruta_csv}")

            contenido_csv = funcion_csv(grupo)
            if isinstance(contenido_csv, str): contenido_csv = contenido_csv.encode("utf-8")
            zip_salida.writestr(ruta_csv, contenido_csv)
            nombres_usados.add(ruta_csv)

            # No se fusionan libros con macros de distintos orígenes.
            origenes = grupo["ArchivoOrigen"].dropna().unique().tolist()
            if len(origenes) != 1:
                raise ValueError(
                    f"{periodo} {responsable} aparece en varios Excel. "
                    "El CSV puede consolidarse, pero la solicitud XLSM requiere conservar cada libro por separado."
                )

            contenido_excel, extension = crear_solicitud_excel(grupo, archivos_originales)
            nombre_excel = nombre_archivo_altas(periodo, responsable, "SOLICITUDES", extension)
            ruta_excel = f"SOLICITUDES/{nombre_excel}"
            if ruta_excel in nombres_usados: raise ValueError(f"Nombre Excel duplicado: {ruta_excel}")

            zip_salida.writestr(ruta_excel, contenido_excel)
            nombres_usados.add(ruta_excel)

    return buffer.getvalue()


# ============================================================
# PESTAÑA 1: VALIDACIÓN DE ALTAS, PA Y GENERACIÓN DE CSV
# ============================================================

with tab1:
    st.header("⚙️ Validación de ALTAS y Generación de CSV")

    # 1. REINICIAR
    if st.button("🔄 Limpiar / Recomenzar", key="limpiar_t1"):
        prefijos = ("manual_", "edit_nom_", "edit_met_")
        claves = [
            "cat_ext_t1", "pa_t1", "restr_t1", "altas_t1", "modo_csv_t1",
            "raw_altas", "res_auditoria", "cat_pa_cache", "advertencias_pa",
            "reglas_restricciones", "errores_restricciones", "errores_secciones",
            "df_secciones_corregidas", "ready_for_download", "zip_file_bytes",
            "csv_consolidado_bytes", "df_manual_fijo", "cat_avanzado_cache",
            "cat_avanzado_firma", "indice_nombres_avanzado"
        ]
        claves += [k for k in st.session_state if k.startswith(prefijos)]
        for k in claves: st.session_state.pop(k, None)
        st.rerun()

    # 2. CARGA DE ARCHIVOS
    st.subheader("📁 Archivos de entrada")
    c1, c2, c3, c4 = st.columns(4)
    file_cat_ext = c1.file_uploader("📚 Catálogo Avanzado", type=["csv", "xlsx"], key="cat_ext_t1")
    file_pa = c2.file_uploader("📊 PA (Opcional)", type=["csv"], key="pa_t1")
    file_restr = c3.file_uploader("📋 Restricciones (Opcional)", type=["xlsx"], key="restr_t1")
    files_altas = c4.file_uploader("📁 Archivos ALTAS", type=["xlsx", "xlsm"], accept_multiple_files=True, key="altas_t1")

    st.info("Con PA se reservan los grupos existentes. Sin PA se revisan todas las ALTAS juntas.")
    modo_csv = st.radio("Salida CSV", ["Un CSV por cada Excel", "Un solo CSV consolidado"],
                        horizontal=True, key="modo_csv_t1") if files_altas else "Un CSV por cada Excel"

    COLUMNAS_ALTAS = ["Periodo", "Campus", "Subject", "Course", "Nivel", "Nombre de la Materia",
                      "Parte de Periodo", "Estatus", "Capacidad", "Sección", "Tipo de Horario",
                      "Método Educativo", "Modo de Calificar", "Sesion", "Clúster"]
    COLUMNAS_CSV = ["PERIODO", "SEDE", "SUBJ", "COURSE", "PARTEPERIODO", "STATUS", "CAPACIDAD",
                    "GRUPOS", "SECCION", "TIPODEHORARIO", "METODO_EDUCATIVO",
                    "SOCIODEINTEGRACION", "MODODECALIFICAR", "SESION", "datocomplementario"]
    ALIAS = {"area": "Subject", "subj": "Subject", "nocurso": "Course", "crse": "Course",
             "materia": "Nombre de la Materia", "grupo": "Sección", "sede": "Campus",
             "cupo": "Capacidad", "status": "Estatus"}
    MAPA = {normalizar_para_busqueda(c): c for c in COLUMNAS_ALTAS}
    MAPA.update(ALIAS)

    CLUSTERS = ["Ingenieria", "Bachillerato", "Negocios", "Ciencias Exactas", "Posgrado Online",
                "Humanidades", "Idiomas y ADN", "TJYG", "Smart Cities", "Ejecutivas", "EGEL",
                "Intercambio", "Consejeria", "Posgrado"]

    def cluster_oficial(valor):
        original = limpiar_clave_texto(valor)
        return next((c for c in CLUSTERS if normalizar_para_busqueda(c) ==
                     normalizar_para_busqueda(original)), original)

    # 3. CATÁLOGO AVANZADO
    def cargar_catalogo_avanzado():
        if file_cat_ext is None: return {}
        firma = (file_cat_ext.name, file_cat_ext.size, hash(file_cat_ext.getvalue()))
        if st.session_state.get("cat_avanzado_firma") == firma:
            return st.session_state.get("cat_avanzado_cache") or {}

        datos = io.BytesIO(file_cat_ext.getvalue())
        df = pd.read_csv(datos, dtype=str, encoding="utf-8-sig", keep_default_na=False) \
            if file_cat_ext.name.lower().endswith(".csv") else pd.read_excel(datos, dtype=str).fillna("")
        df.columns = [str(c).strip().upper() for c in df.columns]
        faltantes = set(["SCBCRSE_SUBJ_CODE", "SCBCRSE_CRSE_NUMB"]) - set(df.columns)
        if faltantes: raise ValueError(f"Catálogo Avanzado: faltan {sorted(faltantes)}")

        catalogo = {}
        for _, f in df.iterrows():
            llave = (sin_espacios(f.get("SCBCRSE_SUBJ_CODE")), sin_espacios(f.get("SCBCRSE_CRSE_NUMB")))
            if not all(llave): continue
            info = catalogo.setdefault(llave, {"titles": set(), "schd": set(), "insm": set(),
                                                "gmod": set(), "pares": set()})
            for campo in ["SCBCRSE_TITLE", "SCRSYLN_LONG_COURSE_TITLE"]:
                titulo = limpiar_espacios_y_mayusculas(f.get(campo, ""))
                if titulo and titulo != "NAN": info["titles"].add(titulo)

            h, m = sin_espacios(f.get("SCRSCHD_SCHD_CODE")), sin_espacios(f.get("SCRSCHD_INSM_CODE"))
            g = sin_espacios(f.get("SCRGMOD_GMOD_CODE"))
            if h: info["schd"].add(h)
            if m: info["insm"].add(m)
            if g: info["gmod"].add(g)
            if h or m: info["pares"].add((h, m))

        st.session_state.cat_avanzado_cache = catalogo
        st.session_state.cat_avanzado_firma = firma
        return catalogo

    # 4. VALIDACIÓN DE MATERIAS
    def validar_materia(fila, catalogo):
        subj, crse = sin_espacios(fila.get("Subject")), sin_espacios(fila.get("Course"))
        nombre = limpiar_espacios_y_mayusculas(fila.get("Nombre de la Materia"))
        horario, metodo = sin_espacios(fila.get("Tipo de Horario")), sin_espacios(fila.get("Método Educativo"))
        modo = sin_espacios(fila.get("Modo de Calificar"))

        r = {"Luz Verde": False, "Materia Excel": nombre, "Materia Catálogo": "",
             "Subj Original": subj, "Crse Original": crse, "Subj Sugerido": subj, "Crse Sugerido": crse,
             "Horario Original": horario, "Horario Sugerido": horario,
             "Método Original": metodo, "Método Sugerido": metodo,
             "Modo de Calificar Original": modo, "Modo de Calificar Sugerido": modo,
             "Comentario Nombres": "", "Comentario Horario": "", "Comentario Método": "",
             "Comentario Modo de Calificar": ""}

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
            r["Comentario Nombres"] = f"Claves sugeridas (similitud {p:.0%})"

        info = catalogo[llave]
        titulos = sorted(info["titles"])
        r["Materia Catálogo"] = titulos[0] if titulos else ""
        if not r["Comentario Nombres"]:
            nombres = {normalizar_para_cruce(t) for t in titulos}
            r["Comentario Nombres"] = "Todo correcto" if normalizar_para_cruce(nombre) in nombres \
                else "Clave OK, pero Nombre difiere"

        validaciones = [
            ("schd", horario, "Horario", "Comentario Horario", "Horario Sugerido"),
            ("insm", metodo, "Método", "Comentario Método", "Método Sugerido"),
            ("gmod", modo, "Modo de calificar", "Comentario Modo de Calificar", "Modo de Calificar Sugerido")
        ]
        for campo, actual, etiqueta, comentario, sugerido in validaciones:
            permitidos = info[campo]
            if not permitidos: r[comentario] = "Sin restricciones en catálogo"
            elif actual in permitidos: r[comentario] = f"{etiqueta} OK"
            else:
                r[comentario] = "Error. Permitidos: " + ", ".join(sorted(permitidos))
                r[sugerido] = next(iter(permitidos)) if len(permitidos) == 1 else ""

        if horario and metodo and info["pares"] and (horario, metodo) not in info["pares"]:
            r["Comentario Método"] += " | Combinación horario-método no encontrada"
        return r

    # 5. VALIDACIÓN MASIVA
    if files_altas and file_cat_ext and st.button("⚡ Ejecutar Validación Inteligente", type="primary", key="validar_t1"):
        try:
            catalogo = cargar_catalogo_avanzado()
            if not catalogo: raise ValueError("Catálogo Avanzado sin materias válidas.")
            pa, avisos = leer_csv_pa(file_pa) if file_pa else (None, [])
            reglas = leer_restricciones_excel(file_restr) if file_restr else {}

            # Leer ALTAS y conservar los Excel originales con macros
            piezas, errores_archivo = [], []
            st.session_state.original_files_bytes = {}
            
            for archivo in files_altas:
                st.session_state.original_files_bytes[archivo.name] = archivo.getvalue()
                libro = pd.ExcelFile(io.BytesIO(archivo.getvalue()))
                hojas = [h for h in libro.sheet_names if h.strip().upper() == HOJA_ALTAS]
            
                if not hojas:
                    errores_archivo.append(f"{archivo.name}: falta hoja ALTAS")
                    continue
            
                df = libro.parse(hojas[0], dtype=str)
                df["_FilaExcel"] = range(2, len(df) + 2)
                df.columns = [MAPA.get(normalizar_para_busqueda(c), c) for c in df.columns]
                df = df.dropna(how="all").copy()
            
                faltantes = set(["Periodo", "Subject", "Course", "Nivel", "Sección", "Responsable"]) - set(df.columns)
                if faltantes or df.columns.duplicated().any():
                    errores_archivo.append(f"{archivo.name}: columnas faltantes {sorted(faltantes)} o duplicadas")
                    continue
            
                df = df.dropna(subset=["Periodo", "Subject", "Course"], how="all").copy()
                df["ArchivoOrigen"] = archivo.name
                piezas.append(df)

            if errores_archivo: raise ValueError("\n".join(errores_archivo))
            if not piezas: raise ValueError("No hay registros ALTAS válidos.")

            total = pd.concat(piezas, ignore_index=True)
            auditoria = []
            for i, fila in total.iterrows():
                r = validar_materia(fila, catalogo)
                r.update({"idx": i, "Archivo": fila["ArchivoOrigen"]})
                auditoria.append(r)

            st.session_state.raw_altas = total
            st.session_state.res_auditoria = pd.DataFrame(auditoria)
            st.session_state.cat_pa_cache = pa
            st.session_state.advertencias_pa = avisos
            st.session_state.reglas_restricciones = reglas
            st.session_state.errores_restricciones = auditar_restricciones(total, reglas)
            st.session_state.ready_for_download = False
            st.success(f"Validación terminada: {len(total)} registros en {len(piezas)} archivo(s).")
        except Exception as e:
            st.error(f"Error de validación: {e}")

    # 6. MESA DE CONTROL
    if st.session_state.get("res_auditoria") is not None:
        st.divider()
        st.subheader("⚖️ Mesa de Control")
        auditoria = st.session_state.res_auditoria

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

            with st.expander(f"⚠️ {archivo} — {len(pendientes)} advertencias"):
                with st.form(f"form_{archivo}"):
                    t1, t2 = st.tabs(["Nombres y Claves", "Horarios y Métodos"])
                    cols_nom = ["Luz Verde", "Materia Excel", "Materia Catálogo", "Comentario Nombres",
                                "Subj Original", "Crse Original", "Subj Sugerido", "Crse Sugerido"]
                    cols_met = ["Luz Verde", "Materia Excel", "Horario Original", "Horario Sugerido",
                                "Comentario Horario", "Método Original", "Método Sugerido",
                                "Comentario Método", "Modo de Calificar Original",
                                "Modo de Calificar Sugerido", "Comentario Modo de Calificar"]
                    with t1:
                        nom = st.data_editor(pendientes[cols_nom], hide_index=True, use_container_width=True,
                                             disabled=cols_nom[1:6], key=f"edit_nom_{archivo}")
                    with t2:
                        met = st.data_editor(pendientes[cols_met], hide_index=True, use_container_width=True,
                                             disabled=["Materia Excel", "Horario Original", "Comentario Horario",
                                                       "Método Original", "Comentario Método",
                                                       "Modo de Calificar Original", "Comentario Modo de Calificar"],
                                             key=f"edit_met_{archivo}")

                    if st.form_submit_button("💾 Confirmar correcciones"):
                        for i in pendientes.index:
                            auditoria.at[i, "Luz Verde"] = bool(nom.at[i, "Luz Verde"] or met.at[i, "Luz Verde"])
                            for c in ["Subj Sugerido", "Crse Sugerido"]: auditoria.at[i, c] = nom.at[i, c]
                            for c in ["Horario Sugerido", "Método Sugerido", "Modo de Calificar Sugerido"]:
                                auditoria.at[i, c] = met.at[i, c]
                        st.session_state.res_auditoria = auditoria
                        st.session_state.ready_for_download = False
                        st.rerun()

        for titulo, datos in [
            ("⚠️ Advertencias PA", st.session_state.get("advertencias_pa")),
            ("📋 Restricciones", st.session_state.get("errores_restricciones"))
        ]:
            if isinstance(datos, list) and datos: st.subheader(titulo); st.dataframe(pd.DataFrame(datos), hide_index=True)
            elif isinstance(datos, pd.DataFrame) and not datos.empty:
                st.subheader(titulo); st.dataframe(datos, hide_index=True, use_container_width=True)

        # 7. APLICAR CORRECCIONES Y SECCIONES GLOBALES
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
        st.subheader("🔢 Validación global de secciones")
        if st.button("🔄 Recalcular secciones", key="recalcular_t1"):
            try:
                df_sec, errores = construir_altas_corregidas()
                st.session_state.df_secciones_corregidas = df_sec
                st.session_state.errores_secciones = errores
                st.session_state.ready_for_download = False
            except Exception as e: st.error(str(e))

        df_sec = st.session_state.get("df_secciones_corregidas")
        if isinstance(df_sec, pd.DataFrame):
            cols = ["ArchivoOrigen", "Periodo", "Subject", "Course", "Nivel",
                    "Sección Original", "Sección", "Sección Modificada"]
            st.dataframe(df_sec[[c for c in cols if c in df_sec]], hide_index=True, use_container_width=True)
            st.metric("Secciones modificadas", int(df_sec["Sección Modificada"].sum()))

        errores_sec = st.session_state.get("errores_secciones")
        if isinstance(errores_sec, pd.DataFrame) and not errores_sec.empty:
            st.error("Hay secciones sin asignar."); st.dataframe(errores_sec, hide_index=True)

        # 8. GENERACIÓN CSV
        def preparar_csv_banner(df):
            r = pd.DataFrame(index=df.index)
            equivalencias = {
                "PERIODO": "Periodo", "SEDE": "Campus", "SUBJ": "Subject", "COURSE": "Course",
                "PARTEPERIODO": "Parte de Periodo", "STATUS": "Estatus",
                "CAPACIDAD": "Capacidad", "SECCION": "Sección", "TIPODEHORARIO": "Tipo de Horario",
                "METODO_EDUCATIVO": "Método Educativo", "MODODECALIFICAR": "Modo de Calificar",
                "SESION": "Sesion"
            }
            for destino, origen in equivalencias.items(): r[destino] = df[origen].apply(format_r_string)
            for c in ["SUBJ", "COURSE", "TIPODEHORARIO", "METODO_EDUCATIVO"]:
                r[c] = r[c].apply(sin_espacios)
            r["SECCION"] = df["Sección"].apply(clave_seccion)
            r["CAPACIDAD"] = pd.to_numeric(r["CAPACIDAD"], errors="coerce").astype("Int64")
            r["GRUPOS"], r["SOCIODEINTEGRACION"] = 1, "D2L"
            r["datocomplementario"] = df.apply(
                lambda f: "Bachillerato" if normalizar_para_cruce(f["Nivel"]) == "BACHILLERATO"
                else cluster_oficial(f.get("Clúster")), axis=1)
            return r[COLUMNAS_CSV].fillna("").to_csv(**CSV_KWARGS_R)

        if st.button("💾 Generar CSV", type="primary", key="generar_csv_t1"):
            try:
                corregido, errores = construir_altas_corregidas()
                if not errores.empty: raise ValueError("Hay secciones sin asignar. Revisa los errores.")
                faltantes = set(COLUMNAS_ALTAS) - set(corregido.columns)
                if faltantes: raise ValueError(f"Faltan columnas: {sorted(faltantes)}")
                st.session_state.zip_file_bytes = None
                st.session_state.csv_consolidado_bytes = None

                if modo_csv == "Un CSV por cada Excel":
                    buffer = io.BytesIO()
                    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
                        for nombre, sub in corregido.groupby("ArchivoOrigen", sort=False):
                            z.writestr(nombre.rsplit(".", 1)[0] + ".csv",
                                       preparar_csv_banner(sub).encode("utf-8"))
                    st.session_state.zip_file_bytes = buffer.getvalue()
                else:
                    st.session_state.csv_consolidado_bytes = preparar_csv_banner(corregido).encode("utf-8")

                st.session_state.ready_for_download = True
                st.session_state.modo_salida_csv_generado = modo_csv
                st.success("CSV generados.")
            except Exception as e: st.error(str(e))

        if st.session_state.get("ready_for_download"):
            if st.session_state.modo_salida_csv_generado == "Un CSV por cada Excel":
                st.download_button("📥 Descargar ZIP", st.session_state.zip_file_bytes,
                                   file_name="archivos_carga_banner.zip", mime="application/zip")
            else:
                st.download_button("📥 Descargar CSV", st.session_state.csv_consolidado_bytes,
                                   file_name="archivos_carga_banner.csv", mime="text/csv")

    # 9. BUSCADOR MANUAL
    st.divider()
    st.subheader("🔍 Buscador y creación manual")
    if "df_manual_fijo" not in st.session_state:
        st.session_state.df_manual_fijo = pd.DataFrame(columns=COLUMNAS_CSV)
    if "manual_candidatos" not in st.session_state: st.session_state.manual_candidatos = []

    with st.form("buscar_manual_t1"):
        b1, b2, b3 = st.columns(3)
        nombre = b1.text_input("Nombre", key="manual_nombre")
        subj = b2.text_input("SUBJ", key="manual_subj")
        crse = b3.text_input("COURSE", key="manual_crse")
        buscar = st.form_submit_button("🪄 Buscar opciones")

    if buscar:
        if not file_cat_ext: st.warning("Carga el Catálogo Avanzado.")
        elif not any([nombre, subj, crse]): st.warning("Escribe al menos un criterio.")
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
        llave = st.selectbox("Materia correcta", opciones, format_func=lambda x: etiquetas[x], key="manual_materia")
        confirmar = st.checkbox(f"Confirmo: {etiquetas[llave]}", key=f"manual_confirmar_{llave}")
        info = cargar_catalogo_avanzado()[llave]

        d1, d2 = st.columns(2)
        with d1:
            horario = st.selectbox("Tipo de horario", [""] + sorted(info["schd"]), key="manual_horario")
            periodo = st.text_input("Periodo", key="manual_periodo")
            parte = st.text_input("Parte de periodo", key="manual_parte")
            capacidad = st.text_input("Capacidad", key="manual_capacidad")
            seccion = st.text_input("Sección", key="manual_seccion")
        with d2:
            metodos = sorted({m for h, m in info["pares"] if h == horario and m}) if horario else sorted(info["insm"])
            opciones_metodo = [""] + metodos
            if st.session_state.get("manual_metodo", "") not in opciones_metodo:
                st.session_state.manual_metodo = ""
            metodo = st.selectbox("Método educativo", opciones_metodo, key="manual_metodo")
            sede = st.text_input("Sede", key="manual_sede")
            estatus = st.text_input("Estatus", key="manual_estatus")
            modo = st.selectbox("Modo de calificar", [""] + sorted(info["gmod"]), key="manual_modo")
            sesion = st.text_input("Sesión", key="manual_sesion")

        e1, e2 = st.columns(2)
        nivel = e1.selectbox("Nivel", ["LICENCIATURA", "BACHILLERATO", "POSGRADO"], key="manual_nivel")
        cluster = e2.selectbox("Clúster", CLUSTERS, index=None, key="manual_cluster")

        if st.button("➕ Agregar materia", key="manual_agregar"):
            if not confirmar or not cluster: st.warning("Confirma la materia y selecciona clúster.")
            else:
                nuevo = dict(zip(COLUMNAS_CSV, [
                    periodo, sede, llave[0], llave[1], parte, estatus, capacidad, "1",
                    seccion, horario, metodo, "D2L", modo, sesion,
                    "Bachillerato" if nivel == "BACHILLERATO" else cluster
                ]))
                st.session_state.df_manual_fijo = pd.concat(
                    [st.session_state.df_manual_fijo, pd.DataFrame([nuevo])], ignore_index=True)
                st.success("Materia agregada.")

    # 10. TABLA MANUAL Y ACCIONES
    st.subheader("📋 Tabla manual")
    manual = st.session_state.df_manual_fijo.copy()
    manual.insert(0, "Seleccionar", False)
    editado = st.data_editor(manual, hide_index=True, num_rows="fixed",
                             use_container_width=True, key="manual_editor")
    seleccion = editado.index[editado["Seleccionar"]].tolist()
    st.session_state.df_manual_fijo = editado.drop(columns="Seleccionar").copy()

    a1, a2, a3 = st.columns(3)
    if a1.button("Nuevo renglón", key="manual_nuevo"):
        nuevo = {c: "" for c in COLUMNAS_CSV}
        nuevo.update({"GRUPOS": "1", "SOCIODEINTEGRACION": "D2L"})
        st.session_state.df_manual_fijo = pd.concat(
            [st.session_state.df_manual_fijo, pd.DataFrame([nuevo])], ignore_index=True)
        st.rerun()
    if a2.button("Copiar seleccionados", key="manual_copiar"):
        if seleccion:
            st.session_state.df_manual_fijo = pd.concat(
                [st.session_state.df_manual_fijo, st.session_state.df_manual_fijo.loc[seleccion]],
                ignore_index=True)
            st.rerun()
        else: st.warning("Selecciona un renglón.")
    if a3.button("Eliminar seleccionados", key="manual_eliminar"):
        if seleccion:
            st.session_state.df_manual_fijo = st.session_state.df_manual_fijo.drop(
                index=seleccion).reset_index(drop=True)
            st.rerun()
        else: st.warning("Selecciona un renglón.")

    # 11. DESCARGA MANUAL
    nombre_manual = st.text_input("Nombre del CSV", "carga_manual.csv", key="manual_archivo")
    df_out = st.session_state.df_manual_fijo.copy()
    errores_manual = pd.DataFrame()

    if not df_out.empty:
        temporal = pd.DataFrame({
            "Periodo": df_out["PERIODO"], "Subject": df_out["SUBJ"],
            "Course": df_out["COURSE"], "Sección": df_out["SECCION"],
            "Nivel": df_out["datocomplementario"].apply(corregir_nivel_por_cluster_csv)
        })
        temporal, errores_manual = asignar_secciones_globales(
            temporal, st.session_state.get("cat_pa_cache"))
        df_out["SECCION"] = temporal["Sección"].values

    for c in ["SUBJ", "COURSE", "TIPODEHORARIO", "METODO_EDUCATIVO"]:
        df_out[c] = df_out[c].apply(sin_espacios)
    df_out["GRUPOS"], df_out["SOCIODEINTEGRACION"] = "1", "D2L"
    df_out["CAPACIDAD"] = pd.to_numeric(df_out["CAPACIDAD"], errors="coerce").astype("Int64")
    df_out["SECCION"] = df_out["SECCION"].apply(clave_seccion)

    if not errores_manual.empty: st.error("Hay secciones sin asignar."); st.dataframe(errores_manual)
    st.download_button("📥 Descargar CSV Manual", df_out[COLUMNAS_CSV].to_csv(**CSV_KWARGS_R).encode("utf-8"),
                       file_name=nombre_manual if nombre_manual.endswith(".csv") else nombre_manual + ".csv",
                       mime="text/csv", disabled=not errores_manual.empty)

# ============================================================
# FIN DE PESTAÑA 1
# ============================================================



# ============================================================
# PESTAÑA 2: REPORTE DE ERRORES Y ENSAMBLAJE FINAL
# ============================================================

with tab_err:
    st.header("⚠️ Reporte de Errores y Ensamblaje Final")

    # 1. REINICIAR PESTAÑA
    col_tit, col_btn = st.columns([4, 1])
    with col_btn:
        if st.button("🔄 Limpiar / Recomenzar", key="limpiar_t2", use_container_width=True):
            claves = ["df_delta_cache", "nombre_delta_cache", "llave_control_archivos",
                      "archivo_final_bytes", "archivo_final_nombre", "ext_base_1", "ext_err_1",
                      "suf_v1", "modo_1", "ed_vivo_1", "iny_base_2", "iny_err_2",
                      "iny_corr_2", "suf_v2", "firma_delta_t2", "firma_final_t2"]
            for k in claves: st.session_state.pop(k, None)
            st.rerun()

    st.info("Extrae los registros con error de Banner, corrígelos y vuelve a integrarlos al CSV original.")

    # 2. FUNCIONES AUXILIARES
    def leer_csv_banner(archivo):
        return pd.read_csv(io.BytesIO(archivo.getvalue()), dtype=str, encoding="utf-8-sig",
                           keep_default_na=False, skip_blank_lines=True)

    def leer_errores_banner(archivo, total_filas):
        df = pd.read_excel(io.BytesIO(archivo.getvalue()), skiprows=2, dtype=str)
        df.columns = [limpiar_nombre_columna(c) for c in df.columns]
        columna = next((c for c in df.columns if "linea" in normalizar_para_busqueda(c)), None)
        if columna is None: raise ValueError("No se encontró la columna 'Línea' en el reporte de Banner.")

        indices, invalidas = [], []
        for valor in df[columna].dropna().unique():
            try:
                numero = float(str(valor).strip())
                if not numero.is_integer(): raise ValueError()
                indice = int(numero) - 2  # Banner cuenta encabezado como línea 1
                if 0 <= indice < total_filas: indices.append(indice)
                else: invalidas.append(str(valor))
            except (ValueError, TypeError):
                invalidas.append(str(valor))

        indices = list(dict.fromkeys(indices))
        if invalidas:
            raise ValueError("El reporte contiene líneas inválidas o fuera del CSV: " +
                             ", ".join(invalidas[:15]))
        if not indices: raise ValueError("No se encontraron líneas válidas con errores.")
        return indices

    def nombre_base_t2(nombre):
        return re.sub(r"(?i)_(base|final|v\d+)$", "", nombre.rsplit(".", 1)[0])

    # ========================================================
    # 3. EXTRAER O EDITAR REGISTROS CON ERRORES
    # ========================================================
    st.subheader("✂️ 1. Extraer registros con errores")

    c1, c2, c3 = st.columns(3)
    file_base_ext = c1.file_uploader("📁 Archivo Base (.csv)", type=["csv"], key="ext_base_1")
    file_err_ext = c2.file_uploader("📊 Reporte de Errores Banner (.xlsx)", type=["xlsx"], key="ext_err_1")
    sufijo = c3.text_input("Versión del fragmento", value="V1", key="suf_v1")

    firma_delta = (
        (file_base_ext.name, file_base_ext.getvalue(), file_err_ext.name, file_err_ext.getvalue(), sufijo)
        if file_base_ext and file_err_ext else None
    )

    if st.session_state.get("firma_delta_t2") != firma_delta:
        for k in ["df_delta_cache", "nombre_delta_cache", "ed_vivo_1"]:
            st.session_state.pop(k, None)
        st.session_state.firma_delta_t2 = firma_delta

    if file_base_ext and file_err_ext:
        if st.button("🔍 Cargar y Procesar Reporte de Errores", key="procesar_delta_t2", use_container_width=True):
            try:
                df_base = leer_csv_banner(file_base_ext)
                indices = leer_errores_banner(file_err_ext, len(df_base))
                df_delta = df_base.iloc[indices].copy().reset_index(drop=True)

                st.session_state.df_delta_cache = df_delta
                st.session_state.nombre_delta_cache = f"{nombre_base_t2(file_base_ext.name)}_{sufijo}"
                st.success(f"Se extrajeron {len(indices)} registros con errores de un total de {len(df_base)}.")
            except Exception as e:
                st.session_state.df_delta_cache = None
                st.error(f"No se pudo extraer el fragmento: {e}")

    df_delta = st.session_state.get("df_delta_cache")
    if isinstance(df_delta, pd.DataFrame):
        st.markdown("#### 📋 Fragmento identificado")
        modo = st.radio("¿Cómo deseas corregirlo?", ["Excel (.xlsx)", "CSV (.csv)", "Editar en vivo"],
                        horizontal=True, key="modo_1")
        nombre_delta = st.session_state.nombre_delta_cache

        if modo == "Excel (.xlsx)":
            buffer = io.BytesIO()
            df_delta.to_excel(buffer, index=False, engine="openpyxl")
            st.download_button("📥 Descargar fragmento Excel", buffer.getvalue(),
                               file_name=f"{nombre_delta}.xlsx", use_container_width=True)

        elif modo == "CSV (.csv)":
            st.download_button("📥 Descargar fragmento CSV",
                               df_delta.to_csv(**CSV_KWARGS_R).encode("utf-8"),
                               file_name=f"{nombre_delta}.csv", use_container_width=True)

        else:
            st.caption("Modifica las celdas y descarga el fragmento corregido.")
            df_editado = st.data_editor(df_delta, hide_index=True, use_container_width=True,
                                        num_rows="fixed", key="ed_vivo_1")
            st.download_button("📥 Descargar parche corregido",
                               df_editado.to_csv(**CSV_KWARGS_R).encode("utf-8"),
                               file_name=f"{nombre_delta}.csv", use_container_width=True)

    # ========================================================
    # 4. INYECTAR CORRECCIONES Y GENERAR CSV FINAL
    # ========================================================
    st.divider()
    st.subheader("💉 2. Inyectar correcciones y generar archivo final")

    c1, c2, c3 = st.columns(3)
    file_base_iny = c1.file_uploader("📁 Archivo Base (.csv)", type=["csv"], key="iny_base_2")
    file_err_iny = c2.file_uploader("📊 Reporte de Errores (.xlsx)", type=["xlsx"], key="iny_err_2")
    file_corr_iny = c3.file_uploader("📝 Fragmento Corregido", type=["csv", "xlsx"], key="iny_corr_2")
    etiqueta = st.text_input("Etiqueta del archivo final", value="final", key="suf_v2")

    firma_final = (
        tuple((f.name, f.getvalue()) for f in [file_base_iny, file_err_iny, file_corr_iny])
        if all([file_base_iny, file_err_iny, file_corr_iny]) else None
    )

    if st.session_state.get("firma_final_t2") != (firma_final, etiqueta):
        st.session_state.archivo_final_bytes = None
        st.session_state.archivo_final_nombre = None
        st.session_state.firma_final_t2 = (firma_final, etiqueta)

    if all([file_base_iny, file_err_iny, file_corr_iny]):
        if st.button("🚀 Ensamblar Archivo Final", type="primary",
                     key="ensamblar_t2", use_container_width=True):
            try:
                base = leer_csv_banner(file_base_iny)
                indices = leer_errores_banner(file_err_iny, len(base))

                if file_corr_iny.name.lower().endswith(".xlsx"):
                    corregidas = pd.read_excel(io.BytesIO(file_corr_iny.getvalue()), dtype=str,
                                               keep_default_na=False)
                else:
                    corregidas = leer_csv_banner(file_corr_iny)

                corregidas = corregidas.fillna("").reset_index(drop=True)

                if len(indices) != len(corregidas):
                    raise ValueError(f"El reporte tiene {len(indices)} registros con error, pero el parche "
                                     f"contiene {len(corregidas)}. Deben coincidir exactamente.")

                faltantes = set(base.columns) - set(corregidas.columns)
                extras = set(corregidas.columns) - set(base.columns)
                if faltantes or extras:
                    raise ValueError(f"Columnas diferentes. Faltantes: {sorted(faltantes)}. "
                                     f"Adicionales: {sorted(extras)}.")

                if corregidas.columns.duplicated().any():
                    raise ValueError("El fragmento corregido contiene columnas duplicadas.")

                final = base.copy()
                final.iloc[indices, :] = corregidas[base.columns].to_numpy()

                if len(final) != len(base):
                    raise ValueError("El archivo final no conserva la cantidad de registros originales.")

                nombre_final = f"{nombre_base_t2(file_base_iny.name)}_{etiqueta}.csv"
                st.session_state.archivo_final_bytes = final.to_csv(**CSV_KWARGS_R).encode("utf-8")
                st.session_state.archivo_final_nombre = nombre_final
                st.success(f"Archivo ensamblado: {len(indices)} registros corregidos, {len(final)} filas totales.")

            except Exception as e:
                st.session_state.archivo_final_bytes = None
                st.error(f"Error al ensamblar: {e}")

    # 5. DESCARGA FINAL
    if st.session_state.get("archivo_final_bytes") is not None:
        st.download_button(f"📥 DESCARGAR {st.session_state.archivo_final_nombre}",
                           data=st.session_state.archivo_final_bytes,
                           file_name=st.session_state.archivo_final_nombre,
                           mime="text/csv", type="primary", use_container_width=True,
                           key="descargar_final_t2")

# ============================================================
# FIN DE PESTAÑA 2
# ============================================================


# ============================================================
# PESTAÑA 3: INYECCIÓN DE NRCs Y CRUCES CON ARGOS
# ============================================================

with tab3:
    st.header("📊 Inyección de NRCs y Cruces con ARGOS")

    # 1. REINICIAR PESTAÑA
    col_tit, col_btn = st.columns([4, 1])
    with col_btn:
        if st.button("🔄 Limpiar / Recomenzar", key="limpiar_t3", use_container_width=True):
            claves = ["modo_inyeccion_t3", "arg_c", "csv_c", "xls_c", "arg_r", "csv_r",
                      "final_argos_zip", "df_cruce_rapido", "columnas_copia_rapida",
                      "alertas_argos_t3", "firma_argos_t3", "firma_rapido_t3"]
            for k in claves: st.session_state.pop(k, None)
            st.rerun()

    modo = st.radio("🛠️ Selecciona el tipo de proceso:", [
        "📦 Completo (ARGOS + CSV Final + Excel Original)",
        "⚡ Rápido (ARGOS + CSV Final)"
    ], horizontal=True, key="modo_inyeccion_t3")

    st.divider()

    # ========================================================
    # 2. FUNCIONES AUXILIARES
    # ========================================================

    def leer_argos(archivo):
        df = pd.read_csv(io.BytesIO(archivo.getvalue()), dtype=str, encoding="utf-8-sig",
                         keep_default_na=False, on_bad_lines="skip")
        df.columns = [re.sub(r"\.+", ".", str(c).replace('"', "").replace("'", "").strip())
                      for c in df.columns]

        def buscar_columna(palabra):
            return next((c for c in df.columns if palabra in normalizar_para_busqueda_t3(c)), None)

        columnas = {
            "periodo": buscar_columna("periodo"),
            "nivel": buscar_columna("nivel"),
            "cluster": buscar_columna("cluster"),
            "subj": buscar_columna("area"),
            "crse": buscar_columna("curso"),
            "grupo": buscar_columna("grupo"),
            "nrc": buscar_columna("nrc")
        }

        faltantes = [k for k, v in columnas.items() if v is None]
        if faltantes: raise ValueError(f"ARGOS: no se encontraron columnas {faltantes}")

        # Identificar columna Grupo exacta, evitando Grupo LC
        grupo_exacto = next((c for c in df.columns if normalizar_para_busqueda_t3(c) == "grupo"), None)
        if grupo_exacto: columnas["grupo"] = grupo_exacto

        for nombre, columna in columnas.items():
            if nombre == "grupo":
                df["_grupo"] = df[columna].apply(ultra_limpiar_seccion)
            else:
                df[f"_{nombre}"] = df[columna].apply(ultra_limpiar)

        df["_llave"] = (
            df["_periodo"] + "_" + df["_nivel"] + "_" + df["_cluster"] + "_" +
            df["_subj"] + "_" + df["_crse"] + "_" + df["_grupo"]
        )

        # Evitar que NRC vacíos o claves repetidas generen cruces incorrectos
        df = df[df["_nrc"].ne("") & df["_periodo"].ne("") & df["_subj"].ne("") &
                df["_crse"].ne("") & df["_grupo"].ne("")].copy()

        conflictos = df.groupby("_llave")["_nrc"].nunique()
        conflictos = conflictos[conflictos > 1].index.tolist()

        if conflictos:
            raise ValueError(
                f"ARGOS contiene {len(conflictos)} combinaciones con NRC diferentes. "
                f"Ejemplos: {conflictos[:5]}. Revisa antes de inyectar."
            )

        df = df.drop_duplicates(subset=["_llave"])
        mapa = dict(zip(df["_llave"], df[columnas["nrc"]]))
        return df, mapa

    def nivel_desde_cluster(cluster):
        c = normalizar_para_cruce(cluster)
        if "POSGRADO" in c: return "POSGRADO"
        if "BACHILLERATO" in c: return "BACHILLERATO"
        return "LICENCIATURA"

    def generar_llaves_csv(df):
        obligatorias = ["PERIODO", "SUBJ", "COURSE", "SECCION", "datocomplementario"]
        faltantes = [c for c in obligatorias if c not in df.columns]
        if faltantes: raise ValueError(f"CSV Final: faltan columnas {faltantes}")

        cluster = df["datocomplementario"].apply(ultra_limpiar)
        nivel = df["datocomplementario"].apply(nivel_desde_cluster)

        return (
            df["PERIODO"].apply(ultra_limpiar) + "_" + nivel + "_" + cluster + "_" +
            df["SUBJ"].apply(ultra_limpiar) + "_" + df["COURSE"].apply(ultra_limpiar) + "_" +
            df["SECCION"].apply(ultra_limpiar_seccion)
        )

    def buscar_nrc(df_csv, mapa, llaves_argos):
        llaves = generar_llaves_csv(df_csv)
        nrc = llaves.map(mapa)
        alertas = []

        for llave in llaves[nrc.isna()].unique():
            parecidas = difflib.get_close_matches(str(llave), llaves_argos, n=1, cutoff=0.5)
            sugerencia = parecidas[0] if parecidas else "Sin coincidencia cercana"
            alertas.append({"Llave sin NRC": llave, "Coincidencia ARGOS": sugerencia})

        return nrc, alertas

    def formato_excel_nrc(ws, columna_nrc="NRC"):
        fuente = Font(name="Calibri", size=11)
        fuente_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        fondo_header = PatternFill(start_color="1F4E78", fill_type="solid")
        fondo_nrc = PatternFill(start_color="DDEBF7", fill_type="solid")
        centro = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        encabezados = {str(c.value): c.column for c in ws[1]}
        idx_nrc = encabezados.get(columna_nrc)

        for celda in ws[1]:
            celda.font, celda.fill, celda.alignment = fuente_header, fondo_header, centro

        for fila in ws.iter_rows(min_row=2):
            for celda in fila:
                celda.font = fuente
                celda.alignment = Alignment(horizontal="center", vertical="center")
                if idx_nrc and celda.column == idx_nrc:
                    celda.font = Font(name="Calibri", size=11, bold=True)
                    celda.fill = fondo_nrc

        for columna in ws.columns:
            ancho = max((len(str(c.value)) for c in columna if c.value is not None), default=8)
            ws.column_dimensions[columna[0].column_letter].width = min(max(ancho + 3, 11), 45)

    def leer_csv_final(archivo):
        return pd.read_csv(io.BytesIO(archivo.getvalue()), dtype=str, encoding="utf-8-sig",
                           keep_default_na=False).dropna(how="all").reset_index(drop=True)

    # ========================================================
    # 3. MODO COMPLETO: ARGOS + CSV + EXCEL ORIGINAL
    # ========================================================

    if modo.startswith("📦"):
        st.subheader("📦 Inyección masiva de NRC en Excel")
        c1, c2, c3 = st.columns(3)
        file_argos = c1.file_uploader("📊 Reporte ARGOS (.csv)", type=["csv"], key="arg_c")
        files_csv = c2.file_uploader("📝 CSV Finales", type=["csv"], accept_multiple_files=True, key="csv_c")
        files_excel = c3.file_uploader("📁 Excel Originales", type=["xlsx"], accept_multiple_files=True, key="xls_c")

        if file_argos and files_csv and files_excel:
            if st.button("🚀 PROCESAR Y GENERAR EXCELS CON NRC", type="primary",
                         use_container_width=True, key="procesar_argos_completo"):
                try:
                    argos, mapa = leer_argos(file_argos)
                    llaves_argos = list(mapa.keys())
                    buffer_zip = io.BytesIO()
                    alertas, procesados = [], 0
                    csv_usados = set()

                    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as z:
                        for fx in files_excel:
                            nombre_excel = simplificar_nombre(fx.name)
                            parejas = [fc for fc in files_csv if
                                       simplificar_nombre(fc.name) == nombre_excel]

                            if len(parejas) != 1:
                                alertas.append(f"{fx.name}: se encontraron {len(parejas)} CSV compatibles.")
                                continue

                            fc = parejas[0]
                            if fc.name in csv_usados:
                                alertas.append(f"{fx.name}: el CSV {fc.name} ya fue utilizado.")
                                continue

                            df_csv = leer_csv_final(fc)
                            wb = openpyxl.load_workbook(io.BytesIO(fx.getvalue()))

                            hoja = next((h for h in wb.sheetnames if h.strip().upper() == HOJA_ALTAS), None)
                            if hoja is None:
                                alertas.append(f"{fx.name}: no tiene pestaña ALTAS.")
                                continue

                            df_excel = pd.read_excel(io.BytesIO(fx.getvalue()), sheet_name=hoja, dtype=str)
                            df_excel = df_excel.dropna(how="all").reset_index(drop=True)
                            df_excel.columns = [str(c).strip() for c in df_excel.columns]

                            if len(df_excel) != len(df_csv):
                                alertas.append(f"{fx.name}: Excel {len(df_excel)} filas vs CSV {len(df_csv)} filas.")
                                continue

                            nrc, faltantes = buscar_nrc(df_csv, mapa, llaves_argos)
                            for f in faltantes:
                                alertas.append(f"{fx.name}: {f['Llave sin NRC']} → {f['Coincidencia ARGOS']}")

                            # Se conserva el Excel original y se agrega una nueva hoja NRC
                            df_nrc = df_excel.copy()
                            equivalencias = {
                                "Periodo": "PERIODO", "Campus": "SEDE", "Subject": "SUBJ",
                                "Course": "COURSE", "Parte de Periodo": "PARTEPERIODO",
                                "Estatus": "STATUS", "Capacidad": "CAPACIDAD", "Sección": "SECCION",
                                "Tipo de Horario": "TIPODEHORARIO", "Método Educativo": "METODO_EDUCATIVO",
                                "Modo de Calificar": "MODODECALIFICAR", "Sesion": "SESION"
                            }
                            for origen, destino in equivalencias.items():
                                if origen in df_nrc.columns and destino in df_csv.columns:
                                    df_nrc[origen] = df_csv[destino].values

                            df_nrc.insert(0, "NRC", nrc.values)
                            df_nrc["Grupos"] = "1"
                            df_nrc["Socio de Integración"] = "D2L"

                            if HOJA_SALIDA_NRC in wb.sheetnames: del wb[HOJA_SALIDA_NRC]
                            ws = wb.create_sheet(HOJA_SALIDA_NRC)
                            ws.append(list(df_nrc.columns))

                            for fila in df_nrc.itertuples(index=False, name=None):
                                ws.append([None if pd.isna(v) else v for v in fila])

                            formato_excel_nrc(ws)
                            salida = io.BytesIO()
                            wb.save(salida)

                            nombre_salida = fc.name.rsplit(".", 1)[0] + "_con_NRC.xlsx"
                            z.writestr(nombre_salida, salida.getvalue())
                            csv_usados.add(fc.name)
                            procesados += 1

                    st.session_state.final_argos_zip = buffer_zip.getvalue() if procesados else None
                    st.session_state.alertas_argos_t3 = alertas

                    if procesados: st.success(f"Se procesaron {procesados} Excel correctamente.")
                    else: st.error("No se pudo procesar ningún archivo.")

                except Exception as e:
                    st.session_state.final_argos_zip = None
                    st.error(f"Error en la inyección de NRC: {e}")

        if st.session_state.get("alertas_argos_t3"):
            with st.expander("🔍 Ver discrepancias y NRC no encontrados"):
                for alerta in st.session_state.alertas_argos_t3: st.warning(alerta)

        if st.session_state.get("final_argos_zip"):
            st.download_button("📥 DESCARGAR EXCELS CON NRC (.ZIP)",
                               st.session_state.final_argos_zip, file_name="Excels_Finales_con_NRC.zip",
                               mime="application/zip", type="primary", use_container_width=True)

    # ========================================================
    # 4. MODO RÁPIDO: ARGOS + CSV FINAL
    # ========================================================

    else:
        st.subheader("⚡ Cruce rápido de NRC")
        c1, c2 = st.columns(2)
        file_argos = c1.file_uploader("📊 Reporte ARGOS (.csv)", type=["csv"], key="arg_r")
        files_csv = c2.file_uploader("📝 CSV Finales", type=["csv"], accept_multiple_files=True, key="csv_r")

        if file_argos and files_csv:
            if st.button("⚡ Cruzar NRC y Generar Tabla", type="primary",
                         use_container_width=True, key="procesar_argos_rapido"):
                try:
                    argos, mapa = leer_argos(file_argos)
                    resultados, alertas = [], []

                    for archivo in files_csv:
                        df = leer_csv_final(archivo)
                        nrc, faltantes = buscar_nrc(df, mapa, list(mapa.keys()))
                        df.insert(0, "NRC", nrc.values)
                        df = df.rename(columns={
                            "TIPODEHORARIO": "TIPO DE HORARIO",
                            "METODO_EDUCATIVO": "METODO_ED",
                            "datocomplementario": "Cluster"
                        })

                        columnas = ["NRC", "PERIODO", "SUBJ", "COURSE", "CAPACIDAD",
                                    "SECCION", "TIPO DE HORARIO", "METODO_ED", "Cluster"]
                        resultados.append(df[[c for c in columnas if c in df.columns]])
                        alertas += [{"Archivo": archivo.name, **f} for f in faltantes]

                    st.session_state.df_cruce_rapido = pd.concat(resultados, ignore_index=True)
                    st.session_state.alertas_argos_t3 = alertas
                    st.session_state.pop("columnas_copia_rapida", None)
                    st.success(f"Cruce completado: {len(st.session_state.df_cruce_rapido)} registros.")

                except Exception as e:
                    st.session_state.df_cruce_rapido = None
                    st.error(f"Error en cruce rápido: {e}")

        if st.session_state.get("alertas_argos_t3"):
            with st.expander("⚠️ NRC no encontrados"):
                st.dataframe(pd.DataFrame(st.session_state.alertas_argos_t3),
                             hide_index=True, use_container_width=True)

        df_resultado = st.session_state.get("df_cruce_rapido")
        if isinstance(df_resultado, pd.DataFrame):
            st.subheader("📋 Resultados del cruce")

            columnas = st.multiselect("Selecciona columnas para copiar o descargar",
                                      options=list(df_resultado.columns),
                                      default=list(df_resultado.columns),
                                      key="columnas_copia_rapida")

            if columnas:
                df_mostrar = df_resultado[columnas].copy()
                st.dataframe(df_mostrar, hide_index=True, use_container_width=True)

                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("#### 📋 Copiar a Excel")
                    st.caption("Copia el contenido del cuadro y pégalo directamente en Excel.")
                    st.code(df_mostrar.to_csv(index=False, sep="\t"), language="text")

                with c2:
                    st.markdown("#### 📥 Descargar Excel")
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        df_mostrar.to_excel(writer, index=False, sheet_name="Cruce_NRC")
                        formato_excel_nrc(writer.sheets["Cruce_NRC"])

                    st.download_button("📥 Descargar tabla formateada (.xlsx)",
                                       buffer.getvalue(), file_name="Cruce_Rapido_NRC.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       type="primary", use_container_width=True)
            else:
                st.warning("Selecciona al menos una columna.")

# ============================================================
# FIN DE PESTAÑA 3
# ============================================================
