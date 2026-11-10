
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

def nombre_csv_altas(periodo, responsable, cluster):
    periodo = nombre_seguro(periodo)
    responsable = responsable_normalizado(responsable)
    cluster = nombre_seguro(cluster_oficial(cluster))
    if not periodo or not responsable or not cluster:
        raise ValueError("Falta Periodo, Responsable o Clúster para nombrar CSV_ALTAS.")
    return f"{periodo} {responsable} {cluster} CSV_ALTAS.csv"

# ============================================================
# 5. CATÁLOGO PA OPCIONAL
# ============================================================
def leer_csv_pa(archivos):
    # Acepta un CSV, una lista de CSV o None. La identidad es PERIODO + NRC.
    if archivos is None: return None, []
    if not isinstance(archivos, (list, tuple)): archivos = [archivos]
    if not archivos: return None, []
    partes = []
    for archivo in archivos:
        nombre = getattr(archivo, "name", "PA")
        datos = archivo.getvalue() if hasattr(archivo, "getvalue") else archivo
        tabla = None
        for enc in ("utf-8-sig", "utf-8", "cp1252"):
            try:
                tabla = pd.read_csv(io.BytesIO(datos), dtype=str, encoding=enc, sep=None,
                                    engine="python", keep_default_na=False)
                break
            except (UnicodeError, pd.errors.ParserError): continue
        if tabla is None: raise ValueError(f"No se pudo leer la PA: {nombre}")
        tabla.columns = [normalizar_para_busqueda(str(c).replace("Ã", "A").replace("Ã¡", "a")) for c in tabla.columns]
        requeridas = ["periodo", "area", "nocurso", "grupo", "nrc"]
        faltantes = [c for c in requeridas if c not in tabla.columns]
        if faltantes: raise ValueError(f"{nombre}: faltan columnas PA {faltantes}")
        tabla = tabla[requeridas].copy()
        for c in ("periodo", "area", "nocurso", "nrc"): tabla[c] = tabla[c].apply(normalizar_para_cruce)
        tabla["grupo"] = tabla["grupo"].apply(clave_seccion)
        tabla = tabla[tabla["nrc"].ne("") & tabla["periodo"].ne("")].copy()
        partes.append(tabla)
    pa = pd.concat(partes, ignore_index=True)
    if pa.empty: return pa, []
    # Los mismos NRC en periodos distintos son independientes.
    # Si un PERIODO+NRC tiene datos contradictorios, reservar todas las secciones
    # asociadas, sin seleccionar silenciosamente una de ellas.
    detalle = ["area", "nocurso", "grupo"]
    conflictos = pa.groupby(["periodo", "nrc"], sort=False)[detalle].nunique(dropna=False).gt(1).any(axis=1)
    advertencias = [
        {"Tipo": "PA", "Periodo": periodo, "NRC": nrc,
         "Detalle": f"Periodo {periodo}, NRC {nrc}: registros contradictorios; se reservan las secciones reportadas."}
        for (periodo, nrc) in conflictos[conflictos].index
    ]
    pa_reserva = pa.drop_duplicates(subset=["periodo", "nrc"] + detalle).reset_index(drop=True)
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
# ============================================================
# 10. SEPARACION Y EXPORTACION DE SOLICITUDES + CSV_ALTAS
# ============================================================
MODOS_SEPARACION = ["Sin separar", "Por periodo", "Por clúster", "Por periodo y clúster"]

def grupos_exportacion_altas(df_final, modo="Sin separar"):
    if modo not in MODOS_SEPARACION: raise ValueError(f"Modo de separación desconocido: {modo}")
    df = df_final.copy()
    obligatorias = ["Periodo", "Responsable", "Clúster", "ArchivoOrigen", "_FilaExcel"]
    faltantes = [c for c in obligatorias if c not in df.columns]
    if faltantes: raise ValueError(f"Faltan columnas para exportar: {faltantes}")
    df["Periodo"] = df["Periodo"].apply(limpiar_clave_texto)
    df["Responsable"] = df["Responsable"].apply(responsable_normalizado)
    df["Clúster"] = df["Clúster"].apply(cluster_oficial)
    for c in ["Periodo", "Responsable", "Clúster"]:
        if df[c].apply(limpiar_clave_texto).eq("").any(): raise ValueError(f"Hay ALTAS sin {c}.")
    # Conservar siempre la procedencia para no fusionar solicitudes de distintos libros.
    claves = ["ArchivoOrigen", "Responsable"]
    if modo in ("Por periodo", "Por periodo y clúster"): claves.append("Periodo")
    if modo in ("Por clúster", "Por periodo y clúster"): claves.append("Clúster")
    return [(k, g.copy()) for k, g in df.groupby(claves, sort=False, dropna=False)]

def nombre_csv_grupo_altas(df_grupo):
    periodos = sorted({nombre_seguro(x) for x in df_grupo["Periodo"] if nombre_seguro(x)})
    responsables = responsables_unicos(df_grupo["Responsable"])
    clusters = sorted({nombre_seguro(cluster_oficial(x)) for x in df_grupo["Clúster"] if nombre_seguro(x)}, key=normalizar_para_busqueda)
    if not periodos or not responsables or not clusters: raise ValueError("Faltan datos para nombrar CSV_ALTAS.")
    unir = lambda xs: " y ".join(xs) if len(xs) == 2 else ", ".join(xs)
    return f"{'-'.join(periodos)} {unir(responsables)} {unir(clusters)} CSV_ALTAS.csv"

def ruta_zip_unica(carpeta, nombre, rutas):
    base, extension = nombre.rsplit(".", 1)
    candidato, n = f"{carpeta}/{nombre}", 2
    while candidato.casefold() in rutas:
        candidato = f"{carpeta}/{base} ({n}).{extension}"
        n += 1
    rutas.add(candidato.casefold())
    return candidato

def crear_excel_normal_altas(df_grupo, archivos_originales):
    # Excel sin VBA: solo las filas del grupo, en el mismo orden de ALTAS.
    origenes = df_grupo["ArchivoOrigen"].dropna().unique().tolist()
    if len(origenes) != 1: raise ValueError("Un Excel separado debe tener un solo origen.")
    origen = origenes[0]
    wb_origen = openpyxl.load_workbook(io.BytesIO(archivos_originales[origen]), read_only=True, data_only=False)
    hoja = next((h for h in wb_origen.sheetnames if h.strip().upper() == HOJA_ALTAS), None)
    if hoja is None: raise ValueError(f"{origen}: falta hoja ALTAS.")
    ws_origen = wb_origen[hoja]
    encabezados = [limpiar_clave_texto(c.value) for c in ws_origen[1]]
    # Preservar encabezados reales de ALTAS y sus 22 campos; nunca añadir columnas técnicas.
    columnas = {normalizar_para_busqueda(c): i for i, c in enumerate(encabezados)}
    equivalencias = {"periodo": "Periodo", "responsable": "Responsable", "cluster": "Clúster",
                    "subj": "Subject", "course": "Course", "seccion": "Sección",
                    "tipohorario": "Tipo de Horario", "metodoeducativo": "Método Educativo",
                    "modocalificar": "Modo de Calificar", "parteperiodo": "Parte de Periodo",
                    "status": "Estatus", "capacidad": "Capacidad", "sesion": "Sesion", "sede": "Campus"}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = HOJA_ALTAS
    ws.append(encabezados)
    for _, fila in df_grupo.iterrows():
        numero = int(fila["_FilaExcel"])
        valores = [ws_origen.cell(numero, j + 1).value for j in range(len(encabezados))]
        for cabecera, campo in equivalencias.items():
            if cabecera in columnas and campo in fila.index:
                valor = fila[campo]
                valores[columnas[cabecera]] = None if pd.isna(valor) else str(valor)
        ws.append(valores)
    wb_origen.close()
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for c in ws[1]:
        c.fill = PatternFill(fill_type="solid", fgColor="7030A0")
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in ws.columns:
        ancho = min(40, max(12, max(len(str(c.value or "")) for c in col) + 2))
        ws.column_dimensions[col[0].column_letter].width = ancho
    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue()

def generar_zip_altas(df_final, archivos_originales, funcion_csv, modo_separacion="Sin separar"):
    grupos = grupos_exportacion_altas(df_final, modo_separacion)
    df = df_final.copy()
    rutas, buffer = set(), io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        # Cada grupo exporta un CSV y, si hay separación, un Excel normal equivalente.
        for _, grupo in grupos:
            nombre_csv = nombre_csv_grupo_altas(grupo)
            contenido = funcion_csv(grupo)
            if isinstance(contenido, str): contenido = contenido.encode("utf-8")
            z.writestr(ruta_zip_unica("CSV_ALTAS", nombre_csv, rutas), contenido)
            if modo_separacion != "Sin separar":
                nombre_excel = nombre_solicitud(grupo, ".xlsx")
                contenido_excel = crear_excel_normal_altas(grupo, archivos_originales)
                z.writestr(ruta_zip_unica("ALTAS_SEPARADAS", nombre_excel, rutas), contenido_excel)
        # Unicamente devolver el libro con macros cuando el nombre original sea incorrecto.
        for origen, grupo in df.groupby("ArchivoOrigen", sort=False):
            extension = ".xlsm" if origen.lower().endswith(".xlsm") else ".xlsx"
            if extension != ".xlsm": continue
            esperado = nombre_solicitud(grupo, extension)
            if normalizar_para_busqueda(origen) == normalizar_para_busqueda(esperado): continue
            contenido, _ = crear_solicitud_excel(grupo, archivos_originales)
            z.writestr(ruta_zip_unica("MACROS_RENOMBRADAS", esperado, rutas), contenido)
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

tab1, tab_err, tab3, tab4 = st.tabs([
    "1️⃣ Proceso: Validación y Generar CSVs",
    "⚠️ Reporte de Errores (Extraer Delta)",
    "2️⃣ Proceso: Inyección de NRCs Masiva (ARGOS)",
    "3️⃣ Proceso: Gestión de Cruces"
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
                  "t1_firma_archivos", "t1_catalogo", "t1_version_editor", "t1_modo_separacion", "t1_modo_generado"]
        claves += [k for k in st.session_state if k.startswith(("nom_t1_", "met_t1_", "manual_t1_", "t1_editor_"))]
        for k in claves: st.session_state.pop(k, None)
        st.rerun()

    # 2. CARGA DE ARCHIVOS
    st.subheader("📁 Archivos de entrada")
    c1, c2, c3 = st.columns(3)
    file_cat_ext = c1.file_uploader("📚 Catálogo Avanzado (SCBCRSE)", type=["csv", "xlsx"], key="cat_ext_t1")
    files_pa = c2.file_uploader("📊 Catálogos PA (Opcionales, varios periodos)", type=["csv"], accept_multiple_files=True, key="pa_t1")
    files_altas = c3.file_uploader("📁 Archivos ALTAS", type=["xlsx", "xlsm"], accept_multiple_files=True, key="altas_t1")
    st.caption("Puedes cargar una PA por periodo o un único CSV que contenga varios periodos. La PA es opcional; si falta un periodo se mostrará una advertencia.")

    # Invalidar resultados cuando cambien los archivos.
    firma_actual = tuple((f.name, hashlib.sha256(f.getvalue()).hexdigest())
                         for f in ([file_cat_ext] + list(files_pa or []) + list(files_altas or [])) if f is not None)
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

    # 4. VALIDACIÓN UNIFICADA: MATERIA, PAREJA HORARIO/MÉTODO Y MODO INDEPENDIENTE
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
             "Comentario Modo de Calificar": "", "Pendientes Alta": [], "Requiere Corrección": False}
        info = catalogo.get((subj, crse))
        if info is None:
            candidatos = []
            for (s, c), datos in catalogo.items():
                pn = max((similitud(normalizar_para_cruce(nombre), normalizar_para_cruce(t)) for t in datos["titles"]), default=0)
                puntuacion = pn * .55 + similitud(subj, s) * .25 + similitud(crse, c) * .20
                if puntuacion >= .63: candidatos.append((puntuacion, s, c))
            if candidatos:
                p, s, c = max(candidatos)
                r["Subj Sugerido"], r["Crse Sugerido"] = s, c
                r["Comentario Nombres"] = f"SUBJ/COURSE no encontrados. Posible materia: {s} {c} ({p:.0%})."
            else: r["Comentario Nombres"] = "Materia no encontrada; revisar clave o alta de materia."
            r["Requiere Corrección"] = True
            return r
        titulos = sorted(info["titles"])
        r["Materia Catálogo"] = titulos[0] if titulos else ""
        if not titulos: r["Comentario Nombres"] = "Sin título verificable en catálogo."
        elif normalizar_para_cruce(nombre) in {normalizar_para_cruce(t) for t in titulos}:
            r["Comentario Nombres"] = "Correcto"
        else:
            r["Comentario Nombres"] = "Nombre diferente al catálogo; revisar."
            r["Requiere Corrección"] = True
        if not horario or not metodo:
            r["Comentario Horario"] = "Tipo de horario o método educativo vacío."
            r["Comentario Método"] = "Revisar la pareja solicitada."
            r["Requiere Corrección"] = True
        elif (horario, metodo) in info["pares"]:
            r["Comentario Horario"] = r["Comentario Método"] = "Combinación correcta"
        else:
            r["Comentario Horario"] = f"Combinación {horario} + {metodo} no registrada para {subj} {crse}."
            r["Comentario Método"] = "Pendiente de alta de combinación horario/método."
            r["Pendientes Alta"].append(("Tipo de Horario + Método Educativo", horario, metodo, "Combinación no registrada"))
        if not modo:
            r["Comentario Modo de Calificar"] = "Modo de calificar vacío; revisar."
            r["Requiere Corrección"] = True
        elif modo in info["gmod"]: r["Comentario Modo de Calificar"] = "Correcto"
        else:
            r["Comentario Modo de Calificar"] = f"Modo de calificar {modo} no registrado para {subj} {crse}."
            r["Pendientes Alta"].append(("Modo de Calificar", "", "", modo))
        return r

    # 5. REPORTE DE ALTAS PENDIENTES: MOSTRAR LA PAREJA COMPLETA
    def obtener_pendientes_altas(df, auditoria):
        columnas = ["Periodo", "SUBJ", "COURSE", "Materia", "Configuración", "Tipo de Horario",
                    "Método Educativo", "Modo de Calificar", "Motivo", "Responsable", "Archivo"]
        registros = []
        for _, r in auditoria.iterrows():
            fila = df.iloc[int(r["idx"])]
            for tipo, horario, metodo, detalle in r["Pendientes Alta"]:
                registros.append({"Periodo": limpiar_clave_texto(fila.get("Periodo")),
                    "SUBJ": sin_espacios(fila.get("Subject")), "COURSE": sin_espacios(fila.get("Course")),
                    "Materia": limpiar_clave_texto(fila.get("Nombre de la Materia")), "Configuración": tipo,
                    "Tipo de Horario": horario, "Método Educativo": metodo,
                    "Modo de Calificar": detalle if tipo == "Modo de Calificar" else sin_espacios(fila.get("Modo de Calificar")),
                    "Motivo": detalle if tipo != "Modo de Calificar" else "Modo de calificar no registrado",
                    "Responsable": responsable_normalizado(fila.get("Responsable")), "Archivo": fila.get("ArchivoOrigen", "")})
        return pd.DataFrame(registros, columns=columnas)

    # 6. VALIDACIÓN INTELIGENTE
    if files_altas and file_cat_ext and st.button("⚡ Ejecutar Validación Inteligente", type="primary",
                                                   use_container_width=True, key="validar_t1"):
        try:
            with st.spinner("Revisando materias, configuraciones y secciones..."):
                catalogo = cargar_catalogo_avanzado()
                if not catalogo: raise ValueError("El Catálogo Avanzado no contiene materias válidas.")
                pa, avisos = leer_csv_pa(files_pa) if files_pa else (None, [])
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
                # La cobertura se comprueba por PERIODO, nunca solo por NRC.
                periodos_altas = {normalizar_para_cruce(p) for p in total["Periodo"] if limpiar_clave_texto(p)}
                periodos_pa = {normalizar_para_cruce(p) for p in pa["periodo"]} if pa is not None else set()
                sin_pa = sorted(periodos_altas - periodos_pa)
                for periodo_faltante in sin_pa:
                    avisos.append({"Tipo": "PA sin cobertura", "Periodo": periodo_faltante,
                                  "Detalle": f"No se cargó PA del periodo {periodo_faltante}. Las secciones se asignarán sin conocer las ocupadas en Banner."})
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
            resumen_altas = pendientes.drop_duplicates(subset=["SUBJ", "COURSE", "Configuración", "Tipo de Horario", "Método Educativo", "Modo de Calificar"])
            resumen_altas = resumen_altas.sort_values(["SUBJ", "COURSE", "Configuración"])
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

        # 7.2 UNA SOLA MESA DE CONTROL; SOLO ERRORES QUE REQUIEREN CORRECCIÓN
        corregibles = auditoria[auditoria["Requiere Corrección"] == True].copy()
        if not corregibles.empty:
            st.subheader("⚖️ Mesa de Control — Revisión unificada")
            st.caption("Las combinaciones pendientes de alta se muestran arriba y no se reemplazan automáticamente.")
            columnas_editor = ["Luz Verde", "Materia Excel", "Materia Catálogo", "Subj Original",
                "Crse Original", "Subj Sugerido", "Crse Sugerido", "Comentario Nombres",
                "Horario Original", "Horario Sugerido", "Método Original", "Método Sugerido",
                "Comentario Horario", "Modo de Calificar Original", "Modo de Calificar Sugerido",
                "Comentario Modo de Calificar"]
            bloqueadas = ["Materia Excel", "Materia Catálogo", "Subj Original", "Crse Original",
                "Comentario Nombres", "Horario Original", "Método Original", "Comentario Horario",
                "Modo de Calificar Original", "Comentario Modo de Calificar"]
            for archivo in corregibles["Archivo"].unique():
                sub = corregibles[corregibles["Archivo"] == archivo]
                with st.expander(f"⚠️ {archivo} — {len(sub)} registros", expanded=True):
                    with st.form(f"revision_{hashlib.md5(archivo.encode()).hexdigest()[:10]}"):
                        editado = st.data_editor(sub[columnas_editor], hide_index=True, use_container_width=True,
                            disabled=bloqueadas, key=f"t1_editor_{hashlib.md5(archivo.encode()).hexdigest()[:10]}_{st.session_state.t1_version_editor}")
                        guardar = st.form_submit_button("💾 Guardar correcciones", use_container_width=True)
                    if guardar:
                        for i in sub.index:
                            auditoria.at[i, "Luz Verde"] = bool(editado.at[i, "Luz Verde"])
                            for c in ["Subj Sugerido", "Crse Sugerido", "Horario Sugerido",
                                      "Método Sugerido", "Modo de Calificar Sugerido"]:
                                auditoria.at[i, c] = editado.at[i, c]
                        st.session_state.res_auditoria = auditoria
                        st.session_state.zip_solicitudes_csv_altas = None
                        st.session_state.t1_version_editor += 1
                        st.rerun()

        # 7.3 ADVERTENCIAS
        if avisos_pa:
            st.subheader("⚠️ Advertencias PA y periodos sin cobertura")
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
        st.subheader("📁 Separación de CSV_ALTAS y Excel")
        modo_separacion = st.radio("¿Cómo deseas separar los archivos?", MODOS_SEPARACION,
            index=3, horizontal=True, key="t1_modo_separacion")
        st.caption("La misma separación aplica a los CSV y a los Excel normales. Las secciones ya se asignaron globalmente.")
        if st.session_state.get("t1_modo_generado") != modo_separacion:
            st.session_state.zip_solicitudes_csv_altas = None
        st.subheader("📋 Vista previa de archivos")
        try:
            vista, _ = construir_altas_corregidas()
            previstos, usados = [], set()
            for _, grupo in grupos_exportacion_altas(vista, modo_separacion):
                nombre_csv = nombre_csv_grupo_altas(grupo)
                nombre_excel = nombre_solicitud(grupo, ".xlsx") if modo_separacion != "Sin separar" else "—"
                previstos.append({"Origen": ", ".join(grupo["ArchivoOrigen"].unique()),
                    "CSV_ALTAS": nombre_csv, "Excel normal": nombre_excel, "Registros": len(grupo)})
            st.dataframe(pd.DataFrame(previstos), hide_index=True, use_container_width=True)
            st.caption("Los .xlsm originales solo se entregarán renombrados cuando su nombre no corresponda al contenido.")
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
                                                  preparar_csv_banner, modo_separacion=modo_separacion)
                st.session_state.zip_solicitudes_csv_altas = zip_final
                st.session_state.t1_modo_generado = modo_separacion
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
        st.session_state.df_manual_fijo = pd.DataFrame(columns=COLUMNAS_CSV + ["Responsable", "Nivel", "Clúster"])
    if "Clúster" not in st.session_state.df_manual_fijo.columns:
        st.session_state.df_manual_fijo["Clúster"] = st.session_state.df_manual_fijo["datocomplementario"].apply(cluster_oficial)
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
                nuevo.update({"Responsable": responsable_normalizado(responsable), "Nivel": nivel, "Clúster": cluster})
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
        if (df_out["PERIODO"].apply(limpiar_clave_texto).eq("").any()
                or df_out["Responsable"].apply(limpiar_clave_texto).eq("").any()
                or df_out["Clúster"].apply(limpiar_clave_texto).eq("").any()):
            st.warning("Hay registros manuales sin Periodo, Responsable o Clúster.")
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

                df_out["_ClusterArchivo"] = df_out["Clúster"].apply(cluster_oficial)
                buffer = io.BytesIO()
                with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
                    for (periodo, resp, cluster_archivo), grupo in df_out.groupby(
                        ["PERIODO", "Responsable", "_ClusterArchivo"], sort=False
                    ):
                        nombre = nombre_csv_altas(periodo, resp, cluster_archivo)
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
        claves = ["t3_modo", "t3_argos_completo", "t3_csv_completo", "t3_excel_completo", "t3_argos_rapido", "t3_csv_rapido", "final_argos_zip", "df_cruce_rapido", "t3_alertas", "t3_resultado_editado", "t3_editor_version", "t3_columnas", "t3_firma_completo", "t3_firma_rapido", "t3_tabla_revision", "t3_informe_zip", "t3_diferencias", "t3_match", "t3_firma_revision", "t3_salida", "t3_incluir_original"]
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
        conflictos = df.groupby(["_periodo", "_nrc"])[llave_nrc].nunique().gt(1).any(axis=1)
        avisos = [{"Tipo": "NRC contradictorio en ARGOS", "Periodo": p, "NRC": n} for p, n in conflictos[conflictos].index]
        df = df.drop_duplicates(subset=["_periodo", "_nrc"] + llave_nrc).copy()
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

    # 3. FUNCIONES DE RECONCILIACION
    T3_COLUMNAS_FINAL = ["PERIODO", "CLUSTER", "NIVEL", "REQUERIMIENTO", "RESPONSABLE", "NRC", "SUBJ", "COURSE", "NOMBRE_MATERIA", "SEDE", "PARTE_PERIODO", "STATUS", "CAPACIDAD", "SECCION", "TIPO_HORARIO", "METODO_EDUCATIVO", "SOCIO_INTEGRACION", "MODO_CALIFICAR", "SESION"]
    T3_MAPA = {"PERIODO": "PERIODO", "SUBJ": "SUBJ", "COURSE": "COURSE", "SEDE": "SEDE", "PARTE_PERIODO": "PARTEPERIODO", "STATUS": "STATUS", "CAPACIDAD": "CAPACIDAD", "SECCION": "SECCION", "TIPO_HORARIO": "TIPODEHORARIO", "METODO_EDUCATIVO": "METODO_EDUCATIVO", "SOCIO_INTEGRACION": "SOCIODEINTEGRACION", "MODO_CALIFICAR": "MODODECALIFICAR", "SESION": "SESION"}
    T3_ALIAS = {"PERIODO": ["periodo"], "CLUSTER": ["cluster"], "NIVEL": ["nivel"], "REQUERIMIENTO": ["requerimiento"], "RESPONSABLE": ["responsable"], "SUBJ": ["subj", "subject"], "COURSE": ["course", "crse"], "NOMBRE_MATERIA": ["nombremateria", "nombredelamateria"], "SEDE": ["sede", "campus"], "PARTE_PERIODO": ["parteperiodo"], "STATUS": ["status", "estatus"], "CAPACIDAD": ["capacidad"], "SECCION": ["seccion", "grupo"], "TIPO_HORARIO": ["tipohorario"], "METODO_EDUCATIVO": ["metodoeducativo"], "SOCIO_INTEGRACION": ["sociointegracion", "sociodeintegracion"], "MODO_CALIFICAR": ["modocalificar", "mododecalificar"], "SESION": ["sesion"]}

    def t3_argos_unidos(archivos):
        partes, avisos = [], []
        for f in archivos:
            d, _, a = leer_argos_t3(f)
            partes.append(d)
            avisos.extend(a)
        if not partes: raise ValueError("Carga al menos un ARGOS/PA.")
        d = pd.concat(partes, ignore_index=True)
        campos = ["_periodo", "_subj", "_crse", "_grupo", "_nivel", "_cluster", "_sede"]
        conflictos = d.groupby(["_periodo", "_nrc"])[campos].nunique().gt(1).any(axis=1)
        for p, n in conflictos[conflictos].index:
            avisos.append({"Tipo": "NRC contradictorio entre reportes", "Periodo": p, "NRC": n})
        d = d.drop_duplicates(subset=["_periodo", "_nrc"] + campos)
        d["_llave"] = d[["_periodo", "_nivel", "_cluster", "_subj", "_crse", "_grupo"]].agg("_".join, axis=1)
        mapa = d.groupby("_llave")["_nrc"].agg(lambda x: sorted(set(x))).to_dict()
        return d, mapa, avisos

    def t3_leer_excels(archivos):
        partes = []
        for f in archivos:
            wb = openpyxl.load_workbook(io.BytesIO(f.getvalue()), read_only=True, data_only=True)
            hoja = next((x for x in wb.sheetnames if x.strip().upper() == HOJA_ALTAS), None)
            if hoja is None: raise ValueError(f"{f.name}: no existe la hoja ALTAS.")
            ws = wb[hoja]
            encabezados = [normalizar_para_busqueda(c.value) for c in ws[1]]
            if len(encabezados) != len(set(encabezados)): raise ValueError(f"{f.name}: encabezados duplicados.")
            posiciones = {campo: next((encabezados.index(a) for a in alias if a in encabezados), None) for campo, alias in T3_ALIAS.items()}
            requeridos = ["PERIODO", "SUBJ", "COURSE", "SECCION", "SEDE"]
            if any(posiciones[c] is None for c in requeridos): raise ValueError(f"{f.name}: faltan {', '.join(c for c in requeridos if posiciones[c] is None)}.")
            for nf, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                reg = {campo: limpiar_clave_texto(row[pos]) if pos is not None and pos < len(row) else "" for campo, pos in posiciones.items()}
                if not any(reg[c] for c in requeridos): continue
                reg.update({"_ArchivoExcel": f.name, "_FilaExcel": nf})
                partes.append(reg)
            wb.close()
        if not partes: raise ValueError("Los Excel no tienen filas ALTAS.")
        return pd.DataFrame(partes).reset_index(drop=True)

    def t3_n(x, campo=""):
        return ultra_limpiar_seccion(x) if campo == "SECCION" else ultra_limpiar(x)

    def t3_similares(ex, csv):
        # La diferencia de seccion, horario, metodo, modo o capacidad no impide proponer una pareja.
        campos = ["PERIODO", "SUBJ", "COURSE", "SEDE", "PARTE_PERIODO", "STATUS", "CAPACIDAD", "SECCION", "TIPO_HORARIO", "METODO_EDUCATIVO", "MODO_CALIFICAR", "SESION"]
        iguales = sum(t3_n(ex[c]) == t3_n(csv[T3_MAPA[c]], c) for c in campos)
        periodo_igual = t3_n(ex["PERIODO"]) == t3_n(csv["PERIODO"])
        return iguales, periodo_igual

    def t3_proponer(excel, csv):
        # Primera pasada: coincidencia exacta de todos los campos disponibles, unica.
        pendientes = set(csv.index)
        asignadas, motivos = {}, {}
        campos = ["PERIODO", "SUBJ", "COURSE", "SEDE", "PARTE_PERIODO", "STATUS", "CAPACIDAD", "SECCION", "TIPO_HORARIO", "METODO_EDUCATIVO", "MODO_CALIFICAR", "SESION"]
        def firma_ex(r): return tuple(t3_n(r[c], c) for c in campos)
        def firma_csv(r): return tuple(t3_n(r[T3_MAPA[c]], c) for c in campos)
        grupos_ex, grupos_csv = {}, {}
        for i, r in excel.iterrows(): grupos_ex.setdefault(firma_ex(r), []).append(i)
        for j, r in csv.iterrows(): grupos_csv.setdefault(firma_csv(r), []).append(j)
        for k, ie in grupos_ex.items():
            jc = grupos_csv.get(k, [])
            if len(ie) == len(jc) == 1:
                asignadas[ie[0]], motivos[ie[0]] = jc[0], "Exacta"
                pendientes.discard(jc[0])
        # Segunda pasada: orden dentro de cada periodo, responsable/cluster cuando se conocen.
        # Solo sugerencias, nunca confirmacion automatica.
        propuestas = []
        for i, r in excel.iterrows():
            if i in asignadas: continue
            candidatos = []
            for j in pendientes:
                c = csv.loc[j]
                puntaje, mismo_periodo = t3_similares(r, c)
                if not mismo_periodo: continue
                # Si ambos tienen cluster explicito, deben ser iguales; Banner Bachillerato puede tener alias.
                cl = t3_n(r.get("CLUSTER", ""))
                cc = t3_n(c.get("datocomplementario", ""))
                if cl and cc and cl != cc and not (t3_n(r.get("NIVEL", "")) == "BACHILLERATO" and cc == "BACHILLERATO"):
                    continue
                candidatos.append((puntaje, j))
            candidatos.sort(key=lambda x: (-x[0], x[1]))
            mejor = candidatos[0][0] if candidatos else 0
            sugeridos = [j for p, j in candidatos if p == mejor and p >= 6]
            propuestas.append({"Fila": i, "Archivo Excel": r["_ArchivoExcel"], "Fila Excel": int(r["_FilaExcel"]), "Periodo": r["PERIODO"], "Materia Excel": f'{r["SUBJ"]} {r["COURSE"]}', "Seccion Excel": r["SECCION"], "CSV sugerido": sugeridos[0] if len(sugeridos) == 1 else -1, "Coincidencias": mejor, "Estado": "Sugerencia para confirmar" if len(sugeridos) == 1 else "Revisar manualmente"})
        return asignadas, motivos, pd.DataFrame(propuestas)

    def t3_generar_resultado(excel, csv, asignadas, argos_df, mapa):
        if len(asignadas) != len(excel) or len(set(asignadas.values())) != len(asignadas):
            raise ValueError("Faltan correspondencias unicas entre Excel y CSV.")
        if len(asignadas) != len(csv): raise ValueError("El numero de filas de ALTAS y CSV no coincide. Revisa registros agregados o eliminados.")
        final, diferencias = [], []
        csv_ordenado = csv.loc[[asignadas[i] for i in excel.index]].copy().reset_index(drop=True)
        nrc, avisos = cruzar_nrc_t3(csv_ordenado, argos_df, mapa, "CSV definitivo")
        if nrc.eq("").any(): raise ValueError(f"ARGOS no tiene NRC unico para {int(nrc.eq('').sum())} registros. Revisa las incidencias antes de exportar: {avisos[:5]}")
        for i, e in excel.iterrows():
            c = csv.loc[asignadas[i]]
            registro = {k: limpiar_clave_texto(e.get(k, "")) for k in T3_COLUMNAS_FINAL if k != "NRC"}
            for destino, fuente in T3_MAPA.items():
                viejo = registro.get(destino, "")
                nuevo = limpiar_clave_texto(c[fuente])
                if t3_n(viejo, destino) != t3_n(nuevo, destino):
                    diferencias.append({"Archivo Excel": e["_ArchivoExcel"], "Fila Excel": int(e["_FilaExcel"]), "Archivo CSV": c["_ArchivoCSV"], "Fila CSV": int(c["_FilaCSV"]), "Campo": destino, "Anterior": viejo, "Definitivo": nuevo})
                registro[destino] = nuevo
            registro["NRC"] = nrc.iloc[i]
            registro["_ArchivoExcel"], registro["_FilaExcel"] = e["_ArchivoExcel"], int(e["_FilaExcel"])
            final.append(registro)
        return pd.DataFrame(final), pd.DataFrame(diferencias)

    def t3_nombre_grupo(df, sufijo="CRNs Finales.xlsx"):
        def unir(xs):
            xs = sorted({limpiar_clave_texto(x) for x in xs if limpiar_clave_texto(x)})
            return " y ".join(xs) if len(xs) == 2 else ", ".join(xs)
        p = "-".join(sorted({limpiar_clave_texto(x) for x in df["PERIODO"]}))
        r = unir(df["RESPONSABLE"].map(responsable_normalizado)) or "Sin Responsable"
        c = unir(df["CLUSTER"].map(cluster_oficial)) or "Sin Cluster"
        return f"{p} {r} {c} {sufijo}"

    def t3_excel_bytes(df):
        b = io.BytesIO()
        with pd.ExcelWriter(b, engine="openpyxl") as writer:
            df[T3_COLUMNAS_FINAL].to_excel(writer, index=False, sheet_name="CRNs_Finales")
            formatear_excel_t3(writer.sheets["CRNs_Finales"])
        return b.getvalue()

    # 4. MODO COMPLETO: comparar, revisar, corregir y finalmente recuperar NRC.
    if modo.startswith("📦"):
        st.subheader("📦 Conciliacion Excel - CSV - ARGOS")
        c1, c2, c3 = st.columns(3)
        argos = c1.file_uploader("ARGOS / PA por periodo (.csv)", type=["csv"], accept_multiple_files=True, key="t3_argos_completo")
        archivos_csv = c2.file_uploader("CSV definitivos de Banner", type=["csv"], accept_multiple_files=True, key="t3_csv_completo")
        archivos_excel = c3.file_uploader("Excel ALTAS originales o separados", type=["xlsx", "xlsm"], accept_multiple_files=True, key="t3_excel_completo")
        opciones = {"Como venian originalmente": [], "Por periodo": ["PERIODO"], "Por cluster": ["CLUSTER"], "Por responsable": ["RESPONSABLE"], "Por periodo y cluster": ["PERIODO", "CLUSTER"], "Por periodo y responsable": ["PERIODO", "RESPONSABLE"], "Por cluster y responsable": ["CLUSTER", "RESPONSABLE"], "Por periodo, cluster y responsable": ["PERIODO", "CLUSTER", "RESPONSABLE"]}
        salida = st.selectbox("Separacion de Excel finales", list(opciones), key="t3_salida")
        incluir_original = st.checkbox("Incluir tambien los Excel como venian originalmente", value=False, key="t3_incluir_original")
        firma = firma_t3(*(argos or []), *(archivos_csv or []), *(archivos_excel or []))
        if st.session_state.get("t3_firma_completo") != firma:
            for k in ("t3_match", "t3_tabla_revision", "final_argos_zip", "t3_alertas"):
                st.session_state.pop(k, None)
            st.session_state.t3_firma_completo = firma
        if argos and archivos_csv and archivos_excel and st.button("1. Comparar Excel y CSV", type="primary", key="t3_comparar", use_container_width=True):
            try:
                excel = t3_leer_excels(archivos_excel)
                partes = []
                for f in archivos_csv:
                    d = leer_csv_t3(f)
                    d["_ArchivoCSV"], d["_FilaCSV"] = f.name, range(2, len(d) + 2)
                    partes.append(d)
                csv = pd.concat(partes, ignore_index=True)
                if len(excel) != len(csv): raise ValueError(f"Excel: {len(excel)} registros; CSV: {len(csv)}. Deben coincidir antes de relacionar filas.")
                asignadas, motivos, propuestas = t3_proponer(excel, csv)
                st.session_state.t3_match = {"excel": excel, "csv": csv, "asignadas": asignadas, "motivos": motivos}
                st.session_state.t3_tabla_revision = propuestas
                st.session_state.final_argos_zip = None
                st.success(f"{len(asignadas)} coincidencias exactas. {len(propuestas)} filas requieren revision.")
            except Exception as e: st.error(f"Error de comparacion: {e}")
        datos = st.session_state.get("t3_match")
        if datos:
            excel, csv = datos["excel"], datos["csv"]
            asignadas = dict(datos["asignadas"])
            propuestas = st.session_state.get("t3_tabla_revision")
            if isinstance(propuestas, pd.DataFrame) and not propuestas.empty:
                st.subheader("Correspondencias pendientes")
                st.caption("Selecciona la fila CSV correcta. La sugerencia es orientativa; no se aplica sin confirmacion.")
                usados = set(asignadas.values())
                for _, r in propuestas.iterrows():
                    i = int(r["Fila"])
                    if i in asignadas: continue
                    candidatos = [j for j in csv.index if j not in usados and t3_n(csv.loc[j, "PERIODO"]) == t3_n(excel.loc[i, "PERIODO"])]
                    etiquetas = {j: f'{csv.loc[j,"_ArchivoCSV"]} | fila {csv.loc[j,"_FilaCSV"]} | {csv.loc[j,"SUBJ"]} {csv.loc[j,"COURSE"]} | seccion {csv.loc[j,"SECCION"]}' for j in candidatos}
                    if not candidatos:
                        st.error(f'{r["Archivo Excel"]}, fila {r["Fila Excel"]}: no quedan candidatos del mismo periodo.')
                        continue
                    sugerido = int(r["CSV sugerido"])
                    seleccion = st.selectbox(f'{r["Archivo Excel"]} - fila {r["Fila Excel"]} - {r["Materia Excel"]} - seccion {r["Seccion Excel"]}', [-1] + candidatos, index=([-1] + candidatos).index(sugerido) if sugerido in candidatos else 0, format_func=lambda j: "-- Seleccionar CSV --" if j == -1 else etiquetas[j], key=f"t3_match_{i}")
                    if seleccion != -1:
                        if seleccion in usados: st.error("Esta fila CSV ya fue asignada.")
                        else:
                            asignadas[i] = seleccion
                            usados.add(seleccion)
                if st.button("Guardar correspondencias", key="t3_guardar_match", use_container_width=True):
                    if len(asignadas) != len(excel): st.warning(f"Faltan {len(excel)-len(asignadas)} correspondencias.")
                    else:
                        datos["asignadas"] = asignadas
                        st.session_state.t3_match = datos
                        st.session_state.t3_tabla_revision = pd.DataFrame()
                        st.session_state.final_argos_zip = None
                        st.rerun()
            if len(datos["asignadas"]) == len(excel):
                st.success("Correspondencias completas. Puedes generar los NRC.")
                if st.button("2. Aplicar CSV y generar Excel con NRC", type="primary", key="t3_generar", use_container_width=True):
                    try:
                        argos_df, mapa, avisos = t3_argos_unidos(argos)
                        final, diferencias = t3_generar_resultado(excel, csv, datos["asignadas"], argos_df, mapa)
                        st.session_state.t3_alertas = avisos
                        st.session_state.t3_diferencias = diferencias
                        grupos = [(k, g) for k, g in final.groupby(opciones[salida], sort=False, dropna=False)] if opciones[salida] else [(k, g) for k, g in final.groupby("_ArchivoExcel", sort=False)]
                        if incluir_original and opciones[salida]: grupos += [(f"Original {k}", g) for k, g in final.groupby("_ArchivoExcel", sort=False)]
                        b, rutas = io.BytesIO(), set()
                        with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
                            for _, grupo in grupos:
                                nombre = t3_nombre_grupo(grupo)
                                z.writestr(ruta_zip_unica("CRNS_FINALES_SIN_MACROS", nombre, rutas), t3_excel_bytes(grupo))
                            # Conservar el libro con macros solo si corresponde a un origen completo.
                            for f in archivos_excel:
                                if not f.name.lower().endswith(".xlsm"): continue
                                g = final[final["_ArchivoExcel"] == f.name]
                                if g.empty: continue
                                wb = openpyxl.load_workbook(io.BytesIO(f.getvalue()), keep_vba=True)
                                hoja = next((h for h in wb.sheetnames if h.strip().upper() == HOJA_ALTAS), None)
                                if not hoja: continue
                                ws = wb[hoja]
                                encabezados = {normalizar_para_busqueda(c.value): c.column for c in ws[1] if c.value is not None}
                                for _, reg in g.iterrows():
                                    fila = int(reg["_FilaExcel"])
                                    for campo, valor in reg.items():
                                        if campo.startswith("_") or campo == "NRC": continue
                                        col = next((encabezados[a] for a in T3_ALIAS.get(campo, []) if a in encabezados), None)
                                        if col is not None: ws.cell(fila, col).value = valor
                                if HOJA_SALIDA_NRC in wb.sheetnames: del wb[HOJA_SALIDA_NRC]
                                wsn = wb.create_sheet(HOJA_SALIDA_NRC)
                                wsn.append(["Fila ALTAS"] + T3_COLUMNAS_FINAL)
                                for _, reg in g.iterrows(): wsn.append([int(reg["_FilaExcel"])] + [reg[c] for c in T3_COLUMNAS_FINAL])
                                formatear_excel_t3(wsn)
                                buf = io.BytesIO()
                                wb.save(buf)
                                nombre = f.name.rsplit(".", 1)[0] + "_con_NRC.xlsm"
                                z.writestr(ruta_zip_unica("SOLICITUDES_CON_NRC", nombre, rutas), buf.getvalue())
                        st.session_state.final_argos_zip = b.getvalue()
                        st.success(f"Generados {len(grupos)} Excel normales; {len(diferencias)} diferencias corregidas desde CSV.")
                    except Exception as e:
                        st.session_state.final_argos_zip = None
                        st.error(f"No se pudo generar: {e}")
        if isinstance(st.session_state.get("t3_diferencias"), pd.DataFrame) and not st.session_state.t3_diferencias.empty:
            with st.expander("Correcciones aplicadas desde CSV", expanded=True):
                st.dataframe(st.session_state.t3_diferencias, hide_index=True, use_container_width=True)
        if st.session_state.get("t3_alertas"):
            with st.expander("Advertencias ARGOS"):
                st.dataframe(pd.DataFrame(st.session_state.t3_alertas), hide_index=True, use_container_width=True)
        if st.session_state.get("final_argos_zip"):
            st.download_button("Descargar Excel finales (.ZIP)", st.session_state.final_argos_zip, file_name="CRNs_Finales.zip", mime="application/zip", type="primary", use_container_width=True)

    # 4. MODO RAPIDO: tabla editable y descarga.
    else:
        st.subheader("⚡ Cruce rápido de NRC")
        c1, c2 = st.columns(2)
        argos = c1.file_uploader("📊 ARGOS/PA (.csv)", type=["csv"], accept_multiple_files=True, key="t3_argos_rapido")
        archivos_csv = c2.file_uploader("📝 CSV finales", type=["csv"], accept_multiple_files=True, key="t3_csv_rapido")
        firma = firma_t3(*(argos or []), *(archivos_csv or []))
        if st.session_state.get("t3_firma_rapido") != firma:
            st.session_state.df_cruce_rapido = None
            st.session_state.t3_resultado_editado = None
            st.session_state.t3_editor_version = 0
            st.session_state.t3_firma_rapido = firma
            st.session_state.pop("t3_columnas", None)
        if argos and archivos_csv and st.button("⚡ Cruzar NRC y generar tabla", type="primary", use_container_width=True, key="t3_procesar_rapido"):
            try:
                argos_df, mapa, avisos = t3_argos_unidos(argos)
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
# PESTAÑA 4: CRUCES - NUEVOS, CONTINUACIÓN Y CONSOLIDADOS
# Requiere imports y funciones de la parte inicial del Streamlit.
# En la declaración de pestañas agrega tab4 y "3️⃣ Proceso: Cruces".
# ============================================================
with tab4:
    st.header("🔗 Gestión de Cruces")
    st.caption("Plantillas con hoja Cruces; claves AA-ZZ independientes por periodo. No se ejecutan macros.")
    COL_COMPLETO = ["Clave", "NRC", "Tipo", "Responsable", "Periodo", "SUB", "COURSE", "Comentarios"]
    COL_TI = ["Clave", "NRC", "Periodo"]
    COLUMNAS_ORIGEN = ["NO_CRUCE", "PERIODO", "RESPONSABLE", "NRC", "TIPO", "CLUSTER_CRN", "SUBJ", "COURSE", "MATERIA"]
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
        return (ord(c[0])-65)*26 + ord(c[1])-65
    def t4_num_clave(n):
        if not 0 <= n < 676: raise ValueError("Se agotaron las claves AA-ZZ para un periodo.")
        return chr(65+n//26) + chr(65+n%26)
    def t4_leer_tabla(archivo, hoja=None):
        datos = io.BytesIO(archivo.getvalue())
        if archivo.name.lower().endswith(".csv"):
            ultimo = None
            for cod in ("utf-8-sig", "cp1252"):
                try: return pd.read_csv(io.BytesIO(archivo.getvalue()), dtype=str, keep_default_na=False, encoding=cod, sep=None, engine="python")
                except (UnicodeError, pd.errors.ParserError) as e: ultimo = e
            raise ValueError(f"No se pudo leer {archivo.name}: {ultimo}")
        libro = pd.ExcelFile(datos)
        nombre = next((h for h in libro.sheet_names if t4_col(h) == t4_col(hoja)), None) if hoja else next((h for h in libro.sheet_names if t4_col(h) in ("CRUCES", "TI")), libro.sheet_names[0])
        if nombre is None: raise ValueError(f"{archivo.name}: falta la hoja '{hoja}'.")
        return libro.parse(nombre, dtype=str, keep_default_na=False)
    def t4_estandarizar(df, origen=False):
        d = df.copy()
        mapa = {t4_col(c): c for c in (COLUMNAS_ORIGEN if origen else COL_COMPLETO)}
        mapa.update({"SUBJECT":"SUBJ", "AREA":"SUBJ", "CRSE":"COURSE", "NOCURSO":"COURSE"} if origen else {"SUBJ":"SUB", "SUBJECT":"SUB"})
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
            if f.name in nombres: raise ValueError(f"Hay archivos con el mismo nombre: {f.name}. Renómbralos para diferenciarlos.")
            nombres.add(f.name)
            d = t4_estandarizar(t4_leer_tabla(f, "Cruces"), origen=True)
            d = d[d[["NO_CRUCE", "PERIODO", "NRC", "TIPO"]].ne("").any(axis=1)].copy()
            d["_origen"] = f.name
            d["_fila"] = d.index + 2
            d["_orden_archivo"] = pos
            piezas.append(d)
        if not piezas: raise ValueError("No hay archivos de Cruces.")
        return pd.concat(piezas, ignore_index=True)
    def t4_validar_nuevos(df):
        problemas = []
        for i, r in df.iterrows():
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
        if (repetidos > 1).any():
            for periodo, nrc in repetidos[repetidos > 1].index:
                problemas.append(f"Periodo {periodo}: NRC {nrc} aparece más de una vez en las plantillas.")
        return problemas
    def t4_leer_anteriores(archivos):
        completos, tis = [], []
        for f in archivos:
            original = t4_leer_tabla(f)
            completo = set(t4_col(c) for c in COL_COMPLETO).issubset(set(t4_col(c) for c in original.columns))
            d = t4_estandarizar(original)
            d = d[d[["Clave", "NRC", "Periodo"]].ne("").any(axis=1)].copy()
            if d[["Clave", "NRC", "Periodo"]].eq("").any().any():
                raise ValueError(f"{f.name}: hay filas anteriores incompletas.")
            for _, r in d.iterrows(): t4_clave_num(r["Clave"])
            (completos if completo else tis).append(d[COL_COMPLETO if completo else COL_TI])
        full = pd.concat(completos, ignore_index=True) if completos else pd.DataFrame(columns=COL_COMPLETO)
        ti = pd.concat(tis + [full[COL_TI]], ignore_index=True) if tis or not full.empty else pd.DataFrame(columns=COL_TI)
        if not ti.empty:
            ti = ti.drop_duplicates()
            por_nrc = ti.groupby(["Periodo", "NRC"])["Clave"].nunique()
            if (por_nrc > 1).any(): raise ValueError("Los anteriores contienen NRC asociados a claves distintas en el mismo periodo.")
            por_clave = ti.groupby(["Periodo", "Clave"])["NRC"].nunique()
            if (por_clave < 2).any(): raise ValueError("Hay claves anteriores con menos de dos NRC.")
            if not full.empty:
                cruce = full.merge(ti, on=COL_TI, how="outer", indicator=True)
                if (cruce["_merge"] == "left_only").any(): raise ValueError("Inconsistencia entre archivos anteriores.")
        if not full.empty and full.duplicated(subset=COL_TI).any():
            raise ValueError("Los anteriores completos tienen registros duplicados para la misma clave, NRC y periodo.")
        return full.drop_duplicates(), ti
    def t4_generar(nuevos, anteriores_full, anteriores_ti, ultimas):
        problemas = t4_validar_nuevos(nuevos)
        if problemas: raise ValueError("\n".join(problemas[:35]))
        ocupados = {}
        for _, r in anteriores_ti.iterrows():
            p = r["Periodo"]
            ocupados.setdefault(p, set()).add(t4_clave_num(r["Clave"]))
        for p, ultima in ultimas.items():
            if ultima: ocupados.setdefault(p, set()).update(range(t4_clave_num(ultima)+1))
        prev_nrc = set(zip(anteriores_ti["Periodo"], anteriores_ti["NRC"]))
        nuevo_nrc = set(zip(nuevos["PERIODO"], nuevos["NRC"]))
        repetidos = prev_nrc & nuevo_nrc
        if repetidos: raise ValueError(f"Estos NRC ya están en cruces anteriores: {sorted(repetidos)[:15]}. Revisa antes de continuar.")
        filas = []
        for (periodo, archivo, no), g in nuevos.groupby(["PERIODO", "_origen", "NO_CRUCE"], sort=False):
            usadas = ocupados.setdefault(periodo, set())
            siguiente = max(usadas, default=-1) + 1
            if siguiente >= 676: raise ValueError(f"Periodo {periodo}: no quedan claves AA-ZZ.")
            usadas.add(siguiente)
            clave = t4_num_clave(siguiente)
            g = g.assign(_tipo_orden=g["TIPO"].map(lambda x: 0 if t4_normal(x) == "MANDANTE" else 1)).sort_values(["_tipo_orden", "_fila"], kind="stable")
            for _, r in g.iterrows():
                filas.append({"Clave": clave, "NRC": r["NRC"], "Tipo": "Mandante" if t4_normal(r["TIPO"]) == "MANDANTE" else "Cruce",
                             "Responsable": r["RESPONSABLE"], "Periodo": periodo, "SUB": r["SUBJ"], "COURSE": r["COURSE"], "Comentarios": ""})
        full = pd.DataFrame(filas, columns=COL_COMPLETO)
        ti = full[COL_TI].copy()
        if not anteriores_ti.empty:
            cons_ti = pd.concat([anteriores_ti, ti], ignore_index=True).drop_duplicates()
        else: cons_ti = ti.copy()
        cons_full = None
        if anteriores_ti.empty: cons_full = full.copy()
        elif not anteriores_full.empty:
            # Solo es completo cuando todos los NRC anteriores tienen sus 8 campos.
            cubiertos = set(map(tuple, anteriores_full[COL_TI].to_numpy()))
            todos = set(map(tuple, anteriores_ti[COL_TI].to_numpy()))
            if cubiertos == todos: cons_full = pd.concat([anteriores_full, full], ignore_index=True).drop_duplicates()
        return full, ti, cons_full, cons_ti
    def t4_excel(df, hoja):
        salida = io.BytesIO()
        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=hoja, index=False)
            ws = writer.sheets[hoja]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for celda in ws[1]:
                celda.font = Font(bold=True, color="FFFFFF")
                celda.fill = PatternFill(fill_type="solid", fgColor="7030A0")
            for col in ws.columns:
                ancho = min(40, max(12, max(len(str(c.value or "")) for c in col)+2))
                ws.column_dimensions[col[0].column_letter].width = ancho
        return salida.getvalue()

    if st.button("🔄 Limpiar Cruces", key="t4_limpiar"):
        for k in list(st.session_state):
            if k.startswith("t4_"): st.session_state.pop(k, None)
        st.rerun()
    st.subheader("1. Cargar plantillas")
    archivos_nuevos = st.file_uploader("Plantillas con hoja Cruces", type=["xlsx", "xlsm"], accept_multiple_files=True, key="t4_nuevos")
    modo_t4 = st.radio("Continuación de claves", ["Empezar en AA", "Cargar cruces anteriores", "Indicar última clave manualmente"], key="t4_modo")
    archivos_anteriores = []
    if modo_t4 == "Cargar cruces anteriores":
        archivos_anteriores = st.file_uploader("Cruces anteriores: completo o TI (.xlsx, .xlsm, .csv)", type=["xlsx", "xlsm", "csv"], accept_multiple_files=True, key="t4_anteriores") or []
    if modo_t4 == "Indicar última clave manualmente":
        st.caption("Escribe una última clave distinta por periodo. Deja vacío si ese periodo comienza en AA.")
    st.subheader("2. Revisar y asignar claves")
    firma = tuple((f.name, hashlib.sha256(f.getvalue()).hexdigest()) for f in list(archivos_nuevos or []) + list(archivos_anteriores)) + (modo_t4,)
    if modo_t4 == "Indicar última clave manualmente":
        firma += tuple(sorted((k, v) for k, v in st.session_state.items() if k.startswith("t4_ultima_")))
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
                    ultimas[p] = cols[i % len(cols)].text_input(f"Última clave del periodo {p}", key=f"t4_ultima_{p}").strip().upper()
            if modo_t4 == "Cargar cruces anteriores" and not archivos_anteriores:
                st.info("Carga los archivos anteriores para continuar la numeración.")
            else:
                anteriores_full, anteriores_ti = t4_leer_anteriores(archivos_anteriores) if archivos_anteriores else (pd.DataFrame(columns=COL_COMPLETO), pd.DataFrame(columns=COL_TI))
                if modo_t4 == "Empezar en AA" and not anteriores_ti.empty: raise ValueError("No se pueden cargar anteriores en modo AA.")
                if st.button("🔍 Validar y generar claves", type="primary", key="t4_generar"):
                    full, ti, cons_full, cons_ti = t4_generar(nuevos, anteriores_full, anteriores_ti, ultimas)
                    st.session_state.t4_resultados = (full, ti, cons_full, cons_ti)
                    st.session_state.t4_editor_version += 1
                    st.success(f"Se asignaron {full['Clave'].groupby([full['Periodo'], full['Clave']]).ngroups} cruces nuevos.")
        except Exception as e:
            st.error(f"Revisión de cruces: {e}")
    resultado = st.session_state.get("t4_resultados")
    if resultado is not None:
        full, ti, cons_full, cons_ti = resultado
        st.subheader("3. Revisar resultados")
        st.caption("Las claves y NRC no son editables para evitar romper la relación entre Mandante e hijos. Puedes editar los datos descriptivos de los nuevos cruces.")
        with st.form("t4_edicion"):
            editado = st.data_editor(full, hide_index=True, num_rows="fixed", use_container_width=True,
                                     disabled=["Clave", "NRC", "Tipo", "Periodo"], key=f"t4_editor_{st.session_state.t4_editor_version}")
            guardar = st.form_submit_button("💾 Guardar cambios")
        if guardar:
            full = editado.copy()
            ti = full[COL_TI].copy()
            if cons_full is not None:
                anteriores = cons_full.iloc[:len(cons_full)-len(resultado[0])].copy()
                cons_full = pd.concat([anteriores, full], ignore_index=True)
            if len(cons_ti) >= len(resultado[1]):
                anteriores_ti = cons_ti.iloc[:len(cons_ti)-len(resultado[1])].copy()
                cons_ti = pd.concat([anteriores_ti, ti], ignore_index=True)
            st.session_state.t4_resultados = (full, ti, cons_full, cons_ti)
            st.session_state.t4_editor_version += 1
            st.rerun()
        salida = st.radio("¿Qué deseas descargar?", ["Solo nuevos", "Consolidado completo", "Ambas opciones"], index=2, horizontal=True, key="t4_salida")
        archivos = {}
        if salida in ("Solo nuevos", "Ambas opciones"):
            archivos["Cruces_Nuevos_Completo.xlsx"] = t4_excel(full[COL_COMPLETO], "Cruces")
            archivos["Cruces_Nuevos_TI.xlsx"] = t4_excel(ti[COL_TI], "TI")
        if salida in ("Consolidado completo", "Ambas opciones"):
            archivos["Cruces_Consolidado_TI.xlsx"] = t4_excel(cons_ti[COL_TI], "TI")
            if cons_full is not None: archivos["Cruces_Consolidado_Completo.xlsx"] = t4_excel(cons_full[COL_COMPLETO], "Cruces")
            else: st.warning("No se genera consolidado completo de 8 columnas porque los anteriores solo tienen Clave, NRC y Periodo.")
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
            for nombre, datos in archivos.items(): z.writestr(nombre, datos)
        st.download_button("📥 Descargar cruces (.ZIP)", buffer.getvalue(), file_name="CRUCES_PROCESADOS.zip",
                           mime="application/zip", type="primary", use_container_width=True, key="t4_descargar")

