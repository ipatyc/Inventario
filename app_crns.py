
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
    # Ejecutar SOLO al finalizar la revision. No reservar secciones manuales:
    # Licenciatura y Posgrado se numeran con el primer hueco disponible.
    requeridas = ["Periodo", "Subject", "Course", "Nivel", "Sección"]
    faltantes = [c for c in requeridas if c not in df_altas.columns]
    if faltantes: raise ValueError(f"ALTAS: faltan columnas {faltantes}")
    salida = df_altas.copy().reset_index(drop=True)
    salida["Sección Original"] = salida["Sección"].apply(clave_seccion)
    salida["Sección"] = salida["Sección Original"]
    ocupadas, siguiente, errores = {}, {}, []
    if pa is not None:
        for _, fila in pa.iterrows():
            llave = tuple(normalizar_para_cruce(fila[c]) for c in ("periodo", "area", "nocurso"))
            grupo = clave_seccion(fila["grupo"])
            if all(llave) and grupo: ocupadas.setdefault(llave, set()).add(grupo)
    for i, fila in salida.iterrows():
        nivel = normalizar_para_cruce(fila["Nivel"])
        if nivel not in ("LICENCIATURA", "POSGRADO"): continue  # Bachillerato: sin cambio
        llave = tuple(normalizar_para_cruce(fila[c]) for c in ("Periodo", "Subject", "Course"))
        if not all(llave):
            errores.append({"Fila": i + 2, "Tipo": "Sección", "Detalle": "Periodo, SUBJ o COURSE vacío."})
            salida.at[i, "Sección"] = ""
            continue
        usadas = ocupadas.setdefault(llave, set())
        numero = siguiente.get(llave, 1)
        while numero <= 99 and f"{numero:02d}" in usadas: numero += 1
        if numero > 99:
            errores.append({"Fila": i + 2, "Tipo": "Sección", "Detalle": f"Sin secciones disponibles 01-99 para {llave}."})
            salida.at[i, "Sección"] = ""
        else:
            valor = f"{numero:02d}"
            salida.at[i, "Sección"] = valor
            usadas.add(valor)
            siguiente[llave] = numero + 1
    salida["Sección Modificada"] = salida["Sección Original"] != salida["Sección"]
    return salida, pd.DataFrame(errores, columns=["Fila", "Tipo", "Detalle"])

def clasificar_configuracion_pendiente(info, horario, metodo):
    # Para una materia concreta (SUBJ + COURSE), distinguir valores ausentes de pareja ausente.
    horario, metodo = sin_espacios(horario), sin_espacios(metodo)
    if not horario or not metodo: return "Datos incompletos"
    existe_h = horario in info.get("schd", set())
    existe_m = metodo in info.get("insm", set())
    if not existe_h and not existe_m: return "Ambos"
    if not existe_h: return "Tipo de Horario"
    if not existe_m: return "Método Educativo"
    if (horario, metodo) not in info.get("pares", set()): return "Combinación"
    return "Correcto"

def firma_revision_altas(df, pa=None):
    # Firma estable para invalidar resultados finales al cambiar datos o PA.
    columnas = [c for c in df.columns if not c.startswith("_") and c not in ("Sección Original", "Sección Modificada")]
    huella = hashlib.sha256()
    huella.update(df[columnas].fillna("").astype(str).to_csv(index=False).encode("utf-8"))
    if isinstance(pa, pd.DataFrame):
        huella.update(pa.fillna("").astype(str).to_csv(index=False).encode("utf-8"))
    return huella.hexdigest()

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
# Columnas editables y equivalencias con los encabezados reales de ALTAS.
CAMPOS_EXPORTACION_ALTAS = {
    "Año": ["ano"], "Ciclo": ["ciclo"], "Periodo": ["periodo"],
    "Clúster": ["cluster"], "Nivel": ["nivel"], "Requerimiento": ["requerimiento"],
    "Responsable": ["responsable"], "Subject": ["subj", "subject", "area"],
    "Course": ["course", "crse", "nocurso"], "Nombre de la Materia": ["nombremateria", "nombredelamateria"],
    "Campus": ["sede", "campus"], "Parte de Periodo": ["parteperiodo"],
    "Estatus": ["status", "estatus"], "Capacidad": ["capacidad", "cupo"],
    "Sección": ["seccion", "grupo"], "Tipo de Horario": ["tipohorario"],
    "Método Educativo": ["metodoeducativo"], "Socio Integración": ["sociointegracion", "sociodeintegracion"],
    "Modo de Calificar": ["modocalificar", "mododecalificar"], "Sesion": ["sesion"],
    "Comentarios": ["comentarios"], "Fecha Registro": ["fecharegistro"]
}

def cambios_solicitud_excel(df_grupo, archivos_originales):
    origenes = df_grupo["ArchivoOrigen"].dropna().unique().tolist()
    if len(origenes) != 1: raise ValueError("Una solicitud con macros debe tener un solo archivo de origen.")
    origen = origenes[0]
    if origen not in archivos_originales: raise ValueError(f"No se encontró el original: {origen}")
    extension = ".xlsm" if origen.lower().endswith(".xlsm") else ".xlsx"
    wb = openpyxl.load_workbook(io.BytesIO(archivos_originales[origen]), keep_vba=(extension == ".xlsm"))
    hoja = next((h for h in wb.sheetnames if h.strip().upper() == HOJA_ALTAS), None)
    if hoja is None: raise ValueError(f"{origen}: falta la hoja ALTAS")
    ws = wb[hoja]
    encabezados = {normalizar_para_busqueda(c.value): c.column for c in ws[1] if c.value is not None}
    cambios = []
    for _, fila in df_grupo.iterrows():
        numero = fila.get("_FilaExcel")
        if pd.isna(numero): raise ValueError(f"{origen}: falta _FilaExcel")
        for campo, alias in CAMPOS_EXPORTACION_ALTAS.items():
            if campo not in fila.index: continue
            columna = next((encabezados[a] for a in alias if a in encabezados), None)
            if columna is None: continue
            celda = ws.cell(int(numero), columna)
            anterior = limpiar_clave_texto(celda.value)
            nuevo = limpiar_clave_texto(fila[campo])
            comparar = clave_seccion if campo == "Sección" else limpiar_clave_texto
            if comparar(anterior) != comparar(nuevo):
                cambios.append({"Archivo": origen, "Fila": int(numero), "Campo": campo, "Anterior": anterior, "Nuevo": nuevo})
                celda.value = nuevo if nuevo else None
    salida = io.BytesIO()
    wb.save(salida)
    wb.close()
    return salida.getvalue(), extension, cambios

def crear_solicitud_excel(df_grupo, archivos_originales):
    contenido, extension, _ = cambios_solicitud_excel(df_grupo, archivos_originales)
    return contenido, extension

# ============================================================
# 10. GENERAR ZIP SOLICITUDES + CSV_ALTAS
# ============================================================
# ============================================================
# 10. SEPARACION Y EXPORTACION DE SOLICITUDES + CSV_ALTAS
# ============================================================
MODOS_SEPARACION = ["Sin separar", "Por periodo", "Por clúster", "Por responsable",
                    "Por periodo y clúster", "Por periodo y responsable", "Por clúster y responsable",
                    "Por periodo, clúster y responsable"]

CLAVES_SEPARACION = {
    "Sin separar": [], "Por periodo": ["Periodo"], "Por clúster": ["Clúster"],
    "Por responsable": ["Responsable"], "Por periodo y clúster": ["Periodo", "Clúster"],
    "Por periodo y responsable": ["Periodo", "Responsable"],
    "Por clúster y responsable": ["Clúster", "Responsable"],
    "Por periodo, clúster y responsable": ["Periodo", "Clúster", "Responsable"]
}

def grupos_exportacion_altas(df_final, modo="Sin separar"):
    if modo not in CLAVES_SEPARACION: raise ValueError(f"Modo de separación desconocido: {modo}")
    df = df_final.copy()
    obligatorias = ["Periodo", "Responsable", "Clúster", "ArchivoOrigen", "_FilaExcel"]
    faltantes = [c for c in obligatorias if c not in df.columns]
    if faltantes: raise ValueError(f"Faltan columnas para exportar: {faltantes}")
    df["Periodo"] = df["Periodo"].apply(limpiar_clave_texto)
    df["Responsable"] = df["Responsable"].apply(responsable_normalizado)
    df["Clúster"] = df["Clúster"].apply(cluster_oficial)
    for c in ["Periodo", "Responsable", "Clúster"]:
        if df[c].apply(limpiar_clave_texto).eq("").any(): raise ValueError(f"Hay ALTAS sin {c}.")
    # Se mantiene el origen para preservar el orden y evitar mezclar libros distintos.
    claves = ["ArchivoOrigen"] + CLAVES_SEPARACION[modo]
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
    equivalencias = {alias: campo for campo, aliases in CAMPOS_EXPORTACION_ALTAS.items() for alias in aliases}
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
        # Entregar .xlsm si el nombre es incorrecto O hay datos corregidos.
        for origen, grupo in df.groupby("ArchivoOrigen", sort=False):
            if not origen.lower().endswith(".xlsm"): continue
            esperado = nombre_solicitud(grupo, ".xlsm")
            contenido, _, cambios = cambios_solicitud_excel(grupo, archivos_originales)
            nombre_incorrecto = normalizar_para_busqueda(origen) != normalizar_para_busqueda(esperado)
            if not nombre_incorrecto and not cambios: continue
            carpeta = "MACROS_CORREGIDAS" if cambios else "MACROS_RENOMBRADAS"
            nombre = esperado[:-5] + " Corregidos.xlsm" if cambios else esperado
            z.writestr(ruta_zip_unica(carpeta, nombre, rutas), contenido)
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
    "df_manual_fijo": None, "manual_candidatos": [], "manual_busqueda_realizada": False,
    "t1_revision_final_firma": None, "t1_revision_final_df": None,
    "t1_revision_final_errores": None, "t1_historial_deshacer": [],
    "t1_resumen_cambios": None
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
                  "t1_firma_archivos", "t1_catalogo", "t1_version_editor", "t1_modo_separacion", "t1_modo_generado", "t1_altas_editadas", "t1_edicion_version", "t1_masivo_campo", "t1_masivo_valor", "t1_masivo_aplicar", "t1_masivo_seleccion", "t1_historial", "t1_revision_final", "t1_revision_firma", "t1_filtro_cluster", "t1_filtro_subj", "t1_filtro_course", "t1_filtro_seccion", "t1_buscar_materia", "t1_columnas_copia", "t1_consulta_config_pendiente", "t1_aviso_revalidacion", "t1_vista_pendientes", "t1_consulta_archivo", "t1_consulta_periodo", "t1_consulta_cluster", "t1_consulta_responsable", "t1_consulta_materia"]
        claves += [k for k in st.session_state if k.startswith(("nom_t1_", "met_t1_", "manual_t1_", "t1_editor_", "t1_pendientes_editor_"))]
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
                  "confirmar_errores_t1", "t1_confirmar_altas", "t1_altas_editadas", "t1_masivo_seleccion", "t1_historial", "t1_revision_final", "t1_revision_firma", "t1_aviso_revalidacion"]:
            st.session_state.pop(k, None)
        st.session_state.t1_firma_archivos = firma_actual
        st.session_state.t1_version_editor = st.session_state.get("t1_version_editor", 0) + 1
        st.session_state.t1_edicion_version = st.session_state.get("t1_edicion_version", 0) + 1

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
            falta = clasificar_configuracion_pendiente(info, horario, metodo)
            r["Comentario Horario"] = f"{horario} + {metodo}: {falta} pendiente para {subj} {crse}."
            r["Comentario Método"] = f"Acción requerida: {falta}."
            r["Pendientes Alta"].append((falta, horario, metodo, "Configuración no registrada"))
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
        columnas = ["Periodo", "Clúster", "Responsable", "SUBJ", "COURSE", "Nombre de la Materia", "Tipo de Horario", "Método Educativo", "Modo de Calificar", "¿Qué falta?", "Archivo"]
        registros = []
        for _, r in auditoria.iterrows():
            fila = df.iloc[int(r["idx"])]
            for tipo, horario, metodo, detalle in r["Pendientes Alta"]:
                registros.append({"Periodo": limpiar_clave_texto(fila.get("Periodo")), "Clúster": cluster_oficial(fila.get("Clúster")), "Responsable": responsable_normalizado(fila.get("Responsable")), "SUBJ": sin_espacios(fila.get("Subject")), "COURSE": sin_espacios(fila.get("Course")), "Nombre de la Materia": limpiar_clave_texto(fila.get("Nombre de la Materia")), "Tipo de Horario": horario or sin_espacios(fila.get("Tipo de Horario")), "Método Educativo": metodo or sin_espacios(fila.get("Método Educativo")), "Modo de Calificar": sin_espacios(fila.get("Modo de Calificar")), "¿Qué falta?": tipo, "Archivo": fila.get("ArchivoOrigen", "")})
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

                st.session_state.original_files_bytes = originales
                st.session_state.raw_altas = total
                st.session_state.t1_altas_editadas = total.copy()
                st.session_state.t1_edicion_version = st.session_state.get("t1_edicion_version", 0) + 1
                st.session_state.res_auditoria = auditoria
                st.session_state.cat_pa_cache = pa
                st.session_state.t1_pendientes_alta = pendientes
                st.session_state.advertencias_pa = avisos
                st.session_state.errores_restricciones = auditar_restricciones_internas(total)
                st.session_state.df_secciones_corregidas = None
                st.session_state.errores_secciones = None
                st.session_state.t1_revision_final = None
                st.session_state.t1_revision_firma = None
                st.session_state.t1_historial = []
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

        # 7.0 CONSULTA GENERAL: solo lectura, independiente de los filtros de edición.
        with st.expander("📊 Consultar todos los registros ALTAS", expanded=False):
            consulta = st.session_state.get("t1_altas_editadas")
            if not isinstance(consulta, pd.DataFrame):
                consulta = st.session_state.raw_altas
            consulta = consulta.copy()
            columnas_consulta = [
                (nombre, MAPA_COLUMNAS.get(normalizar_para_busqueda(nombre), nombre))
                for nombre in COLUMNAS_ALTAS_REALES
            ]
            columnas_consulta = [(nombre, real) for nombre, real in columnas_consulta if real in consulta.columns]
            c1, c2, c3 = st.columns(3)
            c1.metric("Registros", len(consulta))
            c2.metric("Archivos", consulta["ArchivoOrigen"].nunique())
            c3.metric("Periodos", consulta["Periodo"].apply(limpiar_clave_texto).nunique())

            f1, f2, f3, f4 = st.columns(4)
            def opciones_consulta(columna):
                return ["Todos"] + sorted(v for v in consulta[columna].apply(limpiar_clave_texto).unique() if v)
            archivo_consulta = f1.selectbox("Archivo", opciones_consulta("ArchivoOrigen"), key="t1_consulta_archivo")
            periodo_consulta = f2.selectbox("Periodo", opciones_consulta("Periodo"), key="t1_consulta_periodo")
            cluster_consulta = f3.selectbox("Clúster", opciones_consulta("Clúster"), key="t1_consulta_cluster")
            responsable_consulta = f4.selectbox("Responsable", opciones_consulta("Responsable"), key="t1_consulta_responsable")
            materia_consulta = st.text_input("Buscar materia, SUBJ o COURSE", key="t1_consulta_materia")
            mascara_consulta = pd.Series(True, index=consulta.index)
            for columna, valor in [
                ("ArchivoOrigen", archivo_consulta), ("Periodo", periodo_consulta),
                ("Clúster", cluster_consulta), ("Responsable", responsable_consulta)
            ]:
                if valor != "Todos":
                    mascara_consulta &= consulta[columna].apply(limpiar_clave_texto).eq(valor)
            if materia_consulta.strip():
                termino = normalizar_para_busqueda(materia_consulta)
                mascara_consulta &= (
                    consulta["Nombre de la Materia"].apply(normalizar_para_busqueda).str.contains(termino, regex=False)
                    | consulta["Subject"].apply(normalizar_para_busqueda).str.contains(termino, regex=False)
                    | consulta["Course"].apply(normalizar_para_busqueda).str.contains(termino, regex=False)
                )
            consulta_filtrada = consulta.loc[mascara_consulta]
            vista_consulta = consulta_filtrada[[real for _, real in columnas_consulta]].copy().fillna("")
            vista_consulta.columns = [nombre for nombre, _ in columnas_consulta]
            st.caption(f"Mostrando {len(vista_consulta)} de {len(consulta)} registros. Tabla de solo lectura.")
            st.dataframe(vista_consulta, hide_index=True, use_container_width=True, height=360)
            if isinstance(st.session_state.get("t1_revision_final"), pd.DataFrame):
                st.caption("Las secciones definitivas, después de finalizar la revisión, se consultan en el apartado de revisión final.")

        with st.expander("✏️ Corrección manual de solicitudes ALTAS", expanded=False):
            # 7.0 EDICION MANUAL DE ALTAS: conservar origen y fila, sin tocar columnas tecnicas.
            st.caption("Edita registros o aplica cambios masivos. Las secciones se asignarán al finalizar la revisión.")
            if st.button("↶ Deshacer última corrección", key="t1_deshacer", disabled=not st.session_state.get("t1_historial")):
                anterior = st.session_state.t1_historial.pop()
                st.session_state.t1_altas_editadas = anterior
                st.session_state.t1_revision_final = None
                st.session_state.t1_revision_firma = None
                st.session_state.zip_solicitudes_csv_altas = None
                st.session_state.t1_edicion_version = st.session_state.get("t1_edicion_version", 0) + 1
                st.rerun()
            campos_editables = [c for c in COLUMNAS_ALTAS_REALES if c not in ("AÑO", "CICLO", "FECHA_REGISTRO")]
            campos_editables = list(dict.fromkeys(MAPA_COLUMNAS.get(normalizar_para_busqueda(c), c) for c in campos_editables))
            actual = st.session_state.get("t1_altas_editadas")
            if not isinstance(actual, pd.DataFrame):
                actual = st.session_state.raw_altas.copy()
                st.session_state.t1_altas_editadas = actual.copy()
            visibles = [c for c in campos_editables if c in actual.columns]
            filtros = st.columns(3)
            filtro_archivo = filtros[0].selectbox("Archivo", ["Todos"] + sorted(actual["ArchivoOrigen"].dropna().unique().tolist()), key="t1_filtrar_archivo")
            filtro_periodo = filtros[1].selectbox("Periodo", ["Todos"] + sorted(actual["Periodo"].apply(limpiar_clave_texto).unique().tolist()), key="t1_filtrar_periodo")
            filtro_responsable = filtros[2].selectbox("Responsable", ["Todos"] + sorted(actual["Responsable"].apply(limpiar_clave_texto).unique().tolist()), key="t1_filtrar_responsable")
            f2 = st.columns(4)
            filtro_cluster = f2[0].selectbox("Clúster", ["Todos"] + sorted(actual["Clúster"].apply(limpiar_clave_texto).unique().tolist()), key="t1_filtro_cluster")
            filtro_subj = f2[1].selectbox("SUBJ", ["Todos"] + sorted(actual["Subject"].apply(limpiar_clave_texto).unique().tolist()), key="t1_filtro_subj")
            filtro_course = f2[2].selectbox("COURSE", ["Todos"] + sorted(actual["Course"].apply(limpiar_clave_texto).unique().tolist()), key="t1_filtro_course")
            filtro_seccion = f2[3].selectbox("Sección capturada", ["Todos"] + sorted(actual["Sección"].apply(clave_seccion).unique().tolist()), key="t1_filtro_seccion")
            buscar_materia = st.text_input("Buscar nombre de materia", key="t1_buscar_materia")
            mascara = pd.Series(True, index=actual.index)
            if filtro_archivo != "Todos": mascara &= actual["ArchivoOrigen"].eq(filtro_archivo)
            if filtro_periodo != "Todos": mascara &= actual["Periodo"].apply(limpiar_clave_texto).eq(filtro_periodo)
            if filtro_responsable != "Todos": mascara &= actual["Responsable"].apply(limpiar_clave_texto).eq(filtro_responsable)
            if filtro_cluster != "Todos": mascara &= actual["Clúster"].apply(limpiar_clave_texto).eq(filtro_cluster)
            if filtro_subj != "Todos": mascara &= actual["Subject"].apply(limpiar_clave_texto).eq(filtro_subj)
            if filtro_course != "Todos": mascara &= actual["Course"].apply(limpiar_clave_texto).eq(filtro_course)
            if filtro_seccion != "Todos": mascara &= actual["Sección"].apply(clave_seccion).eq(filtro_seccion)
            if buscar_materia.strip(): mascara &= actual["Nombre de la Materia"].apply(normalizar_para_busqueda).str.contains(normalizar_para_busqueda(buscar_materia), regex=False)
            indices_visibles = actual.index[mascara].tolist()
            st.caption(f"{len(indices_visibles)} registros filtrados")
            tabla = actual.loc[indices_visibles, visibles].copy().fillna("")
            tabla.insert(0, "Seleccionar", False)
            tabla.insert(1, "ArchivoOrigen", actual.loc[indices_visibles, "ArchivoOrigen"])
            tabla.insert(2, "FilaExcel", actual.loc[indices_visibles, "_FilaExcel"])
            with st.form("t1_edicion_form"):
                editado_manual = st.data_editor(tabla, hide_index=True, num_rows="fixed", use_container_width=True,
                    disabled=["ArchivoOrigen", "FilaExcel"], key=f"t1_edicion_tabla_{st.session_state.get('t1_edicion_version', 0)}")
                guardar_edicion = st.form_submit_button("💾 Guardar correcciones de la tabla", use_container_width=True)
            if guardar_edicion:
                nuevo = actual.copy()
                for posicion, indice in enumerate(indices_visibles):
                    for campo in visibles:
                        nuevo.at[indice, campo] = limpiar_clave_texto(editado_manual.iloc[posicion][campo])
                    if "Sección" in visibles and clave_seccion(nuevo.at[indice, "Sección"]) != clave_seccion(actual.at[indice, "Sección"]):
                        nuevo.at[indice, "_SeccionManual"] = True
                cambiadas = [i for i in indices_visibles if any(limpiar_clave_texto(nuevo.at[i, c]) != limpiar_clave_texto(actual.at[i, c]) for c in visibles)]
                for i in cambiadas:
                    st.session_state.res_auditoria.loc[st.session_state.res_auditoria["idx"] == i, "Luz Verde"] = False
                if cambiadas:
                    st.session_state.setdefault("t1_historial", []).append(actual.copy())
                    st.session_state.t1_historial = st.session_state.t1_historial[-10:]
                    st.session_state.t1_revision_final = None
                    st.session_state.t1_revision_firma = None
                st.session_state.t1_altas_editadas = nuevo
                st.session_state.t1_masivo_seleccion = [indices_visibles[i] for i, v in enumerate(editado_manual["Seleccionar"]) if v]
                st.session_state.t1_edicion_version = st.session_state.get("t1_edicion_version", 0) + 1
                st.session_state.zip_solicitudes_csv_altas = None
                st.rerun()
            st.caption("Para cambios masivos, selecciona renglones en la tabla y presiona Guardar correcciones; o aplica a todos los filtrados.")
            with st.form("t1_edicion_masiva"):
                c1, c2, c3 = st.columns(3)
                campo_masivo = c1.selectbox("Campo a modificar", visibles, key="t1_masivo_campo")
                valor_masivo = c2.text_input("Nuevo valor", key="t1_masivo_valor")
                alcance_masivo = c3.radio("Aplicar a", ["Filtrados", "Seleccionados guardados"], key="t1_masivo_aplicar")
                aplicar_masivo = st.form_submit_button("Aplicar cambio masivo")
            if aplicar_masivo:
                indices_cambio = indices_visibles if alcance_masivo == "Filtrados" else [i for i in st.session_state.get("t1_masivo_seleccion", []) if i in indices_visibles]
                if not indices_cambio: st.warning("No hay registros seleccionados para modificar.")
                elif not limpiar_clave_texto(valor_masivo): st.warning("Escribe el valor nuevo. No se permiten cambios masivos a vacío.")
                else:
                    nuevo = actual.copy()
                    for i in indices_cambio:
                        nuevo.at[i, campo_masivo] = limpiar_clave_texto(valor_masivo)
                        if campo_masivo == "Sección": nuevo.at[i, "_SeccionManual"] = True
                    for i in indices_cambio:
                        st.session_state.res_auditoria.loc[st.session_state.res_auditoria["idx"] == i, "Luz Verde"] = False
                    st.session_state.setdefault("t1_historial", []).append(actual.copy())
                    st.session_state.t1_historial = st.session_state.t1_historial[-10:]
                    st.session_state.t1_revision_final = None
                    st.session_state.t1_revision_firma = None
                    st.session_state.t1_altas_editadas = nuevo
                    st.session_state.t1_edicion_version = st.session_state.get("t1_edicion_version", 0) + 1
                    st.session_state.zip_solicitudes_csv_altas = None
                    st.rerun()
        # 4. TABLA PARA TEAMS Y DESCARGA CON BORDES (independiente de edición).
        with st.expander("📋 Copiar y descargar registros para revisión", expanded=False):
            st.caption("Usa los filtros del apartado de corrección para elegir los registros. No modifica ALTAS.")
            columnas_copia = [c for c in COLUMNAS_ALTAS_REALES if MAPA_COLUMNAS.get(normalizar_para_busqueda(c), c) in actual.columns]
            columnas_elegidas = st.multiselect("Columnas para compartir", columnas_copia, default=[c for c in ["PERIODO", "CLUSTER", "RESPONSABLE", "SUBJ", "COURSE", "NOMBRE_MATERIA", "SECCION", "TIPO_HORARIO", "METODO_EDUCATIVO", "MODO_CALIFICAR"] if c in columnas_copia], key="t1_columnas_copia")
            if columnas_elegidas:
                campos_copia = [MAPA_COLUMNAS.get(normalizar_para_busqueda(c), c) for c in columnas_elegidas]
                para_copiar = actual.loc[indices_visibles, campos_copia].copy().fillna("")
                para_copiar.columns = columnas_elegidas
                st.dataframe(para_copiar, hide_index=True, use_container_width=True)
                st.code(para_copiar.to_csv(index=False, sep="\t", lineterminator="\n"), language="text")
                excel_copia = io.BytesIO()
                with pd.ExcelWriter(excel_copia, engine="openpyxl") as writer:
                    para_copiar.to_excel(writer, index=False, sheet_name="Revision")
                    ws = writer.sheets["Revision"]
                    ws.freeze_panes = "A2"
                    ws.auto_filter.ref = ws.dimensions
                    for row in ws.iter_rows():
                        for celda in row:
                            celda.border = openpyxl.styles.Border(left=openpyxl.styles.Side(style="thin", color="C9C9C9"), right=openpyxl.styles.Side(style="thin", color="C9C9C9"), top=openpyxl.styles.Side(style="thin", color="C9C9C9"), bottom=openpyxl.styles.Side(style="thin", color="C9C9C9"))
                            celda.alignment = Alignment(vertical="center", wrap_text=True)
                            if celda.row == 1:
                                celda.fill = PatternFill(fill_type="solid", fgColor="7030A0")
                                celda.font = Font(name="Calibri", bold=True, color="FFFFFF")
                    for columna in ws.columns:
                        ancho = max((len(str(c.value or "")) for c in columna), default=10)
                        ws.column_dimensions[columna[0].column_letter].width = min(max(ancho + 3, 12), 45)
                st.download_button("📥 Descargar Excel de registros filtrados (.xlsx)", excel_copia.getvalue(), file_name="Registros_Filtrados_Revision.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="t1_descargar_filtrados", use_container_width=True)
            else: st.info("Selecciona al menos una columna.")
        # Auditoría derivada exclusivamente de los valores guardados en ALTAS.
        # Las sugerencias antiguas nunca se reaplican sobre correcciones nuevas.
        catalogo_actual = cargar_catalogo_avanzado()
        auditoria_actualizada = []
        for i, fila in actual.iterrows():
            revision = validar_materia(fila, catalogo_actual)
            revision.update({"idx": i, "Archivo": fila["ArchivoOrigen"]})
            auditoria_actualizada.append(revision)
        auditoria = pd.DataFrame(auditoria_actualizada)
        st.session_state.res_auditoria = auditoria
        pendientes = obtener_pendientes_altas(actual, auditoria)
        st.session_state.t1_pendientes_alta = pendientes
        errores_reglas = auditar_restricciones_internas(actual)
        st.session_state.errores_restricciones = errores_reglas
        # Reevaluar periodos PA tras ediciones de periodo.
        pa_actual = st.session_state.get("cat_pa_cache")
        avisos_pa = [a for a in (st.session_state.get("advertencias_pa") or []) if a.get("Tipo") != "PA sin cobertura"]
        periodos_actuales = {normalizar_para_cruce(p) for p in actual["Periodo"] if limpiar_clave_texto(p)}
        periodos_pa = {normalizar_para_cruce(p) for p in pa_actual["periodo"]} if isinstance(pa_actual, pd.DataFrame) else set()
        for p in sorted(periodos_actuales - periodos_pa):
            avisos_pa.append({"Tipo": "PA sin cobertura", "Periodo": p, "Detalle": f"Sin PA del periodo {p}; no se conocen todas las secciones ocupadas."})
        st.session_state.advertencias_pa = avisos_pa

        # 7.1 CONFIGURACIONES PENDIENTES: revisar, corregir y revalidar sin recargar ALTAS.
        st.divider()
        if not pendientes.empty:
            st.subheader("📚 Configuraciones pendientes de alta")
            resumen_altas = pendientes.drop_duplicates(subset=["SUBJ", "COURSE", "Tipo de Horario", "Método Educativo", "Modo de Calificar", "¿Qué falta?"])
            resumen_altas = resumen_altas.sort_values(["SUBJ", "COURSE", "¿Qué falta?"])
            st.warning(f"Se encontraron {len(resumen_altas)} configuraciones por revisar.")
            st.dataframe(resumen_altas["¿Qué falta?"].value_counts().rename_axis("Pendiente").reset_index(name="Cantidad"), hide_index=True, use_container_width=True)
            with st.expander("🔎 Ver y corregir configuraciones pendientes", expanded=True):
                pendientes_idx = sorted({int(r["idx"]) for _, r in auditoria.iterrows() if r["Pendientes Alta"]})
                # Cada registro de ALTAS se muestra una sola vez, aunque tenga varias incidencias.
                columnas_p = ["Archivo", "Fila Excel", "Periodo", "SUBJ", "COURSE", "Nombre de la Materia",
                              "Tipo de Horario", "Método Educativo", "Modo de Calificar", "¿Qué falta?"]
                filas_p = []
                for idx in pendientes_idx:
                    f = actual.loc[idx]
                    rev = auditoria.loc[auditoria["idx"] == idx].iloc[0]
                    filas_p.append({
                        "Archivo": f["ArchivoOrigen"], "Fila Excel": int(f["_FilaExcel"]),
                        "Periodo": limpiar_clave_texto(f["Periodo"]),
                        "SUBJ": sin_espacios(f["Subject"]), "COURSE": sin_espacios(f["Course"]),
                        "Nombre de la Materia": limpiar_clave_texto(f["Nombre de la Materia"]),
                        "Tipo de Horario": sin_espacios(f["Tipo de Horario"]),
                        "Método Educativo": sin_espacios(f["Método Educativo"]),
                        "Modo de Calificar": sin_espacios(f["Modo de Calificar"]),
                        "¿Qué falta?": ", ".join(sorted({str(x[0]) for x in rev["Pendientes Alta"]}))
                    })
                tabla_todos = pd.DataFrame(filas_p, index=pendientes_idx, columns=columnas_p)
                vista_p = st.radio("¿Cómo deseas revisar las configuraciones?",
                    ["Materias únicas", "Todos los registros"], horizontal=True, key="t1_vista_pendientes")
                columnas_clave = ["SUBJ", "COURSE", "Tipo de Horario", "Método Educativo",
                                  "Modo de Calificar", "¿Qué falta?"]
                if vista_p == "Materias únicas":
                    # No unir configuraciones diferentes de una misma materia.
                    # Cada fila representa una configuración exacta y conserva sus índices reales.
                    grupos_p = list(tabla_todos.groupby(columnas_clave, sort=False, dropna=False))
                    filas_unicas, indices_por_fila = [], []
                    for _, grupo_p in grupos_p:
                        fila_p = grupo_p.iloc[0].copy()
                        fila_p["Registros afectados"] = len(grupo_p)
                        filas_unicas.append(fila_p)
                        indices_por_fila.append(grupo_p.index.tolist())
                    columnas_unicas = ["SUBJ", "COURSE", "Nombre de la Materia", "Tipo de Horario",
                                       "Método Educativo", "Modo de Calificar", "¿Qué falta?", "Registros afectados"]
                    tabla_editor_p = pd.DataFrame(filas_unicas).reset_index(drop=True)[columnas_unicas]
                    bloqueadas_p = ["SUBJ", "COURSE", "Nombre de la Materia", "¿Qué falta?", "Registros afectados"]
                    st.caption(f"{len(tabla_editor_p)} configuraciones distintas; {len(tabla_todos)} registros afectados. "
                               "Un cambio en esta vista se aplica a todos los registros de esa misma configuración.")
                else:
                    tabla_editor_p = tabla_todos.reset_index(drop=True)
                    indices_por_fila = [[idx] for idx in pendientes_idx]
                    bloqueadas_p = ["Archivo", "Fila Excel", "Periodo", "SUBJ", "COURSE",
                                    "Nombre de la Materia", "¿Qué falta?"]
                    st.caption(f"{len(tabla_editor_p)} registros afectados. Cada corrección modifica solamente su fila.")
                with st.form("t1_form_pendientes"):
                    ed_p = st.data_editor(tabla_editor_p, hide_index=True, num_rows="fixed",
                        use_container_width=True, disabled=bloqueadas_p,
                        key=f"t1_pendientes_editor_{vista_p}_{st.session_state.get('t1_edicion_version', 0)}")
                    guardar_p = st.form_submit_button("💾 Guardar y volver a validar contra el catálogo",
                                                      use_container_width=True)
                if guardar_p:
                    nuevo = actual.copy()
                    afectados = set()
                    campos_config = ["Tipo de Horario", "Método Educativo", "Modo de Calificar"]
                    for posicion, indices_grupo in enumerate(indices_por_fila):
                        for campo in campos_config:
                            anterior_grupo = limpiar_clave_texto(tabla_editor_p.iloc[posicion][campo])
                            valor_nuevo = limpiar_clave_texto(ed_p.iloc[posicion][campo])
                            if valor_nuevo != anterior_grupo:
                                for idx in indices_grupo:
                                    if limpiar_clave_texto(nuevo.at[idx, campo]) != valor_nuevo:
                                        nuevo.at[idx, campo] = valor_nuevo
                                        afectados.add(idx)
                    if afectados:
                        st.session_state.setdefault("t1_historial", []).append(actual.copy())
                        st.session_state.t1_historial = st.session_state.t1_historial[-10:]
                        st.session_state.t1_altas_editadas = nuevo
                        for idx in afectados:
                            st.session_state.res_auditoria.loc[
                                st.session_state.res_auditoria["idx"] == idx, "Luz Verde"] = False
                        st.session_state.t1_revision_final = None
                        st.session_state.t1_revision_firma = None
                        st.session_state.df_secciones_corregidas = None
                        st.session_state.zip_solicitudes_csv_altas = None
                        st.session_state.t1_confirmar_altas = False
                        st.session_state.t1_edicion_version = st.session_state.get("t1_edicion_version", 0) + 1
                        st.session_state.t1_aviso_revalidacion = (
                            f"Se actualizaron {len(afectados)} registros. "
                            "Se volvieron a comprobar las configuraciones contra el Catálogo Avanzado.")
                        st.rerun()
                    else:
                        st.info("No se detectaron cambios. Las configuraciones pendientes permanecen igual.")
                if pendientes_idx:
                    seleccion_p = st.selectbox("Consultar configuraciones válidas de una materia", pendientes_idx,
                        format_func=lambda i: f'{actual.at[i, "Subject"]} {actual.at[i, "Course"]} — '
                                              f'{actual.at[i, "Nombre de la Materia"]} '
                                              f'(fila {actual.at[i, "_FilaExcel"]})',
                        key="t1_consulta_config_pendiente")
                    fila_sel = actual.loc[seleccion_p]
                    info_sel = catalogo_actual.get((sin_espacios(fila_sel["Subject"]), sin_espacios(fila_sel["Course"])))
                    if info_sel:
                        parejas = sorted(info_sel["pares"])
                        if parejas:
                            st.dataframe(pd.DataFrame(parejas, columns=["Tipo de Horario", "Método Educativo"]),
                                         hide_index=True, use_container_width=True)
                        else:
                            st.info("Esta materia no tiene parejas de horario y método registradas.")
                        st.caption("Modos de calificar registrados: " +
                                   (", ".join(sorted(info_sel["gmod"])) or "Ninguno"))
                    else:
                        st.warning("SUBJ + COURSE no existe en el Catálogo Avanzado; "
                                   "primero revisa la materia en la Mesa de Control.")
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
                file_name="Configuraciones_Pendientes_Alta.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key="descargar_altas_t1")
            st.checkbox("Ya revisé las configuraciones pendientes y deseo continuar con advertencia", key="t1_confirmar_altas")
            st.caption("Al corregir y guardar, los pendientes se recalculan. Si una configuración es nueva, puedes conservarla y continuar con advertencia.")
        else:
            st.success("No hay configuraciones pendientes de alta.")
        if st.session_state.get("t1_aviso_revalidacion"):
            st.success(st.session_state.pop("t1_aviso_revalidacion"))

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
                        # Aplicar las aprobaciones directamente a la fuente única.
                        # No guardar sugerencias para sobrescribirlas después.
                        nuevo = st.session_state.t1_altas_editadas.copy()
                        cambios_aprobados = set()
                        equivalencias_aprobadas = {
                            "Subject": "Subj Sugerido", "Course": "Crse Sugerido",
                            "Tipo de Horario": "Horario Sugerido",
                            "Método Educativo": "Método Sugerido",
                            "Modo de Calificar": "Modo de Calificar Sugerido"
                        }
                        for i in sub.index:
                            if not bool(editado.at[i, "Luz Verde"]):
                                continue
                            idx_real = int(auditoria.at[i, "idx"])
                            for destino, origen in equivalencias_aprobadas.items():
                                valor = limpiar_clave_texto(editado.at[i, origen])
                                if valor and valor != limpiar_clave_texto(nuevo.at[idx_real, destino]):
                                    nuevo.at[idx_real, destino] = valor
                                    cambios_aprobados.add(idx_real)
                        if cambios_aprobados:
                            st.session_state.setdefault("t1_historial", []).append(st.session_state.t1_altas_editadas.copy())
                            st.session_state.t1_historial = st.session_state.t1_historial[-10:]
                            st.session_state.t1_altas_editadas = nuevo
                            st.session_state.t1_revision_final = None
                            st.session_state.t1_revision_firma = None
                            st.session_state.df_secciones_corregidas = None
                            st.session_state.zip_solicitudes_csv_altas = None
                            st.session_state.t1_confirmar_altas = False
                            st.session_state.t1_edicion_version = st.session_state.get("t1_edicion_version", 0) + 1
                            st.session_state.t1_aviso_revalidacion = f"Se guardaron {len(cambios_aprobados)} registros de Mesa de Control."
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
            # Fuente única: todos los editores guardan aquí.
            # No aplicar sugerencias pendientes ni aprobaciones antiguas.
            return st.session_state.t1_altas_editadas.copy()

        # 5. REVISION FINAL: calcular secciones SOLO al pulsar este boton.
        st.divider()
        st.subheader("✅ Finalizar revisión y asignar secciones")
        base_corregida = construir_altas_corregidas()
        firma_actual_revision = firma_revision_altas(base_corregida, st.session_state.get("cat_pa_cache"))
        if st.session_state.get("t1_revision_firma") != firma_actual_revision:
            st.session_state.t1_revision_final = None
            st.session_state.df_secciones_corregidas = None
            st.session_state.zip_solicitudes_csv_altas = None
        originales_df = st.session_state.raw_altas
        comparables = [c for c in COLUMNAS_ALTAS if c in base_corregida.columns and c in originales_df.columns]
        cambios_detalle = []
        for i in base_corregida.index:
            for c in comparables:
                antes, despues = limpiar_clave_texto(originales_df.at[i, c]), limpiar_clave_texto(base_corregida.at[i, c])
                if antes != despues: cambios_detalle.append({"Archivo": base_corregida.at[i, "ArchivoOrigen"], "Fila Excel": int(base_corregida.at[i, "_FilaExcel"]), "Campo": c, "Anterior": antes, "Nuevo": despues})
        # Vista de la fuente que se exportará, antes de asignar secciones.
        with st.expander("🔍 Verificar valores guardados antes de exportar", expanded=False):
            columnas_verificacion = ["ArchivoOrigen", "_FilaExcel", "Periodo", "Subject", "Course",
                                     "Tipo de Horario", "Método Educativo", "Modo de Calificar"]
            st.dataframe(base_corregida[[c for c in columnas_verificacion if c in base_corregida.columns]],
                         hide_index=True, use_container_width=True)
        with st.expander(f"📋 Resumen de modificaciones ({len(cambios_detalle)} campos)", expanded=False):
            if cambios_detalle: st.dataframe(pd.DataFrame(cambios_detalle), hide_index=True, use_container_width=True)
            else: st.info("No se modificaron campos antes de asignar secciones.")
        if st.button("✅ Finalizar revisión y asignar secciones", type="primary", key="t1_finalizar_revision", use_container_width=True):
            try:
                if not pendientes.empty and not st.session_state.get("t1_confirmar_altas"):
                    raise ValueError("Confirma que revisaste las configuraciones pendientes.")
                if hay_otras_advertencias and not st.session_state.get("confirmar_errores_t1"):
                    raise ValueError("Confirma las advertencias pendientes.")
                with st.spinner("Asignando secciones definitivas..."):
                    df_sec, errores_sec = asignar_secciones_globales(base_corregida, st.session_state.get("cat_pa_cache"))
                if not errores_sec.empty:
                    st.session_state.errores_secciones = errores_sec
                    raise ValueError(f"Hay {len(errores_sec)} secciones que no pudieron asignarse.")
                st.session_state.t1_revision_final = df_sec
                st.session_state.df_secciones_corregidas = df_sec
                st.session_state.t1_revision_firma = firma_actual_revision
                st.session_state.errores_secciones = errores_sec
                st.session_state.zip_solicitudes_csv_altas = None
                st.success(f"Revisión finalizada: {len(df_sec)} registros, {int(df_sec['Sección Modificada'].sum())} secciones ajustadas.")
            except Exception as e: st.error(f"No se pudo finalizar: {e}")
        revision_lista = isinstance(st.session_state.get("t1_revision_final"), pd.DataFrame) and st.session_state.get("t1_revision_firma") == firma_actual_revision
        if revision_lista:
            with st.expander("🔢 Ver secciones definitivas", expanded=False):
                cols_sec = ["ArchivoOrigen", "Periodo", "Responsable", "Subject", "Course", "Sección Original", "Sección", "Sección Modificada"]
                st.dataframe(st.session_state.t1_revision_final[[c for c in cols_sec if c in st.session_state.t1_revision_final.columns]], hide_index=True, use_container_width=True)
        elif isinstance(st.session_state.get("errores_secciones"), pd.DataFrame) and not st.session_state.errores_secciones.empty:
            st.dataframe(st.session_state.errores_secciones, hide_index=True, use_container_width=True)

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

        # 6. GENERAR ARCHIVOS: utilizar exclusivamente la revision final vigente.
        st.divider()
        st.subheader("📁 Generación de CSV_ALTAS y Excel")
        modo_separacion = st.selectbox("¿Cómo deseas separar los archivos?", MODOS_SEPARACION, index=7 if len(MODOS_SEPARACION) > 7 else 0, key="t1_modo_separacion")
        st.caption("La misma separación aplica a CSV y Excel normales. Los originales .xlsm se devuelven solo si requieren corrección o renombrado.")
        if st.session_state.get("t1_modo_generado") != modo_separacion: st.session_state.zip_solicitudes_csv_altas = None
        if revision_lista:
            vista = st.session_state.t1_revision_final
            previstos = []
            for _, grupo in grupos_exportacion_altas(vista, modo_separacion):
                previstos.append({"Carpeta": "CSV_ALTAS", "Archivo": nombre_csv_grupo_altas(grupo), "Registros": len(grupo)})
                if modo_separacion != "Sin separar": previstos.append({"Carpeta": "ALTAS_SEPARADAS", "Archivo": nombre_solicitud(grupo, ".xlsx"), "Registros": len(grupo)})
            for origen, grupo in vista.groupby("ArchivoOrigen", sort=False):
                if not origen.lower().endswith(".xlsm"): continue
                esperado = nombre_solicitud(grupo, ".xlsm")
                # Verificar cambios sin generar el libro completo para la vista previa.
                original = st.session_state.raw_altas
                filas_origen = original[original["ArchivoOrigen"] == origen].set_index("_FilaExcel")
                cambiado = any(any(limpiar_clave_texto(f.get(c)) != limpiar_clave_texto(filas_origen.loc[int(f["_FilaExcel"]), c]) for c in CAMPOS_EXPORTACION_ALTAS if c in grupo.columns and c in filas_origen.columns) for _, f in grupo.iterrows())
                nombre_mal = normalizar_para_busqueda(origen) != normalizar_para_busqueda(esperado)
                if cambiado or nombre_mal: previstos.append({"Carpeta": "MACROS_CORREGIDAS" if cambiado else "MACROS_RENOMBRADAS", "Archivo": esperado, "Registros": len(grupo)})
            st.caption(f"Vista previa del ZIP: {len(previstos)} archivos previstos")
            with st.expander("📦 Ver contenido previsto del ZIP", expanded=False):
                st.dataframe(pd.DataFrame(previstos), hide_index=True, use_container_width=True)
        else: st.info("Primero finaliza la revisión para calcular las secciones y preparar los archivos.")
        if st.button("💾 Generar SOLICITUDES y CSV_ALTAS", type="primary", use_container_width=True, key="generar_salidas_t1", disabled=not revision_lista):
            try:
                with st.spinner("Generando archivos..."):
                    zip_final = generar_zip_altas(st.session_state.t1_revision_final, st.session_state.original_files_bytes, preparar_csv_banner, modo_separacion=modo_separacion)
                st.session_state.zip_solicitudes_csv_altas = zip_final
                st.session_state.t1_modo_generado = modo_separacion
                st.success("Archivos generados correctamente.")
            except Exception as e:
                st.session_state.zip_solicitudes_csv_altas = None
                st.error(f"No se pudieron generar: {e}")
        if st.session_state.get("zip_solicitudes_csv_altas"):
            st.download_button("📥 Descargar SOLICITUDES + CSV_ALTAS (.ZIP)", st.session_state.zip_solicitudes_csv_altas, file_name="ALTAS_PROCESADAS.zip", mime="application/zip", type="primary", use_container_width=True)

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
    st.caption("Un solo fragmento corregido: sirve para reenviar únicamente los errores a Banner y para reconstruir el CSV completo.")

    if st.button("🔄 Limpiar / Recomenzar", key="limpiar_t2"):
        claves = ["t2_base_ext", "t2_errores_ext", "t2_modo", "t2_version", "t2_base_iny", "t2_errores_iny",
                  "t2_corregido_iny", "t2_etiqueta", "t2_delta", "t2_delta_nombre", "t2_delta_firma",
                  "t2_delta_editado", "t2_delta_indices", "t2_editor_version", "t2_final_bytes",
                  "t2_final_nombre", "t2_final_firma", "t2_editor_vista"]
        claves += [k for k in st.session_state if k.startswith("t2_editor_datos_")]
        for k in claves: st.session_state.pop(k, None)
        st.rerun()

    # FUNCIONES AUXILIARES
    def leer_csv_t2(archivo):
        df = pd.read_csv(io.BytesIO(archivo.getvalue()), dtype=str, encoding="utf-8-sig", keep_default_na=False, skip_blank_lines=True)
        if df.columns.duplicated().any(): raise ValueError("El CSV tiene encabezados duplicados.")
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
        # El orden del reporte determina el orden del fragmento y su reintegración.
        return sorted(set(indices))

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
        st.caption("Conserva el orden de las filas y los encabezados. Este mismo fragmento sirve para Banner y para ensamblar el CSV completo.")
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
                                         use_container_width=True,
                                         key=f"t2_editor_datos_{st.session_state.t2_editor_version}")
                guardar = st.form_submit_button("💾 Guardar correcciones", use_container_width=True)
            if guardar:
                if list(editado.columns) != list(st.session_state.t2_delta.columns) or len(editado) != len(st.session_state.t2_delta):
                    raise ValueError("No se pueden alterar las columnas ni la cantidad de registros.")
                st.session_state.t2_delta_editado = editado.copy()
                st.session_state.t2_editor_version += 1
                st.success("Cambios guardados correctamente.")
                st.rerun()
            st.download_button("📥 Descargar CSV corregido (Banner y ensamblaje)", exportar_delta_t2(st.session_state.t2_delta_editado, "csv"),
                               file_name=f"{nombre_delta}.csv", mime="text/csv", use_container_width=True)

    # 2. ENSAMBLAR ARCHIVO FINAL
    st.divider()
    st.subheader("💉 2. Inyectar correcciones y generar CSV final")
    st.caption("Sube el mismo fragmento corregido de la parte 1. Se ubica cada fila por el reporte de errores, en orden ascendente de línea.")
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
            if len(corregidas) != len(indices):
                raise ValueError(f"El reporte tiene {len(indices)} líneas distintas, pero el fragmento tiene {len(corregidas)} registros.")
            if list(corregidas.columns) != list(base.columns):
                raise ValueError("El fragmento debe tener exactamente los mismos encabezados y orden de columnas que el CSV original.")
            # El fragmento se genera en orden ascendente de línea del CSV.
            # No ordenar, insertar ni eliminar filas del fragmento durante la corrección.
            final = base.copy()
            final.iloc[indices, :] = corregidas.to_numpy()
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
        claves = ["t3_modo", "t3_argos_completo", "t3_csv_completo", "t3_excel_completo", "t3_argos_rapido", "t3_csv_rapido", "final_argos_zip", "df_cruce_rapido", "t3_alertas", "t3_resultado_editado", "t3_editor_version", "t3_columnas", "t3_firma_completo", "t3_firma_rapido", "t3_tabla_revision", "t3_informe_zip", "t3_diferencias", "t3_match", "t3_firma_revision", "t3_salida", "t3_incluir_original", "t3_pendientes", "t3_manual_version", "t3_salida_procesada", "t3_nrc_previos", "t3_nrc_pendientes", "t3_nrc_elecciones", "t3_nrc_firma_previos"]
        claves += [k for k in st.session_state if k.startswith(("t3_editor_", "t3_eleccion_"))]
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
        if faltantes: raise ValueError(f"El archivo {archivo.name} no es un reporte ARGOS/PA válido. Faltan {faltantes}. Carga el reporte de oferta con NRC, Nivel, Cluster, Periodo, Área, No. Curso y Grupo; no el CSV_ALTAS de Banner.")
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

    # NRC utilizados en entregas anteriores: se consultan solo cuando hay ambiguedad.
    def t3_leer_nrc_previos(archivos):
        usados, detalle = set(), []
        for archivo in archivos or []:
            if archivo.name.lower().endswith(".csv"):
                d = pd.read_csv(io.BytesIO(archivo.getvalue()), dtype=str, encoding="utf-8-sig", keep_default_na=False)
            else:
                libro = pd.ExcelFile(io.BytesIO(archivo.getvalue()))
                hoja = next((h for h in libro.sheet_names if normalizar_para_busqueda_t3(h) in ("crnsfinales", "altas", "crns", "nrcs")), None)
                if hoja is None:
                    hoja = next((h for h in libro.sheet_names if "final" in normalizar_para_busqueda_t3(h)), libro.sheet_names[0])
                d = libro.parse(hoja, dtype=str).fillna("")
            columnas = {normalizar_para_busqueda_t3(c): c for c in d.columns}
            cp = next((columnas[k] for k in ("periodo",) if k in columnas), None)
            cn = next((columnas[k] for k in ("nrc", "crn") if k in columnas), None)
            if cp is None or cn is None:
                raise ValueError(f"{archivo.name}: se requieren las columnas PERIODO y NRC.")
            for _, r in d.iterrows():
                periodo, nrc = ultra_limpiar(r[cp]), ultra_limpiar(r[cn])
                if periodo and nrc:
                    usados.add((periodo, nrc))
                    detalle.append({"Archivo": archivo.name, "PERIODO": periodo, "NRC": nrc})
        return usados, pd.DataFrame(detalle)

    def t3_resolver_nrc(df_csv, argos_df, mapa, usados_previos=None, elecciones=None, nombre="CSV definitivo"):
        # Primero reservar coincidencias unicas; despues resolver las ambiguas.
        usados_previos = usados_previos or set()
        elecciones = elecciones or {}
        llaves = llaves_csv_t3(df_csv)
        cortas = argos_df.groupby(["_periodo", "_subj", "_crse", "_grupo"])["_nrc"].agg(lambda x: sorted(set(x))).to_dict()
        registros = []
        for i, (_, fila) in enumerate(df_csv.iterrows()):
            periodo = ultra_limpiar(fila["PERIODO"])
            candidatos = list(mapa.get(llaves.iloc[i], []))
            busqueda = "Llave completa"
            if not candidatos:
                corta = (periodo, ultra_limpiar(fila["SUBJ"]), ultra_limpiar(fila["COURSE"]), ultra_limpiar_seccion(fila["SECCION"]))
                candidatos = list(cortas.get(corta, []))
                busqueda = "Llave corta"
            disponibles = [n for n in candidatos if (periodo, n) not in usados_previos]
            registros.append({"Indice": i, "Archivo": nombre, "Fila CSV": i + 2, "PERIODO": periodo,
                              "SUBJ": fila["SUBJ"], "COURSE": fila["COURSE"], "SECCION": fila["SECCION"],
                              "Nivel": nivel_csv_t3(fila["datocomplementario"]), "NRC candidatos": ", ".join(candidatos),
                              "Disponibles": disponibles, "Busqueda": busqueda, "_candidatos": candidatos})
        asignados, ocupados = {}, set()
        # Las elecciones explicitas se reservan antes de asignar automaticamente.
        for i, elegido in elecciones.items():
            i = int(i)
            if not elegido: continue
            if i < 0 or i >= len(registros): raise ValueError(f"Indice de eleccion invalido: {i}")
            r = registros[i]
            par = (r["PERIODO"], str(elegido))
            if str(elegido) not in r["Disponibles"] or par in ocupados:
                raise ValueError(f"Fila {i+2}: NRC {elegido} no disponible o repetido en el periodo.")
            asignados[i] = str(elegido)
            ocupados.add(par)
        # Reservar las coincidencias forzosas antes de consumir candidatos ambiguos.
        cambios = True
        while cambios:
            cambios = False
            for r in registros:
                i, periodo = r["Indice"], r["PERIODO"]
                if i in asignados: continue
                libres = [n for n in r["Disponibles"] if (periodo, n) not in ocupados]
                if len(libres) == 1:
                    # Si otra fila tambien solo puede usarlo, no asignar arbitrariamente.
                    competidores = sum(1 for otro in registros if otro["Indice"] not in asignados and
                        otro["PERIODO"] == periodo and
                        [n for n in otro["Disponibles"] if (periodo, n) not in ocupados] == libres)
                    if competidores == 1:
                        asignados[i] = libres[0]
                        ocupados.add((periodo, libres[0]))
                        cambios = True
        pendientes = []
        for r in registros:
            if r["Indice"] in asignados: continue
            periodo = r["PERIODO"]
            libres = [n for n in r["Disponibles"] if (periodo, n) not in ocupados]
            estado = "Varios NRC" if len(libres) > 1 else ("NRC compartido: requiere elegir" if len(libres) == 1 else ("Todos utilizados" if r["_candidatos"] else "No encontrado"))
            pendientes.append({k: v for k, v in r.items() if k != "_candidatos"} | {"Disponibles": libres, "Estado": estado})
        return pd.Series([asignados.get(i, "") for i in range(len(registros))], index=df_csv.index, dtype="str"), pendientes

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
        if conflictos.any(): raise ValueError("ARGOS/PA contiene NRC contradictorios dentro de un mismo periodo. Revisa las advertencias y los reportes antes de generar.")
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
        # Exactas unicas primero. Las otras son propuestas, nunca asignaciones arbitrarias.
        pendientes, asignadas, motivos = set(csv.index), {}, {}
        campos = ["PERIODO", "SUBJ", "COURSE", "SEDE", "PARTE_PERIODO", "STATUS", "CAPACIDAD", "SECCION", "TIPO_HORARIO", "METODO_EDUCATIVO", "MODO_CALIFICAR", "SESION"]
        def firma_ex(r): return tuple(t3_n(r[c], c) for c in campos)
        def firma_csv(r): return tuple(t3_n(r[T3_MAPA[c]], c) for c in campos)
        gx, gc = {}, {}
        for i, r in excel.iterrows(): gx.setdefault(firma_ex(r), []).append(i)
        for j, r in csv.iterrows(): gc.setdefault(firma_csv(r), []).append(j)
        for k, ix in gx.items():
            jc = gc.get(k, [])
            if len(ix) == len(jc) == 1:
                asignadas[ix[0]], motivos[ix[0]] = jc[0], "Exacta"
                pendientes.discard(jc[0])
        # Aplicar automaticamente solo una diferencia cuando el emparejamiento sea mutuo y unico.
        pendientes_excel = [i for i in excel.index if i not in asignadas]
        candidaturas = {}
        for i in pendientes_excel:
            r = excel.loc[i]
            opciones = []
            for j in sorted(pendientes):
                c = csv.loc[j]
                puntos, mismo_periodo = t3_similares(r, c)
                if not mismo_periodo or puntos < len(campos) - 1: continue
                cl, cc = t3_n(r.get("CLUSTER", "")), t3_n(c.get("datocomplementario", ""))
                if cl and cc and cl != cc and not (t3_n(r.get("NIVEL", "")) == "BACHILLERATO" and cc == "BACHILLERATO"): continue
                opciones.append(j)
            candidaturas[i] = opciones
        for i, js in candidaturas.items():
            if len(js) == 1 and sum(js[0] in otros for otros in candidaturas.values()) == 1:
                asignadas[i], motivos[i] = js[0], "Una diferencia, correspondencia unica"
                pendientes.discard(js[0])
        propuestas = []
        for i, r in excel.iterrows():
            if i in asignadas: continue
            candidatos = []
            for j in sorted(pendientes):
                c = csv.loc[j]
                puntaje, mismo_periodo = t3_similares(r, c)
                if not mismo_periodo: continue
                cl, cc = t3_n(r.get("CLUSTER", "")), t3_n(c.get("datocomplementario", ""))
                if cl and cc and cl != cc and not (t3_n(r.get("NIVEL", "")) == "BACHILLERATO" and cc == "BACHILLERATO"): continue
                candidatos.append((puntaje, j))
            candidatos.sort(key=lambda x: (-x[0], x[1]))
            mejor = candidatos[0][0] if candidatos else 0
            sugeridos = [j for p, j in candidatos if p == mejor and p >= 6]
            propuestas.append({"Fila": i, "Archivo Excel": r["_ArchivoExcel"], "Fila Excel": int(r["_FilaExcel"]), "Periodo": r["PERIODO"], "Materia Excel": f'{r["SUBJ"]} {r["COURSE"]}', "Seccion Excel": r["SECCION"], "CSV sugerido": sugeridos[0] if len(sugeridos) == 1 else -1, "Coincidencias": mejor, "Estado": "Sugerencia para confirmar" if len(sugeridos) == 1 else "Revisar manualmente"})
        return asignadas, motivos, pd.DataFrame(propuestas)

    def t3_generar_resultado(excel, csv, asignadas, argos_df, mapa, usados_previos=None, elecciones=None):
        if len(asignadas) != len(excel) or len(set(asignadas.values())) != len(asignadas):
            raise ValueError("Faltan correspondencias unicas entre Excel y CSV.")
        if len(asignadas) != len(csv): raise ValueError("El numero de filas de ALTAS y CSV no coincide. Revisa registros agregados o eliminados.")
        final, diferencias = [], []
        csv_ordenado = csv.loc[[asignadas[i] for i in excel.index]].copy().reset_index(drop=True)
        nrc, avisos = t3_resolver_nrc(csv_ordenado, argos_df, mapa, usados_previos, elecciones)
        if avisos: return None, pd.DataFrame(), avisos
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
        return pd.DataFrame(final), pd.DataFrame(diferencias), []

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

    # 4. MODO COMPLETO: un boton principal; revision solo cuando es necesaria.
    if modo.startswith("📦"):
        st.subheader("📦 Conciliación Excel - CSV - ARGOS")
        c1, c2, c3 = st.columns(3)
        argos = c1.file_uploader("ARGOS / PA actualizado con NRC (.csv)", type=["csv"], accept_multiple_files=True, key="t3_argos_completo")
        archivos_csv = c2.file_uploader("CSV definitivos de Banner", type=["csv"], accept_multiple_files=True, key="t3_csv_completo")
        archivos_excel = c3.file_uploader("Excel ALTAS originales o separados", type=["xlsx", "xlsm"], accept_multiple_files=True, key="t3_excel_completo")
        opciones = {"Como venían originalmente": [], "Por periodo": ["PERIODO"], "Por clúster": ["CLUSTER"], "Por responsable": ["RESPONSABLE"], "Por periodo y clúster": ["PERIODO", "CLUSTER"], "Por periodo y responsable": ["PERIODO", "RESPONSABLE"], "Por clúster y responsable": ["CLUSTER", "RESPONSABLE"], "Por periodo, clúster y responsable": ["PERIODO", "CLUSTER", "RESPONSABLE"]}
        salida = st.selectbox("¿Cómo deseas recibir los Excel finales?", list(opciones), key="t3_salida")
        firma = firma_t3(*(argos or []), *(archivos_csv or []), *(archivos_excel or []))
        if st.session_state.get("t3_firma_completo") != firma:
            for k in ("t3_match", "t3_tabla_revision", "final_argos_zip", "t3_alertas", "t3_diferencias", "t3_pendientes", "t3_nrc_pendientes", "t3_nrc_elecciones"):
                st.session_state.pop(k, None)
            st.session_state.t3_firma_completo = firma
            st.session_state.t3_manual_version = st.session_state.get("t3_manual_version", 0) + 1
        if st.session_state.get("t3_salida_procesada") != salida:
            st.session_state.final_argos_zip = None

        def t3_exportar(datos, argos_df, mapa, usados_previos=None, elecciones=None):
            excel, csv, asignadas = datos["excel"], datos["csv"], datos["asignadas"]
            final, diferencias, pendientes_nrc = t3_generar_resultado(excel, csv, asignadas, argos_df, mapa, usados_previos, elecciones)
            if pendientes_nrc:
                st.session_state.t3_nrc_pendientes = pendientes_nrc
                st.session_state.final_argos_zip = None
                return False
            st.session_state.t3_nrc_pendientes = []
            st.session_state.t3_diferencias = diferencias
            grupos = [(k, g) for k, g in final.groupby(opciones[salida], sort=False, dropna=False)] if opciones[salida] else [(k, g) for k, g in final.groupby("_ArchivoExcel", sort=False)]
            b, rutas, macros_corregidas = io.BytesIO(), set(), 0
            with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
                for _, grupo in grupos:
                    nombre = t3_nombre_grupo(grupo)
                    z.writestr(ruta_zip_unica("CRNS_FINALES_SIN_MACROS", nombre, rutas), t3_excel_bytes(grupo))
                # Solo entregar el XLSM si realmente hay cambios frente al CSV.
                for f in archivos_excel:
                    if not f.name.lower().endswith(".xlsm"): continue
                    g = final[final["_ArchivoExcel"] == f.name]
                    cambios = diferencias[diferencias["Archivo Excel"] == f.name] if not diferencias.empty else pd.DataFrame()
                    if g.empty or cambios.empty: continue
                    wb = openpyxl.load_workbook(io.BytesIO(f.getvalue()), keep_vba=True)
                    hoja = next((h for h in wb.sheetnames if h.strip().upper() == HOJA_ALTAS), None)
                    if hoja is None: raise ValueError(f"{f.name}: no tiene hoja ALTAS.")
                    ws = wb[hoja]
                    encabezados = {normalizar_para_busqueda(c.value): c.column for c in ws[1] if c.value is not None}
                    # Solo cambiar los campos realmente distintos; preservar los otros y el VBA.
                    for _, cambio in cambios.iterrows():
                        campo = cambio["Campo"]
                        col = next((encabezados[a] for a in T3_ALIAS.get(campo, []) if a in encabezados), None)
                        if col is not None: ws.cell(int(cambio["Fila Excel"]), col).value = cambio["Definitivo"]
                    if HOJA_SALIDA_NRC in wb.sheetnames: del wb[HOJA_SALIDA_NRC]
                    wsn = wb.create_sheet(HOJA_SALIDA_NRC)
                    wsn.append(["Fila ALTAS"] + T3_COLUMNAS_FINAL)
                    for _, reg in g.iterrows(): wsn.append([int(reg["_FilaExcel"])] + [reg[c] for c in T3_COLUMNAS_FINAL])
                    formatear_excel_t3(wsn)
                    buf = io.BytesIO()
                    wb.save(buf)
                    nombre = f.name.rsplit(".", 1)[0] + " CRNs Corregidos.xlsm"
                    z.writestr(ruta_zip_unica("MACROS_CORREGIDAS", nombre, rutas), buf.getvalue())
                    macros_corregidas += 1
            st.session_state.final_argos_zip = b.getvalue()
            st.session_state.t3_salida_procesada = salida
            st.success(f"Listo: {len(grupos)} Excel finales, {len(diferencias)} cambios desde CSV y {macros_corregidas} libros con macros corregidos.")
            return True

        if argos and archivos_csv and archivos_excel and st.button("🚀 Procesar y generar Excel finales con NRC", type="primary", key="t3_procesar", use_container_width=True):
            try:
                with st.spinner("Comparando Excel, CSV y recuperando NRC..."):
                    argos_df, mapa, avisos = t3_argos_unidos(argos)
                    excel = t3_leer_excels(archivos_excel)
                    partes = []
                    for f in archivos_csv:
                        d = leer_csv_t3(f)
                        d["_ArchivoCSV"], d["_FilaCSV"] = f.name, range(2, len(d) + 2)
                        partes.append(d)
                    csv = pd.concat(partes, ignore_index=True)
                    if len(excel) != len(csv): raise ValueError(f"Excel: {len(excel)} registros; CSV: {len(csv)}. Revisa que hayas cargado todos los CSV correspondientes, sin duplicarlos.")
                    asignadas, motivos, propuestas = t3_proponer(excel, csv)
                    datos = {"excel": excel, "csv": csv, "asignadas": asignadas, "motivos": motivos}
                    st.session_state.t3_match = datos
                    st.session_state.t3_tabla_revision = propuestas
                    st.session_state.t3_alertas = avisos
                    st.session_state.final_argos_zip = None
                    if len(asignadas) == len(excel): t3_exportar(datos, argos_df, mapa)
                    else: st.warning(f"Se encontraron {len(asignadas)} coincidencias exactas y {len(excel)-len(asignadas)} registros para confirmar. Revisa las correspondencias que aparecen abajo.")
            except Exception as e:
                st.session_state.final_argos_zip = None
                st.error(f"No se pudo procesar: {e}")

        datos = st.session_state.get("t3_match")
        if datos and len(datos["asignadas"]) < len(datos["excel"]):
            excel, csv = datos["excel"], datos["csv"]
            propuestas = st.session_state.get("t3_tabla_revision")
            st.subheader("⚠️ Revisar correspondencias pendientes")
            st.caption("Solo aparecen cuando no es posible identificar una fila con certeza. El CSV definitivo será la referencia para las correcciones.")
            asignadas = dict(datos["asignadas"])
            usados = set(asignadas.values())
            with st.form("t3_form_correspondencias"):
                elecciones = {}
                for _, r in propuestas.iterrows():
                    i = int(r["Fila"])
                    if i in asignadas: continue
                    candidatos = [j for j in csv.index if j not in usados and t3_n(csv.loc[j, "PERIODO"]) == t3_n(excel.loc[i, "PERIODO"])]
                    cl = t3_n(excel.loc[i].get("CLUSTER", ""))
                    if cl:
                        candidatos = [j for j in candidatos if t3_n(csv.loc[j, "datocomplementario"]) == cl or (t3_n(excel.loc[i].get("NIVEL", "")) == "BACHILLERATO" and t3_n(csv.loc[j, "datocomplementario"]) == "BACHILLERATO")]
                    etiquetas = {j: f'{csv.loc[j,"_ArchivoCSV"]} | fila {csv.loc[j,"_FilaCSV"]} | {csv.loc[j,"SUBJ"]} {csv.loc[j,"COURSE"]} | sección {csv.loc[j,"SECCION"]}' for j in candidatos}
                    sugerido = int(r["CSV sugerido"])
                    opciones_fila = [-1] + candidatos
                    elecciones[i] = st.selectbox(f'{r["Archivo Excel"]}, fila {r["Fila Excel"]} — {r["Materia Excel"]}, sección {r["Seccion Excel"]}', opciones_fila, index=opciones_fila.index(sugerido) if sugerido in candidatos else 0, format_func=lambda j: "-- Seleccionar fila CSV --" if j == -1 else etiquetas[j], key=f't3_eleccion_{st.session_state.t3_manual_version}_{i}')
                confirmar = st.form_submit_button("💾 Confirmar correspondencias y generar", use_container_width=True)
            if confirmar:
                try:
                    seleccionadas = [j for j in elecciones.values() if j != -1]
                    if len(seleccionadas) != len(elecciones): raise ValueError("Faltan filas por seleccionar.")
                    if len(seleccionadas) != len(set(seleccionadas)) or set(seleccionadas) & usados: raise ValueError("Una misma fila CSV no puede asignarse a dos ALTAS.")
                    asignadas.update(elecciones)
                    datos["asignadas"] = asignadas
                    st.session_state.t3_match = datos
                    argos_df, mapa, avisos = t3_argos_unidos(argos)
                    st.session_state.t3_alertas = avisos
                    t3_exportar(datos, argos_df, mapa)
                    st.rerun()
                except Exception as e: st.error(f"No se pudo confirmar: {e}")

        # Solicitar antecedentes solamente cuando haya coincidencias ambiguas, en cualquier nivel.
        pendientes_nrc = st.session_state.get("t3_nrc_pendientes") or []
        if pendientes_nrc:
            st.subheader("⚠️ NRC pendientes de identificar")
            st.dataframe(pd.DataFrame([{k: v for k, v in r.items() if k not in ("Disponibles", "Indice")} for r in pendientes_nrc]), hide_index=True, use_container_width=True)
            ambiguos = [r for r in pendientes_nrc if len(r["Disponibles"]) > 1 or r["Estado"] == "NRC compartido: requiere elegir"]
            if ambiguos:
                periodos_necesarios = sorted({r["PERIODO"] for r in ambiguos})
                st.warning("Hay NRC ambiguos en los periodos: " + ", ".join(periodos_necesarios) + ". Carga los Excel de CRNs previamente asignados para descartar los ya repartidos.")
                previos = st.file_uploader("📁 CRNs ya asignados (puedes cargar varios Excel o CSV)", type=["xlsx", "xlsm", "csv"], accept_multiple_files=True, key="t3_nrc_previos")
                if previos:
                    try:
                        usados_previos, detalle_previos = t3_leer_nrc_previos(previos)
                        st.caption(f"Se identificaron {len(usados_previos)} pares PERIODO + NRC previamente asignados.")
                        with st.expander("Ver NRC previamente asignados", expanded=False):
                            st.dataframe(detalle_previos, hide_index=True, use_container_width=True)
                        datos_actuales = st.session_state.get("t3_match")
                        if datos_actuales and len(datos_actuales["asignadas"]) == len(datos_actuales["excel"]):
                            csv_ordenado = datos_actuales["csv"].loc[[datos_actuales["asignadas"][i] for i in datos_actuales["excel"].index]].reset_index(drop=True)
                            argos_df, mapa, avisos = t3_argos_unidos(argos)
                            _, pendientes_despues = t3_resolver_nrc(csv_ordenado, argos_df, mapa, usados_previos)
                            opciones_manual = [r for r in pendientes_despues if r["Disponibles"]]
                            elecciones = {}
                            for r in opciones_manual:
                                elecciones[r["Indice"]] = st.selectbox(
                                    f'Periodo {r["PERIODO"]} | {r["Nivel"]} | {r["SUBJ"]} {r["COURSE"]} | sección {r["SECCION"]} | fila {r["Fila CSV"]}',
                                    [""] + r["Disponibles"], key=f't3_nrc_eleccion_{r["Indice"]}')
                            sin_opciones = [r for r in pendientes_despues if not r["Disponibles"]]
                            if sin_opciones:
                                st.error(f"{len(sin_opciones)} registros no tienen NRC disponible; revisa los archivos previos o ARGOS.")
                            if st.button("🔍 Revisar NRC utilizados y generar Excel finales", type="primary", key="t3_revisar_nrc", use_container_width=True):
                                if sin_opciones or any(not n for n in elecciones.values()):
                                    st.error("Resuelve todas las asignaciones antes de exportar.")
                                else:
                                    try:
                                        seleccionados = [(next(r["PERIODO"] for r in opciones_manual if r["Indice"] == i), n) for i, n in elecciones.items()]
                                        if len(seleccionados) != len(set(seleccionados)):
                                            raise ValueError("Un NRC no puede asignarse dos veces en el mismo periodo.")
                                        if t3_exportar(datos_actuales, argos_df, mapa, usados_previos, elecciones): st.rerun()
                                        else: st.warning("Aún existen asignaciones sin resolver.")
                                    except Exception as e: st.error(f"No se pudo resolver la asignación: {e}")
                    except Exception as e:
                        st.error(f"No se pudieron leer o revisar los NRC anteriores: {e}")
            else:
                st.warning("No hay NRC disponibles o no se encontraron coincidencias en ARGOS. Revisa los registros señalados.")

        if isinstance(st.session_state.get("t3_diferencias"), pd.DataFrame) and not st.session_state.t3_diferencias.empty:
            with st.expander("Cambios aplicados desde el CSV", expanded=False):
                st.dataframe(st.session_state.t3_diferencias, hide_index=True, use_container_width=True)
        if st.session_state.get("t3_alertas"):
            with st.expander("Advertencias ARGOS"):
                st.dataframe(pd.DataFrame(st.session_state.t3_alertas), hide_index=True, use_container_width=True)
        if st.session_state.get("final_argos_zip"):
            st.download_button("📥 Descargar Excel finales (.ZIP)", st.session_state.final_argos_zip, file_name="CRNs_Finales.zip", mime="application/zip", type="primary", use_container_width=True)

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

