


# -*- coding: utf-8 -*-
"""
================================================================================
ANÁLISIS DE EXPOSICIÓN DE CARTERA A RIESGOS FÍSICOS (CENEPRED) — AFP INTEGRA
================================================================================
Qué hace este script
--------------------
1. Lee el Excel de PRODUCCIÓN MINERA (titular, unidad minera, depto/prov/distrito,
   unidad de medida y producción acumulada del año).
2. Convierte toda la producción a una unidad común: Tonelada Métrica Fina (TMF).
3. Geocodifica cada distrito a su UBIGEO usando un GeoJSON distrital del Perú.
4. Lee la hoja "CENEPRED" (score 1-5 por distrito para Mov. de Masa, Inundación
   y Sequías Severas).
5. Lee la hoja "Valorización de Instrumentos" y suma la exposición (S/) por
   empresa, aplicando una tabla de equivalencias de nombres (editable).
6. Distribuye la exposición de cada empresa entre los distritos donde produce
   (proporcional a su producción en TMF) y la cruza con el score CENEPRED para
   estimar la "plata en riesgo" por tipo de desastre y por ubicación.
7. Genera:
      - Un EXCEL de salida (una hoja por empresa + resúmenes).
      - Un HTML interactivo con el mapa coroplético del Perú.

Filosofía de blindaje
---------------------
- Las columnas se detectan por NOMBRE de encabezado (no por posición), con
  respaldo por letra de columna. Si la estructura cambia, el script aborta con
  un MENSAJE CLARO que indica qué buscaba, qué encontró y dónde mirar en el
  archivo para arreglarlo.
- El nombre del archivo es FLEXIBLE: se busca por patrón dentro de la carpeta.
- Toda la parametrización (rutas, equivalencias, correcciones de nombres de
  distritos, método de asignación) está en el bloque CONFIG de abajo, pensada
  para editarse a futuro sin tocar la lógica.
================================================================================
"""

import os
import re
import sys
import glob
import json
import unicodedata
from datetime import datetime

import pandas as pd
import numpy as np

# ==============================================================================
# 1) CONFIG  — EDITAR AQUÍ
# ==============================================================================

# Carpeta base del proyecto. Si la variable de entorno RF_BASE está definida,
# se usa esa (útil para pruebas); si no, se usa la ruta de AFP Integra.
BASE_DIR = os.environ.get(
    "RF_BASE",
    r"C:\Users\usuario\OneDrive\Desktop\AFP INTEGRA\ESG\Riesgos Fisicos\2026",
)

# Subcarpeta donde vive el Excel de producción minera (nombre flexible).
PROD_DIR = os.path.join(BASE_DIR, "Output")
# Patrones aceptados para el archivo de producción (el más reciente gana).
PROD_PATTERNS = ["Producci*Minera*.xls*", "*Producci*Minera*.xls*", "*Minera*.xls*"]

# Archivo de Riesgos Físicos (CENEPRED + Valorización). Nombre flexible y
# acepta .xlsx o .xlsm.
RISK_DIR = BASE_DIR
# El archivo real es "Riesgos Fisicos (Mapa Distrital) - 2026.xlsx" (el año varía).
# Apuntamos a él de forma específica para NO capturar los outputs del script
# (Riesgos_Fisicos_Cartera_*), que no contienen "Mapa Distrital".
RISK_PATTERNS = ["Riesgos Fisicos (Mapa Distrital) - *.xls*",
                 "Riesgos*Fisicos*Mapa*Distrital*.xls*",
                 "*Mapa*Distrital*.xls*"]

# Carpeta de salida (Excel + HTML).
OUT_DIR = os.path.join(BASE_DIR, "Inputs")

# GeoJSON distrital del Perú (con propiedades IDDIST/NOMBDEP/NOMBPROV/NOMBDIST).
# Se guarda/lee localmente; si no existe, se descarga de la URL.
GEOJSON_LOCAL = os.path.join(OUT_DIR, "peru_distritos.geojson")
GEOJSON_URL = ("https://raw.githubusercontent.com/juaneladio/peru-geojson/"
               "master/peru_distrital_simple.geojson")

# Para cambiar de fuente sin tocar código, edita GEOJSON_URL arriba y borra el
# cache GEOJSON_LOCAL para que se rebaje. Como los GeoJSON nacionales libres
# derivan del shapefile MINAM ~2016 (~1,834 distritos) y NO incluyen distritos
# creados después (Mi Perú 070107, Veintiséis de Octubre 200115, etc.), abajo se
# define un SUPLEMENTO que se fusiona con la base para rellenar esos huecos.
# Puedes añadir más distritos: pega features con properties.IDDIST/NOMBDEP/
# NOMBPROV/NOMBDIST y su geometry. Santa Anita y La Punta llevan polígono real
# (IGN); Mi Perú y 26 de Octubre, una celda ubicada en su centroide (aproximada).
# También se fusiona, si existe, un archivo externo opcional en la carpeta de
# salida llamado "peru_distritos_suplemento.geojson" (tiene prioridad).
GEOJSON_SUPLEMENTO_FILE = os.path.join(OUT_DIR, "peru_distritos_suplemento.geojson")
GEOJSON_SUPLEMENTO_BUILTIN = r'''{"type":"FeatureCollection","features":[{"type":"Feature","properties":{"IDDIST":"150137","NOMBDEP":"LIMA","NOMBPROV":"LIMA","NOMBDIST":"SANTA ANITA"},"geometry":{"type":"Polygon","coordinates":[[[-76.9379,-12.0439],[-76.94,-12.0386],[-76.9431,-12.0309],[-76.9428,-12.0283],[-76.9438,-12.0275],[-76.9558,-12.0296],[-76.9588,-12.03],[-76.9648,-12.0329],[-76.9665,-12.0341],[-76.9742,-12.0399],[-76.976,-12.0404],[-76.9774,-12.0405],[-76.98,-12.0406],[-76.9814,-12.0405],[-76.9832,-12.0401],[-76.9844,-12.0394],[-76.988,-12.0395],[-76.987,-12.0416],[-76.9855,-12.0435],[-76.9859,-12.0474],[-76.9891,-12.0506],[-76.9856,-12.0571],[-76.985,-12.0574],[-76.9829,-12.0564],[-76.978,-12.055],[-76.9753,-12.0539],[-76.9659,-12.0554],[-76.9575,-12.0518],[-76.9379,-12.0439]]]}},{"type":"Feature","properties":{"IDDIST":"070105","NOMBDEP":"CALLAO","NOMBPROV":"CALLAO","NOMBDIST":"LA PUNTA"},"geometry":{"type":"Polygon","coordinates":[[[-77.1584,-12.0666],[-77.1601,-12.0679],[-77.1623,-12.069],[-77.1653,-12.0701],[-77.1663,-12.0701],[-77.1669,-12.0676],[-77.1682,-12.0686],[-77.1688,-12.0712],[-77.1695,-12.072],[-77.1706,-12.0722],[-77.1698,-12.0745],[-77.1664,-12.0742],[-77.1648,-12.0746],[-77.1632,-12.0754],[-77.159,-12.0741],[-77.1581,-12.0729],[-77.1571,-12.0725],[-77.1579,-12.0712],[-77.1568,-12.0686],[-77.1584,-12.0666]]]}},{"type":"Feature","properties":{"IDDIST":"070107","NOMBDEP":"CALLAO","NOMBPROV":"CALLAO","NOMBDIST":"MI PERU"},"geometry":{"type":"Polygon","coordinates":[[[-77.1225,-11.83417],[-77.1025,-11.85417],[-77.1225,-11.87417],[-77.1425,-11.85417],[-77.1225,-11.83417]]]}},{"type":"Feature","properties":{"IDDIST":"200115","NOMBDEP":"PIURA","NOMBPROV":"PIURA","NOMBDIST":"VEINTISEIS DE OCTUBRE"},"geometry":{"type":"Polygon","coordinates":[[[-80.68083,-5.1569400000000005],[-80.66083,-5.17694],[-80.68083,-5.19694],[-80.70083,-5.17694],[-80.68083,-5.1569400000000005]]]}}]}'''

# Nombres de hojas (con detección flexible si no se encuentran exactos).
SHEET_CENEPRED_HINTS = ["CENEPRED"]
SHEET_VALORIZ_HINTS = ["Valoriz", "Valorización", "Valorizacion", "Instrumentos"]

# --- Equivalencia de UNIDADES de medida a TMF (Tonelada Métrica Fina) ---------
#   TMF = Tonelada Métrica Fina            -> factor 1
#   Gramos finos (Grs. f.)                 -> TMF = gramos / 1_000_000
#   Kilogramos finos (Kg. f.)              -> TMF = kg / 1_000
# La clave se normaliza (sin espacios/puntos/acentos, en minúsculas).
UNIDAD_A_TMF = {
    "tmf": 1.0,
    "grsf": 1.0 / 1_000_000.0,   # gramos finos  -> TMF
    "kgf": 1.0 / 1_000.0,        # kilogramos finos -> TMF
}
# Etiquetas "bonitas" para mostrar la unidad original en el HTML.
UNIDAD_LABEL = {"tmf": "TMF", "grsf": "Grs. f.", "kgf": "Kg. f.",
                "credmiles": "Miles S/", "mw": "MW", "mt": "MT", "pct": "%"}

# --- Equivalencia de NOMBRES de empresa (lado PRODUCCIÓN: titular legal) -------
# Editar/añadir libremente. Clave = nombre legal tal cual aparece en producción;
# Valor = nombre canónico (corto) con el que se reportará y se buscará exposición.
EQUIV_EMPRESAS_PRODUCCION = {
    "COMPAÑIA DE MINAS BUENAVENTURA S.A.A.": "Buenaventura",
    "SOCIEDAD MINERA EL BROCAL S.A.A.": "Buenaventura",
    "VOLCAN COMPAÑIA MINERA S.A.A.": "Volcan",
    "MINSUR S.A.": "Minsur",
    "NEXA RESOURCES PERU S.A.A.": "Nexa Resources Perú",
    "NEXA RESOURCES ATACOCHA S.A.A.": "Nexa Resources Perú",
    "NEXA RESOURCES EL PORVENIR S.A.C.": "Nexa Resources Perú",
    "NEXA RESOURCES CAJAMARQUILLA S.A.": "Nexa Resources Perú",
    "COMPAÑIA MINERA ARES S.A.C.": "Hochschild",
    "SOCIEDAD MINERA CERRO VERDE S.A.A.": "Cerro Verde",
    "MARCOBRE S.A.C.": "Minsur",
    "HUDBAY PERU S.A.C.": "Hudbay",
}

# --- Equivalencia de NOMBRES de empresa (lado EXPOSICIÓN: columna "Nombre (G&P)")
# Si la columna H trae variantes, mapéalas aquí al mismo nombre canónico.
# Los nombres que ya coinciden con el canónico no necesitan estar listados.
EQUIV_EMPRESAS_EXPOSICION = {
    "Nexa Resources": "Nexa Resources Perú",
}

# --- Correcciones manuales de geocodificación (cuando el nombre de distrito en
#     producción no calza con el GeoJSON). Clave = (DEP, PROV, DIST) normalizado;
#     valor = UBIGEO de 6 dígitos. Suele NO ser necesario gracias al respaldo
#     (departamento + distrito), pero queda aquí por si a futuro hace falta.
OVERRIDE_UBIGEO = {
    # ("ICA", "NASCA", "MARCONA"): "110304",
    ("LIMA", "LIMA", "PUEBLO LIBRE (MAGDALENA VIEJA)"): "150121",
    ("LIMA", "LIMA", "PUEBLO LIBRE"): "150121",
}

# --- Alias de DISTRITO (nombre popular <-> nombre oficial en el GeoJSON). El
# geocode prueba también estos equivalentes. Clave y valor en texto normal; la
# comparación es robusta a mayúsculas/tildes.
ALIAS_DISTRITO = {
    "PUEBLO LIBRE": "MAGDALENA VIEJA",   # Lima: nombre popular vs. oficial
    "MAGDALENA VIEJA": "PUEBLO LIBRE",
}

# Departamentos/zonas que NO son ubicables en el mapa del Perú (se reportan aparte
# y se excluyen limpiamente, sin contar como error de geocodificación).
DEPTOS_NO_MAPEABLES = {"EXTRANJERO", "EXTERIOR", "FUERA DEL PAIS"}

# --- Método de asignación de la exposición (S/) de cada empresa a sus distritos
#   "tmf_share"  : proporcional a la producción (TMF) de la empresa en cada distrito (recomendado)
#   "equal"      : repartida en partes iguales entre los distritos de la empresa
#   "units_share": proporcional al número de unidades mineras por distrito
METODO_ASIGNACION = "tmf_share"

# Empresas canónicas que son el universo de análisis (orden de reporte).
EMPRESAS_CANONICAS = ["Buenaventura", "Volcan", "Minsur", "Nexa Resources Perú",
                      "Hochschild", "Cerro Verde", "Hudbay"]

# Etiquetas de sector (se usan en el dashboard y el Excel).
SECTOR_MINERIA = "Minería"
SECTOR_FINANCIERO = "Financiero"

# Prefijos de los archivos que ESTE script genera. find_file los ignora siempre
# para no confundir un output (que puede caer en la misma carpeta) con un input.
EXCLUDE_OUTPUTS = ["Riesgos_Fisicos_Cartera", "Mapa_Riesgos_Fisicos"]

# ==============================================================================
# 1-bis) CONFIG DEL SECTOR FINANCIERO (BANCOS)
# ==============================================================================
# Analogía con minería: el "Total_Creditos_Suma" por distrito (en el archivo ya
# procesado B-2358-resumen-*.xlsx) cumple el mismo rol que la PRODUCCIÓN (TMF) en
# minería: es el PESO con el que se reparte geográficamente la exposición de
# cartera de cada banco entre los distritos donde coloca créditos.
#
# Interruptor maestro: pon en False para volver al análisis solo-minería.
INCLUIR_FINANCIERO = True

# Carpeta donde vive el resumen de bancos. Por defecto la raíz "AFP INTEGRA"
# (un nivel arriba de ...\ESG\Riesgos Fisicos\2026). Override con la variable de
# entorno RF_FIN si lo tienes en otro lado.
FIN_DIR = os.environ.get(
    "RF_FIN",
    os.path.dirname(os.path.dirname(os.path.dirname(BASE_DIR))),
)
# Patrones de nombre aceptados (el más reciente gana). El archivo ESPERADO es el
# RESUMEN ya procesado por tu script Financials (columnas Empresa, Departamento,
# Provincia, Distrito, Total_Creditos_Suma), NO el crudo B-2358 de la SBS.
FIN_PATTERNS = ["B-2358-resumen-*.xlsx", "B-2358-resumen-*.xls*",
                "*resumen*B-2358*.xls*", "*resumen*creditos*.xls*"]

# Nombres ESPERADOS de columna en el resumen de bancos (se buscan de forma robusta
# a mayúsculas/minúsculas y tildes; si no se encuentran, el script dice cuál falta).
FIN_COLS = {
    "empresa": ["Empresa", "Banco", "Entidad"],
    "depto":   ["Departamento", "Depto", "Region"],
    "prov":    ["Provincia"],
    "dist":    ["Distrito"],
    "monto":   ["Total_Creditos_Suma", "Total Creditos Suma", "Total Creditos",
                "Total_Creditos", "Creditos", "Monto"],
}

# --- Equivalencia de NOMBRES de banco (lado CRÉDITOS: columna A del resumen) ---
# Clave = nombre tal cual aparece en la columna "Empresa" del resumen de bancos;
# Valor = nombre canónico del grupo en el portafolio. El emparejamiento es robusto
# a mayúsculas/minúsculas, tildes y espacios extra (se normaliza antes de comparar).
# Si la SBS cambia un nombre en el futuro, el script te avisará por consola qué
# bancos del archivo NO están mapeados aquí para que los agregues.
EQUIV_EMPRESAS_FINANCIERO = {
    "BANCO DE CREDITO": "Credicorp",
    "MIBANCO": "Credicorp",        # Mibanco es subsidiaria de Credicorp (microfinanzas)
    "INTERBANK": "Intercorp Perú",
    "BBVA": "BBVA",
}

# Bancos canónicos que forman el universo financiero de análisis (orden de reporte).
EMPRESAS_FIN_CANONICAS = ["Credicorp", "Intercorp Perú", "BBVA"]

# --- Equivalencia de NOMBRES (lado EXPOSICIÓN: columna H "Nombre (G&P)") --------
# Igual que en minería, en "Valorización de Instrumentos" un mismo grupo económico
# puede aparecer bajo varias razones sociales. Aquí decides QUÉ líneas de la col H
# se SUMAN a cada banco canónico. Es un juicio de negocio: revísalo.
#   OJO con los homónimos que NO son del grupo peruano y por eso quedan FUERA por
#   defecto (descoméntalos solo si decides incluirlos):
#     - "BBVA Bancomer"                  -> BBVA México (no BBVA Perú)
#     - "Banco de Crédito e Inversiones" -> BCI de Chile (no BCP/Credicorp)
# El script imprime una auditoría con lo que sumó a cada grupo y los nombres de la
# col H que parecen bancarios pero quedaron sin mapear, para que ajustes esta lista.
EQUIV_EXPOSICION_FINANCIERO = {
    "Credicorp": "Credicorp",
    "Banco de Crédito del Perú": "Credicorp",
    "Intercorp Perú": "Intercorp Perú",
    "Interbank": "Intercorp Perú",
    "BBVA": "BBVA",
    # "BBVA Bancomer": "BBVA",                        # ← BBVA México (revisar)
    # "Banco de Crédito e Inversiones": "Credicorp",  # ← BCI Chile (revisar)
}

# Pistas para detectar, en la auditoría, nombres de col H que "parecen banco" y
# podrías querer mapear (solo para avisarte; no afecta el cálculo).
PISTAS_BANCO = ["BANCO", "BANK", "CREDICORP", "INTERCORP", "INTERBANK", "BBVA",
                "SCOTIA", "FINANCI", "CREDITO", "BCP"]

# ==============================================================================
# 1-ter) CONFIG DEL SECTOR UTILITIES (GENERACIÓN ELÉCTRICA)
# ==============================================================================
# Analogía: la POTENCIA EFECTIVA (MW) de cada central cumple el rol del peso de
# reparto (producción en minería, créditos en banca). La exposición de cartera a
# cada empresa sale de "Valorización de Instrumentos" (col H / col M) y se reparte
# entre los distritos donde la empresa tiene centrales, en proporción a sus MW.
INCLUIR_UTILITIES = True
SECTOR_UTILITIES = "Utilities"
# Carpeta de las bases macro (BD Utilities / BD Construcción). Por defecto
# ...\AFP INTEGRA\Docs privados\Proyecto Inputs Macro. Override con RF_MACRO.
MACRO_DIR = os.environ.get(
    "RF_MACRO",
    os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(BASE_DIR))),
                 "Docs privados", "Proyecto Inputs Macro"))
UTIL_DIR = MACRO_DIR
UTIL_PATTERNS = ["BD Utilities*.xls*", "BD_Utilities*.xls*", "*Utilities*.xls*"]
UTIL_SHEET_HINTS = ["BD ptf", "BD PTF", "ptf"]
UTIL_TIPO = "POTENCIA EFECTIVA"     # valor de la columna B (Tipo) que se usa
# Equivalencia de NOMBRES de empresa (col C de "BD ptf") -> canónico de cartera.
EQUIV_EMPRESAS_UTILITIES = {
    "ORYGEN": "Orygen",
    "ENGIE": "Engie",
    "ORAZUL ENERGY PERÚ": "Orazen",
}
EMPRESAS_UTIL_CANONICAS = ["Orygen", "Engie", "Orazen"]
# Equivalencia lado EXPOSICIÓN (col H de Valorización) -> canónico. Aquí coinciden.
EQUIV_EXPOSICION_UTILITIES = {
    "Orygen": "Orygen", "Engie": "Engie", "Orazen": "Orazen",
}
# Hoja del archivo de Riesgos que ubica cada central -> Dep/Prov/Distrito.
CENTRALES_SHEET_HINTS = ["Centrales Electricas", "Centrales Eléctricas", "Centrales"]

# ==============================================================================
# 1-quater) CONFIG DEL SECTOR CONSTRUCCIÓN (CEMENTO)
# ==============================================================================
# Analogía: los DESPACHOS de cemento (MT = miles de toneladas) son el peso de
# reparto. Exposición desde "Valorización" (col H / M), repartida entre los
# distritos donde la empresa tiene plantas, en proporción a los despachos.
INCLUIR_CONSTRUCCION = True
SECTOR_CONSTRUCCION = "Construcción"
CONS_DIR = MACRO_DIR
CONS_PATTERNS = ["BD Construcc*.xls*", "BD_Construcc*.xls*", "*Construcc*.xls*"]
CONS_SHEET_HINTS = ["Despachos"]
# Equivalencia de NOMBRES de empresa (parte antes del " - " en col C) -> canónico.
EQUIV_EMPRESAS_CONSTRUCCION = {
    "UNACEM": "Unacem",
    "PACASMAYO": "Pacasmayo",
}
EMPRESAS_CONS_CANONICAS = ["Unacem", "Pacasmayo"]
EQUIV_EXPOSICION_CONSTRUCCION = {
    "Unacem": "Unacem", "Pacasmayo": "Pacasmayo",
}
# Ubicación de cada PLANTA de cemento. Ahora vive en la MISMA hoja de cruce que
# las centrales eléctricas ("Centrales Electricas": columna 'Central' = nombre de la
# planta, + Empresa/Dep/Prov/Distrito). Se busca primero una hoja dedicada y, si no
# existe, esa. UBIC_PLANTAS_CEMENTO queda solo como última red si faltara la hoja.
PLANTAS_SHEET_HINTS = ["Plantas Cemento", "Plantas Cementeras",
                       "Centrales Electricas", "Centrales Eléctricas", "Centrales"]
# Filas de "Despachos" que agregan DOS plantas ("A + B"): cómo desglosarlas en los
# nombres EXACTOS de la hoja de ubicación, repartiendo el valor por igual (50/50).
# Si cambia el texto en la BD, ajusta la clave; si cambia el nombre en la hoja de
# ubicación, ajusta los valores. El script avisa si una planta no se ubica.
MAPEO_PLANTAS_DESPACHOS = {
    "Pacasmayo - Pacasmayo + Piura": ["Planta de Pacasmayo", "Pacasmayo - Piura"],
}
# Última red por si faltara por completo la hoja de ubicación (NO se usa si la hoja
# 'Centrales Electricas' ya trae las plantas, como es el caso actual).
UBIC_PLANTAS_CEMENTO = {
    "Unacem - Atocongo": ("Lima", "Lima", "Villa María del Triunfo"),
    "Unacem - Andino":   ("Junín", "Tarma", "La Unión"),
    "Planta de Pacasmayo": ("La Libertad", "Pacasmayo", "Pacasmayo"),
    "Pacasmayo - Piura": ("Piura", "Piura", "Veintiseis de Octubre"),
    "Pacasmayo - Rioja": ("San Martín", "Rioja", "Elias Soplin Vargas"),
}

# --- Ventana temporal para Utilities y Construcción ---------------------------
# Se SUMAN los valores del SEMESTRE ANTERIOR completo (6 meses). Si hoy es
# 11/06/2026 -> se suman Jul..Dic 2025. Si hoy cae en el 2º semestre, sería el
# 1er semestre del mismo año. SEMESTRE_REF: None = fecha de hoy; o fija una fecha
# "YYYY-MM-DD" para reproducir un cierre concreto sin importar cuándo se ejecute.
SEMESTRE_REF = None

# ==============================================================================
# METADATOS DE SECTOR (etiquetas y unidades, para Excel y dashboard)
# ==============================================================================
SECTOR_CONSUMO = "Consumo"
SECTOR_SALUD = "Salud"
SECTOR_INFRA = "Infraestructura"
SECTOR_EDU = "Educación"
SECTOR_ENERGIA = "Energía"
_META_PCT = {"peso": "Distribución", "unidad": "%", "activo": "Ubicaciones",
             "detalle": "ubicación", "pseudo": "Distribución"}
SECTOR_META = {
    SECTOR_MINERIA:     {"peso": "Producción", "unidad": "TMF",
                         "activo": "Unidades mineras", "detalle": "mineral",
                         "pseudo": None},
    SECTOR_FINANCIERO:  {"peso": "Créditos colocados", "unidad": "Miles S/",
                         "activo": "Oficinas/Agencias", "detalle": "colocaciones",
                         "pseudo": "Créditos"},
    SECTOR_UTILITIES:   {"peso": "Potencia efectiva", "unidad": "MW",
                         "activo": "Centrales eléctricas", "detalle": "central",
                         "pseudo": "Potencia efectiva"},
    SECTOR_CONSTRUCCION:{"peso": "Despachos de cemento", "unidad": "MT (miles t)",
                         "activo": "Plantas de cemento", "detalle": "planta",
                         "pseudo": "Despachos cemento"},
    SECTOR_CONSUMO:     dict(_META_PCT, activo="Predios / locales"),
    SECTOR_SALUD:       dict(_META_PCT, activo="Sedes / clínicas"),
    SECTOR_INFRA:       dict(_META_PCT, activo="Activos / concesiones"),
    SECTOR_EDU:         dict(_META_PCT, activo="Sedes / colegios"),
    SECTOR_ENERGIA:     dict(_META_PCT, activo="Activos / operaciones"),
}
ORDEN_SECTORES = [SECTOR_MINERIA, SECTOR_FINANCIERO, SECTOR_UTILITIES,
                  SECTOR_CONSTRUCCION, SECTOR_CONSUMO, SECTOR_SALUD,
                  SECTOR_INFRA, SECTOR_EDU, SECTOR_ENERGIA]

# ==============================================================================
# EMPRESAS CON HOJA PROPIA (peso = % de distribución por ubicación)
# ==============================================================================
# Cada una vive en su propia hoja del archivo de Riesgos. La hoja trae, a partir de
# una fila-header con Departamento/Provincia/Distrito y una columna de peso
# (% Porcentaje / Participación % / % Criterio… / Peso…), la base de datos de
# ubicaciones. El peso (decimal) reparte el AUM del emisor (col H de Valorización)
# entre esos distritos, y se cruza con CENEPRED igual que los demás sectores.
INCLUIR_EMPRESAS_HOJA = True
# Encabezados que identifican la columna de PESO (se usa la 1ra tras 'Distrito').
PESO_HINTS = ["DISTRIBUCION", "PARTICIPACION", "PORCENTAJE", "CRITERIO", "PESO", "%"]
# spec: hoja (nombre sin el número), empresa canónica, sector, y nombres en col H.
EMPRESAS_HOJA = [
    {"hoja": "Aenza",            "empresa": "Aenza",        "sector": SECTOR_INFRA,  "exp": ["Aenza"]},
    {"hoja": "Alicorp",          "empresa": "Alicorp",      "sector": SECTOR_CONSUMO,"exp": ["Alicorp"]},
    {"hoja": "Auna",             "empresa": "Auna",         "sector": SECTOR_SALUD,  "exp": ["Auna"]},
    {"hoja": "Ferreycorp",       "empresa": "Ferreycorp",   "sector": SECTOR_INFRA,  "exp": ["Ferreycorp"]},
    {"hoja": "Pluz Energia",     "empresa": "Pluz Energía", "sector": SECTOR_UTILITIES, "exp": ["Pluz Energía"]},
    {"hoja": "Puerto Chancay",   "empresa": "Puerto Chancay","sector": SECTOR_INFRA, "exp": ["Puerto Chancay"]},
    {"hoja": "Casa Andina",      "empresa": "Casa Andina",  "sector": SECTOR_CONSUMO,"exp": ["Casa Andina"]},
    {"hoja": "Colegios Peruanos","empresa": "Colegios Peruanos","sector": SECTOR_EDU,"exp": ["PF Colegios Peruanos"]},
    {"hoja": "Hermes",           "empresa": "Hermes",       "sector": SECTOR_INFRA,  "exp": ["Hermes"]},
    {"hoja": "Hunt Oil",         "empresa": "Hunt Oil",     "sector": SECTOR_ENERGIA,"exp": ["Hunt Oil"]},
    {"hoja": "Inca Rail",        "empresa": "Inca Rail",    "sector": SECTOR_INFRA,  "exp": ["Inca Rail"]},
    {"hoja": "InRetail",         "empresa": "InRetail",     "sector": SECTOR_CONSUMO,
     "exp": ["InRetail Shopping Malls", "InRetail Consumer", "InRetail Peru"]},
    {"hoja": "Intursa",          "empresa": "Intursa",      "sector": SECTOR_CONSUMO,
     "exp": ["Inversiones Nacionales de Turismo"]},
    {"hoja": "Jockey Plaza",     "empresa": "Jockey Plaza", "sector": SECTOR_CONSUMO,"exp": ["Jockey Plaza"]},
    {"hoja": "Lima Expresa",     "empresa": "Lima Expresa", "sector": SECTOR_INFRA,  "exp": ["Línea Amarilla"]},
    {"hoja": "Primax",           "empresa": "Primax",       "sector": SECTOR_CONSUMO,"exp": ["Primax"]},
    {"hoja": "Tecsup",           "empresa": "Tecsup",       "sector": SECTOR_EDU,    "exp": ["Tecsup"]},
    {"hoja": "Rutas de Lima",    "empresa": "Rutas de Lima","sector": SECTOR_INFRA,  "exp": ["Rutas de Lima"]},
]

HAZARDS = ["Movimientos de Masa", "Inundación", "Sequías Severas"]
HAZARD_KEY = {"Movimientos de Masa": "mm", "Inundación": "inu", "Sequías Severas": "seq"}

# Paleta corporativa AFP Integra
AFP_AZUL = "#1E2E6E"
AFP_CYAN = "#00AECB"
AFP_AMAR = "#E3E829"
AFP_GRIS = "#DCDDDE"   # gris claro (fondos, divisores, tarjetas)
AFP_PLOMO = "#7E8083"  # plomo (textos secundarios, montos)

# Logo institucional. Se embebe en el HTML como Base64 al generar (archivo único
# y portable por correo). Si la ruta no existe, se usa un wordmark SVG de respaldo.
LOGO_PATH = os.environ.get(
    "RF_LOGO",
    os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(BASE_DIR))),
                 "Docs privados", "ZODA", "shortcut", "integra_logo.png"))

# Cómo se define el "nivel de riesgo CENEPRED" de un distrito para la tabla pivote:
#   "max"  = el peligro más alto entre Mov. Masa / Inundación / Sequías (conservador)
#   "comb" = el score combinado (promedio) redondeado
NIVEL_CENEPRED_MODO = "max"
# Niveles que se consideran "altos" y entran en la tabla pivote (1 y 2 son leves).
NIVELES_ALTOS = [5, 4, 3]

# Archivo ENFEN (El Niño / La Niña) para la pestaña de Fenómenos Geofísicos.
INCLUIR_ENFEN = True
ENFEN_DIR = os.path.join(os.path.dirname(BASE_DIR), "FEN", "ENFEN")
ENFEN_PATTERNS = ["ENFEN_Comunicados_Resumen*.xls*", "*ENFEN*Resumen*.xls*",
                  "*ENFEN*.xls*"]
ENFEN_SHEET_HINTS = ["Comunicados ENFEN", "Comunicados", "ENFEN"]

# ==============================================================================
# 1-sextus) CONFIG DEL GRÁFICO CLIMA vs. INDICADOR SECTORIAL (pestaña Geofísicos)
# ==============================================================================
INCLUIR_FEN_CHART = True
# --- Indicadores climáticos (varias hojas en un mismo archivo) ---
FEN_IND_DIR = os.path.join(os.path.dirname(BASE_DIR), "FEN", "Indicadores")
CLIMA_PATTERNS = ["indicadores_franco*.xls*", "*indicadores*franco*.xls*",
                  "*indicadores*clima*.xls*"]
CLIMA_SHEETS = ["ICEN", "MEIv2", "ONI", "NINO1+2", "ERSSTv5", "ITCP", "LABCOS"]
CLIMA_DEFAULT = "ICEN"
# El archivo real tiene hojas con estructuras DISTINTAS. Este mapeo explícito dice,
# para cada indicador del dropdown, de qué hoja y con qué modo de lectura sale:
#   ("Hoja", "col:NOMBRE_COLUMNA")  -> hoja con columna 'fecha' + esa columna de valor
#   ("Hoja", "bimestre")            -> columnas anio + bimestre (código de 2 letras)
#   ("Hoja", "wide")                -> año en filas y temporadas (3 meses) en columnas
# Edita aquí si cambian nombres de hoja/columna (p. ej. NINO1+2 sale de ERSST5/ANOM_1+2).
CLIMA_FUENTE = {
    "ICEN":    ("ICEN", "col:ICEN"),
    "ITCP":    ("ITCP", "col:ITCP"),
    "LABCOS":  ("LABCOS", "col:LABCOS"),
    "NINO1+2": ("ERSST5", "col:ANOM_1+2"),   # anomalía Niño 1+2 (costa Perú)
    "ERSSTv5": ("ERSST5", "col:ANOM_3.4"),   # anomalía Niño 3.4 (ENSO, base ERSSTv5)
    "MEIv2":   ("MEIv2", "bimestre"),
    "ONI":     ("ONI", "wide"),
}
# Bimestre MEI (2 letras) -> mes representativo (segundo mes del par).
CLIMA_BIMESTRE_MES = {"DJ": 1, "JF": 2, "FM": 3, "MA": 4, "AM": 5, "MJ": 6,
                      "JJ": 7, "JA": 8, "AS": 9, "SO": 10, "ON": 11, "ND": 12}
# Temporada trimestral ONI/ERSST -> mes central.
CLIMA_TEMPORADA_MES = {"DJF": 1, "JFM": 2, "FMA": 3, "MAM": 4, "AMJ": 5, "MJJ": 6,
                       "JJA": 7, "JAS": 8, "ASO": 9, "SON": 10, "OND": 11, "NDJ": 12}
# Umbrales de categoría (líneas horizontales) por indicador. Editables.
CLIMA_UMBRALES = {
    "ICEN": [("El Niño fuerte", 1.4, "#c0392b"), ("El Niño moderado", 1.0, "#e67e22"),
             ("El Niño débil", 0.4, "#f0b27a"), ("La Niña débil", -1.0, "#85c1e9"),
             ("La Niña moderada", -1.2, "#5dade2"), ("La Niña fuerte", -1.4, "#2471a3")],
    "ONI": [("El Niño fuerte", 1.5, "#c0392b"), ("El Niño moderado", 1.0, "#e67e22"),
            ("El Niño débil", 0.5, "#f0b27a"), ("La Niña débil", -0.5, "#85c1e9"),
            ("La Niña moderada", -1.0, "#5dade2"), ("La Niña fuerte", -1.5, "#2471a3")],
    "ERSSTv5": [("El Niño fuerte", 1.5, "#c0392b"), ("El Niño moderado", 1.0, "#e67e22"),
                ("El Niño débil", 0.5, "#f0b27a"), ("La Niña débil", -0.5, "#85c1e9"),
                ("La Niña moderada", -1.0, "#5dade2"), ("La Niña fuerte", -1.5, "#2471a3")],
    "NINO1+2": [("El Niño fuerte", 1.5, "#c0392b"), ("El Niño moderado", 1.0, "#e67e22"),
                ("El Niño débil", 0.5, "#f0b27a"), ("La Niña débil", -0.5, "#85c1e9"),
                ("La Niña fuerte", -1.5, "#2471a3")],
}
# --- Indicadores financieros SBS ---
SBS_DIR = FEN_IND_DIR
SBS_PATTERNS = ["BD SBS*.xls*", "BD_SBS*.xls*", "*SBS*.xls*"]
SBS_SHEET_HINTS = ["indicadores SBS", "SBS", "Indicadores"]
SBS_IND_DEFAULT = "Créditos Refinanciados y Reestructurados / Créditos Directos"
# Indicadores SBS bien MANTENIDOS en el tiempo (cobertura continua para los bancos
# de interés). Verificado: Refinanciados/Directos y ROEA/ROAA son continuos. El
# clásico "Vencidos y en Cobranza Judicial / Colocaciones Brutas" se descontinuó en
# 2000, por eso NO se usa de default. Amplía si necesitas otros.
SBS_IND_WHITELIST = ["Créditos Refinanciados y Reestructurados / Créditos Directos",
                     "ROEA", "ROAA"]
# Fusión de indicadores que la SBS RENOMBRÓ con el tiempo (mismo concepto) en una
# sola serie continua. Clave = nombre canónico mostrado; valor = variantes a unir
# (se homologan con norm_fuerte, así que basta una redacción aproximada de cada una).
SBS_FUSION = {
    "Morosidad: Créditos Atrasados / Créditos Directos": [
        "Cartera Atrasada / Créditos Directos",
        "Créditos Atrasados (criterio SBS) / Créditos Directos",
        "Créditos Atrasados / Créditos Directos"],
    "Provisiones / Créditos Atrasados": [
        "Provisiones / Cartera Atrasada",
        "Provisiones / Créditos Atrasados"],
}
# Homologación de bancos (col Empresa de SBS) -> nombre canónico del portafolio.
EQUIV_BANCOS_SBS = {
    "BCP": "Credicorp", "BANCO DE CREDITO": "Credicorp",
    "BANCO DE CREDITO DEL PERU": "Credicorp", "MIBANCO": "Credicorp",
    "INTERBANK": "Intercorp Perú", "BANCO INTERNACIONAL DEL PERU": "Intercorp Perú",
    "BBVA": "BBVA Perú", "BBVA CONTINENTAL": "BBVA Perú",
    "BANCO CONTINENTAL": "BBVA Perú", "CONTINENTAL": "BBVA Perú",
}
BANCOS_DEFAULT = ["Credicorp", "BBVA Perú", "Intercorp Perú"]
# Pistas para avisar de bancos "de interés" presentes en SBS pero NO mapeados.
# Solo los relevantes para el portafolio (evita ruido con bancos sin exposición).
PISTAS_BANCO_SBS = ["BCP", "BANCO DE CREDITO", "INTERBANK", "BBVA", "CONTINENTAL",
                    "MIBANCO"]
# --- Producción minera mensual (serie por empresa para el gráfico Geofísicos) ---
# Fuente ÚNICA del gráfico: hoja "BD Prod" de BD Mineria.xlsm (en MACRO_DIR).
# Estructura: fila 1 = headers + fechas mensuales (desde col G); col B MINERAL,
# col C ETAPA, col E TITULAR (empresa), col F UNIDAD. Suma por empresa+mes con
# ETAPA=CONCENTRACIÓN, sin distinguir mineral.
PROD_MENSUAL_DIR = MACRO_DIR
PROD_MENSUAL_PATTERNS = ["BD Mineria*.xls*", "BD_Mineria*.xls*", "*Mineria*.xls*"]
PROD_MENSUAL_SHEET_HINTS = ["BD Prod", "BD PROD", "Prod"]
# Procesos/etapas válidos (col C ETAPA). Solo CONCENTRACIÓN (norm() ignora la tilde).
PROCESOS_VALIDOS = {"CONCENTRACION"}
EQUIV_MINERAL = {"ESTANO": "ESTAÑO"}
# Homologación de empresas mineras en BD Prod (col E TITULAR) -> canónico.
EQUIV_MINERAS_MENSUAL = {
    "NEXA ATACOCHA": "Nexa Resources Perú",
    "NEXA CAJAMARQUILLA": "Nexa Resources Perú",
    "NEXA EL PORVENIR": "Nexa Resources Perú",
    "NEXA PERU": "Nexa Resources Perú",
    "BUENAVENTURA": "Buenaventura",
    "CERRO VERDE": "Cerro Verde",
    "VOLCAN": "Volcan",
    "ARES": "Hochschild",
    "MINSUR": "Minsur",
}
EMPRESAS_MIN_DEFAULT = ["Buenaventura", "Volcan", "Minsur"]
# Meses (español) para sumar el semestre anterior en el archivo anual de producción.
MESES_ES = {1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
            7: "JULIO", 8: "AGOSTO", 9: "SETIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE",
            12: "DICIEMBRE"}
# Eventos extraordinarios de El Niño (bandas verticales). Fechas oficiales:
#  - El Niño extraordinario 1997–1998: fase cálida fuerte reconocida por ENFEN /
#    Comisión Multisectorial ENFEN; pico mar-1997 a may-1998.
#  - El Niño Costero 2017: declarado por ENFEN; ene-2017 a may-2017 (Costero).
EVENTOS_ELNINO = [
    {"nombre": "El Niño extraordinario 1997–1998", "ini": "1997-03", "fin": "1998-06",
     "fuente": "ENFEN / Comisión Multisectorial ENFEN"},
    {"nombre": "El Niño Costero 2017", "ini": "2017-01", "fin": "2017-05",
     "fuente": "ENFEN (Comunicado Oficial El Niño Costero 2017)"},
]


# ==============================================================================
# 2) UTILIDADES
# ==============================================================================

def log(msg, level="INFO"):
    print(f"[{level}] {msg}")


def fatal(msg):
    print("\n" + "=" * 78)
    print("ERROR — el proceso se detuvo. Revisa el detalle de abajo:")
    print("=" * 78)
    print(msg)
    print("=" * 78)
    sys.exit(1)


def norm(s):
    """Normaliza texto: mayúsculas, sin acentos, solo [A-Z0-9 ], colapsa espacios."""
    if s is None:
        return ""
    s = str(s).strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def norm_fuerte(s):
    """Homologación agresiva: solo alfanuméricos (ignora espacios, '/', '*', '(',
    ')', tildes, mayúsculas y sufijos como (1)/(2)/(*)). Sirve para reconocer como
    iguales nombres de bancos/indicadores SBS con variaciones menores de escritura."""
    if s is None:
        return ""
    s = "".join(c for c in unicodedata.normalize("NFKD", str(s).upper())
                if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]", "", s)


def norm_unidad(s):
    """Normaliza la unidad de medida: 'Grs. f.' -> 'grsf', 'Kg.f.' -> 'kgf', 'TMF' -> 'tmf'."""
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", s)


def find_file(folder, patterns, descripcion, exclude=None):
    if not os.path.isdir(folder):
        fatal(f"No existe la carpeta esperada para {descripcion}:\n  {folder}\n"
              f"-> Verifica la ruta en CONFIG (BASE_DIR) o crea/renombra la carpeta.")
    hits = []
    for pat in patterns:
        hits += glob.glob(os.path.join(folder, pat))
    # excluir temporales de Excel (~$...)
    hits = [h for h in sorted(set(hits)) if not os.path.basename(h).startswith("~$")]
    # excluir los archivos que el propio script genera (evita autoconfusión cuando
    # un output cae en la misma carpeta y comparte prefijo de nombre con un input)
    excl = (exclude or []) + EXCLUDE_OUTPUTS
    if excl:
        hits = [h for h in hits
                if not any(norm(s) in norm(os.path.basename(h)) for s in excl)]
    if not hits:
        fatal(f"No se encontró ningún archivo de {descripcion} en:\n  {folder}\n"
              f"Patrones buscados: {patterns}\n"
              f"-> Confirma que el archivo esté en esa carpeta. El NOMBRE puede "
              f"variar, pero debe contener las palabras del patrón.\n"
              f"   (Se ignoran los archivos de salida del propio script: {EXCLUDE_OUTPUTS}.)")
    hits.sort(key=os.path.getmtime, reverse=True)
    if len(hits) > 1:
        log(f"Se encontraron {len(hits)} candidatos para {descripcion}; se usa el "
            f"más reciente: {os.path.basename(hits[0])}", "WARN")
    return hits[0]


def pick_sheet(xls_path, hints, descripcion):
    xl = pd.ExcelFile(xls_path)
    sheets = xl.sheet_names
    for h in hints:
        for s in sheets:
            if norm(h) in norm(s):
                return s
    fatal(f"No se encontró la hoja de {descripcion} en {os.path.basename(xls_path)}.\n"
          f"Hojas disponibles: {sheets}\n"
          f"Pistas buscadas: {hints}\n"
          f"-> Renombra la hoja o agrega su nombre a los *_HINTS en CONFIG.")


def detect_header_row(df_raw, required_tokens, max_scan=20):
    """Busca la fila de encabezados: la que contiene más de los tokens esperados."""
    req = [norm(t) for t in required_tokens]
    best_i, best_score = None, -1
    for i in range(min(max_scan, len(df_raw))):
        cells = [norm(x) for x in df_raw.iloc[i].tolist()]
        score = sum(1 for t in req if any(t in c or c in t for c in cells if c))
        if score > best_score:
            best_i, best_score = i, score
    return best_i, best_score


def resolve_columns(columns, spec):
    """
    spec: {logico: ([substrings de encabezado], letra_respaldo_opcional)}
    Devuelve {logico: nombre_real_de_columna}. Lanza fatal si falta alguno.
    """
    norm_cols = {col: norm(col) for col in columns}
    out = {}
    faltan = []
    for logico, (subs, letra) in spec.items():
        found = None
        for col, ncol in norm_cols.items():
            if any(norm(sub) in ncol for sub in subs):
                found = col
                break
        if found is None and letra:  # respaldo por letra de columna
            idx = column_letter_to_index(letra)
            if idx < len(columns):
                found = columns[idx]
                log(f"Columna '{logico}' no se halló por nombre; se usa respaldo por "
                    f"letra {letra} -> '{found}'.", "WARN")
        if found is None:
            faltan.append((logico, subs, letra))
        else:
            out[logico] = found
    if faltan:
        det = "\n".join(f"   - {lg}: buscaba encabezados {subs} (respaldo letra {lt})"
                        for lg, subs, lt in faltan)
        fatal("No se pudieron ubicar estas columnas en el Excel:\n" + det +
              f"\n\nColumnas detectadas en el archivo:\n   {list(columns)}\n"
              "-> Abre el Excel y revisa que esos encabezados existan (o ajusta los "
              "'subs' en la función correspondiente del script).")
    return out


def column_letter_to_index(letter):
    letter = letter.upper()
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


# ==============================================================================
# 3) CARGA DEL GEOJSON Y CONSTRUCCIÓN DEL ÍNDICE NOMBRE -> UBIGEO
# ==============================================================================

def _merge_suplemento(base):
    """Fusiona en 'base' los distritos del suplemento externo (si existe) y del
    embebido. El suplemento tiene prioridad por IDDIST (reemplaza geometría)."""
    feats = base.get("features", [])
    by_u = {}
    order = []
    for f in feats:
        u = str(f.get("properties", {}).get("IDDIST", "")).zfill(6)
        by_u[u] = f; order.append(u)
    added, replaced = 0, 0
    fuentes = []
    # externo (prioridad máxima)
    if os.path.isfile(GEOJSON_SUPLEMENTO_FILE):
        try:
            with open(GEOJSON_SUPLEMENTO_FILE, "r", encoding="utf-8") as fh:
                fuentes.append(json.load(fh))
        except Exception as e:
            log(f"No se pudo leer el suplemento externo: {e}", "WARN")
    # embebido
    try:
        fuentes.append(json.loads(GEOJSON_SUPLEMENTO_BUILTIN))
    except Exception:
        pass
    for sup in fuentes:
        for f in sup.get("features", []):
            u = str(f.get("properties", {}).get("IDDIST", "")).zfill(6)
            if not u or u == "000000":
                continue
            if u in by_u:
                replaced += 1
            else:
                added += 1; order.append(u)
            by_u[u] = f
    if added or replaced:
        log(f"Suplemento GeoJSON aplicado: {added} distrito(s) añadido(s), "
            f"{replaced} reemplazado(s).")
    base["features"] = [by_u[u] for u in dict.fromkeys(order)]
    return base


def load_geojson():
    if os.path.isfile(GEOJSON_LOCAL):
        log(f"GeoJSON local: {GEOJSON_LOCAL}")
        with open(GEOJSON_LOCAL, "r", encoding="utf-8") as f:
            return _merge_suplemento(json.load(f))
    # descargar
    log("No hay GeoJSON local; intentando descargar...", "WARN")
    try:
        import urllib.request
        os.makedirs(os.path.dirname(GEOJSON_LOCAL), exist_ok=True)
        with urllib.request.urlopen(GEOJSON_URL, timeout=90) as r:
            data = json.loads(r.read().decode("utf-8"))
        with open(GEOJSON_LOCAL, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        log(f"GeoJSON descargado y guardado en {GEOJSON_LOCAL}")
        return _merge_suplemento(data)
    except Exception as e:
        fatal(f"No se pudo cargar ni descargar el GeoJSON distrital.\n"
              f"Detalle: {e}\n"
              f"-> Descarga manualmente este archivo y guárdalo como:\n"
              f"   {GEOJSON_LOCAL}\n"
              f"   Fuente: {GEOJSON_URL}")


def build_geo_index(geojson):
    """Devuelve:
       name2u : {(dep,prov,dist)->ubigeo}, dep_dist2u : {(dep,dist)->[ubigeos]},
       u2name : {ubigeo->(dep,prov,dist)} con nombres ORIGINALES."""
    name2u, dep_dist2u, u2name = {}, {}, {}
    for f in geojson["features"]:
        p = f.get("properties", {})
        u = p.get("IDDIST")
        if not u:
            continue
        u = str(u).strip().zfill(6)
        dep, prov, dist = p.get("NOMBDEP"), p.get("NOMBPROV"), p.get("NOMBDIST")
        name2u[(norm(dep), norm(prov), norm(dist))] = u
        dep_dist2u.setdefault((norm(dep), norm(dist)), []).append(u)
        u2name[u] = (dep, prov, dist)
    return name2u, dep_dist2u, u2name


def _variantes_distrito(dist):
    """Devuelve nombres alternativos a probar para un distrito: el original, el
    contenido fuera y dentro de paréntesis, y sus alias conocidos."""
    if dist is None:
        return []
    base = str(dist).strip()
    cands = [base]
    # "Pueblo Libre (Magdalena Vieja)" -> "Pueblo Libre" + "Magdalena Vieja"
    m = re.match(r"^(.*?)\s*\((.*?)\)\s*$", base)
    if m:
        cands += [m.group(1).strip(), m.group(2).strip()]
    # alias (popular <-> oficial)
    for c in list(cands):
        a = ALIAS_DISTRITO.get(norm(c))
        if a:
            cands.append(a)
    # dedup conservando orden
    seen, out = set(), []
    for c in cands:
        if c and norm(c) not in seen:
            seen.add(norm(c)); out.append(c)
    return out


def geocode(dep, prov, dist, idx):
    """Resuelve UBIGEO: 1) extranjero 2) override 3) (dep,prov,dist) 4) (dep,dist)
       único, probando variantes de nombre (paréntesis/alias).
       Devuelve (ubigeo, metodo) o (None, motivo)."""
    name2u, dep_dist2u, _ = idx
    # 1) zonas no mapeables (extranjero, etc.)
    if norm(dep) in DEPTOS_NO_MAPEABLES or norm(dist) in DEPTOS_NO_MAPEABLES:
        return None, "extranjero"
    variantes = _variantes_distrito(dist)
    # 2) override (sobre el nombre tal cual y sus variantes)
    for dv in variantes:
        k3 = (norm(dep), norm(prov), norm(dv))
        if k3 in OVERRIDE_UBIGEO:
            return OVERRIDE_UBIGEO[k3], "override"
    # 3) match exacto (dep,prov,dist) con variantes
    for dv in variantes:
        k3 = (norm(dep), norm(prov), norm(dv))
        if k3 in name2u:
            return name2u[k3], "exacto"
    if norm(prov) in ("", "-") or norm(dist) in ("", "-"):
        return None, "sin_distrito"
    # 4) respaldo (dep,dist) único, con variantes
    for dv in variantes:
        cands = dep_dist2u.get((norm(dep), norm(dv)), [])
        if len(cands) == 1:
            return cands[0], "dep+distrito"
        if len(cands) > 1:
            return None, "ambiguo"
    return None, "no_encontrado"


# ==============================================================================
# 4) LECTURA DE PRODUCCIÓN MINERA
# ==============================================================================

def load_produccion(path):
    xl = pd.ExcelFile(path)
    sheet = xl.sheet_names[0]
    if len(xl.sheet_names) > 1:
        # preferir una hoja que parezca de "informacion general / produccion"
        for s in xl.sheet_names:
            if norm("INFORMACION") in norm(s) or norm("PRODUCC") in norm(s):
                sheet = s
                break
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    hdr_row, score = detect_header_row(
        raw, ["MINERAL", "UNIDAD DE MEDIDA", "TITULAR", "DEPARTAMENTO",
              "DISTRITO", "ACUM"])
    if score < 4:
        fatal(f"No se reconoció la fila de encabezados en la hoja '{sheet}' de "
              f"producción (coincidencias={score}).\n"
              f"-> Confirma que existan encabezados como MINERAL, UNIDAD DE MEDIDA, "
              f"TITULAR, DEPARTAMENTO, PROVINCIA, DISTRITO y un acumulado anual (ACUM...).")
    df = pd.read_excel(path, sheet_name=sheet, header=hdr_row)
    df = df.dropna(axis=1, how="all")

    spec = {
        "mineral":  (["MINERAL"], "A"),
        "unidad":   (["UNIDAD DE MEDIDA", "UNIDAD"], "B"),
        "titular":  (["TITULAR"], "F"),
        "unidad_minera": (["UNIDAD MINERA"], "G"),
        "depto":    (["DEPARTAMENTO"], "H"),
        "prov":     (["PROVINCIA"], "I"),
        "dist":     (["DISTRITO"], "J"),
        "acum":     (["ACUM"], "W"),
    }
    cols = resolve_columns(df.columns, spec)
    df = df.rename(columns={v: k for k, v in cols.items()})

    # filtrar pies de página / filas sin titular o sin departamento
    df = df[df["titular"].notna() & df["depto"].notna()].copy()

    # PESO = suma de los últimos 6 meses del SEMESTRE ANTERIOR (consistente con
    # Utilities/Construcción). El archivo trae columnas por mes (ENERO..DICIEMBRE);
    # se suman las del semestre previo. Si faltan, se avisa y se usa lo disponible
    # (o el acumulado anual 'ACUM' como última red).
    meses_obj = sorted(set(m for _, m in _meses_semestre_anterior()))
    nombres_obj = [MESES_ES[m] for m in meses_obj]
    col_por_nombre = {norm(c): c for c in df.columns}
    # SETIEMBRE/SEPTIEMBRE: aceptar ambas grafías
    alias = {"SETIEMBRE": "SEPTIEMBRE", "SEPTIEMBRE": "SETIEMBRE"}
    mcols, faltan = [], []
    for nm in nombres_obj:
        c = col_por_nombre.get(norm(nm)) or col_por_nombre.get(norm(alias.get(nm, "")))
        if c is not None:
            mcols.append(c)
        else:
            faltan.append(nm)
    if mcols:
        df["acum"] = df[mcols].apply(lambda s: pd.to_numeric(s, errors="coerce")).sum(axis=1)
        etq = f"{nombres_obj[0].capitalize()}–{nombres_obj[-1].capitalize()}"
        log(f"Producción minera: peso = suma del semestre anterior ({etq}); "
            f"{len(mcols)} mes(es).")
        if faltan:
            log(f"      Meses del semestre sin columna en el Excel anual: "
                f"{', '.join(faltan)}. Se sumó solo lo disponible.", "WARN")
    else:
        df["acum"] = pd.to_numeric(df["acum"], errors="coerce")
        log("Producción minera: no se hallaron columnas del semestre anterior; se usó "
            "la columna ACUM (acumulado anual) como respaldo. Revisa los encabezados "
            "de meses del Excel.", "WARN")
    df["acum"] = df["acum"].fillna(0.0)

    # unidad -> factor TMF
    df["unidad_norm"] = df["unidad"].apply(norm_unidad)
    desconocidas = sorted(set(df.loc[~df["unidad_norm"].isin(UNIDAD_A_TMF), "unidad"]
                              .dropna().astype(str)))
    if desconocidas:
        log(f"Unidades de medida no reconocidas (se excluyen del TMF): {desconocidas}\n"
            f"      -> Si corresponden, agrégalas a UNIDAD_A_TMF en CONFIG.", "WARN")
    df["factor"] = df["unidad_norm"].map(UNIDAD_A_TMF)
    df["tmf"] = df["acum"] * df["factor"]

    # empresa canónica (solo universo de interés)
    eq_norm = {norm(k): v for k, v in EQUIV_EMPRESAS_PRODUCCION.items()}
    df["empresa"] = df["titular"].apply(lambda t: eq_norm.get(norm(t)))
    df_univ = df[df["empresa"].notna()].copy()
    log(f"Producción: {len(df)} filas válidas; {len(df_univ)} de las empresas del "
        f"universo ({df_univ['empresa'].nunique()} empresas).")
    return df_univ, sheet


# ==============================================================================
# 4-bis) LECTURA DEL RESUMEN DE BANCOS (CRÉDITOS POR DISTRITO)
# ==============================================================================

def _resolver_columnas_por_nombre(df_cols, fin_cols, descripcion):
    """Empareja los nombres ESPERADOS (robustos a mayúsculas/tildes) con las
    columnas reales del DataFrame. Devuelve {clave: nombre_real}. Si falta alguna,
    aborta con un mensaje claro de qué columna falta y cómo arreglarlo."""
    norm_real = {norm(c): c for c in df_cols}
    resol = {}
    faltan = []
    for clave, alias in fin_cols.items():
        encontrado = None
        for a in alias:
            if norm(a) in norm_real:
                encontrado = norm_real[norm(a)]
                break
        if encontrado is None:
            faltan.append((clave, alias))
        else:
            resol[clave] = encontrado
    if faltan:
        detalle = "\n".join(
            f"      · '{clave}': se buscó alguno de {alias}" for clave, alias in faltan)
        fatal(f"En {descripcion} no se encontraron estas columnas:\n{detalle}\n"
              f"   Columnas presentes en el archivo: {list(df_cols)}\n"
              f"-> Ajusta los nombres esperados en FIN_COLS dentro de CONFIG, o "
              f"corrige el encabezado del archivo. (No importan mayúsculas ni tildes.)")
    return resol


def load_financiero(path):
    """Lee el resumen de bancos (Empresa, Departamento, Provincia, Distrito,
    Total_Creditos_Suma) y lo devuelve con la MISMA forma que la producción minera
    para reutilizar todo el pipeline: el monto de créditos pasa a la columna 'tmf'
    (el peso de reparto), y se etiqueta como pseudo-mineral 'Créditos'."""
    xl = pd.ExcelFile(path)
    # preferir una hoja que parezca de resumen
    sheet = xl.sheet_names[0]
    for s in xl.sheet_names:
        if norm("RESUMEN") in norm(s):
            sheet = s
            break

    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    hdr_row, _ = detect_header_row(
        raw, ["Empresa", "Departamento", "Provincia", "Distrito"], max_scan=15)
    df = pd.read_excel(path, sheet_name=sheet, header=hdr_row)
    df = df.dropna(axis=1, how="all")

    cols = _resolver_columnas_por_nombre(df.columns, FIN_COLS,
                                         "el resumen de bancos (créditos)")
    df = df.rename(columns={
        cols["empresa"]: "titular", cols["depto"]: "depto",
        cols["prov"]: "prov", cols["dist"]: "dist", cols["monto"]: "acum"})

    df = df[df["titular"].notna() & df["depto"].notna()].copy()
    df["acum"] = pd.to_numeric(df["acum"], errors="coerce").fillna(0.0)

    # --- empresa canónica (robusto a mayúsculas/tildes/espacios) ---
    eq_norm = {norm(k): v for k, v in EQUIV_EMPRESAS_FINANCIERO.items()}
    df["empresa"] = df["titular"].apply(lambda t: eq_norm.get(norm(t)))

    # Aviso de bancos del archivo que NO están mapeados (para que decidas incluirlos)
    no_map = sorted(set(str(t).strip() for t in df.loc[df["empresa"].isna(), "titular"]
                        .dropna()))
    if no_map:
        log(f"Bancos en el resumen NO mapeados al portafolio (se omiten). Si alguno "
            f"debe incluirse, agrégalo a EQUIV_EMPRESAS_FINANCIERO en CONFIG:", "WARN")
        for b in no_map:
            log(f"      · '{b}'", "WARN")

    df_univ = df[df["empresa"].notna()].copy()
    if df_univ.empty:
        log("Ningún banco del resumen coincide con EQUIV_EMPRESAS_FINANCIERO; el "
            "sector financiero quedará vacío. Revisa los nombres en CONFIG.", "WARN")

    # dar forma de 'producción' para reutilizar build_model sin cambios
    df_univ["mineral"] = "Créditos"
    df_univ["unidad"] = "Miles S/"
    df_univ["unidad_minera"] = None
    df_univ["unidad_norm"] = "credmiles"     # etiqueta propia (no entra a UNIDAD_A_TMF)
    df_univ["factor"] = 1.0
    df_univ["tmf"] = df_univ["acum"]          # el crédito ES el peso de reparto

    log(f"Financiero: {len(df)} filas leídas; {len(df_univ)} de bancos del universo "
        f"({df_univ['empresa'].nunique()} bancos).")
    return df_univ[["mineral", "unidad", "titular", "unidad_minera", "depto", "prov",
                    "dist", "acum", "unidad_norm", "factor", "tmf", "empresa"]], sheet


# ==============================================================================
# 5) LECTURA DE CENEPRED
# ==============================================================================

def load_cenepred(path, sheet):
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    # fila de ubigeos = la que tiene más celdas tipo código de 6 dígitos
    ubi_row = None
    for i in range(min(10, len(raw))):
        vals = raw.iloc[i].tolist()[1:]
        cnt = sum(1 for v in vals if re.fullmatch(r"\d{5,6}", str(v).strip().split(".")[0] if v is not None else ""))
        if cnt > 30:
            ubi_row = i
            break
    if ubi_row is None:
        fatal(f"No se encontró la fila de UBIGEOS en la hoja CENEPRED.\n"
              f"-> Debe haber una fila con códigos de distrito de 6 dígitos "
              f"(p.ej. 010201) a lo largo de las columnas.")
    ubigeos = []
    for v in raw.iloc[ubi_row].tolist()[1:]:
        if v is None or str(v).strip() == "":
            ubigeos.append(None)
        else:
            ubigeos.append(str(v).strip().split(".")[0].zfill(6))
    # filas de peligros: las siguientes con etiqueta en la 1a columna
    scores = {}  # ubigeo -> {mm,inu,seq}
    hazard_rows = {}
    for i in range(ubi_row + 1, len(raw)):
        label = raw.iloc[i, 0]
        if label is None or str(label).strip() == "":
            continue
        nl = norm(label)
        for hz in HAZARDS:
            if norm(hz) in nl or nl in norm(hz):
                hazard_rows[hz] = i
    faltan = [h for h in HAZARDS if h not in hazard_rows]
    if faltan:
        fatal(f"En CENEPRED no se hallaron las filas de peligro: {faltan}\n"
              f"-> La primera columna debe contener: {HAZARDS}.")
    for hz, ri in hazard_rows.items():
        rowvals = raw.iloc[ri].tolist()[1:]
        for j, u in enumerate(ubigeos):
            if u is None or j >= len(rowvals):
                continue
            val = rowvals[j]
            try:
                sc = int(float(val))
            except (TypeError, ValueError):
                sc = None
            scores.setdefault(u, {})[HAZARD_KEY[hz]] = sc
    log(f"CENEPRED: {len([u for u in ubigeos if u])} distritos con score.")
    return scores


# ==============================================================================
# 6) LECTURA DE EXPOSICIÓN (VALORIZACIÓN DE INSTRUMENTOS)
# ==============================================================================

def load_exposicion(path, sheet):
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    hdr_row, score = detect_header_row(raw, ["Nombre (G&P)", "VALOR_VECTOR_LOCAL"],
                                       max_scan=10)
    if score < 1:
        fatal(f"No se reconoció la fila de encabezados en '{sheet}'.\n"
              f"-> Debe existir una columna 'Nombre (G&P)' (H) y "
              f"'VALOR_VECTOR_LOCAL' (M).")
    df = pd.read_excel(path, sheet_name=sheet, header=hdr_row)
    df = df.dropna(how="all")
    spec = {
        "nombre": (["Nombre (G&P)", "Nombre G P", "Nombre"], "H"),
        "valor":  (["VALOR_VECTOR_LOCAL", "VALOR VECTOR LOCAL"], "M"),
    }
    cols = resolve_columns(df.columns, spec)
    df = df.rename(columns={cols["nombre"]: "nombre", cols["valor"]: "valor"})
    df["nombre"] = df["nombre"].astype(str).str.strip()
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0.0)

    eq_ex_norm = {norm(k): v for k, v in EQUIV_EMPRESAS_EXPOSICION.items()}
    canon_norm = {norm(c): c for c in EMPRESAS_CANONICAS}

    def to_canon(n):
        nn = norm(n)
        if nn in eq_ex_norm:
            return eq_ex_norm[nn]
        if nn in canon_norm:
            return canon_norm[nn]
        return None

    df["empresa"] = df["nombre"].apply(to_canon)
    exp = (df.dropna(subset=["empresa"]).groupby("empresa")["valor"].sum().to_dict())
    for c in EMPRESAS_CANONICAS:
        exp.setdefault(c, 0.0)
    log(f"Exposición: S/ {sum(exp.values()):,.0f} en {sum(1 for v in exp.values() if v>0)} empresas.")
    return exp


# ==============================================================================
# 7) CONSTRUCCIÓN DEL MODELO (producción por distrito, asignación, riesgo)
# ==============================================================================

def build_model(df_prod, exp, scores, geo_idx, universo=None, sector=SECTOR_MINERIA):
    if universo is None:
        universo = EMPRESAS_CANONICAS
    _, _, u2name = geo_idx
    # geocodificar cada fila
    geocodes, motivos = [], []
    for _, r in df_prod.iterrows():
        u, m = geocode(r["depto"], r["prov"], r["dist"], geo_idx)
        geocodes.append(u)
        motivos.append(m)
    df = df_prod.copy()
    df["ubigeo"] = geocodes
    df["geo_motivo"] = motivos

    no_map = df[df["ubigeo"].isna()].copy()
    if len(no_map):
        # separar "extranjero" (categoría conocida, no es error) del resto
        ext = no_map[no_map["geo_motivo"] == "extranjero"]
        falla = no_map[no_map["geo_motivo"] != "extranjero"]
        peso_lbl = "miles S/" if sector == SECTOR_FINANCIERO else "TMF"
        if len(ext):
            log(f"Excluidas {len(ext)} fila(s) de {sector} en el EXTRANJERO "
                f"(sin riesgo físico en Perú): {ext['tmf'].sum():,.2f} {peso_lbl}.")
        if len(falla):
            resumen = (falla.groupby(["depto", "prov", "dist", "geo_motivo"])
                       .size().reset_index(name="filas"))
            log(f"{len(falla)} fila(s) de {sector} NO geocodificadas "
                f"({falla['tmf'].sum():,.2f} {peso_lbl}). Detalle por motivo:", "WARN")
            for _, rr in resumen.iterrows():
                log(f"      {rr['depto']} / {rr['prov']} / {rr['dist']}  "
                    f"[{rr['geo_motivo']}]  filas={rr['filas']}", "WARN")

    dfm = df[df["ubigeo"].notna()].copy()

    # ---- producción por (empresa, ubigeo) en TMF + detalle por mineral ----
    prod_emp_dist = (dfm.groupby(["empresa", "ubigeo"])["tmf"].sum()
                     .reset_index())
    # unidades mineras por (empresa, ubigeo)
    um = (dfm.groupby(["empresa", "ubigeo"])["unidad_minera"]
          .apply(lambda s: sorted(set(str(x) for x in s.dropna()))).to_dict())
    # detalle por mineral (en unidad original) por (empresa, ubigeo)
    detalle = {}
    for (e, u, mineral, uni_norm), g in dfm.groupby(
            ["empresa", "ubigeo", "mineral", "unidad_norm"]):
        detalle.setdefault((e, u), []).append({
            "mineral": str(mineral),
            "unidad": UNIDAD_LABEL.get(uni_norm, uni_norm),
            "cantidad_original": float(g["acum"].sum()),
            "tmf": float(g["tmf"].sum()),
        })

    # ---- asignación de exposición de cada empresa a sus distritos ----
    df_assign = prod_emp_dist.copy()
    df_assign["exposicion"] = 0.0
    for e in universo:
        sub = df_assign[df_assign["empresa"] == e]
        E = exp.get(e, 0.0)
        if sub.empty or E == 0:
            continue
        if METODO_ASIGNACION == "tmf_share" and sub["tmf"].sum() > 0:
            w = sub["tmf"] / sub["tmf"].sum()
        elif METODO_ASIGNACION == "units_share":
            nun = sub["ubigeo"].map(lambda u: max(1, len(um.get((e, u), []))))
            w = nun / nun.sum()
        else:  # equal o tmf=0
            w = pd.Series(1.0 / len(sub), index=sub.index)
        df_assign.loc[sub.index, "exposicion"] = (E * w).values

    # ---- consolidado por distrito (exposición) ----
    dist_rows = {}
    for _, r in df_assign.iterrows():
        u = r["ubigeo"]
        d = dist_rows.setdefault(u, {"ubigeo": u, "tmf": 0.0, "exposicion": 0.0,
                                     "empresas": {}})
        d["tmf"] += r["tmf"]
        d["exposicion"] += r["exposicion"]
        d["empresas"][r["empresa"]] = d["empresas"].get(r["empresa"], 0.0) + r["exposicion"]

    # ---- desglose por mineral (unidad original) agregado por ubigeo (para tooltip) ----
    min_by_u = {}
    for (e, u), items in detalle.items():
        agg = min_by_u.setdefault(u, {})
        for it in items:
            key = (it["mineral"], it["unidad"])
            a = agg.setdefault(key, {"cantidad_original": 0.0, "tmf": 0.0})
            a["cantidad_original"] += it["cantidad_original"]
            a["tmf"] += it["tmf"]

    # Universo de distritos a publicar: TODOS los que tienen score CENEPRED (y existen
    # en el geojson, para que pinten en el mapa) UNIDOS a los que tienen exposición.
    # Esto permite mostrar el panorama completo de riesgo físico del país y superponer
    # la plata invertida solo donde corresponde.
    ubigeos_score = [u for u in scores if u in u2name]
    todos_u = set(dist_rows) | set(ubigeos_score)
    score_no_geo = [u for u in scores if u not in u2name]
    if score_no_geo:
        log(f"{len(score_no_geo)} distrito(s) con score CENEPRED no existen en el "
            f"geojson y no se pintarán en el mapa: {', '.join(sorted(score_no_geo))}",
            "WARN")

    districts = []
    for u in todos_u:
        dep, prov, dist = u2name.get(u, (None, None, None))
        d = dist_rows.get(u, {"tmf": 0.0, "exposicion": 0.0, "empresas": {}})
        sc = scores.get(u, {})
        mm, inu, seq = sc.get("mm"), sc.get("inu"), sc.get("seq")
        avail = [x for x in (mm, inu, seq) if x is not None]
        comb = round(sum(avail) / len(avail), 2) if avail else None
        minerales = [
            {"mineral": k[0], "unidad": k[1], "sector": sector,
             "cantidad_original": round(v["cantidad_original"], 3),
             "tmf": round(v["tmf"], 4)}
            for k, v in min_by_u.get(u, {}).items()
        ]
        expo = round(d["exposicion"], 2)
        # nivel de riesgo CENEPRED del distrito (para la tabla pivote)
        if NIVEL_CENEPRED_MODO == "max":
            nivel = max(avail) if avail else None
        else:
            nivel = int(round(comb)) if comb is not None else None
        # aportes por emisor en este distrito (monto + activos identificados)
        aportes = []
        for emp_n, monto in d["empresas"].items():
            if monto <= 0:
                continue
            activos = list(um.get((emp_n, u), [])) if isinstance(um, dict) else []
            aportes.append({"emisor": emp_n, "sector": sector,
                            "monto": round(monto, 2), "activos": activos})
        rec = {
            "ubigeo": u, "dep": dep, "prov": prov, "dist": dist,
            "tmf": round(d["tmf"], 4), "exposicion": expo,
            "exp_sec": {sector: expo} if expo > 0 else {},
            "has_exp": expo > 0,
            "has_score": comb is not None,
            "nivel": nivel,
            "score_mm": mm, "score_inu": inu, "score_seq": seq, "score_comb": comb,
            "mar_mm": round(expo * mm, 2) if mm else 0.0,
            "mar_inu": round(expo * inu, 2) if inu else 0.0,
            "mar_seq": round(expo * seq, 2) if seq else 0.0,
            "empresas": {k: round(v, 2) for k, v in d["empresas"].items() if v > 0},
            "minerales": minerales,
            "aportes": aportes,
        }
        districts.append(rec)
    return districts, df_assign, um, detalle, no_map, dfm


# ==============================================================================
# 7-bis) EXPOSICIÓN FINANCIERA + AUDITORÍA + FUSIÓN DE SECTORES
# ==============================================================================

def load_exposicion_sector(path, sheet, equiv, universo, etiqueta, pistas=None):
    """Exposición de cartera (genérica por sector) desde 'Valorización de
    Instrumentos' (col H = nombre, col M = VALOR_VECTOR_LOCAL), usando el mapeo
    editable 'equiv' (nombre col H -> canónico). Imprime auditoría de lo sumado a
    cada grupo y, si se pasan 'pistas', los nombres no mapeados que las contienen."""
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    hdr_row, _ = detect_header_row(raw, ["Nombre (G&P)", "VALOR_VECTOR_LOCAL"],
                                   max_scan=10)
    df = pd.read_excel(path, sheet_name=sheet, header=hdr_row).dropna(how="all")
    spec = {"nombre": (["Nombre (G&P)", "Nombre G P", "Nombre"], "H"),
            "valor":  (["VALOR_VECTOR_LOCAL", "VALOR VECTOR LOCAL"], "M")}
    cols = resolve_columns(df.columns, spec)
    df = df.rename(columns={cols["nombre"]: "nombre", cols["valor"]: "valor"})
    df["nombre"] = df["nombre"].astype(str).str.strip()
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0.0)

    eq_norm = {norm(k): v for k, v in equiv.items()}
    df["empresa"] = df["nombre"].apply(lambda n: eq_norm.get(norm(n)))
    exp = (df.dropna(subset=["empresa"]).groupby("empresa")["valor"].sum().to_dict())
    for c in universo:
        exp.setdefault(c, 0.0)

    log(f"Exposición {etiqueta}: S/ {sum(exp.values()):,.0f} en "
        f"{sum(1 for v in exp.values() if v > 0)} entidad(es).")
    detalle_grupo = (df.dropna(subset=["empresa"])
                     .groupby(["empresa", "nombre"])["valor"].sum())
    for canon in universo:
        if canon in detalle_grupo.index.get_level_values(0):
            partes = detalle_grupo.loc[canon].sort_values(ascending=False)
            comp = "; ".join(f"{nm}: S/ {v:,.0f}" for nm, v in partes.items())
            log(f"      [{canon}] = {comp}")
    if pistas:
        no_map = df[df["empresa"].isna()]
        cand = (no_map[no_map["nombre"].apply(lambda n: any(p in norm(n) for p in pistas))]
                .groupby("nombre")["valor"].sum().sort_values(ascending=False))
        if len(cand):
            log(f"      Nombres en col H que parecen del sector y NO están mapeados "
                f"(revisa la equivalencia de exposición):", "WARN")
            for nm, v in cand.items():
                log(f"        · '{nm}'  (S/ {v:,.0f})", "WARN")
    return exp


def load_exposicion_financiero(path, sheet):
    return load_exposicion_sector(path, sheet, EQUIV_EXPOSICION_FINANCIERO,
                                  EMPRESAS_FIN_CANONICAS, "financiera", PISTAS_BANCO)


# ==============================================================================
# 7-ter) UBICACIÓN DE ACTIVOS (CENTRALES / PLANTAS) Y LECTURA UTILITIES/CONSTRUCCIÓN
# ==============================================================================

def _meses_semestre_anterior(ref=None):
    """Devuelve la lista de (año, mes) del SEMESTRE ANTERIOR completo respecto a
    'ref' (o hoy / SEMESTRE_REF). Hoy en 1er semestre -> Jul..Dic del año previo;
    hoy en 2º semestre -> Ene..Jun del año en curso."""
    import datetime
    if ref is None:
        if SEMESTRE_REF:
            ref = datetime.datetime.strptime(SEMESTRE_REF, "%Y-%m-%d")
        else:
            ref = datetime.datetime.now()
    if ref.month <= 6:
        y, ini = ref.year - 1, 7
    else:
        y, ini = ref.year, 1
    return [(y, m) for m in range(ini, ini + 6)]


def _cols_en_ventana(ws, fila_fecha):
    """Columnas cuyas fechas (en 'fila_fecha') caen en el semestre anterior.
    Devuelve (cols, etiqueta_ventana, meses_faltantes). Reconoce datetime y strings
    tipo fecha. Si no encuentra NINGUNA, hace fallback a los 6 meses más recientes
    disponibles y lo informa (zona de riesgo: estructura/fechas cambiadas)."""
    import datetime
    objetivo = _meses_semestre_anterior()
    obj_set = set(objetivo)
    todas = []   # (date, col)
    for c in range(1, ws.max_column + 1):
        v = ws.cell(fila_fecha, c).value
        d = v if isinstance(v, datetime.datetime) else None
        if d is None and isinstance(v, str):
            for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
                try:
                    d = datetime.datetime.strptime(v.strip().split()[0], fmt); break
                except Exception:
                    pass
        if d is not None:
            todas.append((d, c))
    cols = [c for d, c in todas if (d.year, d.month) in obj_set]
    presentes = {(d.year, d.month) for d, c in todas if (d.year, d.month) in obj_set}
    faltantes = [ym for ym in objetivo if ym not in presentes]
    etq = f"{objetivo[0][0]}-{objetivo[0][1]:02d} a {objetivo[-1][0]}-{objetivo[-1][1]:02d}"
    if not cols:
        # fallback: 6 meses más recientes disponibles
        todas.sort(reverse=True)
        cols = [c for d, c in todas[:6]]
        return cols, etq, faltantes, True   # True = se usó fallback
    return cols, etq, faltantes, False


def _suma_ventana(ws, fila, cols):
    return sum(ws.cell(fila, c).value for c in cols
              if isinstance(ws.cell(fila, c).value, (int, float)))


def load_ubicaciones(risk_path, sheet_hints, fallback, descripcion):
    """Lee una hoja de cruce activo->(Dep,Prov,Distrito) del archivo de Riesgos.
    Devuelve ({norm(activo): (dep,prov,dist)}, nombre_hoja|None). Si no encuentra
    la hoja, usa el diccionario 'fallback' del CONFIG (sin abortar)."""
    xl = pd.ExcelFile(risk_path)
    sheet = None
    for h in sheet_hints:
        for s in xl.sheet_names:
            if norm(h) in norm(s):
                sheet = s; break
        if sheet:
            break
    ubic = {}
    if sheet is not None:
        df = pd.read_excel(risk_path, sheet_name=sheet)
        df = df.dropna(how="all")
        # columnas por nombre robusto
        cols = _resolver_columnas_por_nombre(
            df.columns,
            {"activo": ["Central", "Planta", "Activo", "Unidad"],
             "depto": ["Departamento", "Depto"], "prov": ["Provincia"],
             "dist": ["Distrito"]}, f"la hoja '{sheet}'")
        for _, r in df.iterrows():
            a = r[cols["activo"]]
            if pd.isna(a):
                continue
            ubic.setdefault(norm(a), (r[cols["depto"]], r[cols["prov"]], r[cols["dist"]]))
        log(f"Ubicaciones '{sheet}': {len(ubic)} activos cargados.")
    # añadir fallback solo para los que falten
    for k, v in (fallback or {}).items():
        ubic.setdefault(norm(k), v)
    return ubic, sheet


def _df_actividad(filas):
    """Construye el DataFrame con la forma de producción a partir de filas dict."""
    df = pd.DataFrame(filas, columns=["mineral", "unidad", "titular", "unidad_minera",
                                      "depto", "prov", "dist", "acum", "unidad_norm",
                                      "factor", "tmf", "empresa"])
    return df


def load_utilities(util_path, risk_path):
    """Lee 'BD ptf': suma la POTENCIA EFECTIVA (MW) de cada central sobre el
    SEMESTRE ANTERIOR (6 meses) y ubica cada central con la hoja 'Centrales
    Electricas'. Sumar el semestre maneja altas/bajas: una central de baja aporta
    los meses que estuvo activa; una nueva del semestre actual aporta 0."""
    import openpyxl
    wb = openpyxl.load_workbook(util_path, data_only=True)
    sheet = None
    for s in wb.sheetnames:
        if any(norm(h) == norm(s) or norm(h) in norm(s) for h in UTIL_SHEET_HINTS):
            sheet = s; break
    if sheet is None:
        sheet = wb.sheetnames[0]
    ws = wb[sheet]

    eq_norm = {norm(k): v for k, v in EQUIV_EMPRESAS_UTILITIES.items()}
    # filas POTENCIA EFECTIVA
    filas_pe = [r for r in range(2, ws.max_row + 1)
                if norm(ws.cell(r, 2).value) == norm(UTIL_TIPO)]
    if not filas_pe:
        log(f"Utilities: no se hallaron filas '{UTIL_TIPO}' en la columna B (Tipo). "
            f"Revisa UTIL_TIPO en CONFIG o la hoja '{sheet}'.", "WARN")
    cols, etq, faltan, fb = _cols_en_ventana(ws, 1)
    log(f"Utilities: sumando potencia efectiva del semestre anterior ({etq}); "
        f"{len(cols)} columna(s).")
    if faltan:
        log(f"      Meses del semestre sin columna en 'BD ptf': "
            f"{', '.join(f'{y}-{m:02d}' for y,m in faltan)}. "
            f"{'Se usaron los 6 meses más recientes disponibles.' if fb else ''} "
            f"Revisa la fila 1 de fechas si el periodo no es el esperado.", "WARN")

    ubic, ce_sheet = load_ubicaciones(risk_path, CENTRALES_SHEET_HINTS, None,
                                      "Centrales Eléctricas")
    filas, sin_ubic, emp_prev = [], [], None
    for r in filas_pe:
        emp_raw = ws.cell(r, 3).value or emp_prev   # celdas combinadas -> arrastrar
        emp_prev = emp_raw
        canon = eq_norm.get(norm(emp_raw))
        if canon is None:
            continue
        central = ws.cell(r, 4).value
        mw = _suma_ventana(ws, r, cols)
        if not isinstance(mw, (int, float)) or mw <= 0:
            continue   # central de baja / sin dato en el periodo
        loc = ubic.get(norm(central))
        if loc is None:
            sin_ubic.append((canon, central, mw)); continue
        dep, prov, dist = loc
        filas.append(["Potencia efectiva", "MW", central, central, dep, prov, dist,
                      mw, "mw", 1.0, mw, canon])

    if sin_ubic:
        log(f"{len(sin_ubic)} central(es) de generación SIN ubicación en la hoja "
            f"'{ce_sheet or 'Centrales Electricas'}'. Agrégalas ahí "
            f"(columnas Central/Empresa/Departamento/Provincia/Distrito):", "WARN")
        for e, c, mw in sin_ubic:
            log(f"      · {e} | '{c}' | {mw:,.2f} MW  -> falta su fila en la hoja", "WARN")

    df = _df_actividad(filas)
    log(f"Utilities: {len(df)} central(es) ubicadas de "
        f"{df['empresa'].nunique()} empresa(s).")
    return df, sheet


def _parse_planta_construccion(valor):
    """'Unacem - Atocongo' -> ('Unacem', ['Atocongo']).
    'Pacasmayo - Pacasmayo + Piura' -> ('Pacasmayo', ['Pacasmayo','Piura']).
    Devuelve (None, None) si el valor NO tiene formato 'Empresa - Planta'
    (así se descartan totales del bloque 'Resumen' como 'Unacem' o 'Total')."""
    s = str(valor)
    if " - " in s:
        emp, resto = s.split(" - ", 1)
    elif " – " in s:              # guion largo por si acaso
        emp, resto = s.split(" – ", 1)
    else:
        return None, None
    emp = emp.strip()
    plantas = [p.strip() for p in resto.split("+") if p.strip()]
    return emp, (plantas or [resto.strip()])


def load_construccion(cons_path, risk_path):
    """Lee 'Despachos': suma los despachos de cemento (MT) de cada planta sobre el
    SEMESTRE ANTERIOR (6 meses) y ubica cada planta con la hoja de ubicación
    (Centrales Electricas / Plantas Cemento). Una fila que agrega dos plantas
    ('A + B') se reparte 50/50 según MAPEO_PLANTAS_DESPACHOS."""
    import openpyxl
    wb = openpyxl.load_workbook(cons_path, data_only=True)
    sheet = None
    for s in wb.sheetnames:
        if any(norm(h) in norm(s) for h in CONS_SHEET_HINTS):
            sheet = s; break
    if sheet is None:
        sheet = wb.sheetnames[0]
    ws = wb[sheet]

    eq_norm = {norm(k): v for k, v in EQUIV_EMPRESAS_CONSTRUCCION.items()}
    map_norm = {norm(k): v for k, v in MAPEO_PLANTAS_DESPACHOS.items()}
    # detectar fila de fechas (la que tenga más datetimes en las primeras 8 filas)
    import datetime
    fila_fecha, best = 5, -1
    for rr in range(1, 9):
        n = sum(1 for c in range(1, ws.max_column + 1)
                if isinstance(ws.cell(rr, c).value, datetime.datetime))
        if n > best:
            best, fila_fecha = n, rr
    # filas de plantas: col C con formato 'Empresa - Planta' cuya empresa es del universo
    filas_planta = []
    for r in range(fila_fecha + 1, ws.max_row + 1):
        c3 = ws.cell(r, 3).value
        if c3 is None:
            continue
        emp, _ = _parse_planta_construccion(c3)
        if emp and norm(emp) in eq_norm:
            filas_planta.append(r)
    cols, etq, faltan, fb = _cols_en_ventana(ws, fila_fecha)
    log(f"Construcción: sumando despachos del semestre anterior ({etq}); "
        f"{len(cols)} columna(s).")
    if faltan:
        log(f"      Meses del semestre sin columna en 'Despachos': "
            f"{', '.join(f'{y}-{m:02d}' for y,m in faltan)}. "
            f"{'Se usaron los 6 meses más recientes disponibles.' if fb else ''} "
            f"Revisa la fila de fechas si el periodo no es el esperado.", "WARN")

    ubic, pl_sheet = load_ubicaciones(risk_path, PLANTAS_SHEET_HINTS,
                                      UBIC_PLANTAS_CEMENTO, "Plantas de Cemento")
    if pl_sheet is None:
        log("No se encontró hoja de ubicación de plantas; se usan las ubicaciones por "
            "defecto de UBIC_PLANTAS_CEMENTO (CONFIG). Agrega las plantas como filas en "
            "'Centrales Electricas' (Central=nombre de planta, +Empresa/Dep/Prov/"
            "Distrito) o edita ese diccionario.", "WARN")

    filas, sin_ubic = [], []
    for r in filas_planta:
        desc = str(ws.cell(r, 3).value).strip()
        emp_raw, _ = _parse_planta_construccion(desc)
        canon = eq_norm.get(norm(emp_raw))
        mt = _suma_ventana(ws, r, cols)
        if not isinstance(mt, (int, float)) or mt <= 0:
            continue
        # nombre(s) de planta para ubicar: descriptor completo, o desglose 'A + B'
        destinos = map_norm.get(norm(desc), [desc])
        cuota = mt / len(destinos)       # 'A + B' -> 50/50
        for p in destinos:
            loc = ubic.get(norm(p))
            if loc is None:
                sin_ubic.append((canon, p, cuota)); continue
            dep, prov, dist = loc
            filas.append(["Despachos cemento", "MT", p, p,
                          dep, prov, dist, cuota, "mt", 1.0, cuota, canon])

    if sin_ubic:
        dest = pl_sheet or "Centrales Electricas (o UBIC_PLANTAS_CEMENTO en CONFIG)"
        log(f"{len(sin_ubic)} planta(s) de cemento SIN ubicación. Agrégalas en "
            f"'{dest}' con ese nombre EXACTO de planta:", "WARN")
        for e, p, mt in sin_ubic:
            log(f"      · {e} | planta '{p}' | {mt:,.2f} MT  -> falta su fila/ubicación",
                "WARN")

    df = _df_actividad(filas)
    log(f"Construcción: {len(df)} planta-distrito de "
        f"{df['empresa'].nunique()} empresa(s).")
    return df, sheet


# ==============================================================================
# 7-quater) FUSIÓN DE SECTORES
# ==============================================================================

def _strip_num_hoja(nombre):
    """'1 .Aenza' -> 'AENZA', '28. InRetail' -> 'INRETAIL' (normalizado)."""
    return norm(re.sub(r"^\s*\d+\s*[.\-]?\s*", "", str(nombre)))


def _buscar_hoja(sheetnames, objetivo):
    """Encuentra la hoja real cuyo nombre (sin número de prefijo) coincide."""
    t = norm(objetivo)
    for s in sheetnames:
        if _strip_num_hoja(s) == t:
            return s
    for s in sheetnames:                       # tolerante: coincidencia parcial
        if t and (t in _strip_num_hoja(s) or _strip_num_hoja(s) in t):
            return s
    return None


def load_empresa_hoja(risk_path, spec, sheet_real):
    """Lee la hoja propia de una empresa y devuelve un DataFrame de ubicaciones
    (mismo formato que _df_actividad) con el peso de distribución como 'tmf'.
    Detecta la fila-header por Departamento/Provincia/Distrito y toma la 1ra
    columna de peso (PESO_HINTS) que aparezca tras 'Distrito' (primer bloque, a la
    izquierda). Si el peso viene vacío, reparte equitativo (1/N)."""
    import openpyxl
    wb = openpyxl.load_workbook(risk_path, data_only=True)
    ws = wb[sheet_real]
    emp = spec["empresa"]
    # 1) fila-header: contiene Departamento + Provincia + Distrito
    hdr = None
    for r in range(1, min(ws.max_row, 40) + 1):
        vals = {norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)}
        if {"DEPARTAMENTO", "PROVINCIA", "DISTRITO"} <= vals:
            hdr = r
            break
    if hdr is None:
        log(f"   [{emp}] no se halló fila-header (Departamento/Provincia/Distrito). "
            f"Hoja omitida.", "WARN")
        return _df_actividad([])

    def col_de(nombre, desde=1):
        for c in range(desde, ws.max_column + 1):
            if norm(ws.cell(hdr, c).value) == nombre:
                return c
        return None
    cD = col_de("DEPARTAMENTO")
    cP = col_de("PROVINCIA", cD + 1)
    cDist = col_de("DISTRITO", cP + 1) if cP else None
    if not (cD and cP and cDist):
        log(f"   [{emp}] header incompleto. Hoja omitida.", "WARN")
        return _df_actividad([])
    # 2) columna de PESO: primera tras 'Distrito' cuyo header tenga '%' o una pista
    #    de texto (DISTRIBUCION/PARTICIPACION/PORCENTAJE/CRITERIO/PESO). OJO: norm('%')
    #    queda vacío, por eso el '%' se evalúa sobre el texto CRUDO.
    pistas_txt = [norm(p) for p in PESO_HINTS if norm(p)]
    def _es_peso(raw):
        if raw is None:
            return False
        if "%" in str(raw):
            return True
        h = norm(raw)
        return any(p in h for p in pistas_txt)
    cW = None
    for c in range(cDist + 1, ws.max_column + 1):
        if _es_peso(ws.cell(hdr, c).value):
            cW = c
            break
    peso_hdr = ws.cell(hdr, cW).value if cW else None
    # 3) columna de 'activo' (nombre de la ubicación), si existe antes de Departamento
    cAct = None
    nombres_act = {"ACTIVO", "PREDIO O ACTIVO", "PREDIO", "CENTRO COMERCIAL",
                   "SOCIEDAD", "SEDE", "MAPA DE ESCUELAS", "SUB BLOQUE", "LINEA"}
    for c in range(1, cD):
        if norm(ws.cell(hdr, c).value) in nombres_act:
            cAct = c
            break
    if cAct is None:                            # cualquier columna de texto a la izquierda
        for c in range(1, cD):
            if isinstance(ws.cell(hdr, c).value, str):
                cAct = c
                break
    # 4) leer filas hasta el primer hueco total de Dep/Prov/Dist
    filas, pesos = [], []
    empezo = False
    for r in range(hdr + 1, ws.max_row + 1):
        dep = ws.cell(r, cD).value
        prov = ws.cell(r, cP).value
        dist = ws.cell(r, cDist).value
        if dep is None and prov is None and dist is None:
            if empezo:
                break
            continue
        if dist is None or dep is None:
            continue
        empezo = True
        w = ws.cell(r, cW).value if cW else None
        try:
            w = float(w)
        except (TypeError, ValueError):
            w = None
        act = ws.cell(r, cAct).value if cAct else None
        if act is None or not isinstance(act, str):
            act = str(dist)
        filas.append({"depto": dep, "prov": prov, "dist": dist, "activo": act})
        pesos.append(w)
    if not filas:
        log(f"   [{emp}] sin filas de datos bajo el header.", "WARN")
        return _df_actividad([])
    # 5) pesos: si todos vacíos -> equitativo; si algunos vacíos -> 0
    if all(p is None for p in pesos):
        n = len(pesos)
        pesos = [1.0 / n] * n
        log(f"   [{emp}] columna de peso vacía; reparto equitativo (1/{n}).", "WARN")
    else:
        pesos = [p if p is not None else 0.0 for p in pesos]

    out = []
    for f, w in zip(filas, pesos):
        out.append({
            "mineral": "Distribución", "unidad": "%", "titular": emp,
            "unidad_minera": f["activo"], "depto": f["depto"], "prov": f["prov"],
            "dist": f["dist"], "acum": w, "unidad_norm": "pct", "factor": 1.0,
            "tmf": w, "empresa": emp})
    df = _df_actividad(out)
    log(f"   [{emp}] {len(out)} ubicación(es) | peso='{peso_hdr}' | "
        f"suma={sum(pesos):.4f}")
    return df


# ==============================================================================
# 7-quinquies) FUSIÓN DE SECTORES
# ==============================================================================

def merge_sectores(listas_districts):
    """Fusiona varias listas de 'districts' (una por sector) en una sola por ubigeo.
    Suma exposición y plata-en-riesgo, combina empresas y concatena el desglose.
    Los scores CENEPRED son idénticos por ubigeo entre sectores."""
    out = {}
    for districts in listas_districts:
        for d in districts:
            u = d["ubigeo"]
            if u not in out:
                out[u] = {
                    "ubigeo": u, "dep": d["dep"], "prov": d["prov"], "dist": d["dist"],
                    "tmf": 0.0, "exposicion": 0.0, "exp_sec": {},
                    "has_exp": False, "has_score": d["has_score"],
                    "nivel": d.get("nivel"),
                    "score_mm": d["score_mm"], "score_inu": d["score_inu"],
                    "score_seq": d["score_seq"], "score_comb": d["score_comb"],
                    "mar_mm": 0.0, "mar_inu": 0.0, "mar_seq": 0.0,
                    "empresas": {}, "minerales": [], "aportes": [],
                }
            o = out[u]
            o["tmf"] += d.get("tmf", 0.0)
            o["exposicion"] = round(o["exposicion"] + d["exposicion"], 2)
            for s, v in d.get("exp_sec", {}).items():
                o["exp_sec"][s] = round(o["exp_sec"].get(s, 0.0) + v, 2)
            o["has_exp"] = o["has_exp"] or d["has_exp"]
            for hz in ("mar_mm", "mar_inu", "mar_seq"):
                o[hz] = round(o[hz] + d.get(hz, 0.0), 2)
            for k, v in d.get("empresas", {}).items():
                o["empresas"][k] = round(o["empresas"].get(k, 0.0) + v, 2)
            o["minerales"].extend(d.get("minerales", []))
            o["aportes"].extend(d.get("aportes", []))
    return list(out.values())


# ==============================================================================
# 8) EXCEL DE SALIDA
# ==============================================================================

def export_excel(out_path, df_assign, exp, scores, geo_idx, um, detalle,
                 districts, no_map, universo=None, emp_sector=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    if universo is None:
        universo = EMPRESAS_CANONICAS
    if emp_sector is None:
        emp_sector = {e: SECTOR_MINERIA for e in universo}

    def smeta(e):
        return SECTOR_META.get(emp_sector.get(e, SECTOR_MINERIA),
                               SECTOR_META[SECTOR_MINERIA])

    def es_fin(e):
        return emp_sector.get(e) == SECTOR_FINANCIERO

    emp_pct = {spec["empresa"] for spec in EMPRESAS_HOJA}

    def peso_lbl(e, largo=False):
        if e in emp_pct:                 # empresas con peso de distribución (%)
            return "Distribución (%)"
        m = smeta(e)
        return f"{m['peso']} ({m['unidad']})"

    _, _, u2name = geo_idx
    wb = Workbook()
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor=AFP_AZUL.replace("#", ""))
    hdr_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    title_font = Font(name="Calibri", bold=True, color=AFP_AZUL.replace("#", ""), size=14)
    sub_font = Font(name="Calibri", italic=True, color="595959", size=9)

    def style_header(ws, row, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

    total_exp = sum(exp.values())

    # ---------- Hoja RESUMEN ----------
    ws = wb.active
    ws.title = "Resumen"
    ws["A1"] = "Exposición de cartera a riesgos físicos — Resumen por empresa"
    ws["A1"].font = title_font
    ws["A2"] = (f"Generado: {datetime.now():%Y-%m-%d %H:%M}  ·  Minería en TMF "
                f"(Tonelada Métrica Fina) · Banca en miles de S/ colocados")
    ws["A2"].font = sub_font
    headers = ["Empresa (canónica)", "Sector", "Producción / Créditos", "N° distritos",
               "Exposición cartera (S/)", "% de la exposición total"]
    ws.append([])
    ws.append(headers)
    hrow = ws.max_row
    style_header(ws, hrow, len(headers))
    prod_emp = df_assign.groupby("empresa")["tmf"].sum().to_dict()
    ndist_emp = df_assign[df_assign["tmf"] > 0].groupby("empresa")["ubigeo"].nunique().to_dict()
    for e in universo:
        ws.append([e, emp_sector.get(e, ""), round(prod_emp.get(e, 0.0), 2),
                   int(ndist_emp.get(e, 0)), round(exp.get(e, 0.0), 2), None])
        r = ws.max_row
    # total + porcentajes con fórmulas (A=Empresa B=Sector C=Prod D=NDist E=Exp F=%)
    first = hrow + 1
    last = ws.max_row
    ws.append(["TOTAL", "", f"=SUM(C{first}:C{last})", f"=SUM(D{first}:D{last})",
               f"=SUM(E{first}:E{last})", None])
    trow = ws.max_row
    for i, e in enumerate(universo):
        rr = first + i
        ws.cell(rr, 6).value = f"=IF($E${trow}=0,0,E{rr}/$E${trow})"
        ws.cell(rr, 6).number_format = "0.0%"
    ws.cell(trow, 6).value = f"=IF($E${trow}=0,0,SUM(E{first}:E{last})/$E${trow})"
    ws.cell(trow, 6).number_format = "0.0%"
    for c in range(1, 7):
        ws.cell(trow, c).font = Font(name="Calibri", bold=True)
    for rr in range(first, trow + 1):
        ws.cell(rr, 3).number_format = "#,##0.00"
        ws.cell(rr, 5).number_format = "#,##0"
    widths = [26, 13, 22, 13, 22, 22]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f"A{hrow + 1}"

    # ---------- Hoja por EMPRESA ----------
    for e in universo:
        sub = df_assign[(df_assign["empresa"] == e)].copy()
        ws = wb.create_sheet(_safe_sheet_name(e))
        m = smeta(e)
        ws["A1"] = f"{e} — {m['peso']} por distrito y exposición asignada"
        ws["A1"].font = title_font
        ws["A2"] = (f"{m['peso']} total: {sub['tmf'].sum():,.2f} {m['unidad']}   ·   "
                    f"Exposición en cartera: S/ {exp.get(e,0):,.0f}   ·   "
                    f"Asignación: {METODO_ASIGNACION}")
        ws["A2"].font = sub_font
        if sub.empty:
            ws["A4"] = "Sin registros para esta empresa en el archivo."
            continue
        col_peso = f"{m['peso']} ({m['unidad']})"
        col_pct = "% de la empresa"
        col_um = m["activo"]
        headers = ["Departamento", "Provincia", "Distrito", "Ubigeo",
                   col_um, col_peso, col_pct,
                   "Exposición asignada (S/)", "Score Mov. Masa", "Score Inundación",
                   "Score Sequías", "Plata en riesgo (S/, comb.)"]
        ws.append([]); ws.append(headers)
        hrow = ws.max_row
        style_header(ws, hrow, len(headers))
        sub = sub.sort_values("tmf", ascending=False)
        first = hrow + 1
        for _, r in sub.iterrows():
            u = r["ubigeo"]
            dep, prov, dist = u2name.get(u, ("", "", ""))
            sc = scores.get(u, {})
            uni = ", ".join(um.get((e, u), [])) or "—"
            avail = [sc.get(k) for k in ("mm", "inu", "seq") if sc.get(k) is not None]
            comb = sum(avail) / len(avail) if avail else 0
            mar = r["exposicion"] * comb
            ws.append([dep, prov, dist, u, uni, round(r["tmf"], 4), None,
                       round(r["exposicion"], 2), sc.get("mm"), sc.get("inu"),
                       sc.get("seq"), round(mar, 2)])
        last = ws.max_row
        tot_tmf_cell = f"=SUM(F{first}:F{last})"
        ws.append(["TOTAL", "", "", "", "", tot_tmf_cell, "", f"=SUM(H{first}:H{last})",
                   "", "", "", f"=SUM(L{first}:L{last})"])
        trow = ws.max_row
        for i in range(first, last + 1):
            ws.cell(i, 7).value = f"=IF($F${trow}=0,0,F{i}/$F${trow})"
            ws.cell(i, 7).number_format = "0.0%"
            ws.cell(i, 6).number_format = "#,##0.0000"
            ws.cell(i, 8).number_format = "#,##0"
            ws.cell(i, 12).number_format = "#,##0"
        for c in (1, 6, 8, 12):
            ws.cell(trow, c).font = Font(bold=True)
        ws.cell(trow, 6).number_format = "#,##0.0000"
        ws.cell(trow, 8).number_format = "#,##0"
        ws.cell(trow, 12).number_format = "#,##0"
        for col, w in zip("ABCDEFGHIJKL",
                          [16, 16, 18, 9, 34, 16, 13, 20, 13, 13, 12, 20]):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = f"A{first}"

        # detalle por mineral / colocaciones (debajo)
        ws.append([]); ws.append([])
        det_t = (f"Detalle por {smeta(e)['detalle']} ({smeta(e)['unidad']})"
                 if emp_sector.get(e) != SECTOR_MINERIA
                 else "Detalle por mineral (unidad original)")
        ws.append([det_t])
        ws.cell(ws.max_row, 1).font = Font(bold=True, color=AFP_AZUL.replace("#", ""))
        dh = ["Distrito", "Ubigeo", "Mineral", "Unidad", "Cantidad (orig.)", "TMF equiv."]
        ws.append(dh)
        style_header(ws, ws.max_row, len(dh))
        for u in sub["ubigeo"]:
            dep, prov, dist = u2name.get(u, ("", "", ""))
            for it in detalle.get((e, u), []):
                ws.append([dist, u, it["mineral"], it["unidad"],
                           round(it["cantidad_original"], 4), round(it["tmf"], 6)])
                ws.cell(ws.max_row, 5).number_format = "#,##0.0000"
                ws.cell(ws.max_row, 6).number_format = "#,##0.000000"

    # ---------- Hoja EXPOSICIÓN Y RIESGO (por distrito) ----------
    ws = wb.create_sheet("Exposición y Riesgo")
    ws["A1"] = "Exposición y plata en riesgo por distrito (todas las empresas)"
    ws["A1"].font = title_font
    ws["A2"] = ("Plata en riesgo = exposición asignada × score CENEPRED (1-5). "
                "Score combinado = promedio de los disponibles.")
    ws["A2"].font = sub_font
    headers = ["Departamento", "Provincia", "Distrito", "Ubigeo",
               "Exposición (S/)", "Score Mov. Masa", "Score Inundación",
               "Score Sequías", "Score comb.", "Riesgo Mov. Masa (S/)",
               "Riesgo Inundación (S/)", "Riesgo Sequías (S/)"]
    ws.append([]); ws.append(headers)
    style_header(ws, ws.max_row, len(headers))
    for d in sorted(districts, key=lambda x: x["exposicion"], reverse=True):
        ws.append([d["dep"], d["prov"], d["dist"], d["ubigeo"], d["exposicion"],
                   d["score_mm"], d["score_inu"], d["score_seq"], d["score_comb"],
                   d["mar_mm"], d["mar_inu"], d["mar_seq"]])
        for c in (5, 10, 11, 12):
            ws.cell(ws.max_row, c).number_format = "#,##0"
    for col, w in zip("ABCDEFGHIJKL", [16, 16, 18, 9, 18, 13, 13, 12, 11, 18, 18, 18]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A4"

    # ---------- Hoja EQUIVALENCIAS ----------
    ws = wb.create_sheet("Equivalencias")
    ws["A1"] = "Equivalencias de unidades y de nombres de empresa"
    ws["A1"].font = title_font
    ws.append([]); ws.append(["Unidades de medida (todo se lleva a TMF)"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, color=AFP_AZUL.replace("#", ""))
    ws.append(["Abreviatura", "Significado", "Inglés", "Uso típico", "Factor a TMF"])
    style_header(ws, ws.max_row, 5)
    uni_tbl = [
        ["TMF", "Tonelada Métrica Fina", "Metric Fine Ton (MFT)",
         "Cobre, zinc, plomo, molibdeno, hierro, estaño", 1],
        ["Grs. f.", "Gramos finos", "Fine grams", "Oro", "1 / 1,000,000"],
        ["Kg. f.", "Kilogramos finos", "Fine kilograms", "Plata", "1 / 1,000"],
    ]
    for row in uni_tbl:
        ws.append(row)
    ws.append([]); ws.append(["Equivalencia de empresas (producción → canónico)"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, color=AFP_AZUL.replace("#", ""))
    ws.append(["Nombre en producción (TITULAR)", "Nombre canónico"])
    style_header(ws, ws.max_row, 2)
    for k, v in EQUIV_EMPRESAS_PRODUCCION.items():
        ws.append([k, v])
    ws.append([]); ws.append(["Equivalencia de empresas (exposición → canónico)"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, color=AFP_AZUL.replace("#", ""))
    ws.append(["Nombre en Valorización (col. H)", "Nombre canónico"])
    style_header(ws, ws.max_row, 2)
    for k, v in EQUIV_EMPRESAS_EXPOSICION.items():
        ws.append([k, v])
    for col, w in zip("ABCDE", [40, 26, 24, 44, 14]):
        ws.column_dimensions[col].width = w

    # ---------- Hoja NO MAPEADOS ----------
    ws = wb.create_sheet("No mapeados")
    ws["A1"] = "Filas de producción sin ubigeo asignado (excluidas del mapa)"
    ws["A1"].font = title_font
    ws.append([]); ws.append(["Empresa", "Departamento", "Provincia", "Distrito",
                              "Motivo", "Producción (TMF)"])
    style_header(ws, ws.max_row, 6)
    if len(no_map):
        agg = (no_map.groupby(["empresa", "depto", "prov", "dist", "geo_motivo"])["tmf"]
               .sum().reset_index())
        for _, r in agg.iterrows():
            ws.append([r["empresa"], r["depto"], r["prov"], r["dist"],
                       r["geo_motivo"], round(r["tmf"], 4)])
            ws.cell(ws.max_row, 6).number_format = "#,##0.0000"
    else:
        ws.append(["(Todas las filas fueron geocodificadas correctamente)"])
    for col, w in zip("ABCDEF", [22, 18, 18, 20, 16, 16]):
        ws.column_dimensions[col].width = w

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)
    return out_path


def _safe_sheet_name(name):
    n = re.sub(r"[\[\]\:\*\?\/\\]", " ", name).strip()[:31]
    return n


# ==============================================================================
# 9) HTML INTERACTIVO
# ==============================================================================

def _equiv_exposicion_all():
    """Une todas las equivalencias de exposición (col H -> canónico) de los 4
    sectores, más los nombres canónicos que coinciden consigo mismos."""
    eq = {c: c for c in EMPRESAS_CANONICAS}
    eq.update(EQUIV_EMPRESAS_EXPOSICION)
    eq.update(EQUIV_EXPOSICION_FINANCIERO)
    eq.update(EQUIV_EXPOSICION_UTILITIES)
    eq.update(EQUIV_EXPOSICION_CONSTRUCCION)
    # empresas con hoja propia (cada nombre de col H -> su empresa canónica)
    for spec in EMPRESAS_HOJA:
        for nm in spec["exp"]:
            eq[nm] = spec["empresa"]
    return eq


def load_exposicion_clase(risk_path, sheet):
    """Lee 'Valorización de Instrumentos' y devuelve:
    - exp_clase: {emisor_canónico: {'RENTA VARIABLE':x, 'RENTA FIJA':y, 'TOTAL':x+y}}
    - aum_cartera: {'TODOS':sumM, 'RENTA VARIABLE':..., 'RENTA FIJA':...} sobre TODA
      la cartera (todos los instrumentos, estén o no mapeados a un activo físico)."""
    raw = pd.read_excel(risk_path, sheet_name=sheet, header=None)
    hdr_row, _ = detect_header_row(raw, ["Nombre (G&P)", "VALOR_VECTOR_LOCAL"],
                                   max_scan=10)
    df = pd.read_excel(risk_path, sheet_name=sheet, header=hdr_row).dropna(how="all")
    spec = {"nombre": (["Nombre (G&P)", "Nombre"], "H"),
            "valor":  (["VALOR_VECTOR_LOCAL", "VALOR VECTOR LOCAL"], "M"),
            "clase":  (["TIPO_RENTA", "TIPO RENTA", "Tipo Renta"], "S")}
    cols = resolve_columns(df.columns, spec)
    df = df.rename(columns={cols["nombre"]: "nombre", cols["valor"]: "valor",
                            cols["clase"]: "clase"})
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0.0)

    def clase_norm(c):
        n = norm(c)
        if "VARIABLE" in n:
            return "RENTA VARIABLE"
        if "FIJA" in n:
            return "RENTA FIJA"
        return "OTRO"
    df["clase_n"] = df["clase"].apply(clase_norm)

    aum_cartera = {"TODOS": float(df["valor"].sum()),
                   "RENTA VARIABLE": float(df.loc[df["clase_n"] == "RENTA VARIABLE", "valor"].sum()),
                   "RENTA FIJA": float(df.loc[df["clase_n"] == "RENTA FIJA", "valor"].sum())}

    eq_norm = {norm(k): v for k, v in _equiv_exposicion_all().items()}
    df["emisor"] = df["nombre"].apply(lambda n: eq_norm.get(norm(n)))
    exp_clase = {}
    for _, r in df.dropna(subset=["emisor"]).iterrows():
        d = exp_clase.setdefault(r["emisor"], {"RENTA VARIABLE": 0.0,
                                               "RENTA FIJA": 0.0, "TOTAL": 0.0})
        if r["clase_n"] in ("RENTA VARIABLE", "RENTA FIJA"):
            d[r["clase_n"]] += r["valor"]
        d["TOTAL"] += r["valor"]
    for k in exp_clase:
        for kk in exp_clase[k]:
            exp_clase[k][kk] = round(exp_clase[k][kk], 2)
    return exp_clase, {k: round(v, 2) for k, v in aum_cartera.items()}


def load_enfen():
    """Lee el resumen ENFEN (serie mensual de estado de alerta El Niño/La Niña).
    Devuelve (lista_meses, ultimo) o ([], None) si no se encuentra el archivo."""
    try:
        path = find_file(ENFEN_DIR, ENFEN_PATTERNS, "ENFEN (Fenómenos Geofísicos)")
    except SystemExit:
        log("No se encontró el archivo ENFEN; la pestaña de Fenómenos Geofísicos "
            "quedará vacía. Revisa ENFEN_DIR / ENFEN_PATTERNS en CONFIG.", "WARN")
        return [], None
    xl = pd.ExcelFile(path)
    sheet = xl.sheet_names[0]
    for h in ENFEN_SHEET_HINTS:
        for s in xl.sheet_names:
            if norm(h) in norm(s):
                sheet = s; break
    df = pd.read_excel(path, sheet_name=sheet)
    cols = {norm(c): c for c in df.columns}
    def pick(*alts):
        for a in alts:
            if norm(a) in cols:
                return cols[norm(a)]
        return None
    c_fecha = pick("Fecha"); c_estado = pick("Estado de Alerta", "Estado")
    c_conf = pick("Conf. Estado", "Conf", "Confianza")
    if c_fecha is None or c_estado is None:
        log("ENFEN: no se hallaron columnas 'Fecha'/'Estado de Alerta'. Revisa el "
            "archivo.", "WARN")
        return [], None
    serie = []
    for _, r in df.iterrows():
        f = r[c_fecha]
        if pd.isna(f):
            continue
        try:
            f = pd.to_datetime(f)
        except Exception:
            continue
        serie.append({"fecha": f.strftime("%Y-%m"), "anio": f.year, "mes": f.month,
                      "estado": str(r[c_estado]) if not pd.isna(r[c_estado]) else "No Activo",
                      "conf": (str(r[c_conf]) if c_conf and not pd.isna(r[c_conf]) else "")})
    serie.sort(key=lambda x: x["fecha"])
    log(f"ENFEN: {len(serie)} meses leídos ({serie[0]['fecha']} a {serie[-1]['fecha']})."
        if serie else "ENFEN: sin datos.")
    return serie, (serie[-1] if serie else None)


# ==============================================================================
# 9-bis) GRÁFICO CLIMA vs. INDICADOR SECTORIAL (clima, SBS, producción mensual)
# ==============================================================================

def _fecha_ym(v):
    """Normaliza un valor de fecha a 'YYYY-MM' (o None)."""
    try:
        return pd.to_datetime(v).strftime("%Y-%m")
    except Exception:
        return None


def load_clima():
    """Lee el archivo de indicadores climáticos (hojas con estructuras distintas) y
    devuelve ({indicador: [{'f':'YYYY-MM','v':float}]}, orden). Cada indicador del
    dropdown se obtiene según CLIMA_FUENTE (col/bimestre/wide). Robusto: si una hoja
    o columna falta, avisa y omite ese indicador (el resto sigue)."""
    try:
        path = find_file(FEN_IND_DIR, CLIMA_PATTERNS, "Indicadores climáticos")
    except SystemExit:
        log("No se encontró 'indicadores_franco.xlsx'; el panel climático del gráfico "
            "quedará vacío (el resto del gráfico funciona). Revisa FEN_IND_DIR / "
            "CLIMA_PATTERNS en CONFIG.", "WARN")
        return {}, []
    xl = pd.ExcelFile(path)
    sheets_norm = {norm(s): s for s in xl.sheet_names}

    def hoja_real(nombre):
        if norm(nombre) in sheets_norm:
            return sheets_norm[norm(nombre)]
        for n, real in sheets_norm.items():
            if norm(nombre) in n:
                return real
        return None

    def col_real(df, nombre):
        cols = {norm(c): c for c in df.columns}
        if norm(nombre) in cols:
            return cols[norm(nombre)]
        for n, real in cols.items():
            if norm(nombre) in n:
                return real
        return None

    out, orden = {}, []
    for ind in CLIMA_SHEETS:
        fuente = CLIMA_FUENTE.get(ind)
        if not fuente:
            continue
        hoja, modo = fuente
        real = hoja_real(hoja)
        if real is None:
            log(f"Clima: no se encontró la hoja '{hoja}' para '{ind}'; se omite.", "WARN")
            continue
        df = pd.read_excel(path, sheet_name=real)
        if df.empty:
            continue
        serie = []
        if modo.startswith("col:"):
            cval = col_real(df, modo[4:]); cfec = col_real(df, "fecha")
            if cfec is None or cval is None:
                log(f"Clima '{ind}': falta columna fecha o '{modo[4:]}' en '{real}'; "
                    f"se omite.", "WARN")
                continue
            for _, r in df.iterrows():
                ym = _fecha_ym(r[cfec]); v = pd.to_numeric(r[cval], errors="coerce")
                if ym and pd.notna(v):
                    serie.append({"f": ym, "v": float(v)})
        elif modo == "bimestre":
            cano = col_real(df, "anio") or col_real(df, "año")
            cbim = col_real(df, "bimestre")
            cval = [c for c in df.columns if c not in (cano, cbim)]
            cval = cval[0] if cval else None
            if not all([cano, cbim, cval]):
                log(f"Clima '{ind}': estructura año/bimestre/valor no reconocida; se "
                    f"omite.", "WARN"); continue
            for _, r in df.iterrows():
                mes = CLIMA_BIMESTRE_MES.get(norm(r[cbim]))
                v = pd.to_numeric(r[cval], errors="coerce")
                try:
                    anio = int(r[cano])
                except Exception:
                    anio = None
                if mes and anio and pd.notna(v):
                    serie.append({"f": f"{anio:04d}-{mes:02d}", "v": float(v)})
        elif modo == "wide":
            cano = col_real(df, "anio") or col_real(df, "año")
            if cano is None:
                log(f"Clima '{ind}': no se halló columna de año (wide); se omite.", "WARN")
                continue
            for _, r in df.iterrows():
                try:
                    anio = int(r[cano])
                except Exception:
                    continue
                for c in df.columns:
                    if c == cano:
                        continue
                    mes = CLIMA_TEMPORADA_MES.get(norm(c))
                    if not mes:
                        continue
                    v = pd.to_numeric(r[c], errors="coerce")
                    if pd.notna(v):
                        serie.append({"f": f"{anio:04d}-{mes:02d}", "v": float(v)})
        serie = [s for s in serie if s["f"]]
        serie.sort(key=lambda x: x["f"])
        if serie:
            out[ind] = serie; orden.append(ind)
        else:
            log(f"Clima '{ind}': sin datos válidos tras procesar '{real}'.", "WARN")
    log(f"Clima: {len(out)} indicador(es) cargado(s): {', '.join(orden)}."
        if out else "Clima: sin indicadores.")
    return out, orden


def load_sbs():
    """Lee BD SBS (formato largo) y devuelve {banco_canon: {indicador: [{f,v}]}}
    solo para los bancos homologados y los indicadores de la whitelist (para acotar
    el HTML). Audita bancos de interés no mapeados y variantes de indicador."""
    try:
        path = find_file(SBS_DIR, SBS_PATTERNS, "BD SBS (indicadores financieros)")
    except SystemExit:
        log("No se encontró 'BD SBS.xlsx'; el panel financiero del gráfico quedará "
            "vacío. Revisa SBS_DIR / SBS_PATTERNS en CONFIG.", "WARN")
        return {}, []
    xl = pd.ExcelFile(path)
    sheet = xl.sheet_names[0]
    for h in SBS_SHEET_HINTS:
        for s in xl.sheet_names:
            if norm(h) in norm(s):
                sheet = s; break
    df = pd.read_excel(path, sheet_name=sheet)
    cols = {norm(c): c for c in df.columns}
    cF = cols.get("FECHA"); cE = cols.get("EMPRESA")
    cI = cols.get("INDICADOR"); cV = cols.get("VALOR")
    if not all([cF, cE, cI, cV]):
        log("SBS: no se hallaron columnas Fecha/Empresa/Indicador/Valor.", "WARN")
        return {}, []
    eqb = {norm_fuerte(k): v for k, v in EQUIV_BANCOS_SBS.items()}
    df["banco"] = df[cE].apply(lambda x: eqb.get(norm_fuerte(x)))

    # --- resolución de indicador -> nombre canónico (whitelist + fusión) ---
    # mapa: clave_fuerte de cada variante -> nombre canónico a mostrar
    canon_por_clave = {}
    for canon, variantes in SBS_FUSION.items():
        for v in variantes:
            canon_por_clave[norm_fuerte(v)] = canon
    wl_clave = {norm_fuerte(w): w for w in SBS_IND_WHITELIST}  # clave -> nombre
    def canon_ind(ind):
        k = norm_fuerte(ind)
        if k in canon_por_clave:
            return canon_por_clave[k]
        if k in wl_clave:
            return wl_clave[k]
        return None
    df["canon"] = df[cI].apply(canon_ind)

    # --- auditoría: bancos de interés no mapeados ---
    no_map = sorted(set(str(x) for x in df.loc[df["banco"].isna(), cE].dropna()
                        if any(norm_fuerte(p) in norm_fuerte(x) for p in PISTAS_BANCO_SBS)))
    if no_map:
        log("=== ADVERTENCIA DE HOMOLOGACIÓN DE BANCOS (SBS) ===", "WARN")
        log("Nombres que parecen banco de interés y NO están mapeados (agrega a "
            "EQUIV_BANCOS_SBS):", "WARN")
        for b in no_map:
            log(f"   - {b}", "WARN")

    sub = df[df["banco"].notna() & df["canon"].notna()].copy()
    sub["f"] = sub[cF].apply(_fecha_ym)
    sub["v"] = pd.to_numeric(sub[cV], errors="coerce")
    sub = sub.dropna(subset=["f", "v"])
    out = {}
    indicadores = set()
    cobertura = {}
    for (banco, ind), g in sub.groupby(["banco", "canon"]):
        indicadores.add(str(ind))
        serie = (g.groupby("f")["v"].mean().reset_index().sort_values("f"))
        out.setdefault(banco, {})[str(ind)] = [
            {"f": r["f"], "v": float(r["v"])} for _, r in serie.iterrows()]
        cobertura.setdefault(str(ind), []).append((serie["f"].min(), serie["f"].max(),
                                                    len(serie)))
    # default robusto: si el configurado no quedó, usa el de mayor cobertura
    inds = sorted(indicadores)
    default = SBS_IND_DEFAULT if SBS_IND_DEFAULT in inds else (
        max(inds, key=lambda i: min(c[2] for c in cobertura[i])) if inds else "")
    log(f"SBS: {len(out)} banco(s), {len(inds)} indicador(es) homologados. "
        f"Default: '{default}'.")
    for i in inds:
        rng = cobertura[i]
        log(f"      · {i}: {min(r[0] for r in rng)} -> {max(r[1] for r in rng)}")
    return out, inds, default


def load_produccion_mensual():
    """Serie mensual de producción minera por empresa, desde la hoja 'BD Prod' de
    BD Mineria.xlsm. Estructura ancha: fila 1 = fechas (desde col G) + headers
    (B MINERAL, C ETAPA, E TITULAR, F UNIDAD). Suma por empresa+mes con
    ETAPA=CONCENTRACIÓN, sin distinguir mineral. {empresa_canon: [{f,v}]}."""
    try:
        path = find_file(PROD_MENSUAL_DIR, PROD_MENSUAL_PATTERNS,
                         "BD Mineria (serie mensual)")
    except SystemExit:
        log("No se encontró 'BD Mineria.xlsm'; el panel de minería del gráfico quedará "
            "vacío. Revisa PROD_MENSUAL_DIR / PROD_MENSUAL_PATTERNS en CONFIG.", "WARN")
        return {}
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet = next((s for s in wb.sheetnames
                  if any(norm(h) in norm(s) for h in PROD_MENSUAL_SHEET_HINTS)),
                 wb.sheetnames[0])
    ws = wb[sheet]
    # localizar columnas de descriptores por su header en fila 1
    hdr = {norm(ws.cell(1, c).value): c for c in range(1, ws.max_column + 1)
           if isinstance(ws.cell(1, c).value, str)}
    cEt = hdr.get("ETAPA"); cTit = hdr.get("TITULAR")
    if not cEt or not cTit:
        log("BD Prod: no se hallaron columnas ETAPA/TITULAR en la fila 1.", "WARN")
        return {}
    cols_fecha = _cols_fecha_todas(ws, 1)        # (col, 'YYYY-MM')
    proc_ok = {norm(p) for p in PROCESOS_VALIDOS}
    eq = {norm(k): v for k, v in EQUIV_MINERAS_MENSUAL.items()}
    acc = {}
    for r in range(2, ws.max_row + 1):
        if norm(ws.cell(r, cEt).value) not in proc_ok:
            continue
        canon = eq.get(norm(ws.cell(r, cTit).value))
        if not canon:
            continue
        for c, ym in cols_fecha:
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)) and v:
                acc.setdefault(canon, {}); acc[canon][ym] = acc[canon].get(ym, 0.0) + v
    out = {e: [{"f": k, "v": round(val, 4)} for k, val in sorted(d.items())]
           for e, d in acc.items()}
    if out:
        fmin = min(s[0]["f"] for s in out.values())
        fmax = max(s[-1]["f"] for s in out.values())
        log(f"Minería (BD Prod, serie mensual): {len(out)} empresa(s), {fmin} a {fmax}.")
    else:
        log("Minería (BD Prod): no se acumuló producción (revisa ETAPA/TITULAR).", "WARN")
    return out


def _cols_fecha_todas(ws, fila_fecha):
    """Lista de (col, 'YYYY-MM') de todas las columnas-fecha de 'fila_fecha'."""
    import datetime
    out = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(fila_fecha, c).value
        if isinstance(v, datetime.datetime):
            out.append((c, f"{v.year:04d}-{v.month:02d}"))
    return out


def load_serie_utilities_mensual():
    """Serie mensual de POTENCIA EFECTIVA (MW) por empresa, desde BD Utilities
    ('BD ptf'). Cobertura hasta el mes más reciente del archivo. {empresa: [{f,v}]}."""
    try:
        path = find_file(UTIL_DIR, UTIL_PATTERNS, "BD UTILITIES (serie mensual)")
    except SystemExit:
        return {}
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet = next((s for s in wb.sheetnames
                  if any(norm(h) in norm(s) for h in UTIL_SHEET_HINTS)), wb.sheetnames[0])
    ws = wb[sheet]
    eq = {norm(k): v for k, v in EQUIV_EMPRESAS_UTILITIES.items()}
    cols = _cols_fecha_todas(ws, 1)
    acc, emp_prev = {}, None
    for r in range(2, ws.max_row + 1):
        if norm(ws.cell(r, 2).value) != norm(UTIL_TIPO):
            continue
        emp = ws.cell(r, 3).value or emp_prev; emp_prev = emp
        canon = eq.get(norm(emp))
        if not canon:
            continue
        for c, ym in cols:
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)) and v:
                acc.setdefault(canon, {}); acc[canon][ym] = acc[canon].get(ym, 0.0) + v
    out = {e: [{"f": k, "v": round(val, 4)} for k, val in sorted(d.items())]
           for e, d in acc.items()}
    if out:
        fmax = max(s[-1]["f"] for s in out.values())
        log(f"Utilities (serie mensual): {len(out)} empresa(s), hasta {fmax}.")
    return out


def load_serie_construccion_mensual():
    """Serie mensual de despachos de cemento (MT) por empresa, desde BD Construcción
    ('Despachos'). {empresa: [{f,v}]}."""
    try:
        path = find_file(CONS_DIR, CONS_PATTERNS, "BD CONSTRUCCIÓN (serie mensual)")
    except SystemExit:
        return {}
    import openpyxl, datetime
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet = next((s for s in wb.sheetnames
                  if any(norm(h) in norm(s) for h in CONS_SHEET_HINTS)), wb.sheetnames[0])
    ws = wb[sheet]
    fila_fecha, best = 5, -1
    for rr in range(1, 9):
        n = sum(1 for c in range(1, ws.max_column + 1)
                if isinstance(ws.cell(rr, c).value, datetime.datetime))
        if n > best:
            best, fila_fecha = n, rr
    eq = {norm(k): v for k, v in EQUIV_EMPRESAS_CONSTRUCCION.items()}
    cols = _cols_fecha_todas(ws, fila_fecha)
    acc = {}
    for r in range(fila_fecha + 1, ws.max_row + 1):
        c3 = ws.cell(r, 3).value
        if c3 is None:
            continue
        emp, _ = _parse_planta_construccion(c3)
        canon = eq.get(norm(emp)) if emp else None
        if not canon:
            continue
        for c, ym in cols:
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)) and v:
                acc.setdefault(canon, {}); acc[canon][ym] = acc[canon].get(ym, 0.0) + v
    out = {e: [{"f": k, "v": round(val, 4)} for k, val in sorted(d.items())]
           for e, d in acc.items()}
    if out:
        fmax = max(s[-1]["f"] for s in out.values())
        log(f"Construcción (serie mensual): {len(out)} empresa(s), hasta {fmax}.")
    return out


def _enfen_meses_alerta(enfen):
    """Meses ENFEN con condición de ALERTA (homologando Costero == no Costero, y
    descartando 'No identificado'). Devuelve lista de 'YYYY-MM' con tipo niño/niña."""
    out = []
    for s in (enfen or []):
        e = norm(s.get("estado"))
        if "NO IDENTIFICADO" in e or "ALERTA" not in e:
            continue
        tipo = "nino" if "NINO" in e else ("nina" if "NINA" in e else "otro")
        out.append({"f": s["fecha"], "tipo": tipo})
    return out


def _logo_data_uri():
    """Devuelve un data URI con el logo institucional. Usa el PNG de LOGO_PATH si
    existe (embebido en Base64 -> archivo portable); si no, un wordmark SVG."""
    import base64
    if os.path.isfile(LOGO_PATH):
        try:
            with open(LOGO_PATH, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            ext = os.path.splitext(LOGO_PATH)[1].lower().lstrip(".") or "png"
            mime = "image/svg+xml" if ext == "svg" else f"image/{'jpeg' if ext in ('jpg','jpeg') else ext}"
            log(f"Logo embebido desde {LOGO_PATH}.")
            return f"data:{mime};base64,{b64}"
        except Exception as e:
            log(f"No se pudo leer el logo ({e}); se usa wordmark de respaldo.", "WARN")
    else:
        log(f"No se encontró el logo en {LOGO_PATH}; se usa wordmark de respaldo. "
            f"(Coloca el PNG ahí o ajusta LOGO_PATH/RF_LOGO para embeber el real.)",
            "WARN")
    svg = (f"<svg xmlns='http://www.w3.org/2000/svg' width='150' height='40'>"
           f"<rect width='150' height='40' rx='5' fill='{AFP_AZUL}'/>"
           f"<circle cx='26' cy='20' r='11' fill='{AFP_CYAN}'/>"
           f"<text x='46' y='26' font-family='Arial' font-size='18' font-weight='bold' "
           f"fill='white'>integra</text></svg>")
    return "data:image/svg+xml;base64," + base64.b64encode(svg.encode()).decode()


def export_html(out_path, geojson, districts, exp, scores, geo_idx, emp_sector=None,
                exp_clase=None, aum_cartera=None, enfen=None, fen_chart=None):
    _, _, u2name = geo_idx
    if emp_sector is None:
        emp_sector = {e: SECTOR_MINERIA for e in exp}
    exposicion_sector = {}
    for e, v in exp.items():
        s = emp_sector.get(e, SECTOR_MINERIA)
        exposicion_sector[s] = round(exposicion_sector.get(s, 0.0) + v, 2)
    # sectores realmente presentes, en orden canónico
    presentes = [s for s in ORDEN_SECTORES
                 if s in set(emp_sector.values())]
    # fecha de valorización del portafolio (de la hoja Valorización si está)
    fval = datetime.now().strftime("%d/%m/%Y")
    payload = {
        "generado": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "fecha_valoracion": fval,
        "districts": districts,
        "exposicion_empresa": {k: round(v, 2) for k, v in exp.items()},
        "emp_sector": emp_sector,
        "exposicion_sector": exposicion_sector,
        "exp_clase": exp_clase or {},
        "aum_cartera": aum_cartera or {"TODOS": 0, "RENTA VARIABLE": 0, "RENTA FIJA": 0},
        "niveles_altos": NIVELES_ALTOS,
        "nivel_modo": NIVEL_CENEPRED_MODO,
        "enfen": enfen or [],
        "fen_chart": fen_chart or {},
        "sectores": presentes,
        "sector_meta": {s: {"unidad": SECTOR_META[s]["unidad"],
                            "peso": SECTOR_META[s]["peso"],
                            "activo": SECTOR_META[s]["activo"]}
                        for s in ORDEN_SECTORES},
        "metodo": METODO_ASIGNACION,
        "logo": _logo_data_uri(),
        "afp": {"azul": AFP_AZUL, "cyan": AFP_CYAN, "amar": AFP_AMAR,
                "gris": AFP_GRIS, "plomo": AFP_PLOMO},
    }
    geojson_min = {"type": "FeatureCollection", "features": [
        {"type": "Feature",
         "properties": {"IDDIST": str(f["properties"].get("IDDIST", "")).zfill(6),
                        "NOMBDIST": f["properties"].get("NOMBDIST"),
                        "NOMBPROV": f["properties"].get("NOMBPROV"),
                        "NOMBDEP": f["properties"].get("NOMBDEP")},
         "geometry": f["geometry"]}
        for f in geojson["features"] if f.get("geometry")]}

    html = _HTML_TEMPLATE.replace("__PAYLOAD__", json.dumps(payload, ensure_ascii=False))
    html = html.replace("__GEOJSON__", json.dumps(geojson_min, ensure_ascii=False))
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path


_HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="es"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Riesgos Físicos ESG — AFP Integra</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js" charset="utf-8"></script>
<style>
 :root{
   --azul:#1E2E6E; --cyan:#00AECB; --amar:#E3E829; --gris:#DCDDDE; --plomo:#7E8083;
   --grisbg:#f4f5f8; --txt:#1f2533; --linea:#DCDDDE;
 }
 *{box-sizing:border-box;font-family:Calibri,'Segoe UI',Arial,sans-serif;}
 body{margin:0;background:#fff;color:var(--txt);font-size:14px;}
 /* ---------- Header corporativo ---------- */
 .topbar{display:flex;align-items:center;justify-content:space-between;
   padding:14px 26px;border-bottom:3px solid var(--azul);background:#fff;}
 .topbar .logo{height:42px;display:block;}
 .topbar .center{text-align:center;flex:1;}
 .topbar .center h1{margin:0;font-size:19px;color:var(--azul);font-weight:700;
   letter-spacing:.2px;}
 .topbar .center .sub{font-size:12px;color:var(--plomo);margin-top:2px;}
 .topbar .right{text-align:right;font-size:12px;color:var(--plomo);min-width:160px;}
 .topbar .right b{color:var(--azul);font-size:13px;}
 /* ---------- Barra de pestañas y modos ---------- */
 .tabbar{display:flex;align-items:center;gap:4px;background:var(--grisbg);
   padding:8px 22px 0;border-bottom:1px solid var(--linea);flex-wrap:wrap;}
 .tab{padding:9px 18px;border:none;background:transparent;color:var(--azul);
   font-size:14px;font-weight:600;cursor:pointer;border-radius:7px 7px 0 0;
   border-bottom:3px solid transparent;}
 .tab:hover{background:#e8ebf3;}
 .tab.active{background:#fff;border-bottom:3px solid var(--cyan);color:var(--azul);}
 .modos{margin-left:auto;display:flex;align-items:center;gap:8px;padding-bottom:6px;}
 .seg{display:inline-flex;border:1px solid var(--azul);border-radius:7px;overflow:hidden;}
 .seg button{border:none;background:#fff;color:var(--azul);padding:6px 12px;
   font-size:12.5px;cursor:pointer;font-weight:600;}
 .seg button.on{background:var(--azul);color:#fff;}
 .chk{font-size:12px;color:var(--plomo);display:inline-flex;align-items:center;gap:5px;
   cursor:pointer;}
 .wrap{max-width:1360px;margin:0 auto;padding:18px 22px 30px;}
 /* ---------- KPIs ---------- */
 .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:6px 0 18px;}
 .kpi{background:#fff;border:1px solid var(--linea);border-left:4px solid var(--cyan);
   border-radius:8px;padding:12px 14px;}
 .kpi .v{font-size:21px;font-weight:700;color:var(--azul);}
 .kpi .l{font-size:11.5px;color:var(--plomo);margin-top:3px;line-height:1.35;}
 /* ---------- Tarjetas de chart ---------- */
 .card{background:#fff;border:1px solid var(--linea);border-radius:10px;padding:14px 16px;
   margin-bottom:18px;}
 .card h3{margin:0 0 4px;font-size:15px;color:var(--azul);}
 .card .note{font-size:12px;color:var(--plomo);margin:0 0 10px;}
 .controls{display:flex;flex-wrap:wrap;gap:14px;align-items:flex-end;margin-bottom:10px;}
 .controls label{display:block;font-size:11px;color:var(--plomo);margin-bottom:3px;}
 .controls select,.controls input{padding:6px 9px;border:1px solid var(--linea);
   border-radius:6px;font-size:13px;color:var(--txt);background:#fff;}
 /* ---------- Tabla pivote ---------- */
 .pivwrap{overflow:auto;max-height:560px;border:1px solid var(--linea);border-radius:8px;}
 table.piv{border-collapse:collapse;width:100%;font-size:12.5px;}
 table.piv thead th{position:sticky;top:0;background:var(--azul);color:#fff;
   padding:9px 10px;text-align:left;font-weight:600;cursor:pointer;white-space:nowrap;z-index:2;}
 table.piv thead th .ar{opacity:.6;font-size:10px;}
 table.piv td{padding:7px 10px;border-top:1px solid var(--linea);vertical-align:top;}
 table.piv td.riesgo{font-weight:800;font-size:16px;color:#fff;text-align:center;
   vertical-align:middle;width:54px;}
 .r5{background:#c0392b;} .r4{background:#e67e22;} .r3{background:#f1c40f;color:#5a4500 !important;}
 table.piv td.aumtot{background:#eef3ff;vertical-align:middle;font-weight:700;
   color:var(--azul);text-align:right;white-space:nowrap;}
 table.piv td.num{text-align:right;white-space:nowrap;}
 .pct{font-weight:700;color:var(--azul);font-size:13.5px;}
 td.aumtot .pct{font-size:15px;}
 .monto{display:block;font-size:10.5px;color:var(--plomo);font-weight:400;margin-top:1px;}
 body.ocultar-montos .monto{display:none;}
 .activo-list{color:var(--txt);} .geo-list{color:var(--plomo);}
 table.piv tbody tr:hover td:not(.riesgo):not(.aumtot){background:#f7f9ff;}
 .search{margin-bottom:8px;}
 .search input{width:280px;}
 /* ---------- Mapa ---------- */
 #mapwrap,#mapENF{width:100%;height:600px;}
 /* ---------- ENFEN ---------- */
 .enf-banner{display:flex;gap:16px;align-items:center;border-radius:10px;padding:16px 18px;
   color:#fff;margin-bottom:16px;}
 .enf-banner .big{font-size:22px;font-weight:800;}
 .enf-grid{border-collapse:collapse;font-size:11px;}
 .enf-grid th{color:var(--plomo);font-weight:600;padding:3px 6px;text-align:center;}
 .enf-grid td{width:30px;height:22px;text-align:center;border:1px solid #fff;color:#fff;
   font-size:0;}
 .enf-leg{display:flex;flex-wrap:wrap;gap:14px;margin-top:14px;font-size:12px;color:var(--txt);}
 .enf-leg span{display:inline-flex;align-items:center;gap:6px;}
 .enf-leg i{width:14px;height:14px;border-radius:3px;display:inline-block;}
 /* ---------- Footer ---------- */
 footer{border-top:1px solid var(--linea);margin-top:10px;padding:16px 26px;
   color:var(--plomo);font-size:11.5px;line-height:1.6;background:var(--grisbg);}
 footer b{color:var(--azul);}
 /* ---------- Modo comité ---------- */
 body.modo-comite .solo-analista{display:none !important;}
 body.modo-comite .monto{display:none !important;}
 body.modo-comite #mapwrap,body.modo-comite #mapENF{height:680px;}
 body.modo-comite .kpi .v{font-size:24px;}
 .hidden{display:none !important;}
 details.metod{margin-top:6px;} details.metod summary{cursor:pointer;color:var(--azul);
   font-weight:600;font-size:13px;} details.metod p{font-size:12px;color:var(--plomo);}
</style></head>
<body>
<div class="topbar">
  <img id="logo" class="logo" alt="AFP Integra">
  <div class="center">
    <h1>Dashboard de Riesgos Físicos ESG</h1>
    <div class="sub">Mesa de Inversiones · Exposición de cartera a peligros CENEPRED y ENFEN</div>
  </div>
  <div class="right">Información al<br><b id="fval">—</b><br><span id="gen"></span></div>
</div>

<div class="tabbar">
  <button class="tab active" data-tab="resumen">Resumen</button>
  <button class="tab" data-tab="rv">Renta Variable</button>
  <button class="tab" data-tab="rf">Renta Fija</button>
  <button class="tab" data-tab="enfen">Fenómenos Geofísicos</button>
  <div class="modos">
    <label class="chk solo-analista"><input type="checkbox" id="chkMonto"> Mostrar montos (S/)</label>
    <div class="seg">
      <button id="mAnalista" class="on">Vista Analista</button>
      <button id="mComite">Vista Comité</button>
    </div>
  </div>
</div>

<div class="wrap">
  <!-- ====== PANEL DE DATOS (compartido por Resumen / RV / RF) ====== -->
  <div id="panelDatos">
    <div id="kpis" class="kpis"></div>

    <div class="card">
      <h3 id="pivTitle">Emisores por nivel de riesgo CENEPRED</h3>
      <p class="note" id="pivNote"></p>
      <div class="search solo-analista">
        <input type="text" id="pivSearch" placeholder="Buscar emisor, activo, distrito…">
      </div>
      <div class="pivwrap"><table class="piv" id="pivTable"></table></div>
    </div>

    <div class="card">
      <h3>Mapa coroplético — riesgo físico y plata invertida</h3>
      <p class="note">Panorama nacional CENEPRED; la capa de color resalta los distritos con exposición de la cartera (según pestaña y filtros).</p>
      <div class="controls solo-analista">
        <div><label>Sector</label>
          <select id="sector"><option value="todos" selected>Todos los sectores</option></select></div>
        <div><label>Colorear por</label>
          <select id="metric">
            <option value="mar_comb" selected>Plata en riesgo (comb.)</option>
            <option value="exposicion">Exposición asignada</option>
            <option value="score_comb">Score CENEPRED (comb.)</option>
            <option value="score_mm">Score Mov. de Masa</option>
            <option value="score_inu">Score Inundación</option>
            <option value="score_seq">Score Sequías</option>
          </select></div>
        <div><label>Ranking por</label>
          <select id="nivel"><option value="dep">Departamento</option>
            <option value="prov">Provincia</option></select></div>
        <div><label>Filtrar provincia</label>
          <select id="provfilter"><option value="">(todas)</option></select></div>
      </div>
      <div id="mapwrap"></div>
      <div id="bar" style="height:240px;margin-top:6px;"></div>
      <details class="metod"><summary>Metodología y supuestos</summary>
        <p>La exposición de cada emisor (col. M de "Valorización de Instrumentos") se
        reparte entre los distritos donde tiene activos físicos, en proporción a su peso
        sectorial (producción minera en TMF, créditos en miles S/, potencia efectiva en MW,
        despachos de cemento en MT). El split Renta Fija / Renta Variable usa la proporción
        de cada emisor en la col. S (TIPO_RENTA), asumida uniforme en su geografía.
        El "nivel de riesgo" del distrito es <b id="modoNivel"></b> de los peligros CENEPRED.
        "Plata en riesgo" = exposición × score.</p>
      </details>
    </div>
  </div>

  <!-- ====== PANEL ENFEN ====== -->
  <div id="panelENFEN" class="hidden">
    <div id="enfBanner" class="enf-banner"></div>
    <div class="card">
      <h3>Histórico de alertas ENFEN (El Niño / La Niña Costero)</h3>
      <p class="note">Estado del sistema oficial de alerta del ENFEN por mes. El Niño Costero se asocia a mayor peligro de inundaciones en la costa norte.</p>
      <div style="overflow:auto"><table class="enf-grid" id="enfGrid"></table></div>
      <div class="enf-leg" id="enfLeg"></div>
    </div>
    <div id="kpisENF" class="kpis"></div>

    <div class="card" id="cardFen">
      <h3>Clima (El Niño/La Niña) vs. indicador sectorial — ¿coinciden los shocks climáticos con el deterioro?</h3>
      <p class="note">Un solo gráfico de doble eje: izquierda = indicador del sector elegido (banca o producción física); derecha = indicador climático con sus umbrales de categoría (etiquetas al margen). Bandas rojizas = grandes eventos de El Niño; marcas al pie = meses con Alerta ENFEN. Las series pueden empezar en fechas distintas; el eje arranca en el inicio común más reciente.</p>
      <div class="controls">
        <div><label>Indicador climático</label><select id="fenClima"></select></div>
        <div><label>Sector</label><select id="fenSector"></select></div>
        <div id="fenIndWrap"><label>Indicador financiero</label><select id="fenInd"></select></div>
        <div><label id="fenEntLbl">Entidades</label><select id="fenEnt" multiple size="3" style="min-width:200px"></select></div>
      </div>
      <div id="fenChart" style="width:100%;height:560px;"></div>
    </div>
  </div>
</div>

<footer>
  <b>Fuentes:</b> CENEPRED — Mapa de susceptibilidad y peligros · ENFEN — Monitoreo de El Niño / La Niña ·
  AFP Integra — Valorización de Portafolio.<br>
  <span id="footgen"></span> · Documento autocontenido y portable (logo y librerías embebidas / CDN Plotly).
</footer>

<script>
const PAYLOAD = __PAYLOAD__;
const GEO = __GEOJSON__;
const AFP = PAYLOAD.afp;
const D = PAYLOAD.districts;
const EMP_SECTOR = PAYLOAD.emp_sector || {};
const SECTORES = PAYLOAD.sectores || [];
const SECTOR_META = PAYLOAD.sector_meta || {};
const EXP_CLASE = PAYLOAD.exp_clase || {};
const AUM = PAYLOAD.aum_cartera || {TODOS:0,'RENTA VARIABLE':0,'RENTA FIJA':0};
const NIV_ALTOS = PAYLOAD.niveles_altos || [5,4,3];
const byU = {}; D.forEach(d=>byU[d.ubigeo]=d);

document.getElementById('logo').src = PAYLOAD.logo;
document.getElementById('fval').textContent = PAYLOAD.fecha_valoracion || '—';
document.getElementById('gen').textContent = 'Generado ' + PAYLOAD.generado;
document.getElementById('footgen').innerHTML =
  'Última actualización automática: <b>'+PAYLOAD.generado+'</b>';
document.getElementById('modoNivel').textContent =
  (PAYLOAD.nivel_modo==='max'?'el peligro MÁS ALTO':'el promedio (combinado)');

const fmtS = v => 'S/ ' + (v||0).toLocaleString('es-PE',{maximumFractionDigits:0});
const fmtN = v => (v||0).toLocaleString('es-PE',{maximumFractionDigits:2});
const fmtP = v => (v*100).toLocaleString('es-PE',{maximumFractionDigits:1})+'%';

// ---- estado global
let TAB = 'resumen';
const classKey = () => TAB==='rv'?'RENTA VARIABLE' : TAB==='rf'?'RENTA FIJA' : 'TODOS';
const sectorSel = document.getElementById('sector');
const curSector = () => sectorSel.value;

function classFrac(emisor){
  const ck = classKey();
  if(ck==='TODOS') return 1;
  const e = EXP_CLASE[emisor];
  if(!e || !e.TOTAL) return 0;
  return (e[ck]||0)/e.TOTAL;
}
// exposición efectiva de un distrito (filtra clase + sector) desde los aportes
function aporteDist(d){
  const sec=curSector();
  let exp=0; const empresas={};
  (d.aportes||[]).forEach(a=>{
    if(sec!=='todos' && a.sector!==sec) return;
    const amt = a.monto*classFrac(a.emisor);
    if(amt<=0) return;
    exp+=amt; empresas[a.emisor]=(empresas[a.emisor]||0)+amt;
  });
  return {exp, empresas};
}
function effExp(d){ return aporteDist(d).exp; }
function effHasExp(d){ return effExp(d)>0; }
function metricVal(d, metric){
  if(metric.indexOf('score_')===0) return d[metric];
  const e=effExp(d);
  if(metric==='exposicion') return e;
  const sk={mar_mm:'score_mm',mar_inu:'score_inu',mar_seq:'score_seq',mar_comb:'score_comb'}[metric];
  const sc=d[sk]; return sc? e*sc : 0;
}
function activeDistricts(){
  const pf=document.getElementById('provfilter').value;
  return pf? D.filter(d=>d.prov===pf): D;
}

// ================= TABLA PIVOTE (chart 1) =================
let pivSort={col:null,dir:-1};
function buildBuckets(){
  const sec=curSector();
  const bucket={};   // nivel -> emisor -> {amt, activos:[{n,dep,prov,dist}]}
  NIV_ALTOS.forEach(L=>bucket[L]={});
  D.forEach(d=>{
    const L=d.nivel;
    if(NIV_ALTOS.indexOf(L)<0) return;
    (d.aportes||[]).forEach(a=>{
      if(sec!=='todos' && a.sector!==sec) return;
      const amt=a.monto*classFrac(a.emisor);
      if(amt<=0) return;
      const b=bucket[L][a.emisor]||(bucket[L][a.emisor]={amt:0,activos:[],sector:a.sector});
      b.amt+=amt;
      const acts=(a.activos&&a.activos.length)?a.activos:['(activo s/d)'];
      acts.forEach(n=>b.activos.push({n:n,dep:d.dep,prov:d.prov,dist:d.dist}));
    });
  });
  return bucket;
}
function drawPivot(){
  const ck=classKey();
  const denom=AUM[ck]||1;
  const bucket=buildBuckets();
  const q=(document.getElementById('pivSearch').value||'').toLowerCase();
  // armar filas
  let rows=[];
  NIV_ALTOS.forEach(L=>{
    const ems=Object.entries(bucket[L]).sort((a,b)=>b[1].amt-a[1].amt);
    const aumL=ems.reduce((s,e)=>s+e[1].amt,0);
    ems.forEach(([emisor,b])=>{
      // agrupar activos por nombre
      const byA={};
      b.activos.forEach(x=>{const k=x.n; (byA[k]=byA[k]||{deps:new Set(),provs:new Set(),dists:new Set()});
        byA[k].deps.add(x.dep);byA[k].provs.add(x.prov);byA[k].dists.add(x.dist);});
      const activos=Object.keys(byA);
      const deps=[...new Set(b.activos.map(x=>x.dep))];
      const provs=[...new Set(b.activos.map(x=>x.prov))];
      const dists=[...new Set(b.activos.map(x=>x.dist))];
      const emisorTot=(EXP_CLASE[emisor]&&EXP_CLASE[emisor].TOTAL)||0;
      rows.push({nivel:L, aumL, aumLpct:aumL/denom, emisor:emisor,
        emisorPct: emisorTot? b.amt/emisorTot:0, emisorAmt:b.amt,
        activos, deps, provs, dists});
    });
  });
  // búsqueda
  if(q){ rows=rows.filter(r=>(r.emisor+' '+r.activos.join(' ')+' '+r.deps.join(' ')+' '+
      r.provs.join(' ')+' '+r.dists.join(' ')).toLowerCase().includes(q)); }
  // orden opcional (respeta agrupación por nivel salvo que se ordene por una columna)
  if(pivSort.col){
    const f={emisor:r=>r.emisor, aumemisor:r=>r.emisorPct, aumtotal:r=>r.aumLpct,
             riesgo:r=>r.nivel}[pivSort.col];
    if(f) rows.sort((a,b)=>{const x=f(a),y=f(b);return (x<y?-1:x>y?1:0)*pivSort.dir;});
  }
  // construir tabla con rowspan en Riesgo y AUM Total (solo si no se ha roto el orden)
  const agrupado = !pivSort.col || pivSort.col==='riesgo';
  const cols=[['riesgo','Riesgo'],['aumtotal','AUM Total'],['emisor','Emisor'],
    ['aumemisor','AUM por Emisor'],['activo','Activo'],['dep','Departamento'],
    ['prov','Provincia'],['dist','Distrito']];
  let h='<thead><tr>'+cols.map(c=>`<th data-c="${c[0]}">${c[1]} <span class="ar"></span></th>`).join('')+'</tr></thead><tbody>';
  if(!rows.length){ h+=`<tr><td colspan="8" style="padding:18px;color:var(--plomo)">Sin emisores con riesgo CENEPRED 3–5 para este filtro.</td></tr>`; }
  // contar filas por nivel (para rowspan)
  const cntNivel={}, firstNivel={};
  if(agrupado){ rows.forEach((r,i)=>{cntNivel[r.nivel]=(cntNivel[r.nivel]||0)+1;
      if(firstNivel[r.nivel]===undefined) firstNivel[r.nivel]=i;}); }
  rows.forEach((r,i)=>{
    h+='<tr>';
    if(agrupado && firstNivel[r.nivel]===i){
      h+=`<td class="riesgo r${r.nivel}" rowspan="${cntNivel[r.nivel]}">${r.nivel}</td>`;
      h+=`<td class="aumtot" rowspan="${cntNivel[r.nivel]}"><span class="pct">${fmtP(r.aumLpct)}</span><span class="monto">${fmtS(r.aumL)}</span></td>`;
    } else if(!agrupado){
      h+=`<td class="riesgo r${r.nivel}">${r.nivel}</td>`;
      h+=`<td class="aumtot"><span class="pct">${fmtP(r.aumLpct)}</span><span class="monto">${fmtS(r.aumL)}</span></td>`;
    }
    h+=`<td><b>${r.emisor}</b><br><span class="monto" style="color:var(--plomo)">${EMP_SECTOR[r.emisor]||''}</span></td>`;
    h+=`<td class="num"><span class="pct">${fmtP(r.emisorPct)}</span><span class="monto">${fmtS(r.emisorAmt)}</span></td>`;
    h+=`<td class="activo-list">${r.activos.join('<br>')}</td>`;
    h+=`<td class="geo-list">${r.deps.join('<br>')}</td>`;
    h+=`<td class="geo-list">${r.provs.join('<br>')}</td>`;
    h+=`<td class="geo-list">${r.dists.join('<br>')}</td>`;
    h+='</tr>';
  });
  h+='</tbody>';
  const t=document.getElementById('pivTable'); t.innerHTML=h;
  t.querySelectorAll('th').forEach(th=>th.onclick=()=>{
    const c=th.dataset.c; if(c==='activo'||c==='dep'||c==='prov'||c==='dist')return;
    pivSort.dir=(pivSort.col===c)?-pivSort.dir:-1; pivSort.col=c; drawPivot();
  });
}

// ================= MAPA (chart 2) =================
function drawMap(){
  const metric=document.getElementById('metric').value;
  const isScore=metric.indexOf('score_')===0;
  const act=activeDistricts(); const actSet=new Set(act.map(d=>d.ubigeo));
  const baseLoc=[],baseTxt=[],dataLoc=[],dataZ=[],dataTxt=[];
  GEO.features.forEach(f=>{
    const u=f.properties.IDDIST; const d=byU[u]; baseLoc.push(u);
    let val=null;
    if(d && actSet.has(u)){ val=metricVal(d,metric);
      if(!isScore && !effHasExp(d)) val=null; if(val===undefined)val=null; }
    let txt;
    if(d&&(d.has_score||effHasExp(d))){
      const ap=aporteDist(d);
      const emp=Object.entries(ap.empresas).sort((a,b)=>b[1]-a[1])
        .map(e=>e[0]+' ('+(EMP_SECTOR[e[0]]||'')+'): '+fmtS(e[1])).join('<br>');
      txt=`<b>${d.dist}</b><br>${d.prov}, ${d.dep}<br><b>Nivel CENEPRED:</b> ${d.nivel??'s/d'} `+
        `(MM ${d.score_mm??'-'} · Inu ${d.score_inu??'-'} · Seq ${d.score_seq??'-'})`;
      if(ap.exp>0){ txt+=`<br>Exposición (${classKey()}): ${fmtS(ap.exp)}`+
        `<br>Plata en riesgo: ${fmtS(metricVal(d,'mar_comb'))}`+ (emp?`<br><b>Emisores:</b><br>${emp}`:''); }
      else txt+=`<br><i>sin exposición en esta vista</i>`;
    } else txt=`<b>${f.properties.NOMBDIST}</b><br>${f.properties.NOMBPROV}, ${f.properties.NOMBDEP}<br><i>sin score ni exposición</i>`;
    baseTxt.push(txt);
    if(val!==null){dataLoc.push(u);dataZ.push(val);dataTxt.push(txt);}
  });
  const colorscale=isScore?[[0,'#2c7bb6'],[.25,'#abd9e9'],[.5,'#ffffbf'],[.75,'#fdae61'],[1,'#d7191c']]
    :[[0,'#eaf6f8'],[.5,AFP.cyan],[1,AFP.azul]];
  const base={type:'choropleth',geojson:GEO,locations:baseLoc,z:baseLoc.map(()=>0),
    text:baseTxt,hoverinfo:'text',featureidkey:'properties.IDDIST',
    colorscale:[[0,'#eef1f6'],[1,'#eef1f6']],showscale:false,marker:{line:{color:'#b9c0cf',width:.4}}};
  const data={type:'choropleth',geojson:GEO,locations:dataLoc,z:dataZ,text:dataTxt,
    hoverinfo:'text',featureidkey:'properties.IDDIST',colorscale,
    zmin:isScore?1:undefined,zmax:isScore?5:undefined,marker:{line:{color:'#fff',width:.6}},
    colorbar:{title:{text:isScore?'Score':'S/',side:'right'},thickness:12,len:.8}};
  Plotly.react('mapwrap',[base,data],{geo:{fitbounds:'locations',visible:false,bgcolor:'rgba(0,0,0,0)'},
    margin:{l:0,r:0,t:6,b:0},paper_bgcolor:'#fff'},{displayModeBar:false,responsive:true});
}
function drawBar(){
  const metric=document.getElementById('metric').value;
  const isScore=metric.indexOf('score_')===0;
  const nivel=document.getElementById('nivel').value;
  const act=activeDistricts(); const agg={};
  act.forEach(d=>{const key=nivel==='dep'?d.dep:(d.dep+' / '+d.prov);
    if(!agg[key])agg[key]={sum:0,ssc:0,n:0};
    if(isScore){const v=d[metric];if(v!=null){agg[key].ssc+=v;agg[key].n+=1;}}
    else{const e=effExp(d);if(e>0)agg[key].sum+=metricVal(d,metric)||0;}});
  let rows=Object.entries(agg).map(([k,o])=>[k,isScore?(o.n?o.ssc/o.n:0):o.sum])
    .filter(r=>r[1]>0).sort((a,b)=>b[1]-a[1]).slice(0,10).reverse();
  Plotly.react('bar',[{type:'bar',orientation:'h',x:rows.map(r=>r[1]),y:rows.map(r=>r[0]),
    marker:{color:isScore?'#d7191c':AFP.azul},
    text:rows.map(r=>isScore?r[1].toFixed(2):fmtS(r[1])),textposition:'auto',
    hovertemplate:'%{y}<br>'+(isScore?'score %{x:.2f}':'S/ %{x:,.0f}')+'<extra></extra>'}],
    {margin:{l:150,r:14,t:6,b:22},paper_bgcolor:'#fff',plot_bgcolor:'#fff',
     xaxis:{showgrid:false,zeroline:false,range:isScore?[0,5]:undefined},yaxis:{automargin:true},
     title:{text:(nivel==='dep'?'Top departamentos':'Top provincias'),font:{size:12,color:AFP.azul},x:.01,y:.98}},
    {displayModeBar:false,responsive:true});
}

// ================= KPIs =================
function renderKPIs(){
  const ck=classKey();
  // exposición total ubicada en esta vista
  let expVista=0; D.forEach(d=>expVista+=effExp(d));
  // % en alto riesgo (3-5)
  let altoR=0; D.forEach(d=>{ if(NIV_ALTOS.indexOf(d.nivel)>=0) altoR+=effExp(d); });
  // depto con más plata en riesgo
  const byDep={}; D.forEach(d=>{const m=metricVal(d,'mar_comb'); if(m>0)byDep[d.dep]=(byDep[d.dep]||0)+m;});
  let topDep='—',tv=-1; Object.entries(byDep).forEach(([k,v])=>{if(v>tv){tv=v;topDep=k;}});
  const aumV=AUM[ck]||0;
  const k=[
    ['AUM con activos físicos', fmtS(expVista), (aumV? fmtP(expVista/aumV):'')+' del AUM '+(ck==='TODOS'?'total':ck.toLowerCase())],
    ['AUM en riesgo alto (CENEPRED 3–5)', fmtS(altoR), expVista? fmtP(altoR/expVista)+' de lo ubicado':''],
    ['Mayor plata en riesgo (depto.)', topDep, fmtS(tv)],
    ['AUM de la cartera ('+(ck==='TODOS'?'total':ck.toLowerCase())+')', fmtS(aumV), 'denominador de los %'],
  ];
  document.getElementById('kpis').innerHTML=k.map(x=>
    `<div class="kpi"><div class="v">${x[1]}</div><div class="l">${x[0]}<br><span style="color:#9aa1b3">${x[2]}</span></div></div>`).join('');
}

// ================= ENFEN =================
function enfenColor(estado){
  const e=(estado||'').toLowerCase();
  if(e.includes('niño costero')&&e.includes('alerta')) return '#c0392b';
  if(e.includes('niño')&&e.includes('alerta')) return '#e67e22';
  if(e.includes('niño')&&e.includes('vigil')) return '#f0b27a';
  if(e.includes('niña')&&e.includes('alerta')) return '#2471a3';
  if(e.includes('niña')&&e.includes('vigil')) return '#85c1e9';
  if(e.includes('no activo')||e.includes('no identificado')) return '#d5d8dc';
  return '#bdc3c7';
}
function drawENFEN(){
  const S=PAYLOAD.enfen||[];
  const banner=document.getElementById('enfBanner');
  if(!S.length){ banner.style.background=AFP.plomo;
    banner.innerHTML='<div>Sin datos ENFEN cargados. Revisa ENFEN_DIR en el script.</div>'; 
    document.getElementById('enfGrid').innerHTML=''; return; }
  const last=S[S.length-1];
  banner.style.background=enfenColor(last.estado);
  banner.innerHTML=`<div><div style="font-size:12px;opacity:.9">Estado actual del sistema de alerta (${last.fecha})</div>`+
    `<div class="big">${last.estado}</div><div style="font-size:12px;opacity:.9">Confianza: ${last.conf||'s/d'}</div></div>`;
  // grid año x mes
  const meses=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  const byYear={}; S.forEach(s=>{(byYear[s.anio]=byYear[s.anio]||{})[s.mes]=s;});
  const years=Object.keys(byYear).sort();
  let h='<thead><tr><th></th>'+meses.map(m=>`<th>${m}</th>`).join('')+'</tr></thead><tbody>';
  years.forEach(y=>{ h+=`<tr><th>${y}</th>`;
    for(let m=1;m<=12;m++){ const s=byYear[y][m];
      h+= s? `<td style="background:${enfenColor(s.estado)}" title="${s.fecha}: ${s.estado} (${s.conf})"></td>`
            : `<td style="background:#f4f5f8"></td>`; }
    h+='</tr>'; });
  document.getElementById('enfGrid').innerHTML=h+'</tbody>';
  // leyenda
  const leg=[['#c0392b','Alerta El Niño Costero'],['#e67e22','Alerta El Niño'],
    ['#f0b27a','Vigilancia El Niño'],['#2471a3','Alerta La Niña Costera'],
    ['#85c1e9','Vigilancia La Niña'],['#d5d8dc','No activo']];
  document.getElementById('enfLeg').innerHTML=leg.map(l=>`<span><i style="background:${l[0]}"></i>${l[1]}</span>`).join('');
  // kpis enfen
  const cnt={}; S.forEach(s=>{const c=enfenColor(s.estado);cnt[c]=(cnt[c]||0)+1;});
  const elnino=S.filter(s=>(s.estado||'').toLowerCase().includes('niño')).length;
  const elnina=S.filter(s=>(s.estado||'').toLowerCase().includes('niña')).length;
  const k=[['Estado vigente',last.estado.replace('Alerta de ','').replace('Vigilancia de ',''),last.fecha],
    ['Meses monitoreados',S.length,`${S[0].fecha} → ${last.fecha}`],
    ['Meses con El Niño',elnino,'alerta o vigilancia'],
    ['Meses con La Niña',elnina,'alerta o vigilancia']];
  document.getElementById('kpisENF').innerHTML=k.map(x=>
    `<div class="kpi"><div class="v" style="font-size:17px">${x[1]}</div><div class="l">${x[0]}<br><span style="color:#9aa1b3">${x[2]}</span></div></div>`).join('');
}

// ================= GRÁFICO CLIMA vs SECTOR =================
const FEN = PAYLOAD.fen_chart || {};
const FEN_COLORS = ['#1E2E6E','#00AECB','#e67e22','#27ae60','#8e44ad','#c0392b'];
function fenPopular(){
  const dc=document.getElementById('fenClima');
  (FEN.clima_orden||[]).forEach(k=>{const o=document.createElement('option');o.value=k;o.textContent=k;
    if(k===FEN.clima_default)o.selected=true; dc.appendChild(o);});
  if(!(FEN.clima_orden||[]).length){const o=document.createElement('option');o.textContent='(sin datos de clima)';dc.appendChild(o);}
  const ds=document.getElementById('fenSector');
  const secs=['Financiero'].concat(Object.keys(FEN.sectores||{}));
  secs.forEach(s=>{const o=document.createElement('option');o.value=s;
    o.textContent=(s==='Financiero'?'Financiero (bancos)':s); ds.appendChild(o);});
  const di=document.getElementById('fenInd');
  (FEN.sbs_indicadores||[]).forEach(k=>{const o=document.createElement('option');o.value=k;o.textContent=k;
    if(k===FEN.sbs_default)o.selected=true; di.appendChild(o);});
  fenPopularEnt();
}
function fenPopularEnt(){
  const sec=document.getElementById('fenSector').value;
  const ent=document.getElementById('fenEnt'); ent.innerHTML='';
  document.getElementById('fenIndWrap').style.display = sec==='Financiero'?'':'none';
  document.getElementById('fenEntLbl').textContent = sec==='Financiero'?'Bancos':'Empresas';
  let lista=[], def=[];
  if(sec==='Financiero'){ lista=Object.keys(FEN.sbs||{}); def=FEN.bancos_default||[]; }
  else { lista=Object.keys((FEN.sectores||{})[sec]||{});
         def=(FEN.sectores_default||{})[sec]||lista; }
  lista.sort();
  lista.forEach(k=>{const o=document.createElement('option');o.value=k;o.textContent=k;
    if(def.indexOf(k)>=0)o.selected=true; ent.appendChild(o);});
}
function ymToDate(s){ return s+'-01'; }
function drawFenChart(){
  if(!FEN || (!FEN.clima_orden && !FEN.sbs && !FEN.sectores)){
    document.getElementById('cardFen').style.display='none'; return; }
  const clima=document.getElementById('fenClima').value;
  const sec=document.getElementById('fenSector').value;
  const ind=document.getElementById('fenInd').value;
  const ents=[...document.getElementById('fenEnt').selectedOptions].map(o=>o.value);
  const esFin=(sec==='Financiero');
  const unidad=esFin?'%':(((FEN.sectores_unidad||{})[sec])||'');
  const traces=[];
  // --- series del sector (eje IZQUIERDO y) ---
  const fuente = esFin? (FEN.sbs||{}) : ((FEN.sectores||{})[sec]||{});
  ents.forEach((e,i)=>{
    const serie = esFin? ((fuente[e]||{})[ind]||[]) : (fuente[e]||[]);
    if(!serie||!serie.length) return;
    traces.push({x:serie.map(p=>ymToDate(p.f)),y:serie.map(p=>p.v),type:'scatter',mode:'lines',
      name:e,line:{color:FEN_COLORS[i%FEN_COLORS.length],width:1.8},xaxis:'x',yaxis:'y',
      hovertemplate:'%{x|%Y-%m}<br>'+e+': %{y:.2f} '+unidad+'<extra></extra>'});
  });
  // --- clima (eje DERECHO y2) ---
  const cs=(FEN.clima||{})[clima]||[];
  if(cs.length) traces.push({x:cs.map(p=>ymToDate(p.f)),y:cs.map(p=>p.v),type:'scatter',
    mode:'lines',name:clima+' (clima)',line:{color:'#1E2E6E',width:1.3},xaxis:'x',yaxis:'y2',
    hovertemplate:'%{x|%Y-%m}<br>'+clima+': %{y:.2f}<extra></extra>'});
  const shapes=[], anns=[];
  // --- Alertas ENFEN como ÁREAS SOMBREADAS (agrupando meses consecutivos) ---
  const al=(FEN.alertas||[]).slice().sort((a,b)=>a.f<b.f?-1:1);
  const idx=s=>{const p=s.split('-');return (+p[0])*12+(+p[1]);};
  const addMonth=s=>{let p=s.split('-'),y=+p[0],m=+p[1]+1;if(m>12){m=1;y++;}return y+'-'+String(m).padStart(2,'0');};
  const colFill={nino:'rgba(230,126,34,.16)',nina:'rgba(36,113,163,.16)'};
  let span=null;
  const flush=()=>{ if(span){ shapes.push({type:'rect',xref:'x',yref:'paper',
      x0:ymToDate(span.ini),x1:ymToDate(addMonth(span.fin)),y0:0,y1:1,
      fillcolor:colFill[span.tipo]||'rgba(150,150,150,.12)',line:{width:0},layer:'below'}); span=null; } };
  al.forEach(a=>{
    if(span && a.tipo===span.tipo && idx(a.f)===idx(span.fin)+1) span.fin=a.f;
    else { flush(); span={ini:a.f,fin:a.f,tipo:a.tipo}; }
  });
  flush();
  // proxies de leyenda para las alertas (cuadro de color, sin puntos visibles)
  ['nino','nina'].forEach(tp=>{ if(al.some(a=>a.tipo===tp))
    traces.push({x:[ymToDate(al[0].f)],y:[null],type:'scatter',mode:'markers',
      name:'Alerta ENFEN '+(tp==='nino'?'El Niño':'La Niña'),hoverinfo:'skip',showlegend:true,
      marker:{symbol:'square',size:13,color:tp==='nino'?'rgba(230,126,34,.55)':'rgba(36,113,163,.55)'}}); });
  // --- Eventos extraordinarios de El Niño: RECTÁNGULO SIN RELLENO, con borde ---
  (FEN.eventos||[]).forEach(ev=>{shapes.push({type:'rect',xref:'x',yref:'paper',
    x0:ymToDate(ev.ini),x1:ymToDate(ev.fin),y0:0,y1:1,fillcolor:'rgba(0,0,0,0)',
    line:{color:'#c0392b',width:1.5},layer:'above'});
    anns.push({x:ymToDate(ev.ini),y:1.02,xref:'x',yref:'paper',text:ev.nombre,showarrow:false,
      font:{size:9,color:'#c0392b'},xanchor:'left',yanchor:'bottom'});});
  // --- umbrales de categoría sobre el clima (y2), ETIQUETAS FUERA (margen derecho) ---
  const um=(FEN.clima_umbrales||{})[clima]||[];
  um.forEach(u=>{shapes.push({type:'line',xref:'paper',yref:'y2',x0:0,x1:1,y0:u[1],y1:u[1],
    line:{color:u[2],width:1,dash:'dot'}});
    anns.push({x:1.015,y:u[1],xref:'paper',yref:'y2',text:u[0],showarrow:false,
      font:{size:9,color:u[2]},xanchor:'left',yanchor:'middle'});});
  // --- eje X: arranca en el INICIO COMÚN más reciente entre las series de líneas ---
  const lineas=traces.filter(t=>t.mode==='lines' && t.x && t.x.length);
  let xr=null;
  if(lineas.length){
    const mins=lineas.map(t=>t.x[0]); const maxs=lineas.map(t=>t.x[t.x.length-1]);
    xr=[mins.reduce((a,b)=>a>b?a:b), maxs.reduce((a,b)=>a>b?a:b)];
  }
  const layout={margin:{l:52,r:150,t:24,b:30},paper_bgcolor:'#fff',plot_bgcolor:'#fff',
    showlegend:true,legend:{orientation:'h',y:-0.12,font:{size:10}},
    hovermode:'x unified',shapes:shapes,annotations:anns,
    xaxis:{type:'date',showgrid:false,range:xr},
    yaxis:{title:{text:unidad,font:{size:11,color:'#1E2E6E'}},zeroline:false,
      rangemode:'tozero',showgrid:true,gridcolor:'#eef0f5'},
    yaxis2:{title:{text:'',font:{size:11,color:'#1E2E6E'}},overlaying:'y',side:'right',
      zeroline:true,zerolinecolor:'#cfd4e0',showgrid:false,automargin:false}};
  Plotly.react('fenChart',traces,layout,{displayModeBar:false,responsive:true});
}

// ================= NAV / RENDER =================
function renderDatos(){
  document.getElementById('pivTitle').textContent =
    'Emisores por nivel de riesgo CENEPRED — '+(TAB==='rv'?'Renta Variable':TAB==='rf'?'Renta Fija':'Resumen (todo el universo)');
  const ck=classKey();
  document.getElementById('pivNote').innerHTML =
    'Riesgo = nivel CENEPRED (3–5). <b>AUM Total</b> = % de la plata en ese nivel sobre el AUM '+
    (ck==='TODOS'?'total de la cartera':'de '+ck.toLowerCase())+'. <b>AUM por Emisor</b> = % del AUM del emisor expuesto en ese nivel, sobre su AUM total en cartera. Clic en encabezados para ordenar.';
  renderKPIs(); drawPivot(); drawMap(); drawBar();
}
function showTab(t){
  TAB=t;
  document.querySelectorAll('.tab').forEach(b=>b.classList.toggle('active',b.dataset.tab===t));
  const isENF=(t==='enfen');
  document.getElementById('panelDatos').classList.toggle('hidden',isENF);
  document.getElementById('panelENFEN').classList.toggle('hidden',!isENF);
  if(isENF){ drawENFEN(); drawFenChart(); } else renderDatos();
}
document.querySelectorAll('.tab').forEach(b=>b.onclick=()=>showTab(b.dataset.tab));

// toolbar del gráfico clima-financiero
fenPopular();
['fenClima','fenInd','fenEnt'].forEach(id=>document.getElementById(id).addEventListener('change',drawFenChart));
document.getElementById('fenSector').addEventListener('change',()=>{fenPopularEnt();drawFenChart();});

// poblar sector + provincias
SECTORES.forEach(s=>{const o=document.createElement('option');o.value=s;o.textContent=s;sectorSel.appendChild(o);});
(function(){const provs=[...new Set(D.filter(d=>effHasExp(d)||d.has_score).map(d=>d.prov))].filter(Boolean).sort();
  const sel=document.getElementById('provfilter');
  provs.forEach(p=>{const o=document.createElement('option');o.value=p;o.textContent=p;sel.appendChild(o);});})();

['sector','metric','nivel','provfilter'].forEach(id=>document.getElementById(id).addEventListener('change',()=>{renderKPIs();drawPivot();drawMap();drawBar();}));
document.getElementById('pivSearch').addEventListener('input',drawPivot);

// montos toggle
document.getElementById('chkMonto').addEventListener('change',e=>{
  document.body.classList.toggle('ocultar-montos',!e.target.checked);});
document.body.classList.add('ocultar-montos'); // por defecto ocultos (celda limpia)

// modo analista/comité
function setModo(comite){
  document.body.classList.toggle('modo-comite',comite);
  document.getElementById('mComite').classList.toggle('on',comite);
  document.getElementById('mAnalista').classList.toggle('on',!comite);
  if(TAB==='enfen')drawENFEN(); else {drawMap();drawBar();}
}
document.getElementById('mComite').onclick=()=>setModo(true);
document.getElementById('mAnalista').onclick=()=>setModo(false);

showTab('resumen');
</script>
</body></html>
"""


# ==============================================================================
# 10) MAIN
# ==============================================================================

def main():
    log("=" * 60)
    log("ANÁLISIS DE RIESGOS FÍSICOS DE CARTERA (MINERÍA + FINANCIERO) — INICIO")
    log("=" * 60)

    prod_path = find_file(PROD_DIR, PROD_PATTERNS, "PRODUCCIÓN MINERA")
    risk_path = find_file(RISK_DIR, RISK_PATTERNS, "RIESGOS FÍSICOS (CENEPRED + Valorización)")
    log(f"Producción : {prod_path}")
    log(f"Riesgos    : {risk_path}")

    geojson = load_geojson()
    geo_idx = build_geo_index(geojson)

    sheet_cen = pick_sheet(risk_path, SHEET_CENEPRED_HINTS, "CENEPRED")
    sheet_val = pick_sheet(risk_path, SHEET_VALORIZ_HINTS, "Valorización de Instrumentos")
    scores = load_cenepred(risk_path, sheet_cen)

    # ---------- SECTOR MINERÍA ----------
    df_prod, _ = load_produccion(prod_path)
    exp = load_exposicion(risk_path, sheet_val)
    districts_min, df_assign, um, detalle, no_map, dfm = build_model(
        df_prod, exp, scores, geo_idx, EMPRESAS_CANONICAS, SECTOR_MINERIA)

    # mapa empresa -> sector y exposición combinada (arranca con minería)
    emp_sector = {e: SECTOR_MINERIA for e in EMPRESAS_CANONICAS}
    exp_comb = dict(exp)
    listas = [districts_min]

    # acumuladores para el Excel (mantienen separación por nombre de empresa)
    df_assign_all = df_assign.copy()
    um_all = dict(um)
    detalle_all = dict(detalle)
    no_map_all = no_map.copy()
    universo_all = list(EMPRESAS_CANONICAS)

    # ---------- SECTOR FINANCIERO (BANCOS) ----------
    if INCLUIR_FINANCIERO:
        try:
            fin_path = find_file(FIN_DIR, FIN_PATTERNS, "RESUMEN DE BANCOS (créditos)")
            log(f"Financiero : {fin_path}")
            df_fin, _ = load_financiero(fin_path)
            exp_fin = load_exposicion_financiero(risk_path, sheet_val)
            districts_fin, df_assign_f, um_f, detalle_f, no_map_f, _ = build_model(
                df_fin, exp_fin, scores, geo_idx, EMPRESAS_FIN_CANONICAS, SECTOR_FINANCIERO)

            listas.append(districts_fin)
            for e in EMPRESAS_FIN_CANONICAS:
                emp_sector[e] = SECTOR_FINANCIERO
            exp_comb.update(exp_fin)
            df_assign_all = pd.concat([df_assign_all, df_assign_f], ignore_index=True)
            um_all.update(um_f)
            detalle_all.update(detalle_f)
            no_map_all = pd.concat([no_map_all, no_map_f], ignore_index=True)
            universo_all += list(EMPRESAS_FIN_CANONICAS)
        except SystemExit:
            raise
        except Exception as e:
            log(f"No se pudo procesar el sector financiero: {e}\n"
                f"      -> El dashboard se generará solo con minería. Revisa FIN_DIR / "
                f"FIN_PATTERNS / nombres en CONFIG.", "WARN")

    # ---------- helper para enganchar un sector extra ----------
    def _add_sector(df_act, exp_s, universo_s, sector_s):
        nonlocal exp_comb, df_assign_all, no_map_all, universo_all
        d_s, da_s, um_s, det_s, nm_s, _ = build_model(
            df_act, exp_s, scores, geo_idx, universo_s, sector_s)
        listas.append(d_s)
        for e in universo_s:
            emp_sector[e] = sector_s
        exp_comb.update(exp_s)
        df_assign_all = pd.concat([df_assign_all, da_s], ignore_index=True)
        um_all.update(um_s)
        detalle_all.update(det_s)
        no_map_all = pd.concat([no_map_all, nm_s], ignore_index=True)
        universo_all += list(universo_s)

    # ---------- SECTOR UTILITIES (GENERACIÓN ELÉCTRICA) ----------
    if INCLUIR_UTILITIES:
        try:
            util_path = find_file(UTIL_DIR, UTIL_PATTERNS, "BD UTILITIES (potencia)")
            log(f"Utilities  : {util_path}")
            df_util, _ = load_utilities(util_path, risk_path)
            exp_util = load_exposicion_sector(
                risk_path, sheet_val, EQUIV_EXPOSICION_UTILITIES,
                EMPRESAS_UTIL_CANONICAS, "utilities")
            _add_sector(df_util, exp_util, EMPRESAS_UTIL_CANONICAS, SECTOR_UTILITIES)
        except SystemExit:
            raise
        except Exception as e:
            log(f"No se pudo procesar el sector Utilities: {e}\n"
                f"      -> Se continúa sin él. Revisa UTIL_DIR / UTIL_PATTERNS / la hoja "
                f"'BD ptf' / 'Centrales Electricas' / nombres en CONFIG.", "WARN")

    # ---------- SECTOR CONSTRUCCIÓN (CEMENTO) ----------
    if INCLUIR_CONSTRUCCION:
        try:
            cons_path = find_file(CONS_DIR, CONS_PATTERNS, "BD CONSTRUCCIÓN (despachos)")
            log(f"Construcción: {cons_path}")
            df_cons, _ = load_construccion(cons_path, risk_path)
            exp_cons = load_exposicion_sector(
                risk_path, sheet_val, EQUIV_EXPOSICION_CONSTRUCCION,
                EMPRESAS_CONS_CANONICAS, "construcción")
            _add_sector(df_cons, exp_cons, EMPRESAS_CONS_CANONICAS, SECTOR_CONSTRUCCION)
        except SystemExit:
            raise
        except Exception as e:
            log(f"No se pudo procesar el sector Construcción: {e}\n"
                f"      -> Se continúa sin él. Revisa CONS_DIR / CONS_PATTERNS / la hoja "
                f"'Despachos' / ubicación de plantas / nombres en CONFIG.", "WARN")

    # ---------- EMPRESAS CON HOJA PROPIA (% distribución) ----------
    if INCLUIR_EMPRESAS_HOJA:
        try:
            import openpyxl
            wb_r = openpyxl.load_workbook(risk_path, read_only=True)
            sheetnames = list(wb_r.sheetnames)
            wb_r.close()
            equiv_hoja = {nm: spec["empresa"]
                          for spec in EMPRESAS_HOJA for nm in spec["exp"]}
            universo_hoja = [spec["empresa"] for spec in EMPRESAS_HOJA]
            exp_hoja = load_exposicion_sector(
                risk_path, sheet_val, equiv_hoja, universo_hoja, "empresas-hoja")
            log("Integrando empresas con hoja propia (peso = % de distribución):")
            por_sector, sin_hoja, sin_exp = {}, [], []
            for spec in EMPRESAS_HOJA:
                hoja_real = _buscar_hoja(sheetnames, spec["hoja"])
                if not hoja_real:
                    sin_hoja.append(spec["hoja"]); continue
                df_e = load_empresa_hoja(risk_path, spec, hoja_real)
                if df_e.empty:
                    continue
                if exp_hoja.get(spec["empresa"], 0) <= 0:
                    sin_exp.append(spec["empresa"])
                por_sector.setdefault(spec["sector"], {"dfs": [], "univ": []})
                por_sector[spec["sector"]]["dfs"].append(df_e)
                por_sector[spec["sector"]]["univ"].append(spec["empresa"])
            for sector_s, d in por_sector.items():
                df_s = pd.concat(d["dfs"], ignore_index=True)
                _add_sector(df_s, exp_hoja, d["univ"], sector_s)
            n_emp = sum(len(d["univ"]) for d in por_sector.values())
            log(f"Empresas-hoja integradas: {n_emp} en {len(por_sector)} sector(es) "
                f"({', '.join(sorted(por_sector))}).")
            if sin_hoja:
                log(f"      Sin hoja encontrada (revisar nombre): {sin_hoja}", "WARN")
            if sin_exp:
                log(f"      Sin AUM en Valorización (col H): {sin_exp}", "WARN")
        except SystemExit:
            raise
        except Exception as e:
            import traceback
            log(f"No se pudieron integrar las empresas-hoja: {e}\n"
                f"{traceback.format_exc()}", "WARN")

    # ---------- FUSIÓN POR DISTRITO ----------
    districts = merge_sectores(listas)

    # ---------- clase de activo (RF/RV) y AUM de cartera ----------
    exp_clase, aum_cartera = load_exposicion_clase(risk_path, sheet_val)

    # ---------- ENFEN (Fenómenos Geofísicos) ----------
    enfen = []
    if INCLUIR_ENFEN:
        try:
            enfen, _ = load_enfen()
        except Exception as e:
            log(f"No se pudo procesar ENFEN: {e}", "WARN")

    # ---------- Gráfico clima vs. indicador sectorial ----------
    fen_chart = {}
    if INCLUIR_FEN_CHART:
        try:
            clima, clima_orden = load_clima()
            sbs, sbs_inds, sbs_default = load_sbs()
            # series mensuales por sector (para el panel inferior cuando NO es banca)
            sectores_serie = {}
            unidades = {}
            min_mensual = load_produccion_mensual()      # Minería (produccion_unidades, ->2020)
            if min_mensual:
                sectores_serie[SECTOR_MINERIA] = min_mensual; unidades[SECTOR_MINERIA] = "TMF"
            util_mensual = load_serie_utilities_mensual()  # Utilities (BD Utilities, ->2026)
            if util_mensual:
                sectores_serie[SECTOR_UTILITIES] = util_mensual; unidades[SECTOR_UTILITIES] = "MW"
            cons_mensual = load_serie_construccion_mensual()  # Construcción (BD Construcción, ->2026)
            if cons_mensual:
                sectores_serie[SECTOR_CONSTRUCCION] = cons_mensual; unidades[SECTOR_CONSTRUCCION] = "MT"
            defaults_ent = {SECTOR_MINERIA: EMPRESAS_MIN_DEFAULT,
                            SECTOR_UTILITIES: EMPRESAS_UTIL_CANONICAS,
                            SECTOR_CONSTRUCCION: EMPRESAS_CONS_CANONICAS}
            fen_chart = {
                "clima": clima, "clima_orden": clima_orden,
                "clima_default": CLIMA_DEFAULT,
                "clima_umbrales": CLIMA_UMBRALES,
                "sbs": sbs, "sbs_indicadores": sbs_inds,
                "sbs_default": sbs_default, "bancos_default": BANCOS_DEFAULT,
                "sectores": sectores_serie, "sectores_unidad": unidades,
                "sectores_default": defaults_ent,
                "eventos": EVENTOS_ELNINO,
                "alertas": _enfen_meses_alerta(enfen),
            }
        except Exception as e:
            log(f"No se pudo construir el gráfico clima-financiero: {e}", "WARN")

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    xlsx_out = os.path.join(OUT_DIR, f"Riesgos_Fisicos_Cartera_{stamp}.xlsx")
    html_out = os.path.join(OUT_DIR, f"Mapa_Riesgos_Fisicos_{stamp}.html")

    export_excel(xlsx_out, df_assign_all, exp_comb, scores, geo_idx, um_all,
                 detalle_all, districts, no_map_all,
                 universo=universo_all, emp_sector=emp_sector)
    log(f"Excel generado : {xlsx_out}")
    export_html(html_out, geojson, districts, exp_comb, scores, geo_idx,
                emp_sector=emp_sector, exp_clase=exp_clase, aum_cartera=aum_cartera,
                enfen=enfen, fen_chart=fen_chart)
    log(f"HTML generado  : {html_out}")
    
    
    # ====== EXTENSIÓN RF — PÉRDIDA ESPERADA (módulo aparte, 100% aditivo) ======
    try:
        import rf_extension as rfx
        try:
            _enfen_path = find_file(ENFEN_DIR, ENFEN_PATTERNS, "ENFEN (extensión RF)")
        except SystemExit:
            _enfen_path = None
        try:
            _indic_path = find_file(FEN_IND_DIR, CLIMA_PATTERNS, "Indicadores (extensión RF)")
        except SystemExit:
            _indic_path = None
        rfx.generar_extension(
            out_dir=OUT_DIR, districts=districts, exp_clase=exp_clase,
            emp_sector=emp_sector, risk_path=risk_path, sheet_val=sheet_val,
            equiv_canonico=_equiv_exposicion_all(),
            enfen_path=_enfen_path, indic_path=_indic_path,
            stamp=stamp, logfn=lambda m: log(m))
    except Exception as e:
        log(f"No se pudo generar la extensión RF (el dashboard original no se afecta): {e}", "WARN")
        
    # ====== EXTENSIÓN ACTIVOS FÍSICOS — DASHBOARD LEAFLET (módulo aparte, 100% aditivo) ======
    try:
        import activos_extension3 as actx
        actx.generar_dashboard_activos(
            out_dir=OUT_DIR,
            districts=districts,
            exp_clase=exp_clase,
            aum_cartera=aum_cartera,
            emp_sector=emp_sector,
            risk_path=risk_path,
            geojson=geojson,
            enfen=enfen,
            stamp=stamp,
            mitigantes_path=r"C:\Users\usuario\OneDrive\Desktop\AFP INTEGRA\ESG\Riesgos Fisicos\2026\Inputs\Mitigantes.xlsx",
            logfn=lambda m: log(m))
    except Exception as e:
        log(f"No se pudo generar el dashboard de activos (el dashboard original no se afecta): {e}", "WARN")
    
    
            
        
        
        

    log("=" * 60)
    log("PROCESO COMPLETADO")
    log("=" * 60)
    return xlsx_out, html_out

    log("=" * 60)
    log("PROCESO COMPLETADO")
    log("=" * 60)
    return xlsx_out, html_out


if __name__ == "__main__":
    main()