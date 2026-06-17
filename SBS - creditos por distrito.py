

"""
SBS_B-2314_creditos_depositos_distrital.py
==========================================
Descarga y consolida el reporte mensual B-2314 de la SBS:
"Créditos Directos y Depósitos de la Banca Múltiple por Zona Geográfica"
(desglose por Departamento / Provincia / Distrito).

Pensado para CORRERSE CADA MES: solo baja los meses nuevos (y re-baja el último
por si la SBS revisa cifras), y nunca se rompe si todavía no hay archivo nuevo.

Patrón de URL (igual que B-2401, descubierto de los archivos reales):
    https://intranet2.sbs.gob.pe/estadistica/financiera/{AAAA}/{Carpeta}/B-2314-{cod}{suf}.XLS
      {Carpeta} = Enero..Diciembre (con variante Setiembre/Septiembre)
      {cod}     = 2 letras del mes (en, fe, ma, ab, my, jn, jl, ag, se, oc, no, di)
      {suf}     = año de 4 dígitos (reciente) o 2 dígitos (antiguo). Se prueban ambos.

SALIDA — formato LARGO (robusto a que aparezcan/desaparezcan distritos o
indicadores en publicaciones futuras):
    Fecha | Departamento | Provincia | Distrito | Indicador | Moneda | Valor
  Indicador ∈ {Créditos Directos, Depósitos a la Vista, Depósitos de Ahorro,
               Depósitos a Plazo, Depósitos Totales, ...lo que traiga el archivo}
  Moneda    ∈ {MN, ME, Total}

BLINDAJE:
  * Detección dinámica de la fila de encabezado y de los bloques de indicador:
    si la SBS agrega un indicador o cambia el orden, se toma solo.
  * Las filas de subtotal ("Total Amazonas", "Total País", "Total general"...) se
    descartan solas.
  * Si un mes aún no está publicado -> se omite sin romper.
  * Si un archivo viene en el FORMATO ANTIGUO (pre-modernización, por región/ciudad
    sin distrito) -> se avisa y se omite (no se intenta forzar un parseo erróneo).
  * El guardado es incremental y no pisa otras hojas del Excel.

NOTA: el host es de la intranet de la SBS, así que corre desde tu red. El parseo
está probado contra el archivo real de mayo-2019.
"""

import re
import io
import time
import warnings
import requests
import pandas as pd
from pathlib import Path
from datetime import date

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURACIÓN  —  edita solo esto
# ══════════════════════════════════════════════════════════════════════════════

BD_PATH = Path(r"C:\Users\usuario\OneDrive\Desktop\AFP INTEGRA\ESG\Riesgos Fisicos\2026\Inputs\BD SBS B-2314 creditos depositos distrital.xlsx")
HOJA_BD = "B-2314"

ANIO_INI = 1998                       # primera corrida intenta desde aquí
FORZAR_RECARGA_TOTAL = False          # True = re-descarga todo desde ANIO_INI (ignora lo ya guardado)
PAUSA = 0.5                           # segundos entre descargas

URL_BASE = "https://intranet2.sbs.gob.pe/estadistica/financiera/{anio}/{carpeta}/B-2314-{cod}{suf}.XLS"

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                         "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"}

MESES = {
    1: ("Enero", "en"),  2: ("Febrero", "fe"),  3: ("Marzo", "ma"),
    4: ("Abril", "ab"),  5: ("Mayo", "my"),     6: ("Junio", "jn"),
    7: ("Julio", "jl"),  8: ("Agosto", "ag"),   9: ("Setiembre", "se"),
    10: ("Octubre", "oc"), 11: ("Noviembre", "no"), 12: ("Diciembre", "di"),
}
CARPETA_ALT = {"Setiembre": ["Setiembre", "Septiembre"]}


# ══════════════════════════════════════════════════════════════════════════════
#  Helpers
# ══════════════════════════════════════════════════════════════════════════════

def _n(s) -> str:
    return re.sub(r"\s+", " ", str(s).replace("\xa0", " ")).strip()

def _a_float(x) -> float:
    if isinstance(x, (int, float)):
        return float(x) if pd.notna(x) else float("nan")
    t = _n(x).replace(" ", "")
    if t in ("", "-", "--", "---", "n.d.", "nd", "N.D.", "S/I"):
        return float("nan")
    if "," in t and "." in t:        # 1.234,56  (miles . , decimal ,)
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:                   # 1234,56
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return float("nan")


# ── Descarga ──────────────────────────────────────────────────────────────────

def _candidatos_url(anio: int, mes: int):
    carpeta, cod = MESES[mes]
    carpetas = CARPETA_ALT.get(carpeta, [carpeta])
    sufijos = [f"{anio:04d}", f"{anio % 100:02d}"]      # 4 y 2 dígitos
    urls = []
    for carp in carpetas:
        for suf in sufijos:
            urls.append(URL_BASE.format(anio=anio, carpeta=carp, cod=cod, suf=suf))
    return urls

def descargar_xls(anio: int, mes: int):
    """(bytes, url) del primer candidato válido, o (None, None)."""
    for url in _candidatos_url(anio, mes):
        try:
            r = requests.get(url, headers=HEADERS, timeout=30)
        except requests.RequestException:
            continue
        if r.status_code == 200 and r.content[:2] in (b"\xd0\xcf", b"PK"):
            return r.content, url
    return None, None

def leer_primera_hoja(contenido: bytes) -> pd.DataFrame:
    eng = "openpyxl" if contenido[:2] == b"PK" else "xlrd"
    return pd.read_excel(io.BytesIO(contenido), sheet_name=0, header=None, engine=eng)


# ── Parseo del formato MODERNO (Departamento/Provincia/Distrito) ──────────────

def parsear_hoja(df: pd.DataFrame, fecha: pd.Timestamp):
    """Devuelve lista de dicts en formato largo, o None si no es el formato moderno."""
    # 1. fila de encabezado: contiene 'Departamento' y 'Distrito'
    hrow = None
    for r in range(min(20, df.shape[0])):
        vals = [_n(df.iat[r, c]).upper() for c in range(min(5, df.shape[1])) if pd.notna(df.iat[r, c])]
        if any(v.startswith("DEPARTAMENTO") for v in vals) and any("DISTRITO" in v for v in vals):
            hrow = r
            break
    if hrow is None:
        return None     # formato antiguo / desconocido

    # 2. bloques de indicador = celdas no nulas a la derecha del Distrito
    bloques = [(c, _n(df.iat[hrow, c])) for c in range(3, df.shape[1]) if pd.notna(df.iat[hrow, c])]
    if not bloques:
        return None
    MONEDAS = ["MN", "ME", "Total"]      # patrón fijo de la SBS: 3 sub-columnas por bloque

    # 3. recorrer filas de datos con arrastre de Departamento/Provincia
    recs = []
    dep = prov = None
    for r in range(hrow + 2, df.shape[0]):       # +2 salta la fila de sub-encabezados (MN/ME/Total)
        c0 = df.iat[r, 0] if df.shape[1] > 0 else None
        c1 = df.iat[r, 1] if df.shape[1] > 1 else None
        c2 = df.iat[r, 2] if df.shape[1] > 2 else None

        # subtotales y totales generales -> ignorar (y no arrastrar)
        if pd.notna(c0) and _n(c0).upper().startswith("TOTAL"):
            continue
        if pd.notna(c0) and _n(c0):
            dep = _n(c0); prov = None            # nuevo departamento -> resetea provincia
        if pd.notna(c1) and not _n(c1).upper().startswith("TOTAL"):
            prov = _n(c1)

        # fila de dato = tiene Distrito real
        if not (pd.notna(c2) and _n(c2) and not _n(c2).upper().startswith("TOTAL")):
            continue
        distrito = _n(c2)

        for cb, indicador in bloques:
            for k, moneda in enumerate(MONEDAS):
                c = cb + k
                if c < df.shape[1]:
                    val = _a_float(df.iat[r, c])
                    if pd.notna(val):
                        recs.append({
                            "Fecha": fecha, "Departamento": dep, "Provincia": prov,
                            "Distrito": distrito, "Indicador": indicador,
                            "Moneda": moneda, "Valor": val,
                        })
    return recs


# ── Persistencia incremental ──────────────────────────────────────────────────

COLS = ["Fecha", "Departamento", "Provincia", "Distrito", "Indicador", "Moneda", "Valor"]
LLAVE = ["Fecha", "Departamento", "Provincia", "Distrito", "Indicador", "Moneda"]

def cargar_bd() -> pd.DataFrame:
    if BD_PATH.exists() and not FORZAR_RECARGA_TOTAL:
        try:
            df = pd.read_excel(BD_PATH, sheet_name=HOJA_BD)
            df["Fecha"] = pd.to_datetime(df["Fecha"])
            return df[COLS]
        except (ValueError, KeyError):
            pass
    return pd.DataFrame(columns=COLS)

def guardar_bd(df: pd.DataFrame):
    BD_PATH.parent.mkdir(parents=True, exist_ok=True)
    df = df.sort_values(["Fecha", "Departamento", "Provincia", "Distrito", "Indicador", "Moneda"]).reset_index(drop=True)
    if BD_PATH.exists():
        with pd.ExcelWriter(BD_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
            df.to_excel(w, sheet_name=HOJA_BD, index=False)
    else:
        with pd.ExcelWriter(BD_PATH, engine="openpyxl") as w:
            df.to_excel(w, sheet_name=HOJA_BD, index=False)

def pivotar_ancho(df_largo: pd.DataFrame) -> pd.DataFrame:
    """Vista ancha opcional: una columna por (Indicador, Moneda)."""
    return (df_largo.pivot_table(index=["Fecha", "Departamento", "Provincia", "Distrito"],
                                 columns=["Indicador", "Moneda"], values="Valor", aggfunc="first")
                    .sort_index())


# ── Rango de meses a procesar ─────────────────────────────────────────────────

def _meses_a_procesar(bd: pd.DataFrame):
    """Genera (anio, mes) desde donde haga falta hasta el mes actual."""
    hoy = date.today()
    if bd.empty or FORZAR_RECARGA_TOTAL:
        ini_a, ini_m = ANIO_INI, 1
    else:
        ult = bd["Fecha"].max()
        ini_a, ini_m = ult.year, ult.month       # re-baja el último mes por si hubo revisión
    a, m = ini_a, ini_m
    while (a, m) <= (hoy.year, hoy.month):
        yield a, m
        m += 1
        if m > 12:
            m = 1; a += 1


# ══════════════════════════════════════════════════════════════════════════════
#  Main
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print("=" * 70)
    print("  SBS B-2314 – Créditos y Depósitos por Distrito (consolidado mensual)")
    print("=" * 70)

    bd = cargar_bd()
    print(f"BD actual: {len(bd)} filas"
          + (f" | hasta {bd['Fecha'].max():%b-%Y}" if not bd.empty else " (vacía, primera corrida)"))

    nuevos, avisos, omitidos_formato = [], [], []
    for anio, mes in _meses_a_procesar(bd):
        fecha = pd.Timestamp(anio, mes, 1)
        try:
            contenido, url = descargar_xls(anio, mes)
            if contenido is None:
                continue                                   # mes aún no publicado -> se omite en silencio
            df = leer_primera_hoja(contenido)
            regs = parsear_hoja(df, fecha)
            if regs is None:
                omitidos_formato.append(f"{anio}-{mes:02d}")   # formato antiguo/desconocido
                continue
            if not regs:
                avisos.append(f"{anio}-{mes:02d}: archivo leído pero 0 registros.")
                continue
            nuevos.extend(regs)
            print(f"  [{anio}-{mes:02d}] ✓ {len(regs):5} registros  ({url.split('/')[-1]})")
        except Exception as e:
            avisos.append(f"{anio}-{mes:02d}: ERROR {type(e).__name__}: {e}")
        time.sleep(PAUSA)

    if not nuevos:
        print("\nSin novedades: no se descargó data nueva. La BD queda igual.")
        if omitidos_formato:
            print(f"  ({len(omitidos_formato)} meses en formato antiguo/no soportado, omitidos)")
        for a in avisos:
            print("   ⚠", a)
        return bd

    df_nuevos = pd.DataFrame(nuevos)[COLS]
    bd = pd.concat([bd, df_nuevos], ignore_index=True)
    bd = bd.drop_duplicates(subset=LLAVE, keep="last")     # la descarga fresca gana en revisiones

    guardar_bd(bd)

    print("\n" + "-" * 70)
    print(f"Consolidado: {len(bd)} filas | {bd['Fecha'].nunique()} meses "
          f"({bd['Fecha'].min():%b-%Y} → {bd['Fecha'].max():%b-%Y})")
    print(f"Distritos: {bd['Distrito'].nunique()} | Indicadores: {bd['Indicador'].nunique()}")
    print(f"✅ Guardado: {BD_PATH}")
    if omitidos_formato:
        print(f"\nℹ {len(omitidos_formato)} meses en formato antiguo (por región/ciudad, sin distrito) "
              f"fueron omitidos: {omitidos_formato[:6]}{'...' if len(omitidos_formato)>6 else ''}")
        print("  (Si quieres backfillear esa era, te armo un segundo parser; mapea a región/ciudad, no a distrito.)")
    for a in avisos:
        print("   ⚠", a)
    return bd


if __name__ == "__main__":
    bd_sbs = main()