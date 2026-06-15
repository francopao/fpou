
"""
extraer_indicadores_franco.py
==============================
Extrae los índices ENSO desde múltiples fuentes y los exporta a un Excel,
de forma INCREMENTAL (lee lo que ya hay y solo añade/actualiza data nueva,
sin pisar hojas existentes).

Fuentes:
  1. ICEN       – IGP (mensual)
  2. ITCP       – IMARPE SIOFEN (mensual)
  3. LABCOS     – IMARPE SIOFEN (mensual)
  4. MEI v2     – NOAA PSL (bimestral)
  5. ONI v5     – NOAA CPC (trimestral móvil por año)
  6. NINA1      – NOAA PSL correlation (mensual, Niño 1+2)            ← NUEVO
  7. ERSST5     – NOAA CPC (mensual, SST y anomalías 4 regiones)      ← NUEVO

Salida:
  indicadores_franco  → dict de DataFrames
  Excel              → carpeta indicada en OUTPUT_PATH

Lógica incremental (clave del pedido):
  Cada hoja tiene una o más columnas "llave". Al guardar, se lee la hoja
  existente (si la hay), se concatena con lo descargado, se eliminan duplicados
  por la llave QUEDÁNDOSE CON EL ÚLTIMO (la descarga fresca gana, por si NOAA
  revisa cifras viejas) y se reordena. Así nunca se pierde historia y los meses
  nuevos se agregan solos conforme las fuentes los publiquen.
"""

import re
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from pathlib import Path

# ── Configuración ────────────────────────────────────────────────────────────
OUTPUT_PATH = Path(r"C:\Users\usuario\OneDrive\Desktop\AFP INTEGRA\ESG\Riesgos Fisicos\FEN\Indicadores")
OUTPUT_FILE = OUTPUT_PATH / "indicadores_franco.xlsx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0 Safari/537.36"
    )
}

# Llaves por hoja para la deduplicación incremental
LLAVES = {
    "ICEN":   ["fecha"],
    "ITCP":   ["fecha"],
    "LABCOS": ["fecha"],
    "MEIv2":  ["anio", "bimestre"],
    "ONI":    ["anio"],
    "NINA1":  ["fecha"],
    "ERSST5": ["fecha"],
}

# ── Helpers ──────────────────────────────────────────────────────────────────

def fetch_text(url: str) -> str:
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    return r.text


# ── 1. ICEN – IGP ────────────────────────────────────────────────────────────

def parse_icen(url: str) -> pd.DataFrame:
    """Formato (espacio-separado, sin header fijo):  AAAA  MM  valor"""
    raw = fetch_text(url)
    lines = [l.strip() for l in raw.splitlines() if l.strip()]

    records = []
    for line in lines:
        parts = line.split()
        if len(parts) >= 3:
            try:
                anio, mes, valor = int(parts[0]), int(parts[1]), float(parts[2])
                fecha = pd.Timestamp(year=anio, month=mes, day=1)
                records.append({"fecha": fecha, "ICEN": valor})
            except (ValueError, IndexError):
                continue

    return pd.DataFrame(records).sort_values("fecha").reset_index(drop=True)


# ── 2 & 3. ITCP / LABCOS – IMARPE SIOFEN ───────────────────────────────────

def parse_siofen(url: str, col_name: str) -> pd.DataFrame:
    """Archivos SIOFEN: año  mes  valor  (espacios/tabs). Ignora encabezados."""
    raw = fetch_text(url)
    lines = [l.strip() for l in raw.splitlines() if l.strip()]

    records = []
    for line in lines:
        parts = re.split(r"[\s,;]+", line)
        if len(parts) >= 3:
            try:
                anio = int(parts[0])
                mes  = int(parts[1])
                val  = float(parts[2])
                if 1900 < anio < 2100 and 1 <= mes <= 12:
                    fecha = pd.Timestamp(year=anio, month=mes, day=1)
                    records.append({"fecha": fecha, col_name: val})
            except (ValueError, IndexError):
                continue

    return pd.DataFrame(records).sort_values("fecha").reset_index(drop=True)


# ── 4. MEI v2 – NOAA PSL ─────────────────────────────────────────────────────

MEI_SEASONS = ["DJ", "JF", "FM", "MA", "AM", "MJ", "JJ", "JA", "AS", "SO", "ON", "ND"]

def parse_mei(url: str) -> pd.DataFrame:
    """Líneas de datos: año  val1 … val12 (bimestres). Faltantes: -999 / 999."""
    raw = fetch_text(url)
    lines = [l.strip() for l in raw.splitlines() if l.strip()]

    records = []
    for line in lines:
        parts = line.split()
        if len(parts) < 2:
            continue
        try:
            anio = int(parts[0])
            if not (1900 < anio < 2100):
                continue
        except ValueError:
            continue

        values = parts[1:]
        for i, season in enumerate(MEI_SEASONS):
            if i < len(values):
                try:
                    val = float(values[i])
                    if abs(val) > 900:   # missing
                        val = float("nan")
                    records.append({"anio": anio, "bimestre": season, "MEIv2": val})
                except ValueError:
                    pass

    return pd.DataFrame(records).sort_values(["anio", "bimestre"]).reset_index(drop=True)


# ── 5. ONI v5 – NOAA CPC ─────────────────────────────────────────────────────
#
# La página ONI_v5.php es HTML estático pero con varias tablas de maquetación.
# Elegimos la tabla con más celdas-año, aplanamos sus celdas a tokens y
# reconstruimos cada fila anual tomando 12 valores tras cada año.

ONI_SEASONS = ["DJF", "JFM", "FMA", "MAM", "AMJ", "MJJ",
               "JJA", "JAS", "ASO", "SON", "OND", "NDJ"]


def _es_anio(token: str) -> bool:
    return bool(re.fullmatch(r"\d{4}", token)) and 1950 <= int(token) <= 2100


def _parse_oni_html(html: str) -> pd.DataFrame:
    soup = BeautifulSoup(html, "lxml")

    tabla_datos, max_anios = None, 0
    for tabla in soup.find_all("table"):
        n = sum(1 for td in tabla.find_all("td")
                if _es_anio(td.get_text(strip=True)))
        if n > max_anios:
            tabla_datos, max_anios = tabla, n

    if tabla_datos is None or max_anios == 0:
        raise ValueError("No se encontró la tabla de datos ONI. "
                         "¿Cambió la estructura de la página?")

    tokens = [td.get_text(strip=True) for td in tabla_datos.find_all("td")]

    records = []
    anio_actual, valores = None, []

    def _flush():
        if anio_actual is not None:
            fila = {"anio": anio_actual}
            for i, s in enumerate(ONI_SEASONS):
                fila[s] = valores[i] if i < len(valores) else float("nan")
            records.append(fila)

    for tok in tokens:
        if _es_anio(tok):
            _flush()
            anio_actual, valores = int(tok), []
        elif anio_actual is not None and len(valores) < 12:
            try:
                valores.append(float(tok))
            except ValueError:
                pass
    _flush()

    return pd.DataFrame(records).sort_values("anio").reset_index(drop=True)


def parse_oni(url: str) -> pd.DataFrame:
    """Versión por defecto: requests."""
    return _parse_oni_html(fetch_text(url))


def parse_oni_selenium(url: str) -> pd.DataFrame:
    """Respaldo en Chrome visible (no se usa en main); reutiliza _parse_oni_html."""
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options

    opts = Options()
    opts.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=opts)
    try:
        driver.get(url)
        time.sleep(3)
        html = driver.page_source
    finally:
        driver.quit()
    return _parse_oni_html(html)


# ── 6. NINA1 – NOAA PSL correlation (NUEVO) ──────────────────────────────────
#
# Formato PSL "correlation" (el mismo patrón que usan muchos índices):
#   Línea 1:      AÑO_INI  AÑO_FIN          (solo 2 números -> se ignora)
#   Datos:        AÑO  ene feb … dic        (año + 12 valores mensuales)
#   Faltantes:    -99.99   (centinela; también aparece suelto al pie)
#   Pie:          líneas de descripción/URL -> se ignoran solas
#
# Lo dejamos en formato LARGO (fecha, NINA1) para que sea consistente con
# ICEN/ITCP/LABCOS y trivial de deduplicar por fecha.

def parse_psl_correlation(url: str, col_name: str) -> pd.DataFrame:
    raw = fetch_text(url)
    records = []
    for line in raw.splitlines():
        parts = line.split()
        # una fila de datos trae año + 12 meses = 13 tokens como mínimo
        if len(parts) < 13:
            continue
        try:
            anio = int(parts[0])
        except ValueError:
            continue
        if not (1850 < anio < 2100):
            continue
        for mes in range(1, 13):
            try:
                val = float(parts[mes])
            except (ValueError, IndexError):
                val = float("nan")
            if val <= -90:                 # centinela -99.99 = faltante
                val = float("nan")
            records.append({"fecha": pd.Timestamp(anio, mes, 1), col_name: val})

    return pd.DataFrame(records).sort_values("fecha").reset_index(drop=True)


# ── 7. ERSST5 – NOAA CPC (NUEVO) ─────────────────────────────────────────────
#
# Cabecera:  YR  MON  NINO1+2  ANOM  NINO3  ANOM  NINO4  ANOM  NINO3.4  ANOM
# Datos:     año mes + 8 valores (SST y anomalía de las 4 regiones Niño)
# Formato ANCHO mensual con llave 'fecha'.

ERSST_COLS = ["NINO1+2", "ANOM_1+2", "NINO3", "ANOM_3",
              "NINO4", "ANOM_4", "NINO3.4", "ANOM_3.4"]


def parse_ersst5(url: str) -> pd.DataFrame:
    raw = fetch_text(url)
    records = []
    for line in raw.splitlines():
        parts = line.split()
        if len(parts) < 10:                # año + mes + 8 valores
            continue
        try:
            anio = int(parts[0])
            mes  = int(parts[1])
        except ValueError:
            continue                        # salta la cabecera (YR MON …)
        if not (1850 < anio < 2100 and 1 <= mes <= 12):
            continue
        rec = {"fecha": pd.Timestamp(anio, mes, 1)}
        for i, c in enumerate(ERSST_COLS):
            try:
                rec[c] = float(parts[2 + i])
            except (ValueError, IndexError):
                rec[c] = float("nan")
        records.append(rec)

    return pd.DataFrame(records).sort_values("fecha").reset_index(drop=True)


# ── Lógica incremental ───────────────────────────────────────────────────────

def merge_incremental(df_new: pd.DataFrame, sheet: str, key_cols: list,
                      path: Path) -> pd.DataFrame:
    """
    Lee la hoja existente (si el archivo y la hoja existen), la concatena con la
    data fresca y elimina duplicados por las columnas llave quedándose con el
    ÚLTIMO registro (la descarga nueva manda). Devuelve el DataFrame consolidado.
    """
    combinado = df_new.copy()

    if path.exists():
        try:
            df_old = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
            if "fecha" in df_old.columns:
                df_old["fecha"] = pd.to_datetime(df_old["fecha"])
            # concatena viejo + nuevo; keep="last" -> el nuevo gana en revisiones
            combinado = pd.concat([df_old, df_new], ignore_index=True)
            nuevos = len(df_new)
            antes = len(df_old)
        except (ValueError, KeyError):
            antes, nuevos = 0, len(df_new)   # la hoja aún no existe
    else:
        antes, nuevos = 0, len(df_new)

    combinado = (combinado
                 .drop_duplicates(subset=key_cols, keep="last")
                 .sort_values(key_cols)
                 .reset_index(drop=True))

    agregados = len(combinado) - antes
    print(f"      [{sheet}] existentes: {antes} | descargados: {nuevos} | "
          f"total tras merge: {len(combinado)} (+{max(agregados,0)} nuevos)")
    return combinado


def guardar(indicadores: dict, path: Path):
    """
    Escribe cada hoja. Si el archivo ya existe usa modo 'append' reemplazando
    SOLO nuestras hojas (no toca otras que hayas agregado a mano). Si no existe,
    lo crea de cero.
    """
    path.parent.mkdir(parents=True, exist_ok=True)

    if path.exists():
        writer = pd.ExcelWriter(path, engine="openpyxl",
                                mode="a", if_sheet_exists="replace")
    else:
        writer = pd.ExcelWriter(path, engine="openpyxl", mode="w")

    with writer:
        for nombre, df in indicadores.items():
            df.to_excel(writer, sheet_name=nombre, index=False)
            print(f"  ✓ Hoja '{nombre}' guardada ({len(df)} filas)")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("=" * 64)
    print("  Extracción de Indicadores ENSO – Modelo Franco (incremental)")
    print("=" * 64)

    indicadores_franco = {}

    # 1. ICEN
    print("\n[1/7] Descargando ICEN (IGP)...")
    indicadores_franco["ICEN"] = parse_icen("http://met.igp.gob.pe/datos/ICEN.txt")

    # 2. ITCP
    print("[2/7] Descargando ITCP (IMARPE SIOFEN)...")
    indicadores_franco["ITCP"] = parse_siofen(
        "https://siofen-admin.imarpe.gob.pe/img/menuNivel2/SIOFEN_INDICES_ITCP_DATOS.txt",
        "ITCP")

    # 3. LABCOS
    print("[3/7] Descargando LABCOS (IMARPE SIOFEN)...")
    indicadores_franco["LABCOS"] = parse_siofen(
        "https://siofen-admin.imarpe.gob.pe/img/menuNivel2/SIOFEN_INDICES_LABCOS_DATOS.txt",
        "LABCOS")

    # 4. MEI v2
    print("[4/7] Descargando MEI v2 (NOAA PSL)...")
    indicadores_franco["MEIv2"] = parse_mei("https://psl.noaa.gov/enso/mei/data/meiv2.data")

    # 5. ONI v5
    print("[5/7] Descargando ONI v5 (NOAA CPC)...")
    indicadores_franco["ONI"] = parse_oni(
        "https://www.cpc.ncep.noaa.gov/products/analysis_monitoring/ensostuff/ONI_v5.php")

    # 6. NINA1  (NUEVO)
    print("[6/7] Descargando NINA1 (NOAA PSL correlation)...")
    indicadores_franco["NINA1"] = parse_psl_correlation(
        "https://psl.noaa.gov/data/correlation/nina1.data", "NINA1")

    # 7. ERSST5 (NUEVO)
    print("[7/7] Descargando ERSST5 (NOAA CPC)...")
    indicadores_franco["ERSST5"] = parse_ersst5(
        "https://www.cpc.ncep.noaa.gov/data/indices/ersst5.nino.mth.91-20.ascii")

    # ── Merge incremental contra el Excel existente ───────────────────────────
    print("\n── Consolidando con la data ya existente ──")
    consolidado = {}
    for nombre, df in indicadores_franco.items():
        consolidado[nombre] = merge_incremental(df, nombre, LLAVES[nombre], OUTPUT_FILE)

    # ── Guardar ───────────────────────────────────────────────────────────────
    print(f"\nGuardando → {OUTPUT_FILE}")
    guardar(consolidado, OUTPUT_FILE)

    print("\n✅ ¡Listo! Archivo actualizado sin pisar la historia.")
    print(f"   {OUTPUT_FILE}")
    return consolidado


if __name__ == "__main__":
    indicadores_franco = main()