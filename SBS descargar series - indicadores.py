"""
extraer_indicadores_sbs.py
==========================
Descarga los reportes mensuales de Indicadores Financieros de la Banca Múltiple
de la SBS (archivo B-2401), homologa nombres de bancos / indicadores / categorías
(que mutan año a año) y consolida todo en una base LARGA dentro de BD SBS.xlsx.

Patrón de URL (descubierto a partir de los archivos reales):
    https://intranet2.sbs.gob.pe/estadistica/financiera/{AAAA}/{Mes}/B-2401-{cod}{anio}.XLS
    - {Mes}  = nombre de carpeta (Enero, Febrero, ... Setiembre, ...)
    - {cod}  = código de 2 letras del mes (en, fe, ma, ab, my, jn, jl, ag, se, oc, no, di)
    - {anio} = AAAA (4 dígitos) en los años recientes, pero AA (2 dígitos) en los
               más antiguos (ej. 1998 -> 'en98'). Por eso probamos varias variantes.

Formato de salida (hoja 'indicadores SBS'):
    Fecha | Empresa | Tipo de indicador | Indicador | Valor
    -> formato largo: robusto a cambios de nombres y a la corrida mensual.
       Para vista ancha (indicadores en columnas) usa pivotar_ancho() al final.

IMPORTANTE: la parte de DESCARGA apunta a un host de intranet de la SBS, así que
solo correrá desde una máquina con acceso a esa red (la tuya). El parseo y la
homologación sí están probados contra los archivos reales 1998/2006/2025.
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
#  CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════════════════

BD_PATH  = Path(r"C:\Users\usuario\OneDrive\Desktop\AFP INTEGRA\ESG\Riesgos Fisicos\FEN\Indicadores\BD SBS.xlsx")
HOJA_BD  = "indicadores SBS"

ANIO_INI = 1998
ANIO_FIN = date.today().year
MES_FIN  = date.today().month        # hasta el mes actual del año en curso

URL_BASE = "https://intranet2.sbs.gob.pe/estadistica/financiera/{anio}/{carpeta}/B-2401-{cod}{suf}.XLS"

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                         "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"}

# Carpeta (mes en texto) y código de 2 letras del nombre de archivo
MESES = {
    1:  ("Enero",      "en"), 2:  ("Febrero",  "fe"), 3:  ("Marzo",     "ma"),
    4:  ("Abril",      "ab"), 5:  ("Mayo",     "my"), 6:  ("Junio",     "jn"),
    7:  ("Julio",      "jl"), 8:  ("Agosto",   "ag"), 9:  ("Setiembre", "se"),
    10: ("Octubre",    "oc"), 11: ("Noviembre","no"), 12: ("Diciembre", "di"),
}
# Variantes de carpeta que la SBS ha usado en distintas épocas
CARPETA_ALT = {"Setiembre": ["Setiembre", "Septiembre"]}


# ══════════════════════════════════════════════════════════════════════════════
#  HOMOLOGACIÓN  (extiende estos diccionarios cuando la consola lo pida)
# ══════════════════════════════════════════════════════════════════════════════

def _n(s: str) -> str:
    """Normaliza texto: quita NBSP, colapsa espacios, recorta."""
    return re.sub(r"\s+", " ", str(s).replace("\xa0", " ")).strip()

def _sin_tildes_upper(s: str) -> str:
    s = _n(s).upper()
    for a, b in [("Á","A"),("É","E"),("Í","I"),("Ó","O"),("Ú","U"),("Ü","U"),("Ñ","N")]:
        s = s.replace(a, b)
    return s

# ── Bancos ─────────────────────────────────────────────────────────────────────
# canónico -> lista de variantes (tal como aparecen en los archivos a lo largo
# de los años). El matching es por _sin_tildes_upper tras quitar '(con sucursales...)'.
EQUIV_BANCOS = {
    "BBVA":            ["Continental", "B. BBVA Perú", "BBVA", "Banco BBVA", "BBVA Perú"],
    "BCP":             ["Credito", "B. De Crédito del Perú", "De Crédito del Perú", "BCP"],
    "Interbank":       ["Interbank", "Interamericano de Finanzas Interbank"],
    "Scotiabank":      ["Scotiabank Perú", "Wiese Sudameris", "Sudamericano", "Wiese", "Scotiabank"],
    "BanBif":          ["B. Interamericano de Finanzas", "Interamericano de Finanzas",
                        "Interame-ricano", "Interamericano"],
    "Pichincha":       ["B. Pichincha", "Pichincha", "Financiero", "Del Trabajo"],
    "Mibanco":         ["Mibanco"],
    "Comercio / BANCOM":["BANCOM", "De Comercio", "Banco de Comercio"],
    "GNB":             ["B. GNB", "GNB"],
    "Falabella":       ["B. Falabella Perú", "Falabella"],
    "Santander":       ["B. Santander Perú", "Santander"],
    "Ripley":          ["B. Ripley", "Ripley"],
    "Alfin":           ["Alfin Banco", "Alfin"],
    "ICBC":            ["B. ICBC", "ICBC"],
    "Bank of China":   ["Bank of China"],
    "BCI":             ["B. BCI Perú", "BCI"],
    "Compartamos":     ["Compartamos Banco", "Compartamos"],
    "Citibank":        ["Citibank"],
    "Lima":            ["Lima"],
    "PROMEDIO BCA. MULTIPLE": ["Total Banca Múltiple", "Promedio Bca. Multiple",
                               "PROMEDIO BCA. MULTIPLE", "Total Banca Multiple"],
}
# índice invertido: variante_normalizada -> canónico
_BANCO_IDX = {}
for canon, variantes in EQUIV_BANCOS.items():
    for v in variantes:
        _BANCO_IDX[_sin_tildes_upper(v)] = canon

def homologar_banco(raw: str):
    base = _n(raw)
    base = re.sub(r"\(con sucursales.*?\)", "", base, flags=re.I).strip()
    base = re.sub(r"^B\.\s*", "", base)            # quita prefijo 'B. '
    base = base.replace("-", "")                    # une corte por guion (SUDAME-RICANO)
    key  = _sin_tildes_upper(base)
    if key in _BANCO_IDX:
        return _BANCO_IDX[key], True
    # intento por contiene (variante dentro del raw o viceversa)
    for vk, canon in _BANCO_IDX.items():
        if vk and (vk in key or key in vk):
            return canon, True
    return _n(raw), False                            # no reconocido

# ── Categorías ─────────────────────────────────────────────────────────────────
EQUIV_CATEGORIAS = {
    "SOLVENCIA":            ["SOLVENCIA"],
    "CALIDAD DE ACTIVOS":   ["CALIDAD DE ACTIVOS", "DE CALIDAD DE ACTIVOS Y SUFICIENCIA CAPITAL"],
    "EFICIENCIA Y GESTIÓN": ["EFICIENCIA Y GESTION", "DE GESTION"],
    "RENTABILIDAD":         ["RENTABILIDAD", "DE RENTABILIDAD"],
    "LIQUIDEZ":             ["LIQUIDEZ", "DE LIQUIDEZ"],
    "POSICIÓN EN MONEDA EXTRANJERA": ["POSICION EN MONEDA EXTRANJERA"],
    "ESTRUCTURA":           ["DE ESTRUCTURA"],
}
_CAT_IDX = {}
for canon, variantes in EQUIV_CATEGORIAS.items():
    for v in variantes:
        _CAT_IDX[_sin_tildes_upper(v).replace(":", "").replace("*", "").strip()] = canon

def es_categoria(label: str):
    key = _sin_tildes_upper(label).replace(":", "").replace("*", "").strip()
    return _CAT_IDX.get(key)          # devuelve canónico o None

# ── Indicadores ────────────────────────────────────────────────────────────────
# Reglas (regex sobre el nombre normalizado en MAYÚSCULAS sin tildes) -> código corto.
# El primer match gana. Lo no mapeado pasa con su nombre limpio + aviso.
REGLAS_INDICADOR = [
    (r"^UTILIDAD NETA ANUALIZADA ?/ ?PATRIMONIO PROMEDIO", "ROEA"),
    (r"^UTILIDAD NETA ANUALIZADA ?/ ?ACTIVO PROMEDIO",     "ROAA"),
    (r"^RATIO DE CAPITAL GLOBAL",                          "RCG t-1"),   # el (al dd/mm/aaaa) es t-1
    (r"^RATIO DE LIQUIDEZ M\.?N",                          "RL MN"),
    (r"^RATIO DE LIQUIDEZ M\.?E",                          "RL ME"),
]

def homologar_indicador(raw: str):
    limpio = _n(raw).rstrip(". ")                       # quita puntos de relleno finales
    limpio = re.sub(r"\.{2,}", "", limpio).strip()      # quita '......' internos
    clave  = _sin_tildes_upper(limpio)
    for patron, codigo in REGLAS_INDICADOR:
        if re.search(patron, clave):
            return codigo, True
    return limpio, False                                # nombre limpio sin código corto


# ══════════════════════════════════════════════════════════════════════════════
#  DESCARGA  (detección de formato por magic bytes)
# ══════════════════════════════════════════════════════════════════════════════

def _candidatos_url(anio: int, mes: int):
    carpeta_base, cod = MESES[mes]
    sufijos  = [f"{anio:04d}", f"{anio % 100:02d}"]     # 4 dígitos y 2 dígitos
    carpetas = CARPETA_ALT.get(carpeta_base, [carpeta_base])
    urls = []
    for carpeta in carpetas:
        for suf in sufijos:
            urls.append(URL_BASE.format(anio=anio, carpeta=carpeta, cod=cod, suf=suf))
    return urls

def descargar_xls(anio: int, mes: int):
    """Devuelve (bytes, url) del primer candidato válido, o (None, None)."""
    for url in _candidatos_url(anio, mes):
        try:
            r = requests.get(url, headers=HEADERS, timeout=30)
        except requests.RequestException:
            continue
        if r.status_code == 200 and r.content[:2] in (b"\xd0\xcf", b"PK"):
            return r.content, url
    return None, None


def leer_primera_hoja(contenido: bytes) -> pd.DataFrame:
    """Lee la 1ra hoja detectando BIFF (.xls real) vs OOXML (.xlsx disfrazado)."""
    eng = "openpyxl" if contenido[:2] == b"PK" else "xlrd"
    return pd.read_excel(io.BytesIO(contenido), sheet_name=0, header=None, engine=eng)


# ══════════════════════════════════════════════════════════════════════════════
#  PARSEO  de una hoja  ->  registros largos
# ══════════════════════════════════════════════════════════════════════════════

def _is_num(x):
    return isinstance(x, (int, float)) and not pd.isna(x)

def _detectar_layout(df: pd.DataFrame):
    """Devuelve (fila_bancos, col_categoria, col_indicador, {col: nombre_banco})."""
    best_r, best_cnt = None, 0
    for r in range(2, min(12, len(df))):
        cnt = sum(1 for c in range(1, df.shape[1])
                  if isinstance(df.iat[r, c], str) and _n(df.iat[r, c]))
        if cnt > best_cnt:
            best_r, best_cnt = r, cnt
    R = best_r
    fbc = next(c for c in range(1, df.shape[1])
               if isinstance(df.iat[R, c], str) and _n(df.iat[R, c]))
    # layout antiguo (1998): categoría e indicador en columnas distintas
    cat_col, ind_col = (fbc - 2, fbc - 1) if fbc >= 2 else (0, 0)
    bancos = {c: _n(df.iat[R, c]) for c in range(fbc, df.shape[1])
              if isinstance(df.iat[R, c], str) and _n(df.iat[R, c])}
    return R, cat_col, ind_col, bancos


def parsear_hoja(df: pd.DataFrame, fecha: pd.Timestamp, avisos: dict):
    """Convierte una hoja en una lista de dicts (registros largos) homologados."""
    R, cat_col, ind_col, bancos_cols = _detectar_layout(df)

    # Homologar bancos de esta hoja (y registrar los detectados para comparación mensual)
    bancos_homol = {}
    for c, raw in bancos_cols.items():
        canon, ok = homologar_banco(raw)
        bancos_homol[c] = canon
        if not ok:
            avisos["bancos_nuevos"].add(raw)
    avisos["bancos_mes"].update(bancos_homol.values())

    registros = []
    cat_actual = None
    for r in range(R + 1, df.shape[0]):
        etiqueta  = df.iat[r, ind_col] if ind_col < df.shape[1] else None
        cat_label = df.iat[r, cat_col] if cat_col < df.shape[1] else None
        valores   = {bancos_homol[c]: df.iat[r, c]
                     for c in bancos_cols if _is_num(df.iat[r, c])}

        # ¿es fila de categoría? (etiqueta sin valores numéricos)
        if not valores:
            for lab in (cat_label, etiqueta):
                if isinstance(lab, str) and _n(lab):
                    canon = es_categoria(lab)
                    if canon:
                        cat_actual = canon
                    elif _sin_tildes_upper(lab) == _sin_tildes_upper(_n(lab)) and \
                         lab.strip().isupper() and not lab.strip().startswith(("(", "*")) \
                         and "NOTA" not in _sin_tildes_upper(lab) and len(_n(lab)) < 60:
                        # parece un encabezado de categoría desconocido
                        avisos["categorias_nuevas"].add(_n(lab))
            continue

        # fila de indicador con datos
        if isinstance(etiqueta, str) and _n(etiqueta):
            ind_homol, ok = homologar_indicador(etiqueta)
            if not ok:
                avisos["indicadores_sin_codigo"].add(_n(etiqueta))
            for banco, val in valores.items():
                registros.append({
                    "Fecha": fecha,
                    "Empresa": banco,
                    "Tipo de indicador": cat_actual,
                    "Indicador": ind_homol,
                    "Valor": float(val),
                })
    return registros


# ══════════════════════════════════════════════════════════════════════════════
#  CONSOLIDACIÓN INCREMENTAL
# ══════════════════════════════════════════════════════════════════════════════

def cargar_bd():
    """Lee la hoja consolidada si existe; si no, devuelve DataFrame vacío."""
    cols = ["Fecha", "Empresa", "Tipo de indicador", "Indicador", "Valor"]
    if BD_PATH.exists():
        try:
            df = pd.read_excel(BD_PATH, sheet_name=HOJA_BD)
            df["Fecha"] = pd.to_datetime(df["Fecha"])
            return df[cols]
        except (ValueError, KeyError):
            print(f"  (existe {BD_PATH.name} pero sin hoja '{HOJA_BD}'; se creará)")
    return pd.DataFrame(columns=cols)


def guardar_bd(df: pd.DataFrame):
    BD_PATH.parent.mkdir(parents=True, exist_ok=True)
    df = df.sort_values(["Fecha", "Empresa", "Indicador"]).reset_index(drop=True)
    # modo de escritura: si el archivo existe, reemplaza solo la hoja
    if BD_PATH.exists():
        with pd.ExcelWriter(BD_PATH, engine="openpyxl", mode="a",
                            if_sheet_exists="replace") as w:
            df.to_excel(w, sheet_name=HOJA_BD, index=False)
    else:
        with pd.ExcelWriter(BD_PATH, engine="openpyxl") as w:
            df.to_excel(w, sheet_name=HOJA_BD, index=False)


def pivotar_ancho(df_largo: pd.DataFrame) -> pd.DataFrame:
    """Vista ancha: filas (Fecha, Empresa), columnas = indicadores."""
    return (df_largo
            .pivot_table(index=["Fecha", "Empresa"], columns="Indicador",
                         values="Valor", aggfunc="first")
            .reset_index())


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print("=" * 70)
    print("  Consolidación de Indicadores Financieros SBS – Banca Múltiple")
    print("=" * 70)

    bd = cargar_bd()
    fechas_existentes = set(bd["Fecha"].dt.normalize()) if len(bd) else set()
    print(f"BD actual: {len(bd)} filas | {len(fechas_existentes)} meses ya cargados")

    avisos = {"bancos_nuevos": set(), "categorias_nuevas": set(),
              "indicadores_sin_codigo": set(), "bancos_mes": set()}
    bancos_mes_anterior = None
    nuevos = []

    for anio in range(ANIO_INI, ANIO_FIN + 1):
        for mes in range(1, 13):
            if anio == ANIO_FIN and mes > MES_FIN:
                break
            fecha = pd.Timestamp(anio, mes, 1) + pd.offsets.MonthEnd(0)  # fin de mes
            if fecha.normalize() in fechas_existentes:
                continue   # ya está -> no rehacer trabajo

            contenido, url = descargar_xls(anio, mes)
            if contenido is None:
                # silencioso para meses futuros / inexistentes; avisa solo dentro de rango plausible
                if fecha <= pd.Timestamp(date.today()):
                    print(f"  [{anio}-{mes:02d}] sin archivo (probé {len(_candidatos_url(anio,mes))} URLs)")
                continue

            df_hoja = leer_primera_hoja(contenido)
            avisos["bancos_mes"] = set()
            regs = parsear_hoja(df_hoja, fecha, avisos)
            nuevos.extend(regs)

            # comparación de bancos vs mes anterior (entradas/salidas)
            if bancos_mes_anterior is not None:
                salieron = bancos_mes_anterior - avisos["bancos_mes"]
                entraron = avisos["bancos_mes"] - bancos_mes_anterior
                if salieron:
                    print(f"  [{anio}-{mes:02d}] ⚠ bancos que YA NO aparecen: {sorted(salieron)} "
                          f"(¿desaparecieron o cambió su nombre?)")
                if entraron:
                    print(f"  [{anio}-{mes:02d}] ⚠ bancos NUEVOS: {sorted(entraron)} "
                          f"(verifica equivalencias)")
            bancos_mes_anterior = set(avisos["bancos_mes"])

            print(f"  [{anio}-{mes:02d}] ✓ {len(regs)} registros  ({url.split('/')[-1]})")
            time.sleep(0.4)

    # ── Consolidar ────────────────────────────────────────────────────────────
    if nuevos:
        df_nuevos = pd.DataFrame(nuevos)
        bd = pd.concat([bd, df_nuevos], ignore_index=True)
        # dedup duro por si algo se reprocesó
        bd = bd.drop_duplicates(subset=["Fecha", "Empresa", "Indicador"], keep="last")
        guardar_bd(bd)
        print(f"\n✅ Agregados {len(df_nuevos)} registros nuevos. BD total: {len(bd)} filas.")
        print(f"   Guardado en: {BD_PATH}")
    else:
        print("\nNo hubo meses nuevos que procesar. BD intacta.")

    # ── Resumen de avisos para mantenimiento de diccionarios ───────────────────
    if avisos["bancos_nuevos"]:
        print("\n⚠ BANCOS sin homologar (agrégalos a EQUIV_BANCOS):")
        for b in sorted(avisos["bancos_nuevos"]):
            print(f"     - {b}")
    if avisos["categorias_nuevas"]:
        print("\n⚠ CATEGORÍAS no reconocidas (agrégalas a EQUIV_CATEGORIAS):")
        for c in sorted(avisos["categorias_nuevas"]):
            print(f"     - {c}")
    if avisos["indicadores_sin_codigo"]:
        print(f"\nℹ {len(avisos['indicadores_sin_codigo'])} indicadores sin código corto "
              f"(pasaron con su nombre limpio; agrégalos a REGLAS_INDICADOR si quieres siglas):")
        for i in sorted(avisos["indicadores_sin_codigo"])[:30]:
            print(f"     - {i}")

    return bd


if __name__ == "__main__":
    bd_sbs = main()