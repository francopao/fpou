"""
ONPE Segunda Vuelta 2026 — Scraper DOM
Selectores confirmados funcionando:
  - S_VISTA:  mat-select[formcontrolname='region']
  - S_DEPTO:  mat-select[formcontrolname='department']
  - Tarjetas: .tarjeta-candidato--izquierda / --derecha
  - Actas:    ul.leyenda.vertical li

Fix v2: ElementClickInterceptedException
  → cdk-overlay-backdrop intercepta clicks cuando el panel anterior
    no se cerró del todo. Solución: esperar a que desaparezca el
    backdrop antes de cada acción, y usar JS click como fallback.
"""

import time, re, logging
from datetime import datetime
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    StaleElementReferenceException, ElementClickInterceptedException
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
log = logging.getLogger("ONPE")

URL     = "https://resultadosegundavuelta.onpe.gob.pe/main/resumen"
S_VISTA = "mat-select[formcontrolname='region']"
S_DEPTO = "mat-select[formcontrolname='department']"
S_LIMPIAR = "//button[contains(normalize-space(.),'LIMPIAR')]"
S_BACKDROP = "div.cdk-overlay-backdrop"

REGIONES = [
    "AMAZONAS","ÁNCASH","APURÍMAC","AREQUIPA","AYACUCHO",
    "CAJAMARCA","CALLAO","CUSCO","HUANCAVELICA","HUÁNUCO",
    "ICA","JUNÍN","LA LIBERTAD","LAMBAYEQUE","LIMA",
    "LORETO","MADRE DE DIOS","MOQUEGUA","PASCO","PIURA",
    "PUNO","SAN MARTÍN","TACNA","TUMBES","UCAYALI"
]

# ── Parsers ────────────────────────────────────────────────────────

def _pct(txt):
    c = re.sub(r"[^\d.,]", "", txt).replace(",", ".")
    try:    return float(c)
    except: return 0.0

def _votos(txt):
    limpio = txt.replace("'","").replace(".","").replace(",","")
    nums = re.findall(r"\d+", limpio)
    return int("".join(nums)) if nums else 0

def _num_par(txt):
    m = re.search(r"\((\d+)\)", txt)
    return int(m.group(1)) if m else 0

def _txt(parent, css):
    try:    return parent.find_element(By.CSS_SELECTOR, css).text.strip()
    except: return ""

# ── Driver ─────────────────────────────────────────────────────────

def make_driver(headless=True):
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1440,900")
    opts.add_argument("--lang=es-PE")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    return webdriver.Chrome(options=opts)

# ── Utilidades de espera ───────────────────────────────────────────

def wait_for(driver, css, timeout=20):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, css)))
        return True
    except TimeoutException:
        return False

def wait_gone(driver, css, timeout=5):
    """Espera hasta que el elemento CSS desaparezca del DOM."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, css)))
    except TimeoutException:
        pass

def esperar_sin_backdrop(driver, timeout=6):
    """
    Espera hasta que el cdk-overlay-backdrop desaparezca.
    Es lo que causa ElementClickInterceptedException.
    """
    wait_gone(driver, S_BACKDROP, timeout)
    time.sleep(0.2)

# ── Click robusto ──────────────────────────────────────────────────

def click_robusto(driver, element):
    """
    Intenta click normal; si falla por intercepción usa JS click.
    """
    try:
        element.click()
        return True
    except ElementClickInterceptedException:
        log.debug("  click interceptado → usando JS click")
        try:
            driver.execute_script("arguments[0].click();", element)
            return True
        except Exception as e:
            log.debug(f"  JS click falló: {e}")
            return False
    except Exception as e:
        log.debug(f"  click falló: {e}")
        return False

# ── Interacción mat-select ─────────────────────────────────────────

def abrir_select(driver, css, timeout=12):
    """Abre el mat-select y espera el panel de opciones."""
    esperar_sin_backdrop(driver)
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, css)))
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.3)
        if not click_robusto(driver, el):
            return False
        # Esperar panel de opciones
        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "mat-option")))
        time.sleep(0.3)
        return True
    except TimeoutException:
        log.debug(f"  abrir_select timeout: {css}")
        return False
    except Exception as e:
        log.debug(f"  abrir_select error: {e}")
        return False

def elegir(driver, texto, timeout=8):
    """
    Elige la opción por texto en el panel abierto.
    Usa JS click si el click normal es interceptado.
    """
    for xpath in [
        f"//mat-option[normalize-space(.)='{texto}']",
        f"//mat-option[contains(normalize-space(.),'{texto}')]",
    ]:
        try:
            opt = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, xpath)))
            # Scroll a la opción antes de clickear
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", opt)
            time.sleep(0.2)
            if click_robusto(driver, opt):
                # Esperar que el backdrop desaparezca (panel cerrado)
                esperar_sin_backdrop(driver, timeout=5)
                time.sleep(3.0)   # Angular re-renderiza
                return True
        except TimeoutException:
            continue
        except Exception as e:
            log.debug(f"  elegir '{texto}': {e}")
            continue
    log.warning(f"  Opción no encontrada: '{texto}'")
    return False

def limpiar(driver):
    esperar_sin_backdrop(driver)
    try:
        btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, S_LIMPIAR)))
        click_robusto(driver, btn)
        esperar_sin_backdrop(driver)
        time.sleep(2.0)
    except Exception:
        pass

def seleccionar_vista(driver, vista):
    """Selecciona TODOS | PERÚ | EXTRANJERO si no está ya seleccionado."""
    els = driver.find_elements(By.CSS_SELECTOR, S_VISTA)
    if els and vista.upper() in els[0].text.upper():
        return True
    if abrir_select(driver, S_VISTA):
        return elegir(driver, vista)
    return False

# ── Leer DOM ───────────────────────────────────────────────────────

def leer_actas(driver):
    """Lee ul.leyenda.vertical: Contabilizadas, JEE, Pendientes."""
    meta = {"actas_contabilizadas": None, "actas_jee": None,
            "actas_pendientes": None, "actas_total": None,
            "pct_contabilizadas": None}
    try:
        items = driver.find_elements(By.CSS_SELECTOR, "ul.leyenda.vertical li")
        for li in items:
            txt = li.text.strip()
            n   = _num_par(txt)
            tl  = txt.lower()
            if   "contabilizad" in tl: meta["actas_contabilizadas"] = n
            elif "jee" in tl or "envío" in tl or "envio" in tl:
                meta["actas_jee"] = n
            elif "pendiente" in tl:    meta["actas_pendientes"] = n
        vals = [v for v in [meta["actas_contabilizadas"],
                             meta["actas_jee"],
                             meta["actas_pendientes"]] if v is not None]
        if vals:
            meta["actas_total"] = sum(vals)
            meta["pct_contabilizadas"] = round(
                100 * (meta["actas_contabilizadas"] or 0) / meta["actas_total"], 3
            ) if meta["actas_total"] else None
    except Exception:
        pass
    return meta

def leer_tarjetas(driver, ubigeo, nivel, extra=None):
    rows  = []
    ts    = datetime.now().isoformat()
    extra = extra or {}
    actas = leer_actas(driver)
    for css in [".tarjeta-candidato--izquierda", ".tarjeta-candidato--derecha"]:
        try:
            card    = driver.find_element(By.CSS_SELECTOR, css)
            nombre  = _txt(card, ".tarjeta-candidato__nombre")
            pct_txt = _txt(card, ".tarjeta-candidato__porcentaje")
            partido = _txt(card, ".tarjeta-candidato__organizacion")
            vt      = (_txt(card, ".tarjeta-candidato__votos.d-none-movil")
                    or _txt(card, ".tarjeta-candidato__votos"))
            if not nombre:
                continue
            rows.append({
                "nivel":      nivel,
                "ubigeo":     ubigeo,
                "timestamp":  ts,
                "candidato":  nombre,
                "partido":    partido,
                "votos":      _votos(vt),
                "porcentaje": _pct(pct_txt),
                **actas,
                **extra,
            })
        except Exception:
            continue
    return rows

# ══════════════════════════════════════════════════════════════════
# SCRAPING
# ══════════════════════════════════════════════════════════════════

def scrape(headless=True):
    log.info(f"\n{'═'*60}")
    log.info(f"ONPE — Regional + Extranjero [{datetime.now().strftime('%H:%M:%S')}]")
    log.info(f"{'═'*60}")

    driver = make_driver(headless)
    rows_regional   = []
    rows_extranjero = []

    try:
        driver.get(URL)

        # Espera inicial
        log.info("Esperando carga Angular...")
        cargado = False
        for css, t in [(".tarjeta-candidato", 60), ("mat-select", 60), ("app-root", 90)]:
            if wait_for(driver, css, t):
                log.info(f"  SPA lista [{css}]")
                cargado = True
                break
        if not cargado:
            log.error("SPA no cargó — verifica conexión")
            return {"regional": pd.DataFrame(), "extranjero": pd.DataFrame(),
                    "consolidado": pd.DataFrame()}
        time.sleep(2)

        # ── BLOQUE 1: 25 REGIONES ──────────────────────────────────
        log.info(f"\n{'─'*50}")
        log.info("BLOQUE 1: PERÚ → 25 Regiones")
        log.info(f"{'─'*50}")

        # Primera selección de PERÚ
        seleccionar_vista(driver, "PERÚ")
        if not wait_for(driver, S_DEPTO, timeout=15):
            log.error(f"  {S_DEPTO} no disponible")
        else:
            log.info(f"  {S_DEPTO} disponible ✓")

        for i, region in enumerate(REGIONES, 1):
            log.info(f"  [{i:02d}/25] {region}...")

            # Reset
            limpiar(driver)
            seleccionar_vista(driver, "PERÚ")
            if not wait_for(driver, S_DEPTO, timeout=12):
                log.warning(f"    {S_DEPTO} no disponible para {region}")
                continue

            # Abrir select REGIÓN
            if not abrir_select(driver, S_DEPTO):
                log.warning(f"    No se pudo abrir select REGIÓN para {region}")
                continue

            # Elegir región (con JS click fallback)
            if not elegir(driver, region):
                continue

            # Leer resultados
            rows = leer_tarjetas(driver, region, "regional",
                                  {"departamento": region})
            if rows:
                rows_regional.extend(rows)
                for r in rows:
                    log.info(f"    {r['candidato'][:38]:38s} "
                             f"| {r['porcentaje']:6.3f}% "
                             f"| {r['votos']:>12,} votos")
                log.info(f"    Actas → contab={rows[0]['actas_contabilizadas']} "
                         f"JEE={rows[0]['actas_jee']} "
                         f"pend={rows[0]['actas_pendientes']} "
                         f"({rows[0]['pct_contabilizadas']}%)")
            else:
                log.warning(f"    Sin datos DOM para {region}")

        # ── BLOQUE 2: EXTRANJERO ───────────────────────────────────
        log.info(f"\n{'─'*50}")
        log.info("BLOQUE 2: EXTRANJERO")
        log.info(f"{'─'*50}")

        limpiar(driver)
        seleccionar_vista(driver, "EXTRANJERO")
        time.sleep(2.5)

        rows = leer_tarjetas(driver, "EXTRANJERO", "extranjero")
        if rows:
            rows_extranjero.extend(rows)
            for r in rows:
                log.info(f"  {r['candidato'][:38]:38s} "
                         f"| {r['porcentaje']:6.3f}% "
                         f"| {r['votos']:>12,} votos")
        else:
            log.warning("  Sin datos DOM para EXTRANJERO")

    finally:
        driver.quit()
        log.info("\nBrowser cerrado.")

    df_regional   = pd.DataFrame(rows_regional)
    df_extranjero = pd.DataFrame(rows_extranjero)
    partes = [df for df in [df_regional, df_extranjero] if not df.empty]
    df_consolidado = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()

    log.info(f"\n{'═'*60}  RESUMEN:")
    log.info(f"  df_regional  : {len(df_regional):3d} filas | "
             f"{df_regional['ubigeo'].nunique() if not df_regional.empty else 0} regiones")
    log.info(f"  df_extranjero: {len(df_extranjero):3d} filas")
    log.info(f"  df_consolidado:{len(df_consolidado):3d} filas")
    log.info(f"{'═'*60}")

    return {"regional": df_regional, "extranjero": df_extranjero,
            "consolidado": df_consolidado}


def polling(headless=True, intervalo_seg=900, max_iter=None):
    hist = {"regional": [], "extranjero": []}
    it = 0
    log.info(f"Polling — intervalo={intervalo_seg}s")
    try:
        while True:
            if max_iter and it >= max_iter:
                break
            log.info(f"\n══ Snapshot #{it+1}  {datetime.now().strftime('%H:%M:%S')} ══")
            try:
                res = scrape(headless=headless)
                for k in ["regional","extranjero"]:
                    df = res.get(k, pd.DataFrame())
                    if not df.empty:
                        df = df.copy()
                        df["snapshot"]    = it + 1
                        df["snapshot_ts"] = datetime.now().isoformat()
                        hist[k].append(df)
                conso = res.get("consolidado", pd.DataFrame())
                if not conso.empty:
                    print(f"\n=== Snapshot #{it+1} ===")
                    print(conso.to_string(index=False))
            except Exception as e:
                log.error(f"Error snapshot #{it+1}: {e}")
            it += 1
            if not (max_iter and it >= max_iter):
                log.info(f"Próximo en {intervalo_seg}s... (Ctrl+C)")
                time.sleep(intervalo_seg)
    except KeyboardInterrupt:
        log.info("Polling detenido.")
    return {k: pd.concat(v, ignore_index=True) if v else pd.DataFrame()
            for k, v in hist.items()}


# ══════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":

    HEADLESS      = False   # True una vez que funcione bien
    MODO_POLLING  = False
    INTERVALO_SEG = 900
    MAX_ITER      = None

    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 200)
    pd.set_option("display.max_rows", 100)
    pd.set_option("display.float_format", "{:,.3f}".format)

    if MODO_POLLING:
        hist = polling(headless=HEADLESS, intervalo_seg=INTERVALO_SEG,
                       max_iter=MAX_ITER)
        df_regional    = hist["regional"]
        df_extranjero  = hist["extranjero"]
        partes = [df for df in [df_regional, df_extranjero] if not df.empty]
        df_consolidado = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()
    else:
        res            = scrape(headless=HEADLESS)
        df_regional    = res["regional"]
        df_extranjero  = res["extranjero"]
        df_consolidado = res["consolidado"]

    for nombre, df in [("REGIONAL",    df_regional),
                       ("EXTRANJERO",  df_extranjero),
                       ("CONSOLIDADO", df_consolidado)]:
        print(f"\n{'═'*70}")
        print(f"  df_{nombre.lower()}  "
              f"({len(df)} filas × {len(df.columns) if not df.empty else 0} cols)")
        print(f"{'═'*70}")
        if not df.empty:
            print(df.to_string(index=False))
        else:
            print("  (vacío)")
            
            
df_consolidado['timestamp'] = (pd.to_datetime(df_consolidado['timestamp'])
    .dt.floor('s'))

df_consolidado.to_excel(
    r"C:\Users\usuario\OneDrive\Desktop\AFP INTEGRA\Modelo Franco\Elecciones\df_consolidado.xlsx",
    index=False)
print(f"Guardado: {len(df_consolidado)} filas")