# -*- coding: utf-8 -*-
"""
================================================================================
rf_extension.py  —  EXTENSIÓN DE PÉRDIDA ESPERADA POR RIESGO FÍSICO
================================================================================
AFP Integra — módulo adicional del dashboard de Riesgos Físicos.

CONTRATO DE BLINDAJE
--------------------
NO modifica el pipeline (dashboard_riesgos_fisicos_final.py): no importa nada de
él, no redefine sus funciones, no altera sus estructuras. Recibe como INPUT de
solo lectura lo que el main() ya calculó y produce un ARCHIVO HTML COMPAÑERO
independiente. Se invoca AL FINAL del main(), después de export_html().

CONTENIDO
---------
  Cálculo:
    - calcular_perdida_esperada(...)  -> Pestaña 1 (daño de stock MM/Inundación
                                          + canal sequía de flujo + canal indirecto)
    - calcular_sensibilidad(...)      -> Pestaña 2 (damage_ratios ±30%)
    - cargar_enso(...) / calcular_enso(...) -> Pestaña 3 (multiplicador ENSO costa norte)
  Soporte:
    - cargar_activo_patrimonio(...)   -> lee Activo/Patrimonio (cols W/Y de Valorización)
    - detectar_canal_indirecto(...)   -> bancos / soberanos / extranjeros
  Salida:
    - construir_html_companion(...)   -> escribe el HTML autocontenido de 3 pestañas
    - generar_extension(...)          -> orquestador único llamado desde main()

CITAS (nota de proyecto)
------------------------
  Roncalli/Amundi (2026): damage functions Huizinga por activo; tabla hazard x activo.
  Azzone et al. (2025):   RF/RV vía estructura de capital (Merton); asset intensity.
  Bressan et al. (2024):  shock agudo/stock (activo) vs crónico/flujo (EBITDA); sequía=flujo.
  Chitchumnong et al. (2025): canal PD territorial para bancos (casos especiales).
================================================================================
"""

import os
import re
import json
import unicodedata
from datetime import datetime

import pandas as pd


# ==============================================================================
# 1) PARÁMETROS METODOLÓGICOS  — EDITABLES Y TRAZABLES
# ==============================================================================

# --- damage_ratio por BIN de score CENEPRED (1-5) — STOCK (MM, Inundación) ----
# Roncalli/Amundi (2026): el daño sobre el activo se ancla por analogía a las
# curvas de Huizinga (intensidad -> fracción de daño). Como CENEPRED entrega una
# CLASE de susceptibilidad 1-5 (no una intensidad física), se traduce clase ->
# fracción de daño con BINS. La progresión es NO LINEAL (convexa): el daño marginal
# se acelera en la cola alta de peligro (4->5 salta más que 1->2), consistente con
# la forma de las curvas Huizinga. Solo aplica a peligros de STOCK (Bressan 2024).
DAMAGE_RATIO_BINS = {
    "mm":  {1: 0.02, 2: 0.07, 3: 0.15, 4: 0.28, 5: 0.45},   # Movimientos de Masa
    "inu": {1: 0.02, 2: 0.07, 3: 0.15, 4: 0.28, 5: 0.45},   # Inundación
}

# --- Atenuación de Renta Fija (Azzone 2025) -----------------------------------
# Ante una caída del valor del activo, el equity (RV) absorbe el shock amplificado
# por el apalancamiento (canal Merton); la deuda (RF) está más protegida.
# v1 PLACEHOLDER: ratio_RF = K_RF * damage_ratio, K_RF documentado (no estimado),
# SIN apalancamiento Activo/Patrimonio (eso es exclusivo del equity).
K_RF = 0.15

# --- Sequía: factores de severidad de FLUJO sobre EBITDA (Bressan 2024) -------
# La sequía es shock CRÓNICO de FLUJO (afecta EBITDA), no daño de STOCK sobre el
# activo: NO usa DAMAGE_RATIO_BINS ni el valor del activo.
# *** PLACEHOLDER v1 — Pendiente calibración SENAMHI/PISCO ***
# (no se entregó tabla de bins de sequía; estos factores son un hook inicial).
# Además hoy NO hay EBITDA por emisor: sin EBITDA se reporta un PROXY etiquetado.
SEQ_FLOW_BINS = {1: 0.01, 2: 0.03, 3: 0.06, 4: 0.10, 5: 0.15}  # PLACEHOLDER

# --- Multiplicador de estrés ENSO (Pestaña 3) ---------------------------------
# *** PLACEHOLDER v1 — Pendiente calibración ***
# Mapea la Magnitud de El Niño Costero (ENFEN) a un multiplicador sobre el daño de
# MM + Inundación en la costa norte. El ICEN (indicadores_franco) se muestra como
# ancla cuantitativa. Editable; calibrar con histórico de daños 2017/2023.
ENSO_MULT = {"neutra": 1.0, "debil": 1.25, "moderada": 1.6,
             "fuerte": 2.2, "extraordinario": 3.0}

# Departamentos de la costa norte donde aplica el estrés ENSO (spec de proyecto).
COSTA_NORTE_DEPS = {"PIURA", "LAMBAYEQUE", "LA LIBERTAD", "TUMBES", "ICA"}


# ==============================================================================
# 2) UTILIDADES INTERNAS (auto-contenidas, sin importar del pipeline)
# ==============================================================================

def _norm(s):
    """Réplica intencional de norm() del pipeline (mayúsculas, sin acentos, solo
    [A-Z0-9 ], espacios colapsados). Se duplica a propósito para no acoplar este
    módulo al script principal; si cambias norm allá, espeja el cambio aquí."""
    if s is None:
        return ""
    s = str(s).strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _split_rv_rf(emisor, exp_clase):
    """Fracción RV/RF del emisor (split asumido uniforme en su geografía, igual que
    el pipeline). Si no está en exp_clase o TOTAL=0 -> 100% RV (conservador) + flag."""
    d = exp_clase.get(emisor)
    if not d or d.get("TOTAL", 0) <= 0:
        return 1.0, 0.0, False
    total = d["TOTAL"]
    return d.get("RENTA VARIABLE", 0.0) / total, d.get("RENTA FIJA", 0.0) / total, True


def _pe_punto(monto, dr, frac_rv, frac_rf, ap, k_rf):
    """Pérdida esperada de UN punto (emisor x distrito x hazard) de STOCK.
        ratio_RV = min(1, dr * Activo/Patrimonio)   (Azzone 2025, Merton)
        ratio_RF = k_rf * dr                         (Azzone 2025, deuda protegida)
    Si falta Activo/Patrimonio (fondos): RV usa dr sin apalancar (no se excluye).
    Devuelve (pe_rv, pe_rf, exp_rv, exp_rf, ratio_rv)."""
    if ap and ap.get("patrimonio", 0) > 0:
        ratio_rv = min(1.0, dr * (ap["activo"] / ap["patrimonio"]))
    else:
        ratio_rv = dr
    ratio_rf = k_rf * dr
    exp_rv, exp_rf = monto * frac_rv, monto * frac_rf
    return exp_rv * ratio_rv, exp_rf * ratio_rf, exp_rv, exp_rf, ratio_rv


# ==============================================================================
# 3) CARGA DE ACTIVO / PATRIMONIO POR EMISOR (estructura nueva, lectura aparte)
# ==============================================================================

def cargar_activo_patrimonio(risk_path, sheet_val, equiv_canonico):
    """Lee 'Valorización de Instrumentos' y devuelve {emisor: {'activo','patrimonio'}}.
    El pipeline no carga estas columnas; este módulo las lee por su cuenta (solo
    lectura). Tolera la errata real 'PASTRIMONIO_DEL_EMISOR'. Mapea Nombre (G&P) ->
    emisor canónico con equiv_canonico (= _equiv_exposicion_all() del pipeline)."""
    raw = pd.read_excel(risk_path, sheet_name=sheet_val, header=None, nrows=12)
    hdr_row = None
    for i in range(min(12, len(raw))):
        joined = " | ".join(_norm(v) for v in raw.iloc[i].tolist())
        if "NOMBRE" in joined and "ACTIVO" in joined:
            hdr_row = i
            break
    if hdr_row is None:
        raise ValueError("rf_extension: no se halló encabezado con 'Nombre' y 'Activo' "
                         "en Valorización (cols Nombre (G&P), ACTIVO/PATRIMONIO_DEL_EMISOR).")

    df = pd.read_excel(risk_path, sheet_name=sheet_val, header=hdr_row).dropna(how="all")
    cols_norm = {_norm(c): c for c in df.columns}

    def _find(*tokens_sets):
        for tokens in tokens_sets:
            for cn, real in cols_norm.items():
                if all(t in cn for t in tokens):
                    return real
        return None

    c_nombre = _find(["NOMBRE", "G P"], ["NOMBRE"])
    c_activo = _find(["ACTIVO", "EMISOR"], ["ACTIVO"])
    c_patrim = _find(["PATRIMONIO", "EMISOR"], ["PASTRIMONIO", "EMISOR"],
                     ["PATRIMONIO"], ["PASTRIMONIO"])
    faltan = [n for n, c in [("Nombre (G&P)", c_nombre), ("ACTIVO_DEL_EMISOR", c_activo),
                             ("PATRIMONIO_DEL_EMISOR", c_patrim)] if c is None]
    if faltan:
        raise ValueError(f"rf_extension: faltan columnas en Valorización: {faltan}")

    df = df.rename(columns={c_nombre: "nombre", c_activo: "activo", c_patrim: "patrim"})
    df["activo"] = pd.to_numeric(df["activo"], errors="coerce").fillna(0.0)
    df["patrim"] = pd.to_numeric(df["patrim"], errors="coerce").fillna(0.0)

    eq_norm = {_norm(k): v for k, v in equiv_canonico.items()}
    out = {}
    for _, r in df.iterrows():
        emisor = eq_norm.get(_norm(r["nombre"]))
        if emisor is None:
            continue
        d = out.setdefault(emisor, {"activo": 0.0, "patrimonio": 0.0})
        d["activo"] = max(d["activo"], float(r["activo"]))
        d["patrimonio"] = max(d["patrimonio"], float(r["patrim"]))
    return out


# ==============================================================================
# 4) DETECCIÓN DE CANAL INDIRECTO (bancos / soberanos / extranjeros)
# ==============================================================================

def detectar_canal_indirecto(districts, exp_clase, emp_sector, sector_financiero="Financiero"):
    """Marca como 'canal indirecto' al riesgo físico que NO se modela como daño
    directo sobre un activo físico geolocalizado en Perú:
      - BANCOS: canal de PD territorial de su cartera (Chitchumnong 2025).
      - SOBERANOS/EXTRANJEROS/FONDOS: sin cruce distrital (no aparecen en districts).
    Devuelve {emisor: motivo}."""
    presentes = set()
    for d in districts:
        for e, m in d.get("empresas", {}).items():
            if m and m > 0:
                presentes.add(e)
    indirecto = {}
    for emisor in exp_clase.keys():
        if emp_sector.get(emisor) == sector_financiero:
            indirecto[emisor] = "banco (PD territorial — Chitchumnong 2025)"
        elif emisor not in presentes:
            indirecto[emisor] = "sin cruce distrital (soberano/extranjero/fondo)"
    return indirecto


# ==============================================================================
# 5) PESTAÑA 1 — PÉRDIDA ESPERADA
# ==============================================================================

def calcular_perdida_esperada(districts, exp_clase, emp_sector,
                              activo_patrimonio, canal_indirecto,
                              bins=None, k_rf=K_RF, seq_bins=None, ebitda=None):
    """Pérdida Esperada S/ por (emisor, hazard). Todo lectura + agregación nueva.
    STOCK (MM/Inundación): PE = sum_distritos _pe_punto(...). Split RV/RF uniforme.
    SEQUÍA (Bressan 2024): canal de flujo separado; sin EBITDA -> PROXY etiquetado.
    Bancos/soberanos/extranjeros: excluidos del daño de activo (canal indirecto).
    `bins` parametrizado para que la Sensibilidad reutilice esta misma función."""
    if bins is None:
        bins = DAMAGE_RATIO_BINS
    if seq_bins is None:
        seq_bins = SEQ_FLOW_BINS

    HAZARDS_STOCK = [("mm", "score_mm"), ("inu", "score_inu")]
    acc, acc_seq, emisor_total, split_flag = {}, {}, {}, {}

    for d in districts:
        for emisor, monto in d.get("empresas", {}).items():
            if not monto or monto <= 0:
                continue
            emisor_total[emisor] = emisor_total.get(emisor, 0.0) + monto
            if emisor in canal_indirecto:
                continue
            frac_rv, frac_rf, split_ok = _split_rv_rf(emisor, exp_clase)
            split_flag[emisor] = split_ok
            ap = activo_patrimonio.get(emisor)

            for hz, score_key in HAZARDS_STOCK:
                sc = d.get(score_key)
                if sc is None:
                    continue
                try:
                    sc_i = int(round(float(sc)))
                except (TypeError, ValueError):
                    continue
                dr = bins.get(hz, {}).get(sc_i)
                if dr is None:
                    continue
                pe_rv, pe_rf, exp_rv, exp_rf, _ = _pe_punto(monto, dr, frac_rv, frac_rf, ap, k_rf)
                a = acc.setdefault((emisor, hz), {"pe_rv": 0.0, "pe_rf": 0.0,
                                                  "exp_rv_sc": 0.0, "exp_rf_sc": 0.0, "n_dist": 0})
                a["pe_rv"] += pe_rv; a["pe_rf"] += pe_rf
                a["exp_rv_sc"] += exp_rv; a["exp_rf_sc"] += exp_rf; a["n_dist"] += 1

            sc_seq = d.get("score_seq")
            if sc_seq is not None:
                try:
                    sseq = int(round(float(sc_seq)))
                except (TypeError, ValueError):
                    sseq = None
                f = seq_bins.get(sseq) if sseq is not None else None
                if f is not None:
                    s = acc_seq.setdefault(emisor, {"exp_seq": 0.0, "sumprod_factor": 0.0,
                                                    "sumprod_score": 0.0, "n_dist": 0})
                    s["exp_seq"] += monto; s["sumprod_factor"] += monto * f
                    s["sumprod_score"] += monto * sseq; s["n_dist"] += 1

    directa = []
    for (emisor, hz), a in acc.items():
        ap = activo_patrimonio.get(emisor, {})
        patrim, activo = ap.get("patrimonio", 0.0), ap.get("activo", 0.0)
        exp_rv_sc, exp_rf_sc = a["exp_rv_sc"], a["exp_rf_sc"]
        directa.append({
            "emisor": emisor, "sector": emp_sector.get(emisor, ""), "hazard": hz,
            "exp_total": round(emisor_total.get(emisor, 0.0), 2),
            "exp_rv_scored": round(exp_rv_sc, 2), "exp_rf_scored": round(exp_rf_sc, 2),
            "pe_rv": round(a["pe_rv"], 2), "pe_rf": round(a["pe_rf"], 2),
            "pe_total": round(a["pe_rv"] + a["pe_rf"], 2),
            "dr_efectivo_rv": round(a["pe_rv"] / exp_rv_sc, 4) if exp_rv_sc > 0 else None,
            "dr_efectivo_rf": round(a["pe_rf"] / exp_rf_sc, 4) if exp_rf_sc > 0 else None,
            "activo": activo, "patrimonio": patrim,
            "leverage": round(activo / patrim, 3) if patrim and patrim > 0 else None,
            "ap_disponible": bool(patrim and patrim > 0),
            "split_rv_rf_disponible": split_flag.get(emisor, True),
            "canal": "directo", "n_distritos": a["n_dist"],
        })
    directa.sort(key=lambda r: r["pe_total"], reverse=True)

    sequia = []
    for emisor, s in acc_seq.items():
        exp_seq = s["exp_seq"]
        factor_pond = (s["sumprod_factor"] / exp_seq) if exp_seq > 0 else 0.0
        score_pond = (s["sumprod_score"] / exp_seq) if exp_seq > 0 else None
        if ebitda and ebitda.get(emisor):
            pe_flujo, base, nota = round(ebitda[emisor] * factor_pond, 2), "EBITDA", \
                "shock de flujo sobre EBITDA (Bressan 2024)"
        else:
            pe_flujo, base, nota = round(exp_seq * factor_pond, 2), "PROXY", \
                "Pendiente EBITDA real — proxy: exposición x factor"
        sequia.append({
            "emisor": emisor, "sector": emp_sector.get(emisor, ""),
            "exp_en_sequia": round(exp_seq, 2),
            "score_seq_ponderado": round(score_pond, 2) if score_pond is not None else None,
            "factor_flujo_ponderado": round(factor_pond, 4),
            "base_calculo": base, "flujo_en_riesgo": pe_flujo, "nota": nota,
            "n_distritos": s["n_dist"],
        })
    sequia.sort(key=lambda r: r["flujo_en_riesgo"], reverse=True)

    indirecto = [{"emisor": e, "sector": emp_sector.get(e, ""),
                  "exp_total": round(exp_clase.get(e, {}).get("TOTAL", 0.0), 2),
                  "canal": "indirecto", "motivo": m}
                 for e, m in canal_indirecto.items()]
    indirecto.sort(key=lambda r: r["exp_total"], reverse=True)

    meta = {"bins": bins, "k_rf": k_rf, "seq_bins": seq_bins,
            "ebitda_disponible": bool(ebitda)}
    return {"directa": directa, "sequia": sequia, "indirecto": indirecto, "meta": meta}


# ==============================================================================
# 6) PESTAÑA 2 — SENSIBILIDAD (damage_ratios base ±30%)
# ==============================================================================

def calcular_sensibilidad(districts, exp_clase, emp_sector, activo_patrimonio,
                          canal_indirecto, base_bins=None, k_rf=K_RF, pct=0.30):
    """Recalcula la Pérdida Esperada (solo daño de stock) escalando los damage_ratios
    base por (1-pct), 1, (1+pct): conservador / base / estresado. Reutiliza
    calcular_perdida_esperada con bins escalados (el cap min(1,...) se respeta).
    Devuelve filas por (emisor, hazard) con los 3 escenarios."""
    if base_bins is None:
        base_bins = DAMAGE_RATIO_BINS
    escenarios = [("conservador", 1.0 - pct), ("base", 1.0), ("estresado", 1.0 + pct)]
    por_clave = {}   # (emisor, hazard) -> {conservador, base, estresado}
    for label, factor in escenarios:
        scaled = {hz: {s: min(1.0, v * factor) for s, v in tabla.items()}
                  for hz, tabla in base_bins.items()}
        res = calcular_perdida_esperada(districts, exp_clase, emp_sector,
                                        activo_patrimonio, canal_indirecto,
                                        bins=scaled, k_rf=k_rf)
        for r in res["directa"]:
            k = (r["emisor"], r["hazard"])
            por_clave.setdefault(k, {"emisor": r["emisor"], "sector": r["sector"],
                                     "hazard": r["hazard"]})[label] = r["pe_total"]
    filas = list(por_clave.values())
    for f in filas:
        for lab, _ in escenarios:
            f.setdefault(lab, 0.0)
        f["rango"] = round(f["estresado"] - f["conservador"], 2)
    filas.sort(key=lambda r: r["base"], reverse=True)
    return {"filas": filas, "pct": pct}


# ==============================================================================
# 7) PESTAÑA 3 — ENSO / EL NIÑO
# ==============================================================================

def cargar_enso(enfen_path=None, indic_path=None):
    """Lee ENFEN_Comunicados_Resumen (Estado de Alerta + Magnitud El Niño Costero)
    e indicadores_franco (ICEN). Devuelve el estado ENSO vigente y el multiplicador.
    Robusto: si falta un archivo, degrada a estado neutro (multiplicador 1.0)."""
    estado = {"estado": "No disponible", "magnitud": "neutra", "fecha": None,
              "icen": None, "icen_fecha": None, "multiplicador": 1.0,
              "nota": "Pendiente calibración (multiplicador placeholder)"}

    # --- ENFEN: última fila con Magnitud / Estado ---
    if enfen_path and os.path.exists(enfen_path):
        try:
            df = pd.read_excel(enfen_path, sheet_name="Comunicados ENFEN")
            cols = {_norm(c): c for c in df.columns}
            c_fec = cols.get("FECHA")
            c_est = next((cols[k] for k in cols if "ESTADO DE ALERTA" in k), None)
            c_mag = next((cols[k] for k in cols if "MAGNITUD EL NINO COSTERO" in k), None)
            if c_fec is not None:
                df = df.sort_values(c_fec)
            ult = df.dropna(subset=[c_est]).iloc[-1] if c_est else df.iloc[-1]
            estado["estado"] = str(ult[c_est]) if c_est else "No identificado"
            mag_raw = str(ult[c_mag]) if c_mag and not pd.isna(ult[c_mag]) else "neutra"
            estado["magnitud"] = mag_raw
            if c_fec is not None and not pd.isna(ult[c_fec]):
                estado["fecha"] = pd.to_datetime(ult[c_fec]).strftime("%Y-%m")
            # mapear magnitud -> multiplicador (placeholder)
            mn = _norm(mag_raw)
            if "EXTRAORDINARI" in mn:
                estado["multiplicador"] = ENSO_MULT["extraordinario"]
            elif "FUERTE" in mn:
                estado["multiplicador"] = ENSO_MULT["fuerte"]
            elif "MODERAD" in mn:
                estado["multiplicador"] = ENSO_MULT["moderada"]
            elif "DEBIL" in mn:
                estado["multiplicador"] = ENSO_MULT["debil"]
            else:
                estado["multiplicador"] = ENSO_MULT["neutra"]
        except Exception as e:
            estado["nota"] = f"ENFEN no leído ({e}); multiplicador neutro."

    # --- ICEN (indicadores_franco) como ancla cuantitativa ---
    if indic_path and os.path.exists(indic_path):
        try:
            icen = pd.read_excel(indic_path, sheet_name="ICEN")
            cfec = next((c for c in icen.columns if _norm(c) == "FECHA"), icen.columns[0])
            cval = next((c for c in icen.columns if _norm(c) == "ICEN"), icen.columns[-1])
            icen = icen.dropna(subset=[cval]).sort_values(cfec)
            last = icen.iloc[-1]
            estado["icen"] = round(float(last[cval]), 2)
            estado["icen_fecha"] = pd.to_datetime(last[cfec]).strftime("%Y-%m")
        except Exception:
            pass
    return estado


def calcular_enso(districts, exp_clase, emp_sector, activo_patrimonio,
                  canal_indirecto, enso_state, bins=None, k_rf=K_RF,
                  deps_costa_norte=None):
    """Aplica el multiplicador ENSO al daño de stock (MM + Inundación) SOLO en los
    distritos de la costa norte. Devuelve, por emisor, la PE base y la estresada
    (con multiplicador) en esa geografía, para ver el efecto con/sin El Niño."""
    if bins is None:
        bins = DAMAGE_RATIO_BINS
    if deps_costa_norte is None:
        deps_costa_norte = COSTA_NORTE_DEPS
    deps_norm = {_norm(x) for x in deps_costa_norte}
    mult = float(enso_state.get("multiplicador", 1.0))
    HAZARDS_STOCK = [("mm", "score_mm"), ("inu", "score_inu")]

    base_emp = {}   # emisor -> {pe_base, exp_cn, n_dist}
    for d in districts:
        if _norm(d.get("dep")) not in deps_norm:
            continue
        for emisor, monto in d.get("empresas", {}).items():
            if not monto or monto <= 0 or emisor in canal_indirecto:
                continue
            frac_rv, frac_rf, _ = _split_rv_rf(emisor, exp_clase)
            ap = activo_patrimonio.get(emisor)
            pe_pt = 0.0
            for hz, score_key in HAZARDS_STOCK:
                sc = d.get(score_key)
                if sc is None:
                    continue
                try:
                    dr = bins.get(hz, {}).get(int(round(float(sc))))
                except (TypeError, ValueError):
                    dr = None
                if dr is None:
                    continue
                pe_rv, pe_rf, _, _, _ = _pe_punto(monto, dr, frac_rv, frac_rf, ap, k_rf)
                pe_pt += pe_rv + pe_rf
            if pe_pt > 0 or monto > 0:
                e = base_emp.setdefault(emisor, {"pe_base": 0.0, "exp_cn": 0.0, "n_dist": 0})
                e["pe_base"] += pe_pt
                e["exp_cn"] += monto
                e["n_dist"] += 1

    filas = []
    for emisor, e in base_emp.items():
        pe_base = e["pe_base"]
        filas.append({
            "emisor": emisor, "sector": emp_sector.get(emisor, ""),
            "exp_costa_norte": round(e["exp_cn"], 2),
            "pe_base": round(pe_base, 2),
            "pe_estresado": round(pe_base * mult, 2),
            "delta": round(pe_base * (mult - 1.0), 2),
            "n_distritos": e["n_dist"],
        })
    filas.sort(key=lambda r: r["pe_estresado"], reverse=True)
    return {"filas": filas, "enso": enso_state, "multiplicador": mult,
            "deps": sorted(deps_costa_norte)}


# ==============================================================================
# 8) BUILDER DEL HTML COMPAÑERO (autocontenido, 3 pestañas)
# ==============================================================================

_HTML_TEMPLATE = r"""<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Extensión Riesgo Físico — Pérdida Esperada · AFP Integra</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
:root{--azul:#0b2e59;--cyan:#0aa3c2;--gris:#f4f6fa;--linea:#e2e7f0;--rojo:#c0392b;}
*{box-sizing:border-box;} body{margin:0;font-family:Segoe UI,Roboto,Arial,sans-serif;
 color:#1a2433;background:#fff;}
header{background:var(--azul);color:#fff;padding:16px 22px;}
header h1{margin:0;font-size:18px;} header p{margin:4px 0 0;opacity:.85;font-size:12.5px;}
.tabbar{display:flex;gap:4px;background:var(--gris);padding:6px 10px;border-bottom:1px solid var(--linea);}
.tab{padding:9px 16px;border:none;background:transparent;color:var(--azul);font-weight:700;
 font-size:13px;cursor:pointer;border-radius:6px 6px 0 0;}
.tab.active{background:#fff;border-bottom:3px solid var(--cyan);}
.wrap{padding:18px 22px;max-width:1280px;margin:0 auto;}
.panel{display:none;} .panel.active{display:block;}
.card{border:1px solid var(--linea);border-radius:10px;padding:14px 16px;margin-bottom:18px;}
.card h3{margin:0 0 6px;font-size:15px;color:var(--azul);}
.note{font-size:12px;color:#5a6678;margin:0 0 10px;}
.flag{display:inline-block;background:#fff4e5;color:#8a5a00;border:1px solid #f0c887;
 border-radius:6px;padding:2px 8px;font-size:11px;font-weight:700;}
table{border-collapse:collapse;width:100%;font-size:12.5px;}
thead th{position:sticky;top:0;background:var(--azul);color:#fff;padding:8px 10px;text-align:left;
 cursor:pointer;white-space:nowrap;} thead th:hover{background:#13407a;}
td{padding:6px 10px;border-top:1px solid var(--linea);} tbody tr:hover td{background:#f7f9ff;}
.num{text-align:right;white-space:nowrap;} .tag{font-size:11px;font-weight:700;padding:1px 7px;border-radius:10px;}
.tag-rv{background:#e7f4ff;color:#0b66a3;} .tag-rf{background:#eef0f4;color:#52607a;}
.tag-ind{background:#fdeaea;color:#a33;} .banner{padding:12px 14px;border-radius:8px;
 background:#eaf6fb;border:1px solid #bfe4f0;font-size:13px;margin-bottom:14px;}
.banner b{color:var(--azul);} .toggle{margin:8px 0;font-size:13px;}
.kpi{display:inline-block;background:var(--gris);border:1px solid var(--linea);border-radius:8px;
 padding:8px 12px;margin:0 8px 8px 0;font-size:12px;} .kpi b{display:block;font-size:16px;color:var(--azul);}
.scroll{max-height:520px;overflow:auto;border:1px solid var(--linea);border-radius:8px;}
</style></head><body>
<header>
 <h1>Extensión de Riesgo Físico — Pérdida Esperada, Sensibilidad y ENSO</h1>
 <p>AFP Integra · Complemento del dashboard CENEPRED · <span id="stamp"></span></p>
</header>
<div class="tabbar">
 <button class="tab active" data-p="p1">Pérdida Esperada</button>
 <button class="tab" data-p="p2">Sensibilidad ±30%</button>
 <button class="tab" data-p="p3">ENSO / El Niño</button>
</div>
<div class="wrap">
 <section id="p1" class="panel active"></section>
 <section id="p2" class="panel"></section>
 <section id="p3" class="panel"></section>
</div>
<script>
const DATA = __PAYLOAD__;
const HZ = {mm:"Mov. de Masa", inu:"Inundación"};
const fmt = n => (n==null)?"—":"S/ "+Number(n).toLocaleString("es-PE",{maximumFractionDigits:0});
const pct = n => (n==null)?"—":(100*n).toFixed(1)+"%";
function tableHTML(cols, rows, rowf){
 let h="<table><thead><tr>"+cols.map((c,i)=>`<th data-i="${i}">${c.t}</th>`).join("")+"</tr></thead><tbody>";
 h+=rows.map(rowf).join(""); return h+"</tbody></table>";
}
function makeSortable(sec){
 sec.querySelectorAll("table").forEach(tb=>{
  tb.querySelectorAll("th").forEach((th,i)=>{th.onclick=()=>{
   const body=tb.querySelector("tbody"); const rs=[...body.rows];
   const asc=th.dataset.asc!=="1"; th.dataset.asc=asc?"1":"0";
   rs.sort((a,b)=>{let x=a.cells[i].dataset.v??a.cells[i].innerText,
     y=b.cells[i].dataset.v??b.cells[i].innerText; let nx=parseFloat(x),ny=parseFloat(y);
     if(!isNaN(nx)&&!isNaN(ny)){x=nx;y=ny;} return (x>y?1:x<y?-1:0)*(asc?1:-1);});
   rs.forEach(r=>body.appendChild(r));};});
 });
}

/* ---------- Pestaña 1 ---------- */
function renderP1(){
 const d=DATA.perdida; let h="";
 const totPE=d.directa.reduce((s,r)=>s+r.pe_total,0);
 h+=`<div class="card"><div class="kpi"><b>${fmt(totPE)}</b>Pérdida esperada total (MM+Inundación)</div>
   <div class="kpi"><b>${d.directa.length}</b>filas emisor×peligro</div>
   <div class="kpi"><b>${d.indirecto.length}</b>emisores canal indirecto</div></div>`;
 // daño de stock
 const cols=[{t:"Emisor"},{t:"Sector"},{t:"Peligro"},{t:"PE Total"},{t:"PE Renta Variable"},
   {t:"PE Renta Fija"},{t:"dr efectivo RV"},{t:"Leverage A/P"},{t:"Exp. con score"}];
 h+=`<div class="card"><h3>Daño de stock sobre el activo — ranking por Pérdida Esperada</h3>
   <p class="note">Roncalli 2026 (bins Huizinga) · RV = min(1, dr×Activo/Patrimonio), Merton/Azzone 2025 · RF = k_rf×dr (k_rf=${d.meta.k_rf}). Clic en encabezado para ordenar.</p>
   <div class="scroll">`+tableHTML(cols,d.directa,r=>`<tr>
     <td>${r.emisor}</td><td>${r.sector}</td><td>${HZ[r.hazard]||r.hazard}</td>
     <td class="num" data-v="${r.pe_total}">${fmt(r.pe_total)}</td>
     <td class="num" data-v="${r.pe_rv}"><span class="tag tag-rv">RV</span> ${fmt(r.pe_rv)}</td>
     <td class="num" data-v="${r.pe_rf}"><span class="tag tag-rf">RF</span> ${fmt(r.pe_rf)}</td>
     <td class="num" data-v="${r.dr_efectivo_rv??0}">${pct(r.dr_efectivo_rv)}</td>
     <td class="num" data-v="${r.leverage??0}">${r.leverage!=null?r.leverage.toFixed(2):"—"+(r.ap_disponible?"":" <span class='flag'>sin A/P</span>")}</td>
     <td class="num" data-v="${r.exp_rv_scored+r.exp_rf_scored}">${fmt(r.exp_rv_scored+r.exp_rf_scored)}</td></tr>`)+`</div></div>`;
 // sequía
 const cseq=[{t:"Emisor"},{t:"Sector"},{t:"Exp. en sequía"},{t:"Score sequía pond."},
   {t:"Factor flujo"},{t:"Base"},{t:"Flujo en riesgo"}];
 h+=`<div class="card"><h3>Canal de flujo — Sequía sobre EBITDA <span class="flag">Pendiente EBITDA real</span> <span class="flag">Pendiente calibración SENAMHI/PISCO</span></h3>
   <p class="note">Bressan 2024: shock crónico de flujo (no daño de stock). Sin EBITDA, "Flujo en riesgo" es un PROXY = exposición × factor.</p>
   <div class="scroll">`+tableHTML(cseq,d.sequia,r=>`<tr>
     <td>${r.emisor}</td><td>${r.sector}</td>
     <td class="num" data-v="${r.exp_en_sequia}">${fmt(r.exp_en_sequia)}</td>
     <td class="num" data-v="${r.score_seq_ponderado??0}">${r.score_seq_ponderado??"—"}</td>
     <td class="num" data-v="${r.factor_flujo_ponderado}">${pct(r.factor_flujo_ponderado)}</td>
     <td>${r.base_calculo}</td>
     <td class="num" data-v="${r.flujo_en_riesgo}">${fmt(r.flujo_en_riesgo)}</td></tr>`)+`</div></div>`;
 // indirecto
 const cind=[{t:"Emisor"},{t:"Sector"},{t:"Exp. total"},{t:"Motivo"}];
 h+=`<div class="card"><h3>Canal indirecto — sin daño directo de activo</h3>
   <p class="note">Bancos: PD territorial (Chitchumnong 2025). Soberanos/extranjeros/fondos: sin cruce distrital.</p>
   <div class="scroll">`+tableHTML(cind,d.indirecto,r=>`<tr>
     <td>${r.emisor}</td><td>${r.sector}</td>
     <td class="num" data-v="${r.exp_total}">${fmt(r.exp_total)}</td>
     <td><span class="tag tag-ind">${r.motivo}</span></td></tr>`)+`</div></div>`;
 const sec=document.getElementById("p1"); sec.innerHTML=h; makeSortable(sec);
}

/* ---------- Pestaña 2 ---------- */
function renderP2(){
 const d=DATA.sensibilidad; let h="";
 h+=`<div class="card"><h3>Sensibilidad de la Pérdida Esperada a ±${(100*d.pct).toFixed(0)}% en los damage_ratios</h3>
   <p class="note">Escenarios conservador (−${(100*d.pct).toFixed(0)}%), base y estresado (+${(100*d.pct).toFixed(0)}%). El cap min(1,·) puede comprimir el escenario estresado.</p>
   <div id="barSens" style="height:420px;"></div></div>`;
 const cols=[{t:"Emisor"},{t:"Sector"},{t:"Peligro"},{t:"Conservador"},{t:"Base"},{t:"Estresado"},{t:"Rango"}];
 h+=`<div class="card"><div class="scroll">`+tableHTML(cols,d.filas,r=>`<tr>
     <td>${r.emisor}</td><td>${r.sector}</td><td>${HZ[r.hazard]||r.hazard}</td>
     <td class="num" data-v="${r.conservador}">${fmt(r.conservador)}</td>
     <td class="num" data-v="${r.base}">${fmt(r.base)}</td>
     <td class="num" data-v="${r.estresado}">${fmt(r.estresado)}</td>
     <td class="num" data-v="${r.rango}">${fmt(r.rango)}</td></tr>`)+`</div></div>`;
 const sec=document.getElementById("p2"); sec.innerHTML=h; makeSortable(sec);
 const top=d.filas.slice(0,12); const lbl=top.map(r=>r.emisor+" · "+(HZ[r.hazard]||r.hazard));
 Plotly.newPlot("barSens",[
   {x:lbl,y:top.map(r=>r.conservador),name:"Conservador",type:"bar",marker:{color:"#7fb3d5"}},
   {x:lbl,y:top.map(r=>r.base),name:"Base",type:"bar",marker:{color:"#0b2e59"}},
   {x:lbl,y:top.map(r=>r.estresado),name:"Estresado",type:"bar",marker:{color:"#c0392b"}}],
   {barmode:"group",margin:{l:60,r:10,t:10,b:120},legend:{orientation:"h"},
    yaxis:{title:"Pérdida Esperada (S/)"}},{displayModeBar:false,responsive:true});
}

/* ---------- Pestaña 3 ---------- */
function renderP3(){
 const d=DATA.enso, e=d.enso; let h="";
 h+=`<div class="banner">Estado ENFEN vigente: <b>${e.estado}</b> · Magnitud El Niño Costero: <b>${e.magnitud}</b>`
   +(e.fecha?` (${e.fecha})`:"")+(e.icen!=null?` · ICEN ${e.icen} (${e.icen_fecha})`:"")
   +` · Multiplicador aplicado: <b>×${d.multiplicador}</b> <span class="flag">Placeholder — pendiente calibración</span></div>`;
 h+=`<div class="card"><h3>Estrés ENSO sobre MM + Inundación en la costa norte</h3>
   <p class="note">Departamentos: ${d.deps.join(", ")}. El multiplicador escala la Pérdida Esperada de stock solo en esos distritos. Bancos excluidos (canal indirecto).</p>
   <label class="toggle"><input type="checkbox" id="chkEnso" checked> Aplicar multiplicador ENSO (×${d.multiplicador})</label>
   <div id="barEnso" style="height:420px;"></div></div>`;
 const cols=[{t:"Emisor"},{t:"Sector"},{t:"Exp. costa norte"},{t:"PE base"},{t:"PE con ENSO"},{t:"Δ ENSO"}];
 h+=`<div class="card"><div class="scroll">`+tableHTML(cols,d.filas,r=>`<tr>
     <td>${r.emisor}</td><td>${r.sector}</td>
     <td class="num" data-v="${r.exp_costa_norte}">${fmt(r.exp_costa_norte)}</td>
     <td class="num" data-v="${r.pe_base}">${fmt(r.pe_base)}</td>
     <td class="num" data-v="${r.pe_estresado}">${fmt(r.pe_estresado)}</td>
     <td class="num" data-v="${r.delta}">${fmt(r.delta)}</td></tr>`)+`</div></div>`;
 const sec=document.getElementById("p3"); sec.innerHTML=h; makeSortable(sec);
 const draw=(on)=>{const top=d.filas.slice(0,12);
   Plotly.newPlot("barEnso",[
     {x:top.map(r=>r.emisor),y:top.map(r=>r.pe_base),name:"PE base",type:"bar",marker:{color:"#0b2e59"}},
     {x:top.map(r=>r.emisor),y:top.map(r=>on?r.pe_estresado:r.pe_base),name:on?"PE con ENSO":"PE base",
      type:"bar",marker:{color:on?"#c0392b":"#9aa7bd"}}],
     {barmode:"group",margin:{l:60,r:10,t:10,b:120},legend:{orientation:"h"},
      yaxis:{title:"Pérdida Esperada (S/)"}},{displayModeBar:false,responsive:true});};
 draw(true);
 document.getElementById("chkEnso").onchange=ev=>draw(ev.target.checked);
}

document.getElementById("stamp").textContent=DATA.stamp||"";
document.querySelectorAll(".tab").forEach(b=>b.onclick=()=>{
 document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active"));
 document.querySelectorAll(".panel").forEach(x=>x.classList.remove("active"));
 b.classList.add("active"); document.getElementById(b.dataset.p).classList.add("active");
});
renderP1(); renderP2(); renderP3();
</script></body></html>"""


def construir_html_companion(out_path, perdida, sensibilidad, enso, enso_state, stamp=None):
    """Escribe el HTML compañero autocontenido con las 3 pestañas."""
    payload = {"perdida": perdida, "sensibilidad": sensibilidad, "enso": enso,
               "stamp": stamp or datetime.now().strftime("%Y-%m-%d %H:%M")}
    html = _HTML_TEMPLATE.replace("__PAYLOAD__", json.dumps(payload, ensure_ascii=False))
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path


# ==============================================================================
# 9) ORQUESTADOR  — único punto de entrada desde main()
# ==============================================================================

def generar_extension(out_dir, districts, exp_clase, emp_sector,
                      risk_path, sheet_val, equiv_canonico,
                      enfen_path=None, indic_path=None, ebitda=None,
                      stamp=None, logfn=print):
    """Pega todo: lee Activo/Patrimonio, detecta canal indirecto, calcula las 3
    pestañas y escribe el HTML compañero. Devuelve la ruta del HTML generado.
    Todo es aditivo y de solo lectura sobre las estructuras del pipeline."""
    if stamp is None:
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
    ap = cargar_activo_patrimonio(risk_path, sheet_val, equiv_canonico)
    canal = detectar_canal_indirecto(districts, exp_clase, emp_sector)
    perdida = calcular_perdida_esperada(districts, exp_clase, emp_sector, ap, canal, ebitda=ebitda)
    sens = calcular_sensibilidad(districts, exp_clase, emp_sector, ap, canal)
    enso_state = cargar_enso(enfen_path, indic_path)
    enso = calcular_enso(districts, exp_clase, emp_sector, ap, canal, enso_state)
    out_path = os.path.join(out_dir, f"Extension_RF_PerdidaEsperada_{stamp}.html")
    construir_html_companion(out_path, perdida, sens, enso, enso_state,
                             stamp=datetime.now().strftime("%Y-%m-%d %H:%M"))
    logfn(f"Extensión RF generada: {out_path}  "
          f"(emisores daño={len({r['emisor'] for r in perdida['directa']})}, "
          f"indirecto={len(perdida['indirecto'])}, "
          f"ENSO ×{enso_state.get('multiplicador')})")
    return out_path
