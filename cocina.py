"""
cocina.py — Dashboard Electoral Peru 2026
==========================================
USO CORRECTO — al final de onpe_regional.py añadir:
    exec(open(r'C:\...\cocina.py').read())

O desde Spyder, ejecutar DESPUÉS de onpe_regional.py para que
df_consolidado ya esté en el namespace.

Si se ejecuta standalone (sin df_consolidado) genera el dashboard
con S1 vacío pero S2-S5 completos.
"""

import json, warnings, os
from pathlib import Path
from datetime import datetime, date
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# ── Rutas ──────────────────────────────────────────────────────────
BASE      = Path(r"C:\Users\usuario\OneDrive\Desktop\AFP INTEGRA\Modelo Franco\Elecciones")
STOCKS_PATH = BASE / "Stocks_prices.xlsx"
RF_PATH     = BASE / "Renta Fija Local" / "reporte.xlsx"
OUT_PATH    = BASE / "elecciones_2026.html"

# ── Fechas electorales ─────────────────────────────────────────────
ELECCIONES = {
    "1V_2016": date(2016, 4, 10),
    "2V_2016": date(2016, 6, 5),
    "1V_2021": date(2021, 4, 11),
    "2V_2021": date(2021, 6, 6),
    "1V_2026": date(2026, 4, 13),
    "2V_2026": date(2026, 6, 7),
}
VENTANA_DIAS = 30

# ── Obtener df_consolidado — compatible con Spyder/IPython/exec ───
# Spyder corre cada script en namespace propio (vacío con F5),
# pero el kernel IPython expone variables via get_ipython().shell.user_ns
_df_consolidado = None

# 1. Namespace local (exec / runfile)
_df_consolidado = globals().get("df_consolidado", None)

# 2. Kernel IPython/Spyder
if _df_consolidado is None:
    try:
        import IPython
        _ipy = IPython.get_ipython()
        if _ipy is not None:
            _df_consolidado = _ipy.shell.user_ns.get("df_consolidado", None)
    except Exception:
        pass

# 3. __main__
if _df_consolidado is None:
    try:
        import __main__
        _df_consolidado = getattr(__main__, "df_consolidado", None)
    except Exception:
        pass

# 4. ← NUEVO: cargar desde pickle si ninguno de los anteriores funcionó
# 4. Cargar desde Excel si ninguno de los anteriores funcionó
XLSX_PATH = BASE / "df_consolidado.xlsx"
if (_df_consolidado is None or not isinstance(_df_consolidado, pd.DataFrame) or _df_consolidado.empty):
    if XLSX_PATH.exists():
        _df_consolidado = pd.read_excel(XLSX_PATH)
        print(f"  df_consolidado cargado desde Excel: {len(_df_consolidado)} filas")

# Validar
if not isinstance(_df_consolidado, pd.DataFrame) or _df_consolidado.empty:
    _df_consolidado = None

# ══════════════════════════════════════════════════════════════════
# S1 — ONPE REGIONAL
# ══════════════════════════════════════════════════════════════════
def build_s1(df):
    if df is None or df.empty:
        return {"regiones": [], "extranjero": {}, "total_votos_jp": 0,
                "total_votos_fp": 0, "pct_nac_jp": 0, "pct_nac_fp": 0,
                "timestamp": "Sin datos", "n_regiones_jp": 0, "n_regiones_fp": 0}

    JP = "JUNTOS POR EL PERÚ"
    FP = "FUERZA POPULAR"
    df_reg = df[df["nivel"] == "regional"].copy()
    df_ext = df[df["nivel"] == "extranjero"].copy()

    regiones = []
    for ubigeo, grp in df_reg.groupby("ubigeo"):
        rj = grp[grp["partido"] == JP]
        rf = grp[grp["partido"] == FP]
        pj = float(rj["porcentaje"].values[0]) if not rj.empty else 0
        pf = float(rf["porcentaje"].values[0]) if not rf.empty else 0
        vj = int(rj["votos"].values[0])        if not rj.empty else 0
        vf = int(rf["votos"].values[0])        if not rf.empty else 0
        ac = rj["actas_contabilizadas"].values[0] if not rj.empty else None
        at = rj["actas_total"].values[0]          if not rj.empty else None
        pc = rj["pct_contabilizadas"].values[0]   if not rj.empty else None
        ts = str(rj["timestamp"].values[0])        if not rj.empty else ""
        regiones.append({
            "nombre": ubigeo, "pct_jp": round(pj,3), "pct_fp": round(pf,3),
            "votos_jp": vj, "votos_fp": vf,
            "ganador": "JP" if pj > pf else "FP",
            "margen": round(abs(pj-pf),3),
            "actas_contabilizadas": int(ac) if ac is not None and str(ac)!='nan' else 0,
            "actas_total": int(at) if at is not None and str(at)!='nan' else 0,
            "pct_contabilizadas": round(float(pc),1) if pc is not None and str(pc)!='nan' else 0,
            "timestamp": ts[:19],
        })

    ext = {}
    if not df_ext.empty:
        rj = df_ext[df_ext["partido"]==JP]; rf = df_ext[df_ext["partido"]==FP]
        ext = {"pct_jp": float(rj["porcentaje"].values[0]) if not rj.empty else 0,
               "pct_fp": float(rf["porcentaje"].values[0]) if not rf.empty else 0,
               "votos_jp": int(rj["votos"].values[0]) if not rj.empty else 0,
               "votos_fp": int(rf["votos"].values[0]) if not rf.empty else 0}

    tot_jp = sum(r["votos_jp"] for r in regiones)
    tot_fp = sum(r["votos_fp"] for r in regiones)
    tot    = tot_jp + tot_fp
    return {
        "regiones": sorted(regiones, key=lambda x: x["nombre"]),
        "extranjero": ext,
        "total_votos_jp": tot_jp, "total_votos_fp": tot_fp,
        "pct_nac_jp": round(100*tot_jp/tot,3) if tot>0 else 0,
        "pct_nac_fp": round(100*tot_fp/tot,3) if tot>0 else 0,
        "timestamp": str(df_reg["timestamp"].max())[:19] if not df_reg.empty else "",
        "n_regiones_jp": sum(1 for r in regiones if r["ganador"]=="JP"),
        "n_regiones_fp": sum(1 for r in regiones if r["ganador"]=="FP"),
    }

# ══════════════════════════════════════════════════════════════════
# S2 — ENCUESTADORAS
# ══════════════════════════════════════════════════════════════════
def build_s2():
    return {
        "datum": {
            "fuente":"Datum Internacional","fecha":"07/06/2026",
            "muestra":"117,199 votos","margen_error":1.0,
            "jp":50.14,"fp":49.86,"jp_min":49.14,"jp_max":51.14,
            "fp_min":48.86,"fp_max":50.86,"veredicto":"EMPATE ESTADÍSTICO",
            "por_region":{
                "Amazonas":65.26,"Áncash":56.72,"Apurímac":80.87,
                "Arequipa":62.74,"Ayacucho":79.35,"Cajamarca":67.98,
                "Callao":35.87,"Cusco":78.48,"Huancavelica":80.60,
                "Huánuco":64.75,"Ica":46.18,"Junín":53.39,
                "La Libertad":42.69,"Lambayeque":41.08,
                "Lima Met.":35.77,"Lima Prov.":47.64,
                "Loreto":44.92,"Madre de Dios":72.96,"Moquegua":66.81,
                "Pasco":62.70,"Piura":42.76,"Puno":87.84,
                "San Martín":54.04,"Tacna":74.04,"Tumbes":34.86,
                "Ucayali":46.89,"Extranjero":37.33,
            },
            "evolucion":[
                {"label":"Int.voto\n17-20 May","jp":47.8,"fp":52.2},
                {"label":"Int.voto\n26-30 May","jp":47.4,"fp":52.6},
                {"label":"Simulacro\n26-30 May","jp":47.1,"fp":52.9},
                {"label":"Int.voto\n03-04 Jun","jp":48.8,"fp":51.2},
                {"label":"Simulacro\n06 Jun","jp":49.4,"fp":50.6},
                {"label":"Boca Urna\n07 Jun","jp":49.47,"fp":50.53},
                {"label":"C.Rápido\n07 Jun","jp":50.14,"fp":49.86},
            ],
        },
        "ipsos":{
            "fuente":"Ipsos / Transparencia","fecha":"07/06/2026",
            "muestra":"1,037 actas","margen_error":1.9,
            "jp":50.3,"fp":49.7,"veredicto":"EMPATE TÉCNICO",
            "desagregado":{
                "Lima":{"jp":36.4,"fp":63.6},
                "Regiones":{"jp":57.4,"fp":42.6},
                "Urbano":{"jp":46.1,"fp":53.9},
                "Rural":{"jp":69.0,"fp":31.0},
                "Costa":{"jp":39.5,"fp":60.5},
                "Sierra":{"jp":70.2,"fp":29.8},
                "Selva":{"jp":58.6,"fp":41.4},
            },
        },
    }

# ══════════════════════════════════════════════════════════════════
# S3 — STOCKS
# ══════════════════════════════════════════════════════════════════
def build_s3():
    df = pd.read_excel(STOCKS_PATH, sheet_name="Hoja1", parse_dates=["Fecha"])
    df["Fecha"] = pd.to_datetime(df["Fecha"]).dt.date
    df = df.sort_values("Fecha").reset_index(drop=True)

    sect_df  = pd.read_excel(STOCKS_PATH, sheet_name="Hoja2")
    sec_map  = dict(zip(sect_df["Empresa"], sect_df["Sector"]))
    empresas = [c for c in df.columns if c != "Fecha"]
    sectores = sorted(set(sec_map.values()))

    # Series completas
    series = {}
    for emp in empresas:
        s = df[["Fecha",emp]].dropna()
        series[emp] = {
            "fechas":  [str(d) for d in s["Fecha"].tolist()],
            "precios": [round(float(v),4) for v in s[emp].tolist()],
            "sector":  sec_map.get(emp,"Otro"),
        }

    # Ventanas electorales
    ventanas = {}
    for name, elec_date in ELECCIONES.items():
        mask = df["Fecha"].apply(lambda d: abs((d-elec_date).days) <= VENTANA_DIAS)
        sub  = df[mask].copy()
        if sub.empty:
            continue
        emp_data = {}
        for emp in empresas:
            col = sub[["Fecha",emp]].dropna()
            if col.empty: continue
            base = float(col[emp].iloc[0])
            if base == 0: continue
            emp_data[emp] = {
                "dias": [(d-elec_date).days for d in col["Fecha"].tolist()],
                "ret":  [round((v/base-1)*100,3) for v in col[emp].tolist()],
                "sector": sec_map.get(emp,"Otro"),
            }
        ventanas[name] = {"fecha": str(elec_date), "empresas": emp_data}

    # Rendimiento por sector en días clave
    sect_perf = {}
    for name, vdata in ventanas.items():
        ps = {}
        for emp, ed in vdata["empresas"].items():
            sec = ed["sector"]
            for tgt in [-5, 0, 5, 10, 20]:
                if not ed["dias"]: continue
                ci = min(range(len(ed["dias"])), key=lambda i: abs(ed["dias"][i]-tgt))
                if abs(ed["dias"][ci]-tgt) <= 3:
                    ps.setdefault(sec,{}).setdefault(tgt,[]).append(ed["ret"][ci])
        sect_perf[name] = {s:{t:round(np.mean(v),2) for t,v in td.items()}
                           for s,td in ps.items()}

    return {"series":series,"ventanas":ventanas,"sect_performance":sect_perf,
            "empresas":empresas,"sectores":sectores,"sector_map":sec_map,
            "elecciones":{k:str(v) for k,v in ELECCIONES.items()}}

# ══════════════════════════════════════════════════════════════════
# S4 — MODELO BAYESIANO
# ══════════════════════════════════════════════════════════════════
def build_s4(s1, s2):
    from math import erf, sqrt
    p0 = (s2["datum"]["jp"] + s2["ipsos"]["jp"]) / 200.0
    K  = 1000
    a0, b0 = p0*K, (1-p0)*K
    vj = s1.get("total_votos_jp", 0)
    vf = s1.get("total_votos_fp", 0)
    ap, bp = a0+vj, b0+vf
    mu  = ap/(ap+bp)
    var = (ap*bp)/((ap+bp)**2*(ap+bp+1))
    sd  = var**0.5
    z   = (mu-0.5)/sd if sd>0 else 0
    pw  = 0.5*(1+erf(z/sqrt(2)))

    regs = s1.get("regiones",[])
    ac_t = sum(r.get("actas_total",0) for r in regs)
    ac_c = sum(r.get("actas_contabilizadas",0) for r in regs)
    pesc = ac_c/ac_t*100 if ac_t>0 else 0
    tot  = vj+vf
    vpa  = tot/ac_c if ac_c>0 else 300
    pend = ac_t-ac_c
    vpp  = int(vpa*pend)

    evol = []
    aj, bj = a0, b0
    for r in sorted(regs, key=lambda x: x.get("timestamp","")):
        aj += r.get("votos_jp",0); bj += r.get("votos_fp",0)
        t2 = aj+bj
        if t2>0:
            m2 = aj/t2
            s2_ = ((aj*bj)/(t2**2*(t2+1)))**0.5
            z2  = (m2-0.5)/s2_ if s2_>0 else 0
            p2  = 0.5*(1+erf(z2/sqrt(2)))
            evol.append({"region":r["nombre"],"prob_jp":round(p2*100,1),
                         "prob_fp":round((1-p2)*100,1),
                         "pct_esc":round(r.get("pct_contabilizadas",0),1)})

    return {
        "prior_jp":round(p0*100,2),"alpha_prior":round(a0,1),"beta_prior":round(b0,1),
        "votos_obs_jp":vj,"votos_obs_fp":vf,"total_obs":tot,
        "posterior_jp":round(mu*100,3),"posterior_fp":round((1-mu)*100,3),
        "ci_lo":round(max(0,mu-1.96*sd)*100,2),"ci_hi":round(min(1,mu+1.96*sd)*100,2),
        "std":round(sd*100,3),
        "prob_jp_wins":round(pw*100,1),"prob_fp_wins":round((1-pw)*100,1),
        "pct_escrutado":round(pesc,1),"actas_cont":ac_c,"actas_total":ac_t,
        "proy_jp":vj+int(vpp*mu),"proy_fp":vf+int(vpp*(1-mu)),
        "evolucion":evol,"kappa":K,
    }

# ══════════════════════════════════════════════════════════════════
# S5 — RENTA FIJA
# ══════════════════════════════════════════════════════════════════
def build_s5():
    df = pd.read_excel(RF_PATH, sheet_name="Sheet1", header=1)
    df.columns = ["Fecha","Nemo","ISIN","Emisor","Moneda",
                  "P_L_pct","P_S_pct","TIR","Origen","Spread",
                  "P_L_m","P_S_m","IC_m","F_Venc","F_Emis","Cupon",
                  "Marg_L","TIR_SO","Rating","Ult_C","Prox_C","Dur"]
    df = df.dropna(subset=["Fecha","Nemo"])
    df["TIR"]    = pd.to_numeric(df["TIR"],   errors="coerce")
    df["Spread"] = pd.to_numeric(df["Spread"],errors="coerce")
    df["Dur"]    = pd.to_numeric(df["Dur"],   errors="coerce")

    # Parsear fechas string DD/MM/YYYY → date
    def parse_fecha(v):
        try:
            return datetime.strptime(str(v).strip(), "%d/%m/%Y").date()
        except:
            return None
    df["FechaDt"] = df["Fecha"].apply(parse_fecha)
    df = df.dropna(subset=["FechaDt"])
    fechas_sorted = sorted(df["FechaDt"].unique())

    def clasif(r):
        r = str(r).strip()
        if not r or r=="nan": return "Sin Rating"
        if r.startswith("CP"):   return "Corto Plazo"
        if r in ["AAA","AA+","AA","AA-","A+","A","A-","BBB+","BBB","BBB-"]:
            return "Investment Grade"
        return "Speculative Grade"
    df["Grado"] = df["Rating"].apply(clasif)

    tramos = [(0,1,"0-1a"),(1,3,"1-3a"),(3,5,"3-5a"),(5,10,"5-10a"),(10,99,"10a+")]
    by_moneda = {}
    for mon in ["PEN","USD","VAC"]:
        sub = df[df["Moneda"]==mon]
        ultima = fechas_sorted[-1] if fechas_sorted else None

        # Curva yield (última fecha)
        curva = {}
        if ultima:
            sf = sub[sub["FechaDt"]==ultima]
            for lo,hi,lbl in tramos:
                seg = sf[(sf["Dur"]>=lo)&(sf["Dur"]<hi)]["TIR"].dropna()
                curva[lbl] = round(float(seg.median()),3) if not seg.empty else None

        # Por grado (última fecha)
        por_grado = {}
        if ultima:
            sf = sub[sub["FechaDt"]==ultima]
            for g in ["Investment Grade","Speculative Grade","Corto Plazo","Sin Rating"]:
                seg = sf[sf["Grado"]==g]["TIR"].dropna()
                por_grado[g] = {"mediana_tir":round(float(seg.median()),3) if not seg.empty else None,
                                 "n":len(seg)}

        # Evolución TIR mediana por fecha
        evol_tir = {}
        for f in fechas_sorted:
            sf = sub[sub["FechaDt"]==f]["TIR"].dropna()
            evol_tir[str(f)] = round(float(sf.median()),3) if not sf.empty else None

        by_moneda[mon] = {
            "curva_yield": curva,
            "por_grado":   por_grado,
            "evol_tir":    evol_tir,
            "ultima_fecha":str(ultima),
        }

    # Scatter TIR vs Dur (última fecha, todas las monedas)
    scatter = []
    if fechas_sorted:
        sf = df[df["FechaDt"]==fechas_sorted[-1]]
        for _, row in sf.iterrows():
            if pd.notna(row["TIR"]) and pd.notna(row["Dur"]):
                scatter.append({
                    "emisor": str(row["Emisor"])[:18],
                    "nemo":   str(row["Nemo"]),
                    "tir":    round(float(row["TIR"]),3),
                    "dur":    round(float(row["Dur"]),3),
                    "spread": round(float(row["Spread"]),3) if pd.notna(row["Spread"]) else None,
                    "moneda": str(row["Moneda"]),
                    "rating": str(row["Rating"]) if pd.notna(row["Rating"]) else "",
                    "grado":  row["Grado"],
                })

    return {
        "by_moneda":   by_moneda,
        "scatter":     scatter,
        "fechas":      [str(f) for f in fechas_sorted],
        "ultima_fecha":str(fechas_sorted[-1]) if fechas_sorted else "",
    }


# ══════════════════════════════════════════════════════════════════
# HTML
# ══════════════════════════════════════════════════════════════════
def generar_html(s1, s2, s3, s4, s5):
    ts_gen = datetime.now().strftime("%d/%m/%Y %H:%M")
    data_js = json.dumps({
        "s1":s1,"s2":s2,
        "s3v":s3["ventanas"],"s3s":s3["sect_performance"],
        "s3e":s3["empresas"],"s3sec":s3["sectores"],
        "s3sm":s3["sector_map"],"s3el":s3["elecciones"],
        "s3ser":s3["series"],
        "s4":s4,"s5":s5,
    }, ensure_ascii=False, default=str)

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Elecciones Perú 2026 — AFP Integra RVL</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
:root{{--bg:#0a0d14;--panel:#111520;--brd:#1e2535;--jp:#22c55e;--fp:#f97316;
      --txt:#e2e8f0;--mut:#64748b;--acc:#38bdf8;--gld:#fbbf24;}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--txt);font-family:'IBM Plex Sans',sans-serif;font-size:13px;line-height:1.5}}
.mono{{font-family:'IBM Plex Mono',monospace}}
nav{{position:sticky;top:0;z-index:100;background:rgba(10,13,20,.96);backdrop-filter:blur(8px);
     border-bottom:1px solid var(--brd);display:flex;align-items:center;gap:6px;padding:10px 20px;flex-wrap:wrap}}
nav .logo{{font-weight:700;font-size:14px;color:var(--acc);letter-spacing:.06em;margin-right:12px}}
nav a{{color:var(--mut);text-decoration:none;font-size:11px;font-weight:600;letter-spacing:.05em;
       padding:3px 8px;border-radius:3px;transition:.2s}}
nav a:hover{{color:var(--txt);background:var(--brd)}}
nav .ts{{margin-left:auto;font-size:10px;color:var(--mut)}}
section{{padding:28px 20px;border-bottom:1px solid var(--brd);max-width:1380px;margin:0 auto}}
.sec-tag{{font-size:10px;font-weight:700;letter-spacing:.14em;color:var(--acc);text-transform:uppercase;margin-bottom:4px}}
.sec-h{{font-size:20px;font-weight:700;margin-bottom:18px}}
.g2{{display:grid;grid-template-columns:1fr 1fr;gap:14px}}
.g3{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px;margin-bottom:18px}}
.g4{{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px;margin-bottom:18px}}
.panel{{background:var(--panel);border:1px solid var(--brd);border-radius:8px;padding:14px}}
.panel-t{{font-size:12px;font-weight:600;margin-bottom:10px}}
.panel-sub{{font-size:10px;color:var(--mut);margin-bottom:8px}}
.card{{background:var(--panel);border:1px solid var(--brd);border-radius:7px;padding:14px}}
.card-lbl{{font-size:10px;letter-spacing:.08em;text-transform:uppercase;color:var(--mut)}}
.card-val{{font-size:26px;font-weight:700;margin:4px 0}}
.card-sub{{font-size:10px;color:var(--mut)}}
.jp{{color:var(--jp)}} .fp{{color:var(--fp)}}
.scoreboard{{background:var(--panel);border:1px solid var(--brd);border-radius:10px;padding:18px;
             margin-bottom:16px;display:grid;grid-template-columns:1fr 120px 1fr;gap:12px;align-items:center}}
.sc-side{{text-align:center}}
.sc-partido{{font-size:10px;letter-spacing:.1em;font-weight:700;text-transform:uppercase;margin-bottom:4px}}
.sc-pct{{font-size:44px;font-weight:700;line-height:1}}
.sc-sub{{font-size:10px;color:var(--mut);margin-top:3px}}
.sc-mid{{text-align:center}}
.sc-mid .lbl{{font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:var(--mut);margin-bottom:4px}}
.sc-mid .val{{font-size:22px;font-weight:700;color:var(--gld)}}
.bar-wrap{{height:7px;background:var(--brd);border-radius:4px;overflow:hidden;margin:10px 0}}
.bar-jp{{height:100%;background:var(--jp);border-radius:4px;transition:width 1s}}
.tbl{{width:100%;border-collapse:collapse;font-size:11px}}
.tbl th{{text-align:left;padding:5px 8px;font-size:9px;letter-spacing:.08em;text-transform:uppercase;
          color:var(--mut);border-bottom:1px solid var(--brd)}}
.tbl td{{padding:5px 8px;border-bottom:1px solid rgba(30,37,53,.4)}}
.tbl tr:hover td{{background:rgba(255,255,255,.02)}}
.tbl-wrap{{overflow-x:auto;border-radius:7px}}
.bdg{{display:inline-block;padding:1px 5px;border-radius:3px;font-size:9px;font-weight:700}}
.bdg-jp{{background:rgba(34,197,94,.14);color:var(--jp)}}
.bdg-fp{{background:rgba(249,115,22,.14);color:var(--fp)}}
select,button{{background:var(--panel);color:var(--txt);border:1px solid var(--brd);
               border-radius:4px;padding:4px 8px;font-size:11px;cursor:pointer;font-family:inherit}}
select:hover,button:hover{{border-color:var(--acc)}}
.ctrl{{display:flex;gap:7px;flex-wrap:wrap;margin-bottom:10px;align-items:center}}
.ctrl-lbl{{font-size:10px;color:var(--mut)}}
.ref{{background:rgba(56,189,248,.05);border:1px solid rgba(56,189,248,.18);border-radius:6px;
      padding:12px;font-size:11px;line-height:1.7;color:#94a3b8;margin-top:12px}}
canvas{{max-height:320px}}
.up{{color:#22c55e}} .dn{{color:#f87171}} .neu{{color:var(--mut)}}
@media(max-width:780px){{.g2,.scoreboard{{grid-template-columns:1fr}}.sc-pct{{font-size:32px}}}}
</style>
</head>
<body>
<nav>
  <span class="logo">⬡ AFP INTEGRA · RVL</span>
  <a href="#s1">ONPE</a><a href="#s2">Encuestas</a>
  <a href="#s3">Acciones</a><a href="#s4">Modelo</a><a href="#s5">Renta Fija</a>
  <span class="ts mono">Generado: {ts_gen}</span>
</nav>

<!-- S1 ONPE -->
<section id="s1">
  <div class="sec-tag">Sección 1</div>
  <div class="sec-h">Resultados ONPE — Segunda Vuelta 2026 <span id="s1-ts" class="mono" style="font-size:12px;color:var(--mut)"></span></div>
  <div class="scoreboard">
    <div class="sc-side">
      <div class="sc-partido jp">JP · Juntos por el Perú</div>
      <div class="sc-pct jp" id="sc-jp">—</div>
      <div class="sc-sub" id="sc-vj">—</div>
      <div class="sc-sub" id="sc-rj">—</div>
    </div>
    <div class="sc-mid">
      <div class="lbl">Actas</div>
      <div class="val" id="sc-pct-actas">—%</div>
      <div class="sc-sub" id="sc-actas2" style="font-size:9px;color:var(--mut)"></div>
    </div>
    <div class="sc-side">
      <div class="sc-partido fp">FP · Fuerza Popular</div>
      <div class="sc-pct fp" id="sc-fp">—</div>
      <div class="sc-sub" id="sc-vf">—</div>
      <div class="sc-sub" id="sc-rf">—</div>
    </div>
  </div>
  <div class="bar-wrap"><div class="bar-jp" id="bar-jp" style="width:50%"></div></div>
  <div class="g2" style="margin-bottom:14px">
    <div class="panel"><div class="panel-t">% votos por región</div><canvas id="cRegiones"></canvas></div>
    <div class="panel"><div class="panel-t">Avance de actas por región</div><canvas id="cActas"></canvas></div>
  </div>
  <div class="tbl-wrap"><table class="tbl">
    <thead><tr>
      <th>Región</th><th>Gana</th>
      <th style="color:var(--jp)">JP%</th><th style="color:var(--fp)">FP%</th>
      <th>Margen</th><th>Actas Cont.</th><th>Total</th><th>%Actas</th><th>Hora</th>
    </tr></thead>
    <tbody id="tbl-reg"></tbody>
  </table></div>
</section>

<!-- S2 ENCUESTAS -->
<section id="s2">
  <div class="sec-tag">Sección 2</div>
  <div class="sec-h">Conteo Rápido · Datum &amp; Ipsos · 07 Jun 2026</div>
  <div class="g4" id="cards-enc"></div>
  <div class="g2">
    <div class="panel"><div class="panel-t">Datum — Intervalo de confianza (±1%)</div><canvas id="cDatIC"></canvas></div>
    <div class="panel"><div class="panel-t">Ipsos — Desagregado por zona</div><canvas id="cIpsos"></canvas></div>
  </div>
  <div class="panel" style="margin-top:14px"><div class="panel-t">Datum — Evolución del voto JP durante la campaña</div><canvas id="cEvol" style="max-height:200px"></canvas></div>
  <div class="panel" style="margin-top:14px"><div class="panel-t">Datum CR vs ONPE parcial — JP% por región</div><canvas id="cDatReg"></canvas></div>
</section>

<!-- S3 ACCIONES -->
<section id="s3">
  <div class="sec-tag">Sección 3</div>
  <div class="sec-h">Acciones BVL — Ventanas Electorales 2016–2026</div>
  <div class="ctrl">
    <span class="ctrl-lbl">Acción:</span><select id="selEmp"></select>
    <span class="ctrl-lbl" style="margin-left:10px">Elección:</span>
    <select id="selElec">
      <option value="1V_2016">1ª Vuelta 2016 (10 abr)</option>
      <option value="2V_2016">2ª Vuelta 2016 (05 jun)</option>
      <option value="1V_2021">1ª Vuelta 2021 (11 abr)</option>
      <option value="2V_2021">2ª Vuelta 2021 (06 jun)</option>
      <option value="1V_2026">1ª Vuelta 2026 (13 abr)</option>
      <option value="2V_2026" selected>2ª Vuelta 2026 (07 jun)</option>
    </select>
  </div>
  <div class="panel" style="margin-bottom:14px">
    <div class="panel-t" id="t3-title">Retorno ±30 días (día 0 = elección)</div>
    <div class="panel-sub">Retorno relativo al precio del día -30. Línea punteada = promedio del sector.</div>
    <canvas id="cStocks"></canvas>
  </div>
  <div class="panel">
    <div class="panel-t">Retorno por sector en días clave</div>
    <div class="panel-sub">Promedio simple del sector. Días: -5, 0, +5, +10, +20 respecto a la elección.</div>
    <div class="ctrl">
      <span class="ctrl-lbl">Elección:</span>
      <select id="selElec2">
        <option value="1V_2016">1ª Vuelta 2016</option>
        <option value="2V_2016">2ª Vuelta 2016</option>
        <option value="1V_2021">1ª Vuelta 2021</option>
        <option value="2V_2021">2ª Vuelta 2021</option>
        <option value="1V_2026">1ª Vuelta 2026</option>
        <option value="2V_2026" selected>2ª Vuelta 2026</option>
      </select>
    </div>
    <canvas id="cSect"></canvas>
  </div>
</section>

<!-- S4 MODELO -->
<section id="s4">
  <div class="sec-tag">Sección 4 · Beta-Binomial Bayesiano</div>
  <div class="sec-h">Modelo Estadístico — Predicción del Ganador</div>
  <div class="g4" id="cards-mod"></div>
  <div class="g2">
    <div class="panel">
      <div class="panel-t">Distribución posterior θ ~ Beta(α,β)</div>
      <div class="panel-sub mono" id="mod-eq" style="font-size:9px"></div>
      <canvas id="cPost"></canvas>
    </div>
    <div class="panel">
      <div class="panel-t">P(JP gana) conforme llegan regiones</div>
      <canvas id="cEvolP"></canvas>
    </div>
  </div>
  <div class="ref" id="ref-mod"></div>
</section>

<!-- S5 RENTA FIJA -->
<section id="s5">
  <div class="sec-tag">Sección 5</div>
  <div class="sec-h">Renta Fija Local — Mercado Secundario</div>
  <div class="ctrl">
    <span class="ctrl-lbl">Moneda:</span>
    <button class="btn-m" data-m="PEN" id="btnPEN" style="border-color:var(--acc);color:var(--acc)">PEN (Sol)</button>
    <button class="btn-m" data-m="USD" id="btnUSD">USD (Dólar)</button>
    <button class="btn-m" data-m="VAC" id="btnVAC">VAC (Indexado)</button>
  </div>
  <div class="g2" style="margin-bottom:14px">
    <div class="panel"><div class="panel-t">Curva de rendimientos — TIR mediana por duración</div><div class="panel-sub" id="rf-lbl"></div><canvas id="cCurva"></canvas></div>
    <div class="panel"><div class="panel-t">TIR mediana por grado de riesgo</div><canvas id="cGrado"></canvas></div>
  </div>
  <div class="panel">
    <div class="panel-t">TIR vs Duración — todos los instrumentos (scatter)</div>
    <div class="ctrl">
      <span class="ctrl-lbl">Moneda:</span>
      <select id="selMon"><option value="TODAS">Todas</option><option value="PEN">PEN</option><option value="USD">USD</option><option value="VAC">VAC</option></select>
    </div>
    <canvas id="cScatter" style="max-height:380px"></canvas>
  </div>
</section>

<footer style="text-align:center;padding:18px;font-size:10px;color:var(--mut)">
  AFP Integra · Renta Variable Local · {ts_gen} · Fuentes: ONPE / Datum / Ipsos / Bloomberg
</footer>

<script>
const D = {data_js};
const S1=D.s1,S2=D.s2,S3v=D.s3v,S3s=D.s3s,S3e=D.s3e,S3sec=D.s3sec,S3sm=D.s3sm,S3ser=D.s3ser,S4=D.s4,S5=D.s5;
Chart.defaults.color='#64748b';Chart.defaults.borderColor='#1e2535';
Chart.defaults.font.family="'IBM Plex Sans',sans-serif";Chart.defaults.font.size=10;
const CJP='#22c55e',CFP='#f97316',CJP2='rgba(34,197,94,.12)',CFP2='rgba(249,115,22,.12)';
const fmt=(n,d=2)=>n==null?'—':Number(n).toFixed(d);
const fmtN=n=>n==null?'—':Number(n).toLocaleString('es-PE');

// ── S1 ─────────────────────────────────────────────────────
function initS1(){{
  const r=S1;
  if(!r.regiones||!r.regiones.length){{
    document.getElementById('sc-jp').textContent='Sin datos ONPE';
    document.getElementById('sc-fp').textContent='Ejecutar onpe_regional.py primero';
    return;
  }}
  document.getElementById('s1-ts').textContent='Actualizado: '+r.timestamp;
  document.getElementById('sc-jp').textContent=fmt(r.pct_nac_jp,3)+'%';
  document.getElementById('sc-fp').textContent=fmt(r.pct_nac_fp,3)+'%';
  document.getElementById('sc-vj').textContent=fmtN(r.total_votos_jp)+' votos';
  document.getElementById('sc-vf').textContent=fmtN(r.total_votos_fp)+' votos';
  document.getElementById('sc-rj').textContent=r.n_regiones_jp+' regiones';
  document.getElementById('sc-rf').textContent=r.n_regiones_fp+' regiones';
  const ac=r.regiones.reduce((s,x)=>s+(x.actas_contabilizadas||0),0);
  const at=r.regiones.reduce((s,x)=>s+(x.actas_total||0),0);
  document.getElementById('sc-pct-actas').textContent=(at>0?(ac/at*100).toFixed(1):'0')+'%';
  document.getElementById('sc-actas2').textContent=fmtN(ac)+' / '+fmtN(at)+' actas';
  document.getElementById('bar-jp').style.width=(r.pct_nac_jp||50)+'%';

  // Tabla
  const tb=document.getElementById('tbl-reg');
  r.regiones.forEach(reg=>{{
    const tr=document.createElement('tr');
    const g=reg.ganador==='JP'?`<span class="bdg bdg-jp">JP</span>`:`<span class="bdg bdg-fp">FP</span>`;
    const mc=reg.ganador==='JP'?CJP:CFP;
    const pc=reg.pct_contabilizadas>80?'up':reg.pct_contabilizadas>40?'neu':'dn';
    tr.innerHTML=`<td><b>${{reg.nombre}}</b></td><td>${{g}}</td>
      <td class="jp mono">${{fmt(reg.pct_jp,3)}}%</td><td class="fp mono">${{fmt(reg.pct_fp,3)}}%</td>
      <td class="mono" style="color:${{mc}}">+${{fmt(reg.margen,2)}}pp</td>
      <td class="mono">${{fmtN(reg.actas_contabilizadas)}}</td>
      <td class="mono">${{fmtN(reg.actas_total)}}</td>
      <td class="mono ${{pc}}">${{fmt(reg.pct_contabilizadas,1)}}%</td>
      <td class="mono" style="color:var(--mut);font-size:9px">${{reg.timestamp.slice(11,19)||'—'}}</td>`;
    tb.appendChild(tr);
  }});

  // Gráfico barras divergentes
  const regs=[...r.regiones].sort((a,b)=>b.pct_jp-a.pct_jp);
  new Chart(document.getElementById('cRegiones'),{{type:'bar',
    data:{{labels:regs.map(x=>x.nombre),
      datasets:[
        {{label:'JP%',data:regs.map(x=>x.pct_jp),backgroundColor:CJP,borderRadius:2}},
        {{label:'FP%',data:regs.map(x=>-x.pct_fp),backgroundColor:CFP,borderRadius:2}},
      ]}},
    options:{{indexAxis:'y',responsive:true,
      plugins:{{legend:{{position:'top'}},tooltip:{{callbacks:{{label:c=>c.datasetIndex===0?`JP: ${{c.raw.toFixed(2)}}%`:`FP: ${{Math.abs(c.raw).toFixed(2)}}%`}}}}}},
      scales:{{x:{{ticks:{{callback:v=>Math.abs(v)+'%'}}}},y:{{ticks:{{font:{{size:9}}}}}}}}
    }}}});

  // Actas
  const ra=[...r.regiones].sort((a,b)=>b.pct_contabilizadas-a.pct_contabilizadas);
  new Chart(document.getElementById('cActas'),{{type:'bar',
    data:{{labels:ra.map(x=>x.nombre),
      datasets:[{{label:'% actas',data:ra.map(x=>x.pct_contabilizadas),
        backgroundColor:ra.map(x=>x.pct_contabilizadas>80?'rgba(34,197,94,.7)':x.pct_contabilizadas>40?'rgba(251,191,36,.7)':'rgba(248,113,113,.7)'),
        borderRadius:2}}]}},
    options:{{indexAxis:'y',responsive:true,
      plugins:{{legend:{{display:false}}}},
      scales:{{x:{{max:100,ticks:{{callback:v=>v+'%'}}}},y:{{ticks:{{font:{{size:9}}}}}}}}
    }}}});
}}

// ── S2 ─────────────────────────────────────────────────────
function initS2(){{
  const dat=S2.datum,ips=S2.ipsos;
  const ce=document.getElementById('cards-enc');
  [dat,ips].forEach(d=>{{
    ce.innerHTML+=`<div class="card">
      <div class="card-lbl">${{d.fuente}} · ${{d.muestra}} · ±${{d.margen_error}}pp</div>
      <div style="display:flex;justify-content:space-between;margin-top:8px;align-items:flex-end">
        <div><div class="card-lbl">JP</div><div class="card-val jp">${{d.jp.toFixed(2)}}%</div></div>
        <div style="font-size:11px;font-weight:700;color:var(--gld);text-align:center;padding-bottom:6px">${{d.veredicto}}</div>
        <div style="text-align:right"><div class="card-lbl">FP</div><div class="card-val fp">${{d.fp.toFixed(2)}}%</div></div>
      </div>
      <div class="bar-wrap"><div class="bar-jp" style="width:${{d.jp}}%"></div></div>
    </div>`;
  }});

  new Chart(document.getElementById('cDatIC'),{{type:'bar',
    data:{{labels:['JP — Juntos','FP — Fuerza Popular'],
      datasets:[
        {{label:'Resultado CR',data:[dat.jp,dat.fp],backgroundColor:[CJP,CFP],borderRadius:4}},
        {{label:'Mín (−1%)',data:[dat.jp_min,dat.fp_min],backgroundColor:['rgba(34,197,94,.2)','rgba(249,115,22,.2)'],borderRadius:2}},
        {{label:'Máx (+1%)',data:[dat.jp_max,dat.fp_max],backgroundColor:['rgba(34,197,94,.2)','rgba(249,115,22,.2)'],borderRadius:2}},
      ]}},
    options:{{indexAxis:'y',responsive:true,
      scales:{{x:{{min:46,max:54,ticks:{{callback:v=>v+'%'}}}}}},
      plugins:{{legend:{{position:'top'}}}}
    }}}});

  const iZ=Object.keys(ips.desagregado);
  new Chart(document.getElementById('cIpsos'),{{type:'bar',
    data:{{labels:iZ,
      datasets:[
        {{label:'JP',data:iZ.map(z=>ips.desagregado[z].jp),backgroundColor:CJP,borderRadius:3}},
        {{label:'FP',data:iZ.map(z=>ips.desagregado[z].fp),backgroundColor:CFP,borderRadius:3}},
      ]}},
    options:{{responsive:true,plugins:{{legend:{{position:'top'}}}},
      scales:{{y:{{min:0,max:100,ticks:{{callback:v=>v+'%'}}}}}}}}}});

  const ev=dat.evolucion;
  new Chart(document.getElementById('cEvol'),{{type:'line',
    data:{{labels:ev.map(e=>e.label),
      datasets:[
        {{label:'JP',data:ev.map(e=>e.jp),borderColor:CJP,backgroundColor:CJP2,fill:true,tension:.4,pointRadius:5}},
        {{label:'FP',data:ev.map(e=>e.fp),borderColor:CFP,backgroundColor:CFP2,fill:true,tension:.4,pointRadius:5}},
      ]}},
    options:{{responsive:true,plugins:{{legend:{{position:'top'}}}},
      scales:{{y:{{min:44,max:55,ticks:{{callback:v=>v+'%'}}}}}}}}}});

  const dr=dat.por_region,dk=Object.keys(dr).sort();
  const om={{}};(S1.regiones||[]).forEach(r=>om[r.nombre]=r.pct_jp);
  new Chart(document.getElementById('cDatReg'),{{type:'bar',
    data:{{labels:dk,
      datasets:[
        {{label:'JP — Datum CR',data:dk.map(k=>dr[k]),backgroundColor:'rgba(34,197,94,.6)',borderRadius:2}},
        {{label:'JP — ONPE parcial',data:dk.map(k=>om[k]||null),backgroundColor:'rgba(56,189,248,.6)',borderRadius:2}},
      ]}},
    options:{{responsive:true,plugins:{{legend:{{position:'top'}}}},
      scales:{{y:{{min:0,max:100,ticks:{{callback:v=>v+'%'}}}},x:{{ticks:{{font:{{size:8}}}}}}}}
    }}}});
}}

// ── S3 ─────────────────────────────────────────────────────
let cSt=null,cSe=null;
function initS3(){{
  const s=document.getElementById('selEmp');
  S3e.forEach(e=>{{const o=document.createElement('option');o.value=o.textContent=e;s.appendChild(o)}});
  s.addEventListener('change',updSt);
  document.getElementById('selElec').addEventListener('change',updSt);
  document.getElementById('selElec2').addEventListener('change',updSe);
  updSt();updSe();
}}
function updSt(){{
  const emp=document.getElementById('selEmp').value;
  const elec=document.getElementById('selElec').value;
  if(cSt)cSt.destroy();
  const vd=S3v[elec];
  if(!vd||!vd.empresas[emp])return;
  const ed=vd.empresas[emp],sec=ed.sector;
  document.getElementById('t3-title').textContent=`${{emp}} (${{sec}}) — Retorno ±30d · ${{elec.replace('_',' ')}} (${{vd.fecha}})`;
  const allE=Object.keys(vd.empresas).filter(e=>vd.empresas[e].sector===sec);
  const sRet=ed.dias.map(d=>{{
    const vs=allE.map(e=>{{const ei=vd.empresas[e];const i=ei.dias.indexOf(d);return i>=0?ei.ret[i]:null}}).filter(v=>v!=null);
    return vs.length?vs.reduce((a,b)=>a+b)/vs.length:null;
  }});
  const d0i=ed.dias.indexOf(0);
  cSt=new Chart(document.getElementById('cStocks'),{{type:'line',
    data:{{
      labels:ed.dias.map(d=>d===0?`Día 0`:d>0?`+${{d}}d`:`${{d}}d`),
      datasets:[
        {{label:emp,data:ed.ret,borderColor:CJP,backgroundColor:CJP2,fill:true,tension:.3,pointRadius:2,borderWidth:2}},
        {{label:`Sector ${{sec}}`,data:sRet,borderColor:'#38bdf8',borderDash:[4,3],tension:.3,pointRadius:0,borderWidth:1.5}},
      ]}},
    options:{{responsive:true,plugins:{{legend:{{position:'top'}}}},
      scales:{{y:{{ticks:{{callback:v=>v.toFixed(1)+'%'}}}}}}
    }}}});
}}
function updSe(){{
  const elec=document.getElementById('selElec2').value;
  if(cSe)cSe.destroy();
  const sd=S3s[elec];
  if(!sd||!Object.keys(sd).length)return;
  const dias=[-5,0,5,10,20];
  const cols=['#94a3b8','#fbbf24','#34d399','#60a5fa','#f472b6'];
  const sects=Object.keys(sd);
  cSe=new Chart(document.getElementById('cSect'),{{type:'bar',
    data:{{labels:sects,
      datasets:dias.map((d,i)=>({{
        label:d===0?'Día 0':d>0?`+${{d}}d`:`${{d}}d`,
        data:sects.map(s=>sd[s]?.[d]??null),
        backgroundColor:cols[i]+'aa',borderColor:cols[i],borderWidth:1,borderRadius:3
      }}))}},
    options:{{responsive:true,plugins:{{legend:{{position:'top'}}}},
      scales:{{y:{{ticks:{{callback:v=>v!=null?v.toFixed(1)+'%':'—'}}}}}}
    }}}});
}}

// ── S4 ─────────────────────────────────────────────────────
function initS4(){{
  const m=S4;
  const cm=document.getElementById('cards-mod');
  [['P(JP gana)',m.prob_jp_wins+'%',CJP,'Posterior Bayesiano'],
   ['P(FP gana)',m.prob_fp_wins+'%',CFP,'Posterior Bayesiano'],
   ['Posterior JP',m.posterior_jp+'%',CJP,`IC 95%: [${{m.ci_lo}}% — ${{m.ci_hi}}%]`],
   ['Votos ONPE obs.',fmtN(m.total_obs),'#38bdf8',m.pct_escrutado.toFixed(1)+'% actas'],
  ].forEach(([l,v,c,s])=>cm.innerHTML+=`<div class="card"><div class="card-lbl">${{l}}</div><div class="card-val" style="color:${{c}}">${{v}}</div><div class="card-sub">${{s}}</div></div>`);

  document.getElementById('mod-eq').textContent=
    `θ|datos ~ Beta(${{(m.alpha_prior+m.votos_obs_jp).toFixed(0)}}, ${{(m.beta_prior+m.votos_obs_fp).toFixed(0)}})  Prior: κ=${{m.kappa}}, θ₀=${{m.prior_jp}}%`;

  const mu=m.posterior_jp/100,sd=m.std/100;
  const xs=Array.from({{length:200}},(_,i)=>mu-4*sd+i*(8*sd/200));
  const pdf=xs.map(x=>Math.exp(-(x-mu)**2/(2*sd**2))/(sd*Math.sqrt(2*Math.PI)));
  const col=m.prob_jp_wins>=50?CJP:CFP;
  new Chart(document.getElementById('cPost'),{{type:'line',
    data:{{labels:xs.map(x=>(x*100).toFixed(2)+'%'),
      datasets:[{{label:'P(θ|datos)',data:pdf,borderColor:col,
        backgroundColor:m.prob_jp_wins>=50?CJP2:CFP2,fill:true,tension:.4,pointRadius:0,borderWidth:2}}]}},
    options:{{responsive:true,plugins:{{legend:{{display:false}}}},scales:{{x:{{ticks:{{maxTicksLimit:6}}}},y:{{display:false}}}}}}
  }});

  const ev=m.evolucion;
  if(ev&&ev.length){{
    new Chart(document.getElementById('cEvolP'),{{type:'line',
      data:{{labels:ev.map(e=>e.region),
        datasets:[
          {{label:'P(JP gana) %',data:ev.map(e=>e.prob_jp),borderColor:CJP,backgroundColor:CJP2,fill:true,tension:.3,pointRadius:3,borderWidth:2}},
          {{label:'50%',data:ev.map(_=>50),borderColor:'rgba(255,255,255,.15)',borderDash:[4,4],pointRadius:0,borderWidth:1}},
        ]}},
      options:{{responsive:true,plugins:{{legend:{{position:'top'}}}},
        scales:{{y:{{min:0,max:100,ticks:{{callback:v=>v+'%'}}}},x:{{ticks:{{font:{{size:8}}}}}}}}
      }}}});
  }}

  document.getElementById('ref-mod').innerHTML=`
    <b style="color:var(--acc)">Metodología: Beta-Binomial Bayesiano</b><br>
    <b>Prior</b>: promedio Datum (${{S2.datum.jp}}%) + Ipsos (${{S2.ipsos.jp}}%) → θ₀=${{m.prior_jp}}%, concentración κ=${{m.kappa}} votos virtuales.<br>
    <b>Likelihood</b>: ${{fmtN(m.votos_obs_jp)}} votos JP y ${{fmtN(m.votos_obs_fp)}} FP en ONPE (${{m.pct_escrutado}}% actas).<br>
    <b>Posterior</b>: θ=${{m.posterior_jp}}% ± ${{m.std}}pp · IC95%: [${{m.ci_lo}}%–${{m.ci_hi}}%] · <b>P(JP>50%)=${{m.prob_jp_wins}}%</b><br>
    <i>Refs: Jackman (2005) Australian J. Political Science 40(4) 499-517 · Linzer (2013) JASA 108(501) 124-134</i>`;
}}

// ── S5 ─────────────────────────────────────────────────────
let cCurva=null,cGrado=null,cScat=null;
let monAct='PEN';
function initS5(){{
  document.querySelectorAll('.btn-m').forEach(b=>{{
    b.addEventListener('click',()=>{{
      document.querySelectorAll('.btn-m').forEach(x=>{{x.style.borderColor='';x.style.color=''}});
      b.style.borderColor='var(--acc)';b.style.color='var(--acc)';
      monAct=b.dataset.m;updRF();
    }});
  }});
  document.getElementById('selMon').addEventListener('change',updScat);
  updRF();updScat();
}}
function updRF(){{
  const md=S5.by_moneda[monAct];
  if(!md)return;
  document.getElementById('rf-lbl').textContent='Última fecha: '+md.ultima_fecha;
  const curva=md.curva_yield,tr=Object.keys(curva);
  if(cCurva)cCurva.destroy();
  cCurva=new Chart(document.getElementById('cCurva'),{{type:'line',
    data:{{labels:tr,datasets:[{{label:`TIR mediana ${{monAct}}`,data:tr.map(t=>curva[t]),
      borderColor:'#38bdf8',backgroundColor:'rgba(56,189,248,.1)',fill:true,tension:.4,pointRadius:5,borderWidth:2}}]}},
    options:{{responsive:true,scales:{{y:{{ticks:{{callback:v=>v!=null?v.toFixed(2)+'%':'—'}}}}}}}}
  }});
  const gr=md.por_grado,gk=Object.keys(gr).filter(g=>gr[g].n>0);
  const gcol={{'Investment Grade':'#34d399','Speculative Grade':'#f87171','Corto Plazo':'#60a5fa','Sin Rating':'#94a3b8'}};
  if(cGrado)cGrado.destroy();
  cGrado=new Chart(document.getElementById('cGrado'),{{type:'bar',
    data:{{labels:gk,datasets:[{{label:`TIR ${{monAct}}`,data:gk.map(g=>gr[g].mediana_tir),
      backgroundColor:gk.map(g=>(gcol[g]||'#94a3b8')+'aa'),borderColor:gk.map(g=>gcol[g]||'#94a3b8'),
      borderWidth:1,borderRadius:4}}]}},
    options:{{responsive:true,plugins:{{legend:{{display:false}},
      tooltip:{{callbacks:{{afterLabel:ctx=>`n = ${{gr[gk[ctx.dataIndex]].n}} instrumentos`}}}}}},
      scales:{{y:{{ticks:{{callback:v=>v!=null?v.toFixed(2)+'%':'—'}}}}}}}}}});
}}
function updScat(){{
  const mf=document.getElementById('selMon').value;
  let pts=S5.scatter;
  if(mf!=='TODAS')pts=pts.filter(p=>p.moneda===mf);
  const gcol={{'Investment Grade':'rgba(52,211,153,.8)','Speculative Grade':'rgba(248,113,113,.8)','Corto Plazo':'rgba(96,165,250,.8)','Sin Rating':'rgba(148,163,184,.5)'}};
  if(cScat)cScat.destroy();
  cScat=new Chart(document.getElementById('cScatter'),{{type:'scatter',
    data:{{datasets:['Investment Grade','Speculative Grade','Corto Plazo','Sin Rating'].map(g=>{{
      const sub=pts.filter(p=>p.grado===g);
      return{{label:g,data:sub.map(p=>( {{x:p.dur,y:p.tir,emisor:p.emisor,nemo:p.nemo,mon:p.moneda,rat:p.rating,sp:p.spread}} )),
        backgroundColor:gcol[g],pointRadius:5,pointHoverRadius:8}};
    }})}},
    options:{{responsive:true,
      plugins:{{legend:{{position:'top'}},tooltip:{{callbacks:{{label:c=>{{const d=c.raw;
        return[`${{d.emisor}} (${{d.nemo}})`,`TIR: ${{d.y?.toFixed(3)}}% | Dur: ${{d.x?.toFixed(2)}}a`,
               `Spread: ${{d.sp!=null?d.sp.toFixed(3)+'%':'—'}} | ${{d.mon}} | ${{d.rat}}`];}}}}}}}},
      scales:{{x:{{title:{{display:true,text:'Duración (años)'}},min:0}},
               y:{{title:{{display:true,text:'TIR %'}},ticks:{{callback:v=>v.toFixed(2)+'%'}}}}}}
    }}}});
}}

initS1();initS2();initS3();initS4();initS5();
</script></body></html>"""

# ══════════════════════════════════════════════════════════════════
# RUN
# ══════════════════════════════════════════════════════════════════
print("  [1/5] ONPE regional...", end=" ")
s1 = build_s1(_df_consolidado)
print(f"{len(s1.get('regiones',[]))} regiones")

print("  [2/5] Encuestadoras...", end=" ")
s2 = build_s2()
print("OK")

print("  [3/5] Stocks...", end=" ")
s3 = build_s3()
print(f"{len(s3['empresas'])} empresas · {len(s3['ventanas'])} ventanas")

print("  [4/5] Modelo Bayesiano...", end=" ")
s4 = build_s4(s1, s2)
print(f"P(JP gana)={s4['prob_jp_wins']}%")

print("  [5/5] Renta Fija...", end=" ")
s5 = build_s5()
print(f"{len(s5['scatter'])} instrumentos")

print("  Generando HTML...", end=" ")
html = generar_html(s1, s2, s3, s4, s5)
OUT_PATH.write_text(html, encoding="utf-8")
size_kb = OUT_PATH.stat().st_size // 1024
print(f"OK — {size_kb} KB")
print(f"  ✓ {OUT_PATH}")
print(f"{'═'*60}\n")