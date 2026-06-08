"""
cocina.py — Dashboard Electoral Perú 2026
==========================================
Ejecutar en Spyder DESPUÉS de onpe_regional.py.
df_consolidado debe estar en el namespace de Spyder.

Genera: elecciones_2026.html (autocontenido, enviable por email)
Paleta: AFP Integra corporativa, fondo blanco
"""

import json, warnings
from pathlib import Path
from datetime import datetime, date
from math import erf, sqrt
import numpy as np
import pandas as pd
warnings.filterwarnings("ignore")

# ── RUTAS ──────────────────────────────────────────────────────────
BASE        = Path(r"C:\Users\usuario\OneDrive\Desktop\AFP INTEGRA\Modelo Franco\Elecciones")
STOCKS_PATH = BASE / "Stocks_prices.xlsx"
RF_PATH     = BASE / "Renta Fija Local" / "reporte.xlsx"
OUT_PATH    = BASE / "elecciones_2026.html"

# ── PALETA AFP INTEGRA ────────────────────────────────────────────
BRAND = {
    "navy":   "#1E2E6E",
    "dark":   "#002060",
    "teal":   "#00AECB",
    "yellow": "#E3E829",
    "gray1":  "#DCDDDE",
    "gray2":  "#7E8083",
    "white":  "#FFFFFF",
    "green":  "#2ec97a",
    "red":    "#e84343",
    "amber":  "#e0a320",
}
JP_COLOR   = "#22c55e"
FP_COLOR   = "#f97316"
JP_LIGHT   = "rgba(34,197,94,0.15)"
FP_LIGHT   = "rgba(249,115,22,0.15)"

# ── ELECCIONES ────────────────────────────────────────────────────
ELECCIONES = {
    "1V_2016": date(2016, 4, 10),
    "2V_2016": date(2016, 6,  5),
    "1V_2021": date(2021, 4, 11),
    "2V_2021": date(2021, 6,  6),
    "1V_2026": date(2026, 4, 13),
    "2V_2026": date(2026, 6,  7),
}
VENTANA = 30

# ══════════════════════════════════════════════════════════════════
# CAPTURA df_consolidado desde Spyder/IPython
# ══════════════════════════════════════════════════════════════════
_df = None

# 1. Namespace local (exec/runfile)
_df = globals().get("df_consolidado", None)

# 2. Kernel IPython (F5 en Spyder)
if _df is None:
    try:
        _ipy = get_ipython()                  # noqa: F821
        _df = _ipy.shell.user_ns.get("df_consolidado", None)
    except Exception:
        pass

# 3. __main__
if _df is None:
    try:
        import __main__
        _df = getattr(__main__, "df_consolidado", None)
    except Exception:
        pass

# 4. Fallback: Excel guardado
_xlsx = BASE / "df_consolidado.xlsx"
if (_df is None or not isinstance(_df, pd.DataFrame) or _df.empty):
    if _xlsx.exists():
        _df = pd.read_excel(_xlsx)

if not isinstance(_df, pd.DataFrame) or _df.empty:
    _df = None

print(f"\n{'='*60}")
print("COCINA.PY — Dashboard Electoral 2026")
print(f"{'='*60}")
print(f"df_consolidado: {'OK — ' + str(len(_df)) + ' filas' if _df is not None else 'VACÍO'}")

# ══════════════════════════════════════════════════════════════════
# S1 — ONPE REGIONAL
# ══════════════════════════════════════════════════════════════════
def build_s1(df):
    empty = {"regiones":[], "extranjero":{}, "total_votos_jp":0,
             "total_votos_fp":0, "pct_nac_jp":0, "pct_nac_fp":0,
             "timestamp":"", "n_regiones_jp":0, "n_regiones_fp":0}
    if df is None or df.empty:
        return empty
    JP = "JUNTOS POR EL PERÚ"; FP = "FUERZA POPULAR"
    df_r = df[df["nivel"]=="regional"].copy()
    df_e = df[df["nivel"]=="extranjero"].copy()

    def safe_val(series, col, cast=float):
        try:
            v = series[col].values[0]
            return cast(v) if str(v) not in ("nan","None","") else 0
        except: return 0

    regiones = []
    for ubi, grp in df_r.groupby("ubigeo"):
        rj = grp[grp["partido"]==JP]; rf = grp[grp["partido"]==FP]
        pj = safe_val(rj, "porcentaje")
        pf = safe_val(rf, "porcentaje")
        regiones.append({
            "nombre": ubi,
            "pct_jp": round(pj,3), "pct_fp": round(pf,3),
            "votos_jp": safe_val(rj,"votos",int),
            "votos_fp": safe_val(rf,"votos",int),
            "ganador": "JP" if pj>pf else "FP",
            "margen": round(abs(pj-pf),3),
            "actas_cont": safe_val(rj,"actas_contabilizadas",int),
            "actas_jee":  safe_val(rj,"actas_jee",int),
            "actas_pend": safe_val(rj,"actas_pendientes",int),
            "actas_tot":  safe_val(rj,"actas_total",int),
            "pct_actas":  round(safe_val(rj,"pct_contabilizadas"),1),
            "ts": str(rj["timestamp"].values[0])[:19] if not rj.empty else "",
        })

    ext = {}
    if not df_e.empty:
        rj=df_e[df_e["partido"]==JP]; rf=df_e[df_e["partido"]==FP]
        ext={"pct_jp":safe_val(rj,"porcentaje"),"pct_fp":safe_val(rf,"porcentaje"),
             "votos_jp":safe_val(rj,"votos",int),"votos_fp":safe_val(rf,"votos",int)}

    tj=sum(r["votos_jp"] for r in regiones)
    tf=sum(r["votos_fp"] for r in regiones)
    tot=tj+tf
    return {
        "regiones": sorted(regiones, key=lambda x:x["nombre"]),
        "extranjero": ext,
        "total_votos_jp":tj, "total_votos_fp":tf,
        "pct_nac_jp": round(100*tj/tot,3) if tot>0 else 0,
        "pct_nac_fp": round(100*tf/tot,3) if tot>0 else 0,
        "timestamp": str(df_r["timestamp"].max())[:19] if not df_r.empty else "",
        "n_regiones_jp": sum(1 for r in regiones if r["ganador"]=="JP"),
        "n_regiones_fp": sum(1 for r in regiones if r["ganador"]=="FP"),
    }

# ══════════════════════════════════════════════════════════════════
# S2 — ENCUESTADORAS
# ══════════════════════════════════════════════════════════════════
def build_s2():
    return {
        "datum": {
            "fuente":"Datum Internacional","muestra":"117,199 votos","error":1.0,
            "jp":50.14,"fp":49.86,"jp_min":49.14,"jp_max":51.14,
            "fp_min":48.86,"fp_max":50.86,"veredicto":"EMPATE ESTADÍSTICO",
            "por_region":{
                "Amazonas":65.26,"Áncash":56.72,"Apurímac":80.87,"Arequipa":62.74,
                "Ayacucho":79.35,"Cajamarca":67.98,"Callao":35.87,"Cusco":78.48,
                "Huancavelica":80.60,"Huánuco":64.75,"Ica":46.18,"Junín":53.39,
                "La Libertad":42.69,"Lambayeque":41.08,"Lima Met.":35.77,
                "Lima Prov.":47.64,"Loreto":44.92,"Madre de Dios":72.96,
                "Moquegua":66.81,"Pasco":62.70,"Piura":42.76,"Puno":87.84,
                "San Martín":54.04,"Tacna":74.04,"Tumbes":34.86,
                "Ucayali":46.89,"Extranjero":37.33,
            },
            "evolucion":[
                {"l":"Int.Voto\n17-20 May","jp":47.8,"fp":52.2},
                {"l":"Int.Voto\n26-30 May","jp":47.4,"fp":52.6},
                {"l":"Simulacro\n26-30 May","jp":47.1,"fp":52.9},
                {"l":"Int.Voto\n03-04 Jun","jp":48.8,"fp":51.2},
                {"l":"Simulacro\n06 Jun","jp":49.4,"fp":50.6},
                {"l":"Boca Urna\n07 Jun","jp":49.47,"fp":50.53},
                {"l":"C.Rápido\n07 Jun","jp":50.14,"fp":49.86},
            ],
        },
        "ipsos":{
            "fuente":"Ipsos / Transparencia","muestra":"1,037 actas","error":1.9,
            "jp":50.3,"fp":49.7,"jp_min":48.4,"jp_max":52.2,
            "fp_min":47.8,"fp_max":51.6,"veredicto":"EMPATE TÉCNICO",
            "desagregado":{
                "Lima":{"jp":36.4,"fp":63.6},"Regiones":{"jp":57.4,"fp":42.6},
                "Urbano":{"jp":46.1,"fp":53.9},"Rural":{"jp":69.0,"fp":31.0},
                "Costa":{"jp":39.5,"fp":60.5},"Sierra":{"jp":70.2,"fp":29.8},
                "Selva":{"jp":58.6,"fp":41.4},
            },
        },
    }

# ══════════════════════════════════════════════════════════════════
# S3 — ACCIONES BVL
# ══════════════════════════════════════════════════════════════════
def build_s3():
    df = pd.read_excel(STOCKS_PATH, sheet_name="Hoja1", parse_dates=["Fecha"])
    df["Fecha"] = pd.to_datetime(df["Fecha"]).dt.date
    df = df.sort_values("Fecha").reset_index(drop=True)
    sm = pd.read_excel(STOCKS_PATH, sheet_name="Hoja2")
    sec_map = dict(zip(sm.iloc[:,0], sm.iloc[:,1]))
    empresas = [c for c in df.columns if c!="Fecha"]
    sectores = sorted(set(sec_map.values()))

    ventanas = {}
    for name, ed in ELECCIONES.items():
        mask = df["Fecha"].apply(lambda d: abs((d-ed).days)<=VENTANA)
        sub = df[mask].copy()
        if sub.empty: continue
        edata = {}
        for emp in empresas:
            col = sub[["Fecha",emp]].dropna()
            if len(col)<2: continue
            base = float(col[emp].iloc[0])
            if base==0: continue
            edata[emp] = {
                "dias":[(d-ed).days for d in col["Fecha"]],
                "ret":[round((v/base-1)*100,3) for v in col[emp]],
                "sector":sec_map.get(emp,"Otro"),
            }
        ventanas[name] = {"fecha":str(ed),"empresas":edata}

    sect_perf = {}
    for name, vd in ventanas.items():
        ps = {}
        for emp, ed2 in vd["empresas"].items():
            sec = ed2["sector"]
            for tgt in [-5,0,5,10,20]:
                if not ed2["dias"]: continue
                ci = min(range(len(ed2["dias"])), key=lambda i:abs(ed2["dias"][i]-tgt))
                if abs(ed2["dias"][ci]-tgt)<=3:
                    ps.setdefault(sec,{}).setdefault(tgt,[]).append(ed2["ret"][ci])
        sect_perf[name]={s:{t:round(np.mean(v),2) for t,v in td.items()} for s,td in ps.items()}

    return {"ventanas":ventanas,"sect_perf":sect_perf,
            "empresas":empresas,"sectores":sectores,"sec_map":sec_map,
            "elecciones":{k:str(v) for k,v in ELECCIONES.items()}}

# ══════════════════════════════════════════════════════════════════
# S4 — MODELO BAYESIANO BETA-BINOMIAL
# ══════════════════════════════════════════════════════════════════
def build_s4(s1, s2):
    # Prior: promedio encuestadoras
    p0 = (s2["datum"]["jp"] + s2["ipsos"]["jp"]) / 200.0
    K = 500  # concentración moderada
    a0, b0 = p0*K, (1-p0)*K

    vj = s1["total_votos_jp"]
    vf = s1["total_votos_fp"]
    ap, bp = a0+vj, b0+vf
    mu  = ap/(ap+bp)
    var = (ap*bp)/((ap+bp)**2*(ap+bp+1))
    sd  = max(var**0.5, 1e-10)
    z   = (mu-0.5)/sd
    pw  = 0.5*(1+erf(z/sqrt(2)))

    regs = s1.get("regiones",[])
    ac_t = sum(r.get("actas_tot",0) for r in regs)
    ac_c = sum(r.get("actas_cont",0) for r in regs)
    pesc = ac_c/ac_t*100 if ac_t>0 else 0

    # Distribución posterior para gráfico (100 puntos)
    n_pts = 100
    x_min = max(0.35, mu - 6*sd)
    x_max = min(0.65, mu + 6*sd)
    xs = [x_min + i*(x_max-x_min)/n_pts for i in range(n_pts+1)]
    import math
    pdf = [math.exp(-(x-mu)**2/(2*sd**2))/(sd*math.sqrt(2*math.pi)) for x in xs]
    # Normalizar para gráfico
    mx = max(pdf) if pdf else 1
    pdf_norm = [v/mx for v in pdf]

    # Evolución probabilidad por región
    evol = []
    aj, bj = a0, b0
    for r in sorted(regs, key=lambda x:x.get("ts","")):
        aj += r.get("votos_jp",0); bj += r.get("votos_fp",0)
        t2 = aj+bj
        if t2>0:
            m2 = aj/t2
            s2_ = max(((aj*bj)/(t2**2*(t2+1)))**0.5, 1e-10)
            z2 = (m2-0.5)/s2_
            p2 = 0.5*(1+erf(z2/sqrt(2)))
            evol.append({"r":r["nombre"],"p_jp":round(p2*100,1),"p_fp":round((1-p2)*100,1)})

    return {
        "prior_jp":round(p0*100,2),"kappa":K,
        "a0":round(a0,1),"b0":round(b0,1),
        "vj":vj,"vf":vf,"tot":vj+vf,
        "mu":round(mu*100,3),"sd":round(sd*100,4),
        "ci_lo":round(max(0,mu-1.96*sd)*100,2),
        "ci_hi":round(min(1,mu+1.96*sd)*100,2),
        "p_jp":round(pw*100,1),"p_fp":round((1-pw)*100,1),
        "pesc":round(pesc,1),"ac_c":ac_c,"ac_t":ac_t,
        "xs":[round(x*100,3) for x in xs],"pdf":pdf_norm,
        "evol":evol,
        "lider":"FP" if mu<0.5 else "JP",
        "ventaja":round(abs(mu-0.5)*200,3),
    }

# ══════════════════════════════════════════════════════════════════
# S5 — RENTA FIJA
# ══════════════════════════════════════════════════════════════════
def build_s5():
    """
    Renta Fija Local.
    Soberano     → Emisor==GOB.CENTRAL  o  Nemo starts with SB
    BCRP         → Emisor==BCRP         o  Nemo starts with CD
    Corporativo  → resto
    """
    df = pd.read_excel(RF_PATH, header=1)
    df.columns = [str(c).strip() for c in df.columns]

    col_map = {
        "Fecha":"Fecha","Nemónico":"Nemo","ISIN":"ISIN",
        "Emisor":"Emisor","Moneda":"Moneda",
        "TIR %":"TIR","Spreads":"Spread","Duración":"Dur",
        "F. Vencimiento":"F_Venc","Rating":"Rating",
    }
    df = df.rename(columns={k:v for k,v in col_map.items() if k in df.columns})
    df = df.dropna(subset=["Nemo"])
    for c in ["TIR","Spread","Dur"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    def _pv(v):
        try: return datetime.strptime(str(v).strip(), "%d/%m/%Y")
        except: return None
    df["Venc_dt"]   = df["F_Venc"].apply(_pv)
    df["Venc_year"] = df["Venc_dt"].apply(lambda x: x.year if x else None)

    def _pf(v):
        try: return datetime.strptime(str(v).strip(), "%d/%m/%Y").date()
        except: return None
    df["Fdt"] = df["Fecha"].apply(_pf)
    fechas = sorted([f for f in df["Fdt"].unique() if f])
    ult    = fechas[-1] if fechas else None
    dfu    = df[df["Fdt"]==ult].copy() if ult else df.copy()

    def _tipo(row):
        em=str(row.get("Emisor","")).strip(); nm=str(row.get("Nemo","")).strip()
        if em=="GOB.CENTRAL" or nm.startswith("SB"): return "Soberano"
        elif em=="BCRP"      or nm.startswith("CD"): return "BCRP"
        else:                                         return "Corporativo"
    dfu=dfu.copy(); dfu["Tipo"]=dfu.apply(_tipo, axis=1)

    def _grado(r):
        r=str(r).strip()
        if not r or r=="nan": return "Sin Rating"
        if r.startswith("CP"): return "Corto Plazo"
        if r in ["AAA","AA+","AA","AA-","A+","A","A-","BBB+","BBB","BBB-"]:
            return "Investment Grade"
        return "Speculative Grade"
    dfu["Grado"]=dfu["Rating"].apply(_grado)

    # Curva soberana PEN
    sob_p = dfu[(dfu["Tipo"]=="Soberano")&(dfu["Moneda"]=="PEN")&dfu["TIR"].notna()]
    sc = sob_p.groupby("Venc_year")["TIR"].median().reset_index().sort_values("Venc_year")
    curva_sob = {"years":[int(r["Venc_year"]) for _,r in sc.iterrows()],
                 "tirs": [round(float(r["TIR"]),3) for _,r in sc.iterrows()]}

    # Curva soberana VAC
    sob_v = dfu[(dfu["Tipo"]=="Soberano")&(dfu["Moneda"]=="VAC")&dfu["TIR"].notna()]
    vc = sob_v.groupby("Venc_year")["TIR"].median().reset_index().sort_values("Venc_year")
    curva_sob_vac = {"years":[int(r["Venc_year"]) for _,r in vc.iterrows()],
                     "tirs": [round(float(r["TIR"]),3) for _,r in vc.iterrows()]}

    # Curva BCRP CDs PEN
    bcrp = dfu[(dfu["Tipo"]=="BCRP")&(dfu["Moneda"]=="PEN")&dfu["TIR"].notna()&dfu["Venc_year"].notna()]
    bc = bcrp.groupby("Venc_year")["TIR"].median().reset_index().sort_values("Venc_year")
    curva_bcrp = {"years":[int(r["Venc_year"]) for _,r in bc.iterrows()],
                  "tirs": [round(float(r["TIR"]),3) for _,r in bc.iterrows()]}

    # Spread corporativo PEN vs soberano PEN (interpolado si no hay mismo año)
    sob_yr = dict(zip(sc["Venc_year"], sc["TIR"]))
    corp_p = dfu[(dfu["Tipo"]=="Corporativo")&(dfu["Moneda"]=="PEN")
                 &dfu["TIR"].notna()&dfu["Venc_year"].notna()]
    spread_data = []
    for yr, grp in corp_p.groupby("Venc_year"):
        ref = sob_yr.get(yr)
        if ref is None:
            yrs=sorted(sob_yr.keys()); bf=[y for y in yrs if y<=yr]; af=[y for y in yrs if y>yr]
            if bf and af:
                y0,y1=bf[-1],af[0]
                ref=sob_yr[y0]+(sob_yr[y1]-sob_yr[y0])*(yr-y0)/(y1-y0)
            elif bf: ref=sob_yr[bf[-1]]
            elif af: ref=sob_yr[af[0]]
        if ref is None: continue
        cm=float(grp["TIR"].median())
        spread_data.append({"year":int(yr),"corp":round(cm,3),
                             "sob":round(float(ref),3),
                             "spread":round(cm-float(ref),3),"n":int(len(grp))})
    spread_data.sort(key=lambda x:x["year"])

    # Scatter completo por tipo
    scatter=[]
    for _,row in dfu.iterrows():
        if pd.notna(row.get("TIR")) and pd.notna(row.get("Dur")):
            scatter.append({
                "e":  str(row.get("Emisor",""))[:18],
                "n":  str(row.get("Nemo","")),
                "tir":round(float(row["TIR"]),3),
                "dur":round(float(row["Dur"]),3),
                "sp": round(float(row["Spread"]),3) if pd.notna(row.get("Spread")) else None,
                "mon":str(row.get("Moneda","")),
                "rat":str(row.get("Rating","")) if pd.notna(row.get("Rating")) else "",
                "tipo":row["Tipo"],
                "g":  row["Grado"],
                "yr": int(row["Venc_year"]) if pd.notna(row.get("Venc_year")) else None,
            })

    # Resumen estadístico por tipo y moneda
    resumen={}
    for tipo in ["Soberano","BCRP","Corporativo"]:
        resumen[tipo]={}
        for mon in ["PEN","USD","VAC"]:
            sub=dfu[(dfu["Tipo"]==tipo)&(dfu["Moneda"]==mon)&dfu["TIR"].notna()]
            if sub.empty: continue
            resumen[tipo][mon]={"tir_med":round(float(sub["TIR"].median()),3),
                                 "tir_min":round(float(sub["TIR"].min()),3),
                                 "tir_max":round(float(sub["TIR"].max()),3),
                                 "n":int(len(sub))}

    return {"curva_sob":curva_sob,"curva_sob_vac":curva_sob_vac,
            "curva_bcrp":curva_bcrp,"spread_data":spread_data,
            "scatter":scatter,"resumen":resumen,"ult":str(ult) if ult else ""}

def generar_html(s1,s2,s3,s4,s5):
    ts = datetime.now().strftime("%d/%m/%Y %H:%M")
    # Paleta colores
    NAVY=BRAND["navy"]; TEAL=BRAND["teal"]; YELLOW=BRAND["yellow"]
    GRAY1=BRAND["gray1"]; GRAY2=BRAND["gray2"]

    data_js = json.dumps({
        "s1":s1,"s2":s2,
        "s3v":s3["ventanas"],"s3s":s3["sect_perf"],
        "s3e":s3["empresas"],"s3sec":s3["sectores"],
        "s3sm":s3["sec_map"],"s3el":s3["elecciones"],
        "s4":s4,"s5":s5,
    }, ensure_ascii=False, default=str)

    # Generar tabla de regiones como HTML estático (no JS)
    tbl_rows = ""
    if s1["regiones"]:
        for r in s1["regiones"]:
            g_cls = "bdg-jp" if r["ganador"]=="JP" else "bdg-fp"
            g_lbl = r["ganador"]
            mc = JP_COLOR if r["ganador"]=="JP" else FP_COLOR
            pc_cls = "up" if r["pct_actas"]>80 else ("neu" if r["pct_actas"]>40 else "dn")
            tbl_rows += f"""<tr>
<td><b>{r['nombre']}</b></td>
<td><span class="bdg {g_cls}">{g_lbl}</span></td>
<td class="num jp">{r['pct_jp']:.3f}%</td>
<td class="num fp">{r['pct_fp']:.3f}%</td>
<td class="num" style="color:{mc}">+{r['margen']:.2f}pp</td>
<td class="num">{r['actas_cont']:,}</td>
<td class="num">{r['actas_tot']:,}</td>
<td class="num {pc_cls}">{r['pct_actas']:.1f}%</td>
<td class="ts-cell">{r['ts'][11:19] if r['ts'] else '—'}</td>
</tr>"""

    # Scoreboard
    pj = s1["pct_nac_jp"]; pf = s1["pct_nac_fp"]
    bar_w = pj if pj>0 else 50
    winner_color = JP_COLOR if pj>pf else FP_COLOR
    lider = "JP" if pj>pf else "FP"
    lider_nombre = "JUNTOS POR EL PERÚ" if pj>pf else "FUERZA POPULAR"

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Elecciones Perú 2026 — AFP Integra RVL</title>
<!-- Chart.js: múltiples CDN para garantizar carga -->
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script>
if(typeof Chart==='undefined'){{
  document.write('<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"><\\/script>');
}}
</script>
<style>
/* ── AFP Integra Corporate Theme — fondo blanco ── */
:root{{
  --navy:{NAVY}; --teal:{TEAL}; --yellow:{YELLOW};
  --gray1:{GRAY1}; --gray2:{GRAY2};
  --jp:{JP_COLOR}; --fp:{FP_COLOR};
  --jp-l:{JP_LIGHT}; --fp-l:{FP_LIGHT};
  --bg:#F8F9FC; --surface:#FFFFFF;
  --border:#E2E6EF; --text:#1a2340; --muted:#6b7a99;
  --green:{BRAND["green"]}; --red:{BRAND["red"]}; --amber:{BRAND["amber"]};
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--text);font-family:'Segoe UI',Arial,sans-serif;font-size:13px;line-height:1.5}}

/* HEADER */
.top-bar{{height:5px;background:linear-gradient(90deg,var(--navy) 0%,var(--teal) 70%,var(--yellow) 100%)}}
header{{background:var(--navy);color:#fff;padding:12px 24px;display:flex;align-items:center;gap:16px}}
header .logo-text{{font-size:18px;font-weight:700;letter-spacing:.06em;color:#fff}}
header .logo-sub{{font-size:11px;color:rgba(255,255,255,.6);margin-top:2px}}
header .header-right{{margin-left:auto;text-align:right;font-size:11px;color:rgba(255,255,255,.7)}}
header .header-ts{{font-size:10px;color:{TEAL};margin-top:2px;font-family:monospace}}

/* NAV */
nav{{background:var(--surface);border-bottom:2px solid var(--navy);padding:0 24px;
     display:flex;gap:0;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(30,46,110,.08)}}
nav a{{color:var(--muted);text-decoration:none;font-size:12px;font-weight:600;
       padding:12px 16px;display:block;border-bottom:3px solid transparent;transition:.2s}}
nav a:hover{{color:var(--navy);border-bottom-color:var(--teal)}}

/* SECTIONS */
.section{{padding:28px 24px;max-width:1400px;margin:0 auto}}
.section-header{{display:flex;align-items:baseline;gap:12px;margin-bottom:20px;
                  border-left:4px solid var(--teal);padding-left:12px}}
.section-tag{{font-size:10px;font-weight:700;letter-spacing:.14em;color:var(--teal);text-transform:uppercase}}
.section-title{{font-size:20px;font-weight:700;color:var(--navy)}}
.section-sub{{font-size:11px;color:var(--muted);margin-left:auto}}
.sep{{border:none;border-top:1px solid var(--border);margin:0}}

/* CARDS */
.cards{{display:grid;gap:12px;margin-bottom:20px}}
.c2{{grid-template-columns:repeat(auto-fit,minmax(240px,1fr))}}
.c4{{grid-template-columns:repeat(auto-fit,minmax(180px,1fr))}}
.card{{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:16px;
       box-shadow:0 1px 4px rgba(0,0,0,.06)}}
.card-lbl{{font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted)}}
.card-val{{font-size:28px;font-weight:700;margin:4px 0;line-height:1}}
.card-sub{{font-size:11px;color:var(--muted)}}
.card-accent{{border-top:3px solid var(--teal)}}

/* SCOREBOARD */
.scoreboard{{background:var(--navy);color:#fff;border-radius:10px;padding:20px 24px;
             margin-bottom:16px;display:grid;
             grid-template-columns:1fr auto 1fr;gap:16px;align-items:center}}
.sc-side{{text-align:center}}
.sc-partido{{font-size:10px;letter-spacing:.1em;font-weight:700;text-transform:uppercase;
             opacity:.7;margin-bottom:6px}}
.sc-pct{{font-size:48px;font-weight:700;line-height:1}}
.sc-votos{{font-size:11px;opacity:.6;margin-top:4px}}
.sc-regs{{font-size:11px;opacity:.7;margin-top:2px}}
.sc-mid{{text-align:center;padding:0 8px}}
.sc-mid-lbl{{font-size:9px;letter-spacing:.1em;text-transform:uppercase;opacity:.6;margin-bottom:4px}}
.sc-mid-val{{font-size:22px;font-weight:700;color:{YELLOW}}}
.sc-mid-sub{{font-size:9px;opacity:.6;margin-top:2px}}
.bar-container{{margin:12px 0;background:rgba(255,255,255,.15);border-radius:4px;height:8px;overflow:hidden}}
.bar-jp{{height:100%;background:var(--jp);border-radius:4px;transition:width 1s ease}}

/* GRID */
.g2{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}}
.g3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;margin-bottom:16px}}
@media(max-width:900px){{.g2,.g3{{grid-template-columns:1fr}}}}

/* PANEL */
.panel{{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:16px;
        margin-bottom:16px;box-shadow:0 1px 4px rgba(0,0,0,.05)}}
.panel-title{{font-size:13px;font-weight:600;color:var(--navy);margin-bottom:4px}}
.panel-sub{{font-size:10px;color:var(--muted);margin-bottom:12px}}
canvas{{max-height:320px;width:100%!important}}

/* TABLE */
.tbl-wrap{{overflow-x:auto;border-radius:8px;border:1px solid var(--border);margin-bottom:16px}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th{{background:var(--navy);color:#fff;padding:8px 10px;text-align:left;font-size:10px;
    letter-spacing:.06em;text-transform:uppercase;white-space:nowrap}}
td{{padding:7px 10px;border-bottom:1px solid var(--border)}}
tr:hover td{{background:#f0f4ff}}
.num{{font-family:monospace;text-align:right}}
.ts-cell{{font-family:monospace;font-size:10px;color:var(--muted)}}
.jp{{color:var(--jp);font-weight:600}} .fp{{color:var(--fp);font-weight:600}}
.up{{color:var(--green);font-weight:600}} .dn{{color:var(--red)}} .neu{{color:var(--amber)}}
.bdg{{display:inline-block;padding:2px 7px;border-radius:3px;font-size:10px;font-weight:700}}
.bdg-jp{{background:rgba(34,197,94,.15);color:#16a34a}}
.bdg-fp{{background:rgba(249,115,22,.15);color:#c2410c}}

/* CONTROLS */
.ctrl{{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:12px}}
.ctrl-lbl{{font-size:11px;color:var(--muted);font-weight:600}}
select,button{{background:var(--surface);color:var(--text);border:1px solid var(--border);
               border-radius:5px;padding:5px 10px;font-size:12px;cursor:pointer;
               font-family:inherit;transition:.2s}}
select:hover,button:hover{{border-color:var(--teal);color:var(--navy)}}
button.active{{background:var(--navy);color:#fff;border-color:var(--navy)}}

/* MODELO */
.ref-box{{background:#f0f7ff;border:1px solid #c7dcf5;border-left:4px solid var(--navy);
           border-radius:6px;padding:14px;font-size:11px;line-height:1.8;color:var(--text);
           margin-top:12px}}

/* ENCUESTA CARDS */
.enc-card{{background:var(--surface);border:1px solid var(--border);border-radius:8px;
            overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.06)}}
.enc-header{{background:var(--navy);color:#fff;padding:10px 14px}}
.enc-header-name{{font-weight:700;font-size:13px}}
.enc-header-sub{{font-size:10px;opacity:.7;margin-top:2px}}
.enc-body{{padding:14px}}
.enc-veredicto{{text-align:center;font-weight:700;font-size:12px;
                color:{YELLOW};background:var(--navy);
                padding:6px;border-radius:4px;margin-bottom:12px}}
.enc-row{{display:flex;justify-content:space-between;align-items:flex-end}}
.enc-cand{{text-align:center}}
.enc-cand-lbl{{font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted)}}
.enc-cand-pct{{font-size:32px;font-weight:700;line-height:1.1}}
.enc-bar{{height:8px;background:var(--gray1);border-radius:4px;overflow:hidden;margin:10px 0}}
.enc-bar-jp{{height:100%;border-radius:4px}}

/* FOOTER */
footer{{background:var(--navy);color:rgba(255,255,255,.6);text-align:center;padding:16px;
        font-size:10px;margin-top:24px}}
footer .ft-bar{{height:3px;background:linear-gradient(90deg,var(--teal),var(--yellow));margin-bottom:12px}}
</style>
</head>
<body>
<div class="top-bar"></div>
<header>
  <div>
    <div class="logo-text">AFP INTEGRA</div>
    <div class="logo-sub">Renta Variable Local · Modelo Franco</div>
  </div>
  <div style="margin-left:20px;color:rgba(255,255,255,.8);font-size:13px;font-weight:600">
    Monitoreo Electoral — Segunda Vuelta Presidencial 2026
  </div>
  <div class="header-right">
    <div>Generado el</div>
    <div class="header-ts">{ts}</div>
  </div>
</header>
<nav>
  <a href="#s1">📊 ONPE Regional</a>
  <a href="#s2">📋 Encuestas CR</a>
  <a href="#s3">📈 Acciones BVL</a>
  <a href="#s4">🎯 Modelo Pred.</a>
  <a href="#s5">💹 Renta Fija</a>
</nav>

<!-- ═══ S1 ONPE REGIONAL ═══════════════════════════════════════ -->
<div class="section" id="s1">
  <div class="section-header">
    <div>
      <div class="section-tag">Sección 1</div>
      <div class="section-title">Resultados ONPE — Segunda Vuelta 2026</div>
    </div>
    <div class="section-sub" id="s1-ts">{f"Actualizado: {s1['timestamp']}" if s1['timestamp'] else "Sin datos ONPE"}</div>
  </div>

  <div class="scoreboard">
    <div class="sc-side">
      <div class="sc-partido" style="color:rgba(34,197,94,.8)">JP · Juntos por el Perú</div>
      <div class="sc-pct" style="color:{JP_COLOR}">{pj:.3f}%</div>
      <div class="sc-votos">{s1['total_votos_jp']:,} votos</div>
      <div class="sc-regs">{s1['n_regiones_jp']} regiones</div>
    </div>
    <div class="sc-mid">
      <div class="sc-mid-lbl">Actas contab.</div>
      <div class="sc-mid-val" id="sc-actas-pct">—%</div>
      <div class="sc-mid-sub" id="sc-actas-n">—</div>
    </div>
    <div class="sc-side">
      <div class="sc-partido" style="color:rgba(249,115,22,.8)">FP · Fuerza Popular</div>
      <div class="sc-pct" style="color:{FP_COLOR}">{pf:.3f}%</div>
      <div class="sc-votos">{s1['total_votos_fp']:,} votos</div>
      <div class="sc-regs">{s1['n_regiones_fp']} regiones</div>
    </div>
  </div>
  <div class="bar-container">
    <div class="bar-jp" style="width:{bar_w:.2f}%"></div>
  </div>

  <div class="g2">
    <div class="panel">
      <div class="panel-title">Resultado por región — % votos válidos JP vs FP</div>
      <div class="panel-sub">Ordenado de mayor a menor ventaja JP. Verde = JP lidera, Naranja = FP lidera.</div>
      <canvas id="cRegiones"></canvas>
    </div>
    <div class="panel">
      <div class="panel-title">Avance de actas contabilizadas por región</div>
      <div class="panel-sub">Verde >80% · Amarillo >40% · Rojo <40% de actas procesadas</div>
      <canvas id="cActas"></canvas>
    </div>
  </div>

  <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th>Región</th><th>Gana</th>
        <th>JP %</th><th>FP %</th><th>Margen</th>
        <th>Cont.</th><th>Total</th><th>% Actas</th><th>Hora</th>
      </tr></thead>
      <tbody>{tbl_rows if tbl_rows else '<tr><td colspan="9" style="text-align:center;padding:20px;color:#999">Ejecutar onpe_regional.py para obtener datos</td></tr>'}</tbody>
    </table>
  </div>
</div>
<hr class="sep">

<!-- ═══ S2 ENCUESTADORAS ════════════════════════════════════════ -->
<div class="section" id="s2">
  <div class="section-header">
    <div>
      <div class="section-tag">Sección 2</div>
      <div class="section-title">Conteo Rápido · Datum &amp; Ipsos · 07 Jun 2026</div>
    </div>
    <div class="section-sub">Datos de encuestadoras al cierre de jornada electoral</div>
  </div>

  <div class="cards c2" id="enc-cards"></div>

  <div class="g2">
    <div class="panel">
      <div class="panel-title">Datum — Intervalos de confianza (±1.0%)</div>
      <div class="panel-sub">Los rangos se solapan → empate estadístico. Requiere cómputo oficial ONPE para definir ganador.</div>
      <canvas id="cDatIC"></canvas>
    </div>
    <div class="panel">
      <div class="panel-title">Ipsos — JP% desagregado por zona geográfica</div>
      <div class="panel-sub">Muestra clara la fractura Lima (FP) vs Interior (JP)</div>
      <canvas id="cIpsos"></canvas>
    </div>
  </div>
  <div class="panel">
    <div class="panel-title">Datum — Evolución del voto JP a lo largo de la campaña (encuestas previas)</div>
    <div class="panel-sub">La tendencia muestra convergencia hacia 50-50 desde ventaja histórica de FP</div>
    <canvas id="cEvol" style="max-height:200px"></canvas>
  </div>
  <div class="panel">
    <div class="panel-title">Comparativo: Conteo Rápido Datum vs ONPE parcial — JP% por región</div>
    <div class="panel-sub">Permite validar la representatividad del conteo rápido vs actas oficiales procesadas</div>
    <canvas id="cDatReg"></canvas>
  </div>
</div>
<hr class="sep">

<!-- ═══ S3 ACCIONES BVL ══════════════════════════════════════════ -->
<div class="section" id="s3">
  <div class="section-header">
    <div>
      <div class="section-tag">Sección 3</div>
      <div class="section-title">Acciones BVL — Comportamiento en Ventanas Electorales 2016–2026</div>
    </div>
  </div>

  <div class="ctrl">
    <span class="ctrl-lbl">Acción:</span>
    <select id="selEmp"></select>
    <span class="ctrl-lbl" style="margin-left:12px">Elección:</span>
    <select id="selElec">
      <option value="1V_2016">1ª Vuelta 2016 — 10 abr</option>
      <option value="2V_2016">2ª Vuelta 2016 — 05 jun</option>
      <option value="1V_2021">1ª Vuelta 2021 — 11 abr</option>
      <option value="2V_2021">2ª Vuelta 2021 — 06 jun</option>
      <option value="1V_2026">1ª Vuelta 2026 — 13 abr</option>
      <option value="2V_2026" selected>2ª Vuelta 2026 — 07 jun</option>
    </select>
  </div>
  <div class="panel">
    <div class="panel-title" id="t3title">Retorno acumulado ±30 días (Día 0 = día de elección)</div>
    <div class="panel-sub">Retorno relativo al precio de apertura del período. Línea punteada = promedio del sector.</div>
    <canvas id="cStocks"></canvas>
  </div>

  <div class="panel">
    <div class="panel-title">Retorno por sector en días clave alrededor de la elección</div>
    <div class="panel-sub">Días: −5, 0, +5, +10, +20 respecto al día de votación. Promedio simple por sector.</div>
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
</div>
<hr class="sep">

<!-- ═══ S4 MODELO ════════════════════════════════════════════════ -->
<div class="section" id="s4">
  <div class="section-header">
    <div>
      <div class="section-tag">Sección 4 · Beta-Binomial Bayesiano</div>
      <div class="section-title">Modelo Estadístico — Estimación del Ganador</div>
    </div>
  </div>

  <div class="cards c4" id="mod-cards"></div>

  <div class="g2">
    <div class="panel">
      <div class="panel-title">Distribución posterior — P(θ | votos ONPE observados)</div>
      <div class="panel-sub" id="mod-eq" style="font-family:monospace;font-size:10px;color:var(--muted)"></div>
      <canvas id="cPost"></canvas>
    </div>
    <div class="panel">
      <div class="panel-title">Evolución de P(FP gana) conforme llegan regiones</div>
      <div class="panel-sub">Posterior actualizado secuencialmente por orden de timestamp</div>
      <canvas id="cEvolP"></canvas>
    </div>
  </div>
  <div class="ref-box" id="ref-mod"></div>
</div>
<hr class="sep">

<!-- ═══ S5 RENTA FIJA ════════════════════════════════════════════ -->
<div class="section" id="s5">
  <div class="section-header">
    <div>
      <div class="section-tag">Sección 5</div>
      <div class="section-title">Renta Fija Local — Mercado Secundario</div>
    </div>
    <div class="section-sub" id="rf-ult"></div>
  </div>

  <!-- Cards resumen por tipo -->
  <div class="cards c4" id="rf-cards" style="margin-bottom:18px"></div>

  <!-- Plot 1: Curva soberana PEN + VAC + BCRP -->
  <div class="panel">
    <div class="panel-title">Curva de Rendimientos — Soberanos PEN &amp; VAC + Certificados BCRP</div>
    <div class="panel-sub">
      <b>Soberanos</b> (Emisor GOB.CENTRAL / nemónico SB*) = referencia libre de riesgo crediticio local. &nbsp;
      <b>BCRP</b> (nemónico CD*) = señal de política monetaria. &nbsp;
      <b>VAC</b> = tasa real indexada a inflación. No se mezclan entre sí.
    </div>
    <canvas id="cCurvasSob" style="max-height:300px"></canvas>
  </div>

  <!-- Plot 2: Spread corporativo vs soberano PEN -->
  <div class="panel">
    <div class="panel-title">Spread Crediticio — Corporativo PEN vs Soberano PEN (prima de riesgo por año de vencimiento)</div>
    <div class="panel-sub">
      Puntos básicos adicionales que paga un corporativo vs el soberano de igual vencimiento y moneda.
      Eje izquierdo = TIR (%) · Eje derecho = Spread (pb).
    </div>
    <canvas id="cSpread" style="max-height:280px"></canvas>
  </div>

  <!-- Plot 3: Scatter TIR vs Duración por tipo -->
  <div class="panel">
    <div class="panel-title">TIR vs Duración — Universo completo por tipo de instrumento</div>
    <div class="panel-sub">
      <b style="color:#1E2E6E">■ Soberano</b> &nbsp;
      <b style="color:#00AECB">■ BCRP</b> &nbsp;
      <b style="color:#2f81f7">■ Corp IG</b> &nbsp;
      <b style="color:#e0a320">■ Corp Corto Plazo</b> &nbsp;
      <b style="color:#e84343">■ Corp Speculative</b>
    </div>
    <div class="ctrl">
      <span class="ctrl-lbl">Moneda:</span>
      <select id="selMonRF">
        <option value="TODAS">Todas</option>
        <option value="PEN" selected>PEN (Sol)</option>
        <option value="USD">USD (Dólar)</option>
        <option value="VAC">VAC (Indexado)</option>
      </select>
    </div>
    <canvas id="cScatterRF" style="max-height:400px"></canvas>
  </div>
</div>

<footer>
  <div class="ft-bar"></div>
  AFP Integra · Renta Variable Local · {ts} · Fuentes: ONPE / Datum Internacional / Ipsos Perú / Bloomberg
</footer>

<script>
// ═══════════════════ DATOS ═══════════════════════════════════════
const D = {data_js};
const S1=D.s1, S2=D.s2, S3v=D.s3v, S3s=D.s3s, S3e=D.s3e, S3sec=D.s3sec, S3sm=D.s3sm, S4=D.s4, S5=D.s5;

// Paleta AFP Integra
const NAVY='{NAVY}', TEAL='{TEAL}', YELLOW='{YELLOW}', GRAY1='{GRAY1}', GRAY2='{GRAY2}';
const CJP='{JP_COLOR}', CFP='{FP_COLOR}', CJPl='{JP_LIGHT}', CFPl='{FP_LIGHT}';
const CGREEN='{BRAND["green"]}', CRED='{BRAND["red"]}', CAMBER='{BRAND["amber"]}';

Chart.defaults.color='#6b7a99';
Chart.defaults.borderColor='#E2E6EF';
Chart.defaults.font.family="'Segoe UI',Arial,sans-serif";
Chart.defaults.font.size=11;

const fmt=(n,d=2)=>n==null?'—':Number(n).toFixed(d);
const fmtN=n=>n==null?'—':Number(n).toLocaleString('es-PE');

// ── S1 ─────────────────────────────────────────────────────────
function initS1(){{
  const r=S1;
  const ac=r.regiones.reduce((s,x)=>s+(x.actas_cont||0),0);
  const at=r.regiones.reduce((s,x)=>s+(x.actas_tot||0),0);
  if(at>0){{
    document.getElementById('sc-actas-pct').textContent=(ac/at*100).toFixed(1)+'%';
    document.getElementById('sc-actas-n').textContent=fmtN(ac)+' / '+fmtN(at);
  }}

  if(!r.regiones||!r.regiones.length) return;

  const regs=[...r.regiones].sort((a,b)=>b.pct_jp-a.pct_jp);
  new Chart(document.getElementById('cRegiones'),{{
    type:'bar',
    data:{{
      labels:regs.map(x=>x.nombre),
      datasets:[
        {{label:'JP %',data:regs.map(x=>x.pct_jp),backgroundColor:regs.map(x=>x.ganador==='JP'?CJP:'rgba(34,197,94,.35)'),borderRadius:3}},
        {{label:'FP %',data:regs.map(x=>-x.pct_fp),backgroundColor:regs.map(x=>x.ganador==='FP'?CFP:'rgba(249,115,22,.35)'),borderRadius:3}},
      ]
    }},
    options:{{
      indexAxis:'y',responsive:true,
      plugins:{{legend:{{position:'top'}},
        tooltip:{{callbacks:{{label:c=>c.datasetIndex===0?'JP: '+c.raw.toFixed(2)+'%':'FP: '+Math.abs(c.raw).toFixed(2)+'%'}}}}}},
      scales:{{
        x:{{ticks:{{callback:v=>Math.abs(v)+'%'}},grid:{{color:'rgba(0,0,0,.05)'}}}},
        y:{{ticks:{{font:{{size:10}}}}}}
      }}
    }}
  }});

  const ra=[...r.regiones].sort((a,b)=>b.pct_actas-a.pct_actas);
  new Chart(document.getElementById('cActas'),{{
    type:'bar',
    data:{{
      labels:ra.map(x=>x.nombre),
      datasets:[{{
        label:'% actas contabilizadas',
        data:ra.map(x=>x.pct_actas),
        backgroundColor:ra.map(x=>x.pct_actas>80?'rgba(46,201,122,.7)':x.pct_actas>40?'rgba(224,163,32,.7)':'rgba(232,67,67,.7)'),
        borderRadius:3,
      }}]
    }},
    options:{{
      indexAxis:'y',responsive:true,
      plugins:{{legend:{{display:false}}}},
      scales:{{
        x:{{max:100,ticks:{{callback:v=>v+'%'}},grid:{{color:'rgba(0,0,0,.05)'}}}},
        y:{{ticks:{{font:{{size:10}}}}}}
      }}
    }}
  }});
}}

// ── S2 ─────────────────────────────────────────────────────────
function initS2(){{
  const dat=S2.datum, ips=S2.ipsos;
  const ce=document.getElementById('enc-cards');

  [dat,ips].forEach(d=>{{
    const jpFill=d.jp>50?CJP:NAVY;
    const fpFill=d.fp>50?CFP:NAVY;
    ce.innerHTML+=`
    <div class="enc-card">
      <div class="enc-header">
        <div class="enc-header-name">${{d.fuente}}</div>
        <div class="enc-header-sub">${{d.muestra}} · Error ±${{d.error}}pp · 07/06/2026</div>
      </div>
      <div class="enc-body">
        <div class="enc-veredicto">${{d.veredicto}}</div>
        <div class="enc-row">
          <div class="enc-cand">
            <div class="enc-cand-lbl">JP — Juntos por el Perú</div>
            <div class="enc-cand-pct jp">${{d.jp.toFixed(2)}}%</div>
            <div style="font-size:10px;color:var(--muted)">[${{(d.jp_min||d.jp-d.error).toFixed(2)}}% — ${{(d.jp_max||d.jp+d.error).toFixed(2)}}%]</div>
          </div>
          <div style="font-size:24px;font-weight:300;color:var(--muted);padding:0 12px">vs</div>
          <div class="enc-cand" style="text-align:right">
            <div class="enc-cand-lbl">FP — Fuerza Popular</div>
            <div class="enc-cand-pct fp">${{d.fp.toFixed(2)}}%</div>
            <div style="font-size:10px;color:var(--muted)">[${{(d.fp_min||d.fp-d.error).toFixed(2)}}% — ${{(d.fp_max||d.fp+d.error).toFixed(2)}}%]</div>
          </div>
        </div>
        <div class="enc-bar">
          <div class="enc-bar-jp" style="width:${{d.jp}}%;background:linear-gradient(90deg,${{CJP}},${{CFP}})"></div>
        </div>
      </div>
    </div>`;
  }});

  // Datum IC
  new Chart(document.getElementById('cDatIC'),{{
    type:'bar',
    data:{{
      labels:['JP — Juntos por el Perú','FP — Fuerza Popular'],
      datasets:[
        {{label:'Resultado CR',data:[dat.jp,dat.fp],backgroundColor:[CJP,CFP],borderRadius:4,borderSkipped:false}},
        {{label:'IC mín',data:[dat.jp_min,dat.fp_min],backgroundColor:['rgba(34,197,94,.2)','rgba(249,115,22,.2)'],borderRadius:2}},
        {{label:'IC máx',data:[dat.jp_max,dat.fp_max],backgroundColor:['rgba(34,197,94,.15)','rgba(249,115,22,.15)'],borderRadius:2}},
      ]
    }},
    options:{{
      indexAxis:'y',responsive:true,
      plugins:{{legend:{{position:'top'}},
        annotation:{{annotations:{{l50:{{type:'line',xMin:50,xMax:50,borderColor:'rgba(0,0,0,.4)',borderWidth:1,borderDash:[4,4],
          label:{{display:true,content:'50%',position:'start',color:'#555',font:{{size:9}}}}}}}}}}}},
      scales:{{x:{{min:46,max:54,ticks:{{callback:v=>v.toFixed(1)+'%'}},grid:{{color:'rgba(0,0,0,.05)'}}}}}},
    }}
  }});

  // Ipsos desagregado
  const iZ=Object.keys(ips.desagregado);
  new Chart(document.getElementById('cIpsos'),{{
    type:'bar',
    data:{{
      labels:iZ,
      datasets:[
        {{label:'JP',data:iZ.map(z=>ips.desagregado[z].jp),backgroundColor:CJP,borderRadius:3}},
        {{label:'FP',data:iZ.map(z=>ips.desagregado[z].fp),backgroundColor:CFP,borderRadius:3}},
        {{label:'50% umbral',data:iZ.map(_=>50),type:'line',borderColor:'rgba(0,0,0,.3)',
          borderDash:[4,4],pointRadius:0,borderWidth:1}},
      ]
    }},
    options:{{
      responsive:true,plugins:{{legend:{{position:'top'}}}},
      scales:{{y:{{min:0,max:100,ticks:{{callback:v=>v+'%'}},grid:{{color:'rgba(0,0,0,.05)'}}}}}}
    }}
  }});

  // Evolución
  const ev=dat.evolucion;
  new Chart(document.getElementById('cEvol'),{{
    type:'line',
    data:{{
      labels:ev.map(e=>e.l),
      datasets:[
        {{label:'JP',data:ev.map(e=>e.jp),borderColor:CJP,backgroundColor:CJPl,fill:true,tension:.4,pointRadius:5,pointBackgroundColor:CJP}},
        {{label:'FP',data:ev.map(e=>e.fp),borderColor:CFP,backgroundColor:CFPl,fill:true,tension:.4,pointRadius:5,pointBackgroundColor:CFP}},
        {{label:'50%',data:ev.map(_=>50),borderColor:'rgba(0,0,0,.2)',borderDash:[4,4],pointRadius:0,borderWidth:1}},
      ]
    }},
    options:{{
      responsive:true,plugins:{{legend:{{position:'top'}}}},
      scales:{{y:{{min:44,max:56,ticks:{{callback:v=>v+'%'}},grid:{{color:'rgba(0,0,0,.05)'}}}},
               x:{{ticks:{{font:{{size:9}}}}}}}}
    }}
  }});

  // Datum por región vs ONPE
  const dr=dat.por_region, dk=Object.keys(dr).sort();
  const om={{}};(S1.regiones||[]).forEach(r=>om[r.nombre]=r.pct_jp);
  new Chart(document.getElementById('cDatReg'),{{
    type:'bar',
    data:{{
      labels:dk,
      datasets:[
        {{label:'JP — Datum Conteo Rápido',data:dk.map(k=>dr[k]),backgroundColor:'rgba(34,197,94,.65)',borderRadius:2}},
        {{label:'JP — ONPE parcial',data:dk.map(k=>om[k]||null),backgroundColor:'rgba(0,174,203,.65)',borderRadius:2}},
      ]
    }},
    options:{{
      responsive:true,plugins:{{legend:{{position:'top'}}}},
      scales:{{y:{{min:0,max:100,ticks:{{callback:v=>v+'%'}},grid:{{color:'rgba(0,0,0,.05)'}}}},
               x:{{ticks:{{font:{{size:9}}}}}}}}
    }}
  }});
}}

// ── S3 ─────────────────────────────────────────────────────────
let cSt=null, cSe=null;
function initS3(){{
  const s=document.getElementById('selEmp');
  (D.s3e||[]).forEach(e=>{{const o=document.createElement('option');o.value=o.textContent=e;s.appendChild(o)}});
  s.addEventListener('change',updSt);
  document.getElementById('selElec').addEventListener('change',updSt);
  document.getElementById('selElec2').addEventListener('change',updSe);
  updSt(); updSe();
}}
function updSt(){{
  const emp=document.getElementById('selEmp').value;
  const elec=document.getElementById('selElec').value;
  if(cSt)cSt.destroy();
  const vd=S3v[elec];
  if(!vd||!vd.empresas||!vd.empresas[emp]) return;
  const ed=vd.empresas[emp]; const sec=ed.sector;
  document.getElementById('t3title').textContent=emp+' ('+sec+') — Retorno ±30d · '+elec.replace('_',' ')+' ('+vd.fecha+')';
  const allE=Object.keys(vd.empresas).filter(e=>vd.empresas[e].sector===sec);
  const sRet=ed.dias.map(d=>{{
    const vs=allE.map(e=>{{const ei=vd.empresas[e];const i=ei.dias.indexOf(d);return i>=0?ei.ret[i]:null}}).filter(v=>v!=null);
    return vs.length?vs.reduce((a,b)=>a+b)/vs.length:null;
  }});
  cSt=new Chart(document.getElementById('cStocks'),{{
    type:'line',
    data:{{
      labels:ed.dias.map(d=>d===0?'Día 0':d>0?'+'+d+'d':d+'d'),
      datasets:[
        {{label:emp,data:ed.ret,borderColor:NAVY,backgroundColor:'rgba(30,46,110,.08)',fill:true,tension:.3,pointRadius:2,borderWidth:2}},
        {{label:'Sector '+sec+' (prom)',data:sRet,borderColor:TEAL,borderDash:[5,3],tension:.3,pointRadius:0,borderWidth:1.5}},
      ]
    }},
    options:{{
      responsive:true,plugins:{{legend:{{position:'top'}}}},
      scales:{{
        y:{{ticks:{{callback:v=>v.toFixed(1)+'%'}},grid:{{color:'rgba(0,0,0,.05)'}}}},
        x:{{grid:{{color:'rgba(0,0,0,.05)'}}}}
      }}
    }}
  }});
}}
function updSe(){{
  const elec=document.getElementById('selElec2').value;
  if(cSe)cSe.destroy();
  const sd=S3s[elec];
  if(!sd||!Object.keys(sd).length) return;
  const dias=[-5,0,5,10,20];
  const secCols=[NAVY,TEAL,'#2f81f7','#f0883e','#3fb950','#bc8cff'];
  const sects=Object.keys(sd);
  cSe=new Chart(document.getElementById('cSect'),{{
    type:'bar',
    data:{{
      labels:sects,
      datasets:dias.map((d,i)=>({{
        label:d===0?'Día 0 (elección)':d>0?'+'+d+'d':d+'d',
        data:sects.map(s=>sd[s]&&sd[s][d]!=null?sd[s][d]:null),
        backgroundColor:(secCols[i]||'#888')+'bb',borderColor:secCols[i]||'#888',
        borderWidth:1,borderRadius:3,
      }}))
    }},
    options:{{
      responsive:true,plugins:{{legend:{{position:'top'}}}},
      scales:{{y:{{ticks:{{callback:v=>v!=null?v.toFixed(1)+'%':'—'}},grid:{{color:'rgba(0,0,0,.05)'}}}}}}
    }}
  }});
}}

// ── S4 ─────────────────────────────────────────────────────────
function initS4(){{
  const m=S4;
  const lider=m.lider, liderColor=lider==='JP'?CJP:CFP;
  const cm=document.getElementById('mod-cards');
  [
    ['P('+lider+' gana)',m['p_'+lider.toLowerCase()].toFixed(1)+'%',liderColor,'Probabilidad posterior'],
    ['P('+(lider==='JP'?'FP':'JP')+' gana)',m['p_'+(lider==='JP'?'fp':'jp')].toFixed(1)+'%',lider==='JP'?CFP:CJP,'Probabilidad posterior'],
    ['Posterior '+lider,m.mu.toFixed(3)+'%',liderColor,'IC 95%: ['+m.ci_lo+'% — '+m.ci_hi+'%]'],
    ['Votos ONPE obs.',fmtN(m.tot),'#1E2E6E',m.pesc.toFixed(1)+'% actas procesadas'],
  ].forEach(([l,v,c,s])=>cm.innerHTML+=
    '<div class="card card-accent"><div class="card-lbl">'+l+'</div>'+
    '<div class="card-val" style="color:'+c+'">'+v+'</div>'+
    '<div class="card-sub">'+s+'</div></div>');

  document.getElementById('mod-eq').textContent=
    'θ|datos ~ Beta('+(m.a0+m.vj).toFixed(0)+', '+(m.b0+m.vf).toFixed(0)+
    ')  Prior: θ₀='+m.prior_jp+'%, κ='+m.kappa+' votos virtuales';

  // Distribución posterior
  new Chart(document.getElementById('cPost'),{{
    type:'line',
    data:{{
      labels:m.xs.map(x=>x.toFixed(2)+'%'),
      datasets:[{{
        label:'P(θ | datos ONPE)',
        data:m.pdf,
        borderColor:liderColor,
        backgroundColor:lider==='JP'?CJPl:CFPl,
        fill:true,tension:.4,pointRadius:0,borderWidth:2,
      }}]
    }},
    options:{{
      responsive:true,plugins:{{legend:{{display:false}}}},
      scales:{{x:{{ticks:{{maxTicksLimit:8,font:{{size:9}}}},grid:{{color:'rgba(0,0,0,.05)'}}}},y:{{display:false}}}}
    }}
  }});

  // Evolución probabilidad
  const ev=m.evol;
  if(ev&&ev.length){{
    new Chart(document.getElementById('cEvolP'),{{
      type:'line',
      data:{{
        labels:ev.map(e=>e.r),
        datasets:[
          {{label:'P(FP gana) %',data:ev.map(e=>e.p_fp),borderColor:CFP,backgroundColor:CFPl,fill:true,tension:.3,pointRadius:3,borderWidth:2}},
          {{label:'P(JP gana) %',data:ev.map(e=>e.p_jp),borderColor:CJP,backgroundColor:CJPl,fill:true,tension:.3,pointRadius:3,borderWidth:2}},
          {{label:'50%',data:ev.map(_=>50),borderColor:'rgba(0,0,0,.2)',borderDash:[4,4],pointRadius:0,borderWidth:1}},
        ]
      }},
      options:{{
        responsive:true,plugins:{{legend:{{position:'top'}}}},
        scales:{{
          y:{{min:0,max:100,ticks:{{callback:v=>v+'%'}},grid:{{color:'rgba(0,0,0,.05)'}}}},
          x:{{ticks:{{font:{{size:9}}}}}}
        }}
      }}
    }});
  }}

  document.getElementById('ref-mod').innerHTML=
    '<b style="color:var(--navy)">Metodología: Beta-Binomial Bayesiano</b> ' +
    '(Jackman 2005 · Linzer 2013)<br>'+
    '<b>Prior:</b> promedio Datum ('+S2.datum.jp+'%) + Ipsos ('+S2.ipsos.jp+'%) → θ₀='+m.prior_jp+'%, '+
    'concentración κ='+m.kappa+' (≈ votos virtuales de la encuesta).<br>'+
    '<b>Likelihood:</b> '+fmtN(m.vj)+' votos JP y '+fmtN(m.vf)+' FP observados en ONPE ('+m.pesc+'% actas).<br>'+
    '<b>Posterior:</b> θ='+m.mu+'% ± '+m.sd+'pp · IC 95%: ['+m.ci_lo+'%–'+m.ci_hi+'%]<br>'+
    '<b>Conclusión: P('+m.lider+' > 50%) = '+m['p_'+m.lider.toLowerCase()]+'%</b> · '+
    'Ventaja estimada: '+m.ventaja+'pp sobre umbral 50%.<br>'+
    '<i style="color:var(--muted)">Refs: Jackman S. (2005) Australian J. Political Science 40(4):499-517 · '+
    'Linzer D. (2013) JASA 108(501):124-134</i>';
}}

// ── S5 Renta Fija ──────────────────────────────────────────────
let cCurSob=null, cSpr=null, cScatRF=null;

function initS5(){{
  if(D.s5.ult) document.getElementById('rf-ult').textContent='Datos al: '+D.s5.ult;

  // Cards resumen por tipo/moneda
  const rc=document.getElementById('rf-cards');
  const tcols={{Soberano:NAVY,BCRP:TEAL,Corporativo:'#2f81f7'}};
  ['Soberano','BCRP','Corporativo'].forEach(t=>{{
    const d=(D.s5.resumen[t]||{{}})['PEN']; if(!d) return;
    rc.innerHTML+=`<div class="card card-accent">
      <div class="panel-sub">${{t}} · PEN</div>
      <div style="font-size:22px;font-weight:700;color:${{tcols[t]}}">${{d.tir_med.toFixed(3)}}%</div>
      <div class="panel-sub">TIR mediana · n=${{d.n}} · [${{d.tir_min.toFixed(2)}}%–${{d.tir_max.toFixed(2)}}%]</div>
    </div>`;
  }});
  const cu=(D.s5.resumen['Corporativo']||{{}})['USD'];
  if(cu) rc.innerHTML+=`<div class="card card-accent">
    <div class="panel-sub">Corporativo · USD</div>
    <div style="font-size:22px;font-weight:700;color:#2f81f7">${{cu.tir_med.toFixed(3)}}%</div>
    <div class="panel-sub">TIR mediana · n=${{cu.n}} · [${{cu.tir_min.toFixed(2)}}%–${{cu.tir_max.toFixed(2)}}%]</div>
  </div>`;

  // Plot 1: Curvas soberana PEN + VAC + BCRP
  const cs=D.s5.curva_sob, cv=D.s5.curva_sob_vac, cb=D.s5.curva_bcrp;
  cCurSob=new Chart(document.getElementById('cCurvasSob'),{{
    type:'line',
    data:{{datasets:[
      {{label:'Soberano PEN (nominal)',
        data:cs.years.map((y,i)=>({{x:y,y:cs.tirs[i]}})),
        borderColor:NAVY,backgroundColor:'rgba(30,46,110,.08)',
        fill:true,tension:.4,pointRadius:6,pointBackgroundColor:NAVY,borderWidth:2.5}},
      {{label:'Soberano VAC (real)',
        data:cv.years.map((y,i)=>({{x:y,y:cv.tirs[i]}})),
        borderColor:TEAL,fill:false,tension:.4,pointRadius:5,
        pointBackgroundColor:TEAL,borderWidth:2,borderDash:[6,3]}},
      {{label:'BCRP CD (política monetaria)',
        data:cb.years.map((y,i)=>({{x:y,y:cb.tirs[i]}})),
        borderColor:CAMBER,fill:false,tension:.3,pointRadius:7,
        pointBackgroundColor:CAMBER,borderWidth:2,pointStyle:'triangle'}},
    ]}},
    options:{{responsive:true,
      plugins:{{legend:{{position:'top'}},
        tooltip:{{callbacks:{{label:c=>`${{c.dataset.label}}: ${{c.parsed.y.toFixed(3)}}%`}}}}}},
      scales:{{
        x:{{type:'linear',title:{{display:true,text:'Año de Vencimiento'}},
           ticks:{{stepSize:2}},grid:{{color:'rgba(0,0,0,.05)'}}}},
        y:{{title:{{display:true,text:'TIR %'}},
           ticks:{{callback:v=>v.toFixed(2)+'%'}},grid:{{color:'rgba(0,0,0,.05)'}}}}
      }}
    }}
  }});

  // Plot 2: Spread barras + línea eje secundario
  const sd=D.s5.spread_data;
  cSpr=new Chart(document.getElementById('cSpread'),{{
    type:'bar',
    data:{{
      labels:sd.map(d=>d.year),
      datasets:[
        {{label:'TIR Corporativo PEN',data:sd.map(d=>d.corp),
          backgroundColor:'rgba(30,46,110,.65)',borderRadius:4,order:2}},
        {{label:'TIR Soberano PEN (ref)',data:sd.map(d=>d.sob),
          backgroundColor:'rgba(0,174,203,.65)',borderRadius:4,order:2}},
        {{label:'Spread (pb)',data:sd.map(d=>+(d.spread*100).toFixed(1)),
          type:'line',borderColor:CRED,backgroundColor:'transparent',
          pointRadius:5,pointBackgroundColor:CRED,borderWidth:2,yAxisID:'y2',order:1}},
      ]
    }},
    options:{{responsive:true,
      plugins:{{legend:{{position:'top'}},
        tooltip:{{callbacks:{{label:c=>c.dataset.label==='Spread (pb)'
          ?`Spread: ${{c.raw}}pb`:`${{c.dataset.label}}: ${{Number(c.raw).toFixed(3)}}%`}}}}}},
      scales:{{
        x:{{title:{{display:true,text:'Año de Vencimiento'}},grid:{{color:'rgba(0,0,0,.05)'}}}},
        y:{{title:{{display:true,text:'TIR %'}},ticks:{{callback:v=>v.toFixed(2)+'%'}},grid:{{color:'rgba(0,0,0,.05)'}}}},
        y2:{{position:'right',title:{{display:true,text:'Spread (pb)'}},
             ticks:{{callback:v=>v+'pb'}},grid:{{display:false}}}}
      }}
    }}
  }});

  // Plot 3: Scatter por tipo
  document.getElementById('selMonRF').addEventListener('change', updScatRF);
  updScatRF();
}}

function updScatRF(){{
  const mf=document.getElementById('selMonRF').value;
  let pts=(D.s5.scatter||[]);
  if(mf!=='TODAS') pts=pts.filter(p=>p.mon===mf);
  const cmap={{
    'Soberano':          NAVY,
    'BCRP':              TEAL,
    'Corporativo_IG':    '#2f81f7',
    'Corporativo_CP':    CAMBER,
    'Corporativo_Spec':  CRED,
    'Corporativo_NR':    GRAY2,
  }};
  const dsets=[
    {{k:'Soberano',       lbl:'Soberano (GOB.CENTRAL)', fn:p=>p.tipo==='Soberano'}},
    {{k:'BCRP',           lbl:'BCRP (Certificados CD)', fn:p=>p.tipo==='BCRP'}},
    {{k:'Corporativo_IG', lbl:'Corp — Inv.Grade',       fn:p=>p.tipo==='Corporativo'&&p.g==='Investment Grade'}},
    {{k:'Corporativo_CP', lbl:'Corp — Corto Plazo',     fn:p=>p.tipo==='Corporativo'&&p.g==='Corto Plazo'}},
    {{k:'Corporativo_Spec',lbl:'Corp — Speculative',    fn:p=>p.tipo==='Corporativo'&&p.g==='Speculative Grade'}},
    {{k:'Corporativo_NR', lbl:'Corp — Sin Rating',      fn:p=>p.tipo==='Corporativo'&&p.g==='Sin Rating'}},
  ].map(d=>({{
    label:d.lbl,
    data:pts.filter(d.fn).map(p=>({{x:p.dur,y:p.tir,e:p.e,n:p.n,mon:p.mon,rat:p.rat,sp:p.sp,yr:p.yr}})),
    backgroundColor:(cmap[d.k]||GRAY2)+'bb',pointRadius:5,pointHoverRadius:8,
  }}));
  if(cScatRF) cScatRF.destroy();
  cScatRF=new Chart(document.getElementById('cScatterRF'),{{
    type:'scatter',data:{{datasets:dsets}},
    options:{{responsive:true,
      plugins:{{legend:{{position:'top'}},
        tooltip:{{callbacks:{{label:c=>{{
          const d=c.raw;
          return[d.e+' ('+d.n+')',
                 'TIR: '+d.y.toFixed(3)+'% | Dur: '+d.x.toFixed(2)+'a | Venc: '+(d.yr||'—'),
                 'Spread: '+(d.sp!=null?d.sp.toFixed(3)+'%':'—')+' | '+d.mon+' | '+d.rat];
        }}}}}}}}}},
      scales:{{
        x:{{title:{{display:true,text:'Duración (años)'}},min:0,grid:{{color:'rgba(0,0,0,.05)'}}}},
        y:{{title:{{display:true,text:'TIR %'}},ticks:{{callback:v=>v.toFixed(2)+'%'}},grid:{{color:'rgba(0,0,0,.05)'}}}}
      }}
    }}
  }});
}}

// ══ INIT ══════════════════════════════════════════════════════
function waitForChart(cb){{
  if(typeof Chart!=='undefined') cb();
  else setTimeout(()=>waitForChart(cb), 100);
}}
waitForChart(()=>{{
  initS1(); initS2(); initS3(); initS4(); initS5();
}});
</script>
</body>
</html>"""

# ══════════════════════════════════════════════════════════════════
# EJECUTAR
# ══════════════════════════════════════════════════════════════════
print("  [1/5] ONPE regional...", end=" ", flush=True)
s1 = build_s1(_df)
print(f"{len(s1['regiones'])} regiones · JP={s1['pct_nac_jp']:.2f}% · FP={s1['pct_nac_fp']:.2f}%")

print("  [2/5] Encuestadoras...", end=" ", flush=True)
s2 = build_s2()
print("OK")

print("  [3/5] Stocks BVL...", end=" ", flush=True)
s3 = build_s3()
print(f"{len(s3['empresas'])} acciones · {len(s3['ventanas'])} ventanas electorales")

print("  [4/5] Modelo Bayesiano...", end=" ", flush=True)
s4 = build_s4(s1, s2)
print(f"P({s4['lider']} gana)={s4['p_'+s4['lider'].lower()]:.1f}% · θ={s4['mu']:.3f}%")

print("  [5/5] Renta Fija...", end=" ", flush=True)
s5 = build_s5()
print(f"{len(s5['scatter'])} instrumentos")

print("  Generando HTML...", end=" ", flush=True)
html = generar_html(s1, s2, s3, s4, s5)
OUT_PATH.write_text(html, encoding="utf-8")
print(f"OK — {OUT_PATH.stat().st_size//1024} KB")
print(f"  LISTO: {OUT_PATH}")
print(f"{'='*60}\n")