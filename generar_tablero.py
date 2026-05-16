#!/usr/bin/env python3
"""
generar_tablero.py — Campo & Asociados
Uso: python generar_tablero.py

Lee Reporte_22.xlsx (facturación) y Reporte_23.xlsx (pedidos de venta)
y genera index.html con los datos embebidos, listo para GitHub Pages.
"""

import json, sys, os
import pandas as pd
from datetime import datetime

# ── Configuración de objetivos ─────────────────────────────────────────
# Editá estos valores cada mes antes de correr el script
CONFIG = {
    "mes":          "2026-05",
    "mesNombre":    "Mayo 2026",
    "cc_obj":       516634025.81,
    "ej_obj":       145708697.92,
    "obj_pascual":   63721857.15,
    "obj_bautista":  58680901.80,
    "obj_manuel":    30908189.48,
    "cli_ej_obj":   126,
    "cli_cc_obj":   126,
    "ej_ap":      10907853.63,
    "ej_mar":    256487178.50,
    "cc_ap":     395952493.17,
    "cc_mar":    835027428.67,
}

FAC_FILE = "Reporte_22.xlsx"
PED_FILE = "Reporte_23.xlsx"
EJ_NAMES = ["bautista", "pascual", "manuel"]

def is_ej(v):
    return bool(v) and any(n in str(v).lower() for n in EJ_NAMES)

def compute_metrics():
    print("Leyendo archivos...")
    fac = pd.read_excel(FAC_FILE, sheet_name=0)
    ped = pd.read_excel(PED_FILE, sheet_name=0)
    fac["imp"] = pd.to_numeric(fac["Importe mon. principal"], errors="coerce").fillna(0)
    ped["imp"] = pd.to_numeric(ped["Importe mon. principal"], errors="coerce").fillna(0)
    ej_fac = fac[fac["Vendedor"].apply(is_ej)].copy()
    ej_ped = ped[ped["Vendedor"].apply(is_ej)].copy()

    ped_grp = ej_ped.groupby("Doc.-nro. interno").agg(
        total=("imp","sum"), cliente=("Cliente","first"), vendedor=("Vendedor","first")
    ).reset_index()

    def top_prod(g):
        g2 = g.dropna(subset=["Producto"])
        return "" if g2.empty else g2.loc[g2["imp"].idxmax(), "Producto"]

    prods = ej_ped.groupby("Doc.-nro. interno").apply(top_prod).reset_index()
    prods.columns = ["Doc.-nro. interno","Producto"]
    ped_grp = ped_grp.merge(prods, on="Doc.-nro. interno")
    top3 = ped_grp.nlargest(3,"total")[["cliente","total","Producto","vendedor"]].copy()
    top3["total"] = top3["total"].astype(float)
    top3_list = top3.rename(columns={"Producto":"producto"}).to_dict("records")

    return {
        "ej_fac":      float(ej_fac["imp"].sum()),
        "cc_fac":      float(fac["imp"].sum()),
        "ej_pv":       float(ej_ped["imp"].sum()),
        "cc_pv":       float(ped["imp"].sum()),
        "cli_ej_fac":  int(ej_fac["Cliente"].nunique()),
        "cli_cc_fac":  int(fac["Cliente"].nunique()),
        "cli_ej_pv":   int(ej_ped["Cliente"].nunique()),
        "cli_cc_pv":   int(ped["Cliente"].nunique()),
        "by_vend_fac": {k:float(v) for k,v in ej_fac.groupby("Vendedor")["imp"].sum().items()},
        "by_vend_pv":  {k:float(v) for k,v in ej_ped.groupby("Vendedor")["imp"].sum().items()},
        "ped_count":   {k:int(v)   for k,v in ej_ped.groupby("Vendedor")["Doc.-nro. interno"].nunique().items()},
        "fac_cli_cnt": {k:int(v)   for k,v in ej_fac.groupby("Vendedor")["Cliente"].nunique().items()},
        "top3":        top3_list,
        "generado":    datetime.now().strftime("%d/%m/%Y %H:%M"),
    }

def fmt(n):
    if n is None: return "—"
    a = abs(n); s = "-" if n < 0 else ""
    if a >= 1e9: return s + "$" + str(round(a/1e9, 2)) + "B"
    if a >= 1e6: return s + "$" + str(round(a/1e6)) + "M"
    if a >= 1e3: return s + "$" + str(round(a/1e3)) + "K"
    return s + "$" + str(round(a))

def fmt_k(n):
    if n is None: return "—"
    a = abs(n); s = "-" if n < 0 else ""
    if a >= 1e6: return s + str(round(a/1e6,1)) + "M"
    if a >= 1e3: return s + str(round(a/1e3)) + "K"
    return s + str(round(a))

def pct_w(r): return str(min(r*100, 100))[:5]

def badge(ratio):
    if not ratio: return "neutral","—"
    p = (ratio-1)*100
    cls = "up" if ratio>=1 else "down"
    sign = "+" if ratio>=1 else ""
    return cls, sign + str(round(p,1)) + "%"

def get_obj(name, cfg):
    n = name.lower()
    if "pascual" in n: return cfg["obj_pascual"]
    if "bautista" in n: return cfg["obj_bautista"]
    if "manuel" in n: return cfg["obj_manuel"]
    return 0

def make_top3(top3):
    ranks = [("r1","1°"),("r2","2°"),("r3","3°")]
    out = ""
    for i, it in enumerate(top3):
        rc, rn = ranks[i]
        fn = str(it["vendedor"]).split()[0]
        out += '<div class="top3-item">'
        out += '<div class="top3-rank ' + rc + '">' + rn + '</div>'
        out += '<div class="top3-info">'
        out += '<div class="top3-cliente">' + str(it["cliente"]) + '</div>'
        out += '<div class="top3-producto">' + str(it["producto"]) + '</div>'
        out += '</div>'
        out += '<div class="top3-right">'
        out += '<div class="top3-monto">' + fmt(it["total"]) + '</div>'
        out += '<div class="top3-vendedor">' + fn + '</div>'
        out += '</div></div>'
    return out

def make_vend_cards(bvf, bvp, pc, fcc, cfg):
    order = ["Pascual Ibañez","Bautista Loza","Manuel Uranga"]
    out = ""
    for name in order:
        fv = bvf.get(name, 0)
        pv = bvp.get(name, 0)
        obj = get_obj(name, cfg)
        falt = max(obj - fv, 0)
        pct = min(fv/obj, 1) if obj > 0 else 0
        on = fv >= obj
        nc = fcc.get(name, 0)
        np2 = pc.get(name, 0)
        tk = round(pv/np2) if np2 > 0 else 0
        acc = "on-track" if on else ""
        fc = "green" if on else ""
        rc = "green" if falt == 0 else "red"
        ft = "✓" if falt == 0 else fmt(falt)
        fl = "Cumplido" if falt == 0 else "Faltante"
        out += '<div class="vend-card">'
        out += '<div class="vend-card-accent ' + acc + '"></div>'
        out += '<div class="vend-name">' + name + '</div>'
        out += '<div class="vend-metrics">'
        out += '<div class="vend-metric"><div class="vend-metric-val bx">' + str(nc) + '</div><div class="vend-metric-lbl">Clientes</div></div>'
        out += '<div class="vend-metric"><div class="vend-metric-val">$' + fmt_k(tk) + '</div><div class="vend-metric-lbl">Ticket prom.</div></div>'
        out += '<div class="vend-metric"><div class="vend-metric-val ' + fc + '">' + str(round(pct*100)) + '%</div><div class="vend-metric-lbl">vs objetivo</div></div>'
        out += '<div class="vend-metric"><div class="vend-metric-val ' + rc + '">' + ft + '</div><div class="vend-metric-lbl">' + fl + '</div></div>'
        out += '</div>'
        out += '<div class="vend-prog"><div class="vend-prog-bar ' + acc + '" style="width:' + str(round(pct*100)) + '%"></div></div>'
        out += '</div>'
    return out

def make_pvfac(bvf, bvp, cfg):
    order = ["Pascual Ibañez","Bautista Loza","Manuel Uranga"]
    out = ""
    for name in order:
        fv = bvf.get(name, 0)
        pv = bvp.get(name, 0)
        obj = get_obj(name, cfg)
        pf = min(fv/obj, 1) if obj > 0 else 0
        pp = min(max(pv-fv, 0)/obj, 1-pf) if obj > 0 else 0
        al = (fv/obj < 0.5) if obj > 0 else True
        falt = max(obj - fv, 0)
        alh = '<span class="pvfac-warn">⚠ Menos del 50% del objetivo facturado</span>' if al else '<span></span>'
        fh = ('<span class="pvfac-faltante">Falta ' + fmt(falt) + ' para objetivo</span>'
              if falt > 0 else
              '<span style="font-size:0.65rem;font-weight:700;color:var(--green)">✓ Objetivo cumplido</span>')
        pc2 = "amber" if al else "muted"
        out += '<div class="pvfac-row">'
        out += '<div class="pvfac-header"><span class="pvfac-name">' + name + '</span><span class="pvfac-pv-total">PV: ' + fmt(pv) + '</span></div>'
        out += '<div class="pvfac-nums">'
        out += '<span class="pvfac-num green">' + str(round(pf*100)) + '% facturado (' + fmt(fv) + ')</span>'
        out += '<span class="pvfac-num muted"> · </span>'
        out += '<span class="pvfac-num ' + pc2 + '">PV pendiente: ' + fmt(max(pv-fv, 0)) + '</span>'
        out += '</div>'
        out += '<div class="pvfac-track">'
        out += '<div class="pvfac-seg-fac" style="width:' + str(round(pf*100)) + '%"></div>'
        out += '<div class="pvfac-seg-pend" style="width:' + str(round(pp*100)) + '%"></div>'
        out += '</div>'
        out += '<div class="pvfac-footer">' + alh + fh + '</div>'
        out += '</div>'
    return out

LOGO = "data:image/png;base64,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"

def main():
    for f in [FAC_FILE, PED_FILE]:
        if not os.path.exists(f):
            print("ERROR: No se encontro '" + f + "'")
            sys.exit(1)

    m = compute_metrics()
    cfg = CONFIG

    ej_r = m["ej_fac"] / cfg["ej_obj"]
    cc_r = m["cc_pv"]  / cfg["cc_obj"]
    donut_pct = round(m["ej_fac"] / m["cc_fac"] * 100, 1) if m["cc_fac"] else 0

    ej_pill_cls = "white-up" if ej_r>=1 else "white-down"
    ej_dir = "▲" if ej_r>=1 else "▼"
    ej_word = "sobre" if ej_r>=1 else "bajo"
    ej_pill_txt = ej_dir + " " + str(round(abs(ej_r-1)*100,1)) + "% " + ej_word + " objetivo"

    cc_pill_cls = "up" if cc_r>=1 else "down"
    cc_dir = "▲" if cc_r>=1 else "▼"
    cc_word = "sobre" if cc_r>=1 else "bajo"
    cc_pill_txt = cc_dir + " " + str(round(abs(cc_r-1)*100,1)) + "% " + cc_word + " objetivo"

    ej_falt = max(cfg["ej_obj"] - m["ej_fac"], 0)
    ej_falt_pill = ('<span class="pill white-amber" style="margin-left:0.35rem">' + fmt(ej_falt) + ' faltante</span>'
                    if ej_falt > 0 else "")

    ej_ap_cls,  ej_ap_txt  = badge(m["ej_pv"]/cfg["ej_ap"]   if cfg["ej_ap"]  else None)
    ej_mar_cls, ej_mar_txt = badge(m["ej_pv"]/cfg["ej_mar"]  if cfg["ej_mar"] else None)
    ej_obj_cls, ej_obj_txt = badge(m["ej_fac"]/cfg["ej_obj"] if cfg["ej_obj"] else None)

    data_js = json.dumps({"ej_fac":m["ej_fac"],"cc_fac":m["cc_fac"],"cc_pv":m["cc_pv"],"config":cfg})

    css = open(os.path.join(os.path.dirname(__file__), '_styles.css')).read() if os.path.exists('_styles.css') else get_css()
    js  = get_js(data_js)

    parts = []
    parts.append('<!DOCTYPE html>')
    parts.append('<html lang="es"><head>')
    parts.append('<meta charset="UTF-8">')
    parts.append('<meta name="viewport" content="width=device-width, initial-scale=1.0">')
    parts.append('<title>Tablero TV — Campo & Asociados — ' + cfg["mesNombre"] + '</title>')
    parts.append('<link href="https://fonts.googleapis.com/css2?family=Oswald:wght@400;600&family=DM+Sans:wght@400;500&display=swap" rel="stylesheet">')
    parts.append('<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>')
    parts.append('<style>' + css + '</style>')
    parts.append('</head><body>')

    # TOPBAR
    parts.append('<div class="topbar">')
    parts.append('  <div class="topbar-brand"><img src="' + LOGO + '" alt="Campo">')
    parts.append('    <div><div class="topbar-brand-text">campo &amp; asociados</div><div class="topbar-brand-sub">Tablero mensual</div></div>')
    parts.append('  </div>')
    parts.append('  <div class="topbar-center"><div class="topbar-title">VENTAS</div><div class="topbar-mes">' + cfg["mesNombre"] + '</div></div>')
    parts.append('  <div class="topbar-update">Generado: ' + m["generado"] + '</div>')
    parts.append('</div>')

    parts.append('<div class="dash-inner">')

    # ROW 1
    parts.append('<div class="row1">')
    # Hero EJ
    parts.append('<div class="card bx-hero">')
    parts.append('  <div class="card-label">&#9889; Equipo Joven — Facturado</div>')
    parts.append('  <div class="card-value xl white">' + fmt(m["ej_fac"]) + '</div>')
    parts.append('  <div class="card-sub">Objetivo: <strong style="color:#fff">' + fmt(cfg["ej_obj"]) + '</strong></div>')
    parts.append('  <div class="prog-wrap">')
    parts.append('    <div class="prog-header"><div class="prog-label">Avance vs objetivo</div><div class="prog-pct">' + str(round(ej_r*100,1)) + '%</div></div>')
    parts.append('    <div class="prog-track"><div class="prog-bar white-bar" style="width:' + pct_w(ej_r) + '%"></div></div>')
    parts.append('  </div>')
    parts.append('  <div style="margin-top:0.35rem"><span class="pill ' + ej_pill_cls + '">' + ej_pill_txt + '</span>' + ej_falt_pill + '</div>')
    parts.append('</div>')
    # CC card
    parts.append('<div class="card bx-soft">')
    parts.append('  <div class="card-accent" style="background:var(--bx3)"></div>')
    parts.append('  <div class="card-label">Casa Central — PV Total</div>')
    parts.append('  <div class="card-value dark">' + fmt(m["cc_pv"]) + '</div>')
    parts.append('  <div class="card-sub">Objetivo: <strong style="color:var(--text)">' + fmt(cfg["cc_obj"]) + '</strong></div>')
    parts.append('  <div class="prog-wrap">')
    parts.append('    <div class="prog-header"><div class="prog-label">Avance vs objetivo</div><div class="prog-pct" style="color:var(--bx2)">' + str(round(cc_r*100,1)) + '%</div></div>')
    parts.append('    <div class="prog-track"><div class="prog-bar bx-bar" style="width:' + pct_w(cc_r) + '%"></div></div>')
    parts.append('  </div>')
    parts.append('  <span class="pill ' + cc_pill_cls + '" style="margin-top:0.35rem">' + cc_pill_txt + '</span>')
    parts.append('</div>')
    # Donut
    parts.append('<div class="card">')
    parts.append('  <div class="card-accent" style="background:var(--bx2)"></div>')
    parts.append('  <div class="card-label">Participaci&#243;n EJ sobre Casa Central</div>')
    parts.append('  <div class="donut-wrap">')
    parts.append('    <div class="donut-canvas-wrap"><canvas id="chart-donut" width="110" height="110"></canvas>')
    parts.append('      <div class="donut-center"><div class="donut-pct">' + str(donut_pct) + '%</div><div class="donut-sub">del total</div></div>')
    parts.append('    </div>')
    parts.append('    <div class="donut-legend">')
    parts.append('      <div class="donut-legend-item"><div class="donut-dot" style="background:var(--bx2)"></div><span class="donut-name">Equipo Joven</span><div class="donut-val" style="color:var(--bx2)">' + fmt(m["ej_fac"]) + '</div></div>')
    parts.append('      <div class="donut-legend-item"><div class="donut-dot" style="background:var(--s3)"></div><span class="donut-name">Resto</span><div class="donut-val">' + fmt(m["cc_fac"]-m["ej_fac"]) + '</div></div>')
    parts.append('    </div></div>')
    parts.append('</div>')
    parts.append('</div>') # end row1

    # ROW 2
    parts.append('<div class="row2">')
    # Top3 + comparativos
    parts.append('<div class="card"><div class="card-accent" style="background:var(--bx4)"></div>')
    parts.append('<div class="top3-section-label">&#9889; Top 3 ventas — Equipo Joven</div>')
    parts.append(make_top3(m["top3"]))
    parts.append('<hr class="comp-divider">')
    parts.append('<div class="card-label" style="margin-bottom:0.3rem">Comparativos PV</div>')
    parts.append('<div class="comp-row"><span class="comp-period">Mismo mes a&#241;o anterior</span><span class="comp-val">' + fmt(cfg["ej_ap"]) + '</span><span class="comp-badge ' + ej_ap_cls + '">' + ej_ap_txt + '</span></div>')
    parts.append('<div class="comp-row"><span class="comp-period">Mes anterior</span><span class="comp-val">' + fmt(cfg["ej_mar"]) + '</span><span class="comp-badge ' + ej_mar_cls + '">' + ej_mar_txt + '</span></div>')
    parts.append('<div class="comp-row"><span class="comp-period">Objetivo del mes</span><span class="comp-val bx">' + fmt(cfg["ej_obj"]) + '</span><span class="comp-badge ' + ej_obj_cls + '">' + ej_obj_txt + '</span></div>')
    parts.append('</div>')
    # Vend cards
    parts.append('<div class="card"><div class="card-accent" style="background:var(--bx2)"></div>')
    parts.append('<div class="card-label">&#9889; Equipo Joven — individual</div>')
    parts.append('<div class="vend-cards">' + make_vend_cards(m["by_vend_fac"],m["by_vend_pv"],m["ped_count"],m["fac_cli_cnt"],cfg) + '</div>')
    parts.append('</div>')
    # Clientes
    parts.append('<div class="card"><div class="card-accent" style="background:var(--bx3)"></div>')
    parts.append('<div class="card-label">Clientes vendidos</div>')
    parts.append('<div style="font-size:0.58rem;text-transform:uppercase;letter-spacing:0.1em;color:var(--muted);margin-bottom:0.35rem">&#9889; Equipo Joven</div>')
    parts.append('<div class="cli-strip">')
    parts.append('  <div class="cli-box"><div class="cli-n bx">' + str(round(cfg["cli_ej_obj"])) + '</div><div class="cli-l">Objetivo</div></div>')
    parts.append('  <div class="cli-box"><div class="cli-n">' + str(m["cli_ej_pv"]) + '</div><div class="cli-l">Pedidos</div></div>')
    parts.append('  <div class="cli-box"><div class="cli-n green">' + str(m["cli_ej_fac"]) + '</div><div class="cli-l">Facturados</div></div>')
    parts.append('</div><div class="cli-divider"></div>')
    parts.append('<div class="cli-section-label">Casa Central total</div>')
    parts.append('<div class="cli-strip">')
    parts.append('  <div class="cli-box" style="padding:0.5rem 0.35rem"><div class="cli-n" style="font-size:1.45rem">' + str(round(cfg["cli_cc_obj"])) + '</div><div class="cli-l">Objetivo</div></div>')
    parts.append('  <div class="cli-box" style="padding:0.5rem 0.35rem"><div class="cli-n" style="font-size:1.45rem">' + str(m["cli_cc_pv"]) + '</div><div class="cli-l">Pedidos</div></div>')
    parts.append('  <div class="cli-box" style="padding:0.5rem 0.35rem"><div class="cli-n green" style="font-size:1.45rem">' + str(m["cli_cc_fac"]) + '</div><div class="cli-l">Facturados</div></div>')
    parts.append('</div></div>')
    parts.append('</div>') # end row2

    # ROW 3
    parts.append('<div class="row3">')
    parts.append('<div class="card"><div class="card-accent" style="background:var(--bx2)"></div>')
    parts.append('<div class="pvfac-section-label">&#9889; Equipo Joven — Facturado, pendiente y faltante para objetivo</div>')
    parts.append('<div class="pvfac-legend">')
    parts.append('  <div class="pvfac-legend-item"><div class="pvfac-dot" style="background:var(--green)"></div>Facturado</div>')
    parts.append('  <div class="pvfac-legend-item"><div class="pvfac-dot" style="background:rgba(154,98,0,0.45)"></div>PV pendiente</div>')
    parts.append('</div>')
    parts.append(make_pvfac(m["by_vend_fac"],m["by_vend_pv"],cfg))
    parts.append('</div>')
    parts.append('<div class="card"><div class="card-accent" style="background:var(--bx4)"></div>')
    parts.append('<div class="card-label">Equipo Joven — por per&#237;odo</div>')
    parts.append('<div class="chart-wrap"><canvas id="chart-ej"></canvas></div>')
    parts.append('<div style="border-top:1px solid var(--border);padding-top:0.6rem;margin-top:0.7rem">')
    parts.append('<div class="card-label">Casa Central — por per&#237;odo</div>')
    parts.append('<div class="chart-wrap sm"><canvas id="chart-cc"></canvas></div>')
    parts.append('</div></div>')
    parts.append('</div>') # end row3

    parts.append('</div>') # end dash-inner
    parts.append('<script>' + js + '</script>')
    parts.append('</body></html>')

    html = "\n".join(parts)

    with open("index.html","w",encoding="utf-8") as f:
        f.write(html)

    print("Generado: index.html")
    print("  EJ Facturado: " + fmt(m["ej_fac"]) + " / objetivo " + fmt(cfg["ej_obj"]) + " (" + str(round(ej_r*100,1)) + "%)")
    print("  CC PV: " + fmt(m["cc_pv"]) + " / objetivo " + fmt(cfg["cc_obj"]) + " (" + str(round(cc_r*100,1)) + "%)")
    print("  Top 3: " + str([t["cliente"] for t in m["top3"]]))
    print("")
    print("-> Subi 'index.html' a tu repositorio de GitHub y listo.")

def get_css():
    return """
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{--bg:#f5f3f0;--s1:#ffffff;--s2:#f0ecea;--s3:#e8e2de;--border:rgba(0,0,0,0.07);--border2:rgba(0,0,0,0.13);--text:#1a0a08;--muted:#8a7570;--muted2:#6b5550;--bx1:#6b0f1a;--bx2:#8b1a27;--bx3:#a52235;--bx4:#c0394a;--bx5:#d4566a;--bx-light:#fdf0f1;--bx-mid:#f5d0d5;--green:#1a7a4a;--green-bg:#e8f5ee;--red:#c0392b;--red-bg:#fdf0ee;--amber:#9a6200;--amber-bg:#fdf5e0}
body{background:var(--bg);color:var(--text);font-family:"DM Sans",sans-serif;min-height:100vh;overflow-x:hidden}
.topbar{display:flex;align-items:center;justify-content:space-between;padding:0.7rem 1.5rem;background:#fff;border-bottom:1.5px solid var(--border2)}
.topbar-brand{display:flex;align-items:center;gap:0.55rem}
.topbar-brand img{width:34px;height:34px}
.topbar-brand-text{font-family:"Oswald",sans-serif;font-size:1rem;font-weight:600;letter-spacing:0.04em;color:var(--text);line-height:1.1}
.topbar-brand-sub{font-size:0.6rem;color:var(--muted);letter-spacing:0.05em;text-transform:uppercase}
.topbar-center{display:flex;align-items:baseline;gap:0.8rem}
.topbar-title{font-family:"Oswald",sans-serif;font-size:1.2rem;font-weight:600;letter-spacing:0.1em;color:var(--muted2)}
.topbar-mes{font-family:"Oswald",sans-serif;font-size:1.1rem;letter-spacing:0.08em;color:var(--bx2)}
.topbar-update{font-size:0.62rem;color:var(--muted)}
.dash-inner{padding:0.9rem 1.5rem 1.5rem}
.row1{display:grid;grid-template-columns:2fr 1fr 1fr;gap:0.9rem;margin-bottom:0.9rem}
.row2{display:grid;grid-template-columns:1fr 1fr 1fr;gap:0.9rem;margin-bottom:0.9rem}
.row3{display:grid;grid-template-columns:1.2fr 1fr;gap:0.9rem}
.card{background:var(--s1);border:1px solid var(--border);border-radius:14px;padding:1.1rem 1.2rem;position:relative;overflow:hidden;box-shadow:0 1px 5px rgba(0,0,0,0.05)}
.card.bx-hero{background:var(--bx1);border-color:var(--bx1)}
.card.bx-soft{background:var(--bx-light);border-color:var(--bx-mid)}
.card-accent{position:absolute;top:0;left:0;right:0;height:3px;border-radius:14px 14px 0 0}
.card-label{font-size:0.6rem;text-transform:uppercase;letter-spacing:0.13em;color:var(--muted);font-weight:500;margin-bottom:0.45rem}
.card.bx-hero .card-label{color:rgba(255,255,255,0.55)}
.card-value{font-family:"Oswald",sans-serif;font-size:2.4rem;font-weight:600;letter-spacing:0.02em;line-height:1;color:var(--text)}
.card-value.xl{font-size:2.9rem}
.card-value.white{color:#fff}
.card-value.dark{color:var(--bx1)}
.card-sub{font-size:0.7rem;color:var(--muted2);margin-top:0.28rem;line-height:1.5}
.card.bx-hero .card-sub{color:rgba(255,255,255,0.65)}
.prog-wrap{margin-top:0.8rem}
.prog-header{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:0.3rem}
.prog-label{font-size:0.58rem;color:var(--muted);letter-spacing:0.08em;text-transform:uppercase}
.card.bx-hero .prog-label{color:rgba(255,255,255,0.5)}
.prog-pct{font-family:"Oswald",sans-serif;font-size:1.05rem;color:var(--bx2)}
.card.bx-hero .prog-pct{color:#fff}
.prog-track{height:9px;background:rgba(0,0,0,0.07);border-radius:4px;overflow:hidden}
.card.bx-hero .prog-track{background:rgba(255,255,255,0.2)}
.card.bx-soft .prog-track{background:var(--bx-mid)}
.prog-bar{height:100%;border-radius:4px}
.prog-bar.bx-bar{background:linear-gradient(90deg,var(--bx3),var(--bx5))}
.prog-bar.white-bar{background:rgba(255,255,255,0.9)}
.pill{display:inline-flex;align-items:center;gap:3px;font-size:0.67rem;font-weight:600;padding:0.15rem 0.5rem;border-radius:20px;margin-top:0.4rem}
.pill.up{background:var(--green-bg);color:var(--green)}
.pill.down{background:var(--red-bg);color:var(--red)}
.pill.white-up{background:rgba(255,255,255,0.2);color:#fff}
.pill.white-down{background:rgba(255,255,255,0.15);color:rgba(255,255,255,0.85)}
.pill.white-amber{background:rgba(255,255,255,0.15);color:#ffd580}
.donut-wrap{display:flex;align-items:center;gap:1.1rem;margin-top:0.35rem}
.donut-canvas-wrap{position:relative;width:110px;height:110px;flex-shrink:0}
.donut-center{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;pointer-events:none}
.donut-pct{font-family:"Oswald",sans-serif;font-size:1.5rem;font-weight:600;color:var(--bx2)}
.donut-sub{font-size:0.55rem;color:var(--muted);letter-spacing:0.08em;text-transform:uppercase}
.donut-legend{flex:1}
.donut-legend-item{display:flex;align-items:center;gap:0.4rem;margin-bottom:0.5rem}
.donut-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0}
.donut-name{font-size:0.7rem;color:var(--muted2)}
.donut-val{margin-left:auto;font-weight:600;font-size:0.75rem;color:var(--text)}
.top3-section-label{font-size:0.58rem;text-transform:uppercase;letter-spacing:0.12em;color:var(--bx2);font-weight:600;margin-bottom:0.45rem;padding-bottom:0.25rem;border-bottom:1.5px solid var(--bx-mid)}
.top3-item{display:flex;align-items:flex-start;gap:0.5rem;padding:0.38rem 0;border-bottom:1px solid var(--border)}
.top3-item:last-child{border-bottom:none}
.top3-rank{font-family:"Oswald",sans-serif;font-size:1rem;font-weight:600;min-width:18px;line-height:1.2;margin-top:0.05rem}
.top3-rank.r1{color:var(--bx2)}.top3-rank.r2{color:var(--bx3)}.top3-rank.r3{color:var(--bx4)}
.top3-info{flex:1;min-width:0}
.top3-cliente{font-size:0.75rem;font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;line-height:1.2}
.top3-producto{font-size:0.62rem;color:var(--muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:0.08rem}
.top3-right{text-align:right;flex-shrink:0}
.top3-monto{font-family:"Oswald",sans-serif;font-size:0.95rem;font-weight:600;color:var(--bx2);line-height:1.2}
.top3-vendedor{font-size:0.6rem;color:var(--muted2);margin-top:0.05rem}
.comp-divider{border:none;border-top:1.5px solid var(--border2);margin:0.5rem 0}
.comp-row{display:flex;align-items:center;justify-content:space-between;padding:0.5rem 0;border-bottom:1px solid var(--border)}
.comp-row:last-child{border-bottom:none}
.comp-period{font-size:0.7rem;color:var(--muted2)}
.comp-val{font-family:"Oswald",sans-serif;font-size:1.05rem;font-weight:500;color:var(--text)}
.comp-val.bx{color:var(--bx2)}
.comp-badge{font-size:0.62rem;font-weight:700;padding:0.12rem 0.42rem;border-radius:5px;min-width:48px;text-align:center}
.comp-badge.up{background:var(--green-bg);color:var(--green)}
.comp-badge.down{background:var(--red-bg);color:var(--red)}
.comp-badge.neutral{background:var(--s2);color:var(--muted)}
.cli-strip{display:flex;gap:0.6rem}
.cli-box{flex:1;background:var(--s2);border-radius:8px;padding:0.65rem 0.4rem;text-align:center;border:1px solid var(--border)}
.cli-n{font-family:"Oswald",sans-serif;font-size:1.8rem;font-weight:600;color:var(--text)}
.cli-n.bx{color:var(--bx2)}.cli-n.green{color:var(--green)}
.cli-l{font-size:0.58rem;text-transform:uppercase;letter-spacing:0.09em;color:var(--muted);margin-top:0.1rem}
.cli-divider{border-top:1px solid var(--border);margin:0.6rem 0}
.cli-section-label{font-size:0.58rem;text-transform:uppercase;letter-spacing:0.1em;color:var(--muted);margin-bottom:0.35rem}
.vend-cards{display:flex;flex-direction:column;gap:0.65rem}
.vend-card{background:var(--s2);border:1px solid var(--border);border-radius:10px;padding:0.7rem 0.85rem;position:relative;overflow:hidden}
.vend-card-accent{position:absolute;left:0;top:0;bottom:0;width:3px;border-radius:10px 0 0 10px;background:var(--bx3)}
.vend-card-accent.on-track{background:var(--green)}
.vend-name{font-size:0.72rem;font-weight:600;color:var(--text);margin-bottom:0.35rem;padding-left:0.5rem}
.vend-metrics{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:0.4rem;padding-left:0.5rem}
.vend-metric{text-align:center}
.vend-metric-val{font-family:"Oswald",sans-serif;font-size:1.1rem;font-weight:600;line-height:1;color:var(--text)}
.vend-metric-val.bx{color:var(--bx2)}.vend-metric-val.green{color:var(--green)}.vend-metric-val.red{color:var(--red)}
.vend-metric-lbl{font-size:0.55rem;text-transform:uppercase;letter-spacing:0.08em;color:var(--muted);margin-top:0.15rem}
.vend-prog{height:4px;background:rgba(0,0,0,0.08);border-radius:2px;margin-top:0.45rem;overflow:hidden;margin-left:0.5rem}
.vend-prog-bar{height:100%;border-radius:2px;background:var(--bx3)}
.vend-prog-bar.on-track{background:var(--green)}
.pvfac-section-label{font-size:0.6rem;text-transform:uppercase;letter-spacing:0.13em;color:var(--muted);margin-bottom:0.55rem;font-weight:500}
.pvfac-legend{display:flex;gap:1rem;margin-bottom:0.6rem}
.pvfac-legend-item{display:flex;align-items:center;gap:4px;font-size:0.6rem;color:var(--muted2)}
.pvfac-dot{width:7px;height:7px;border-radius:2px;flex-shrink:0}
.pvfac-row{margin-bottom:0.9rem}
.pvfac-header{display:flex;align-items:baseline;justify-content:space-between;margin-bottom:0.22rem}
.pvfac-name{font-size:0.8rem;font-weight:600;color:var(--text)}
.pvfac-pv-total{font-family:"Oswald",sans-serif;font-size:1rem;color:var(--bx2);font-weight:600}
.pvfac-nums{display:flex;gap:0.5rem;align-items:center;flex-wrap:wrap;margin-top:0.1rem}
.pvfac-num{font-size:0.67rem;font-weight:600}
.pvfac-num.green{color:var(--green)}.pvfac-num.amber{color:var(--amber)}.pvfac-num.muted{color:var(--muted)}
.pvfac-track{height:8px;background:var(--s3);border-radius:4px;overflow:hidden;display:flex;margin-top:0.25rem}
.pvfac-seg-fac{height:100%;background:var(--green)}
.pvfac-seg-pend{height:100%;background:rgba(154,98,0,0.3)}
.pvfac-footer{display:flex;justify-content:space-between;align-items:center;margin-top:0.3rem}
.pvfac-warn{font-size:0.6rem;color:var(--amber)}
.pvfac-faltante{font-size:0.65rem;font-weight:700;color:var(--red)}
.chart-wrap{position:relative;width:100%;height:145px;margin-top:0.65rem}
.chart-wrap.sm{height:100px}
"""

def get_js(data_js):
    return """
var DATA = """ + data_js + """;
function fmt(n){
  if(n==null||isNaN(n))return"--";
  var a=Math.abs(n),s=n<0?"-":"";
  if(a>=1e9)return s+"$"+(a/1e9).toFixed(2)+"B";
  if(a>=1e6)return s+"$"+Math.round(a/1e6)+"M";
  if(a>=1e3)return s+"$"+Math.round(a/1e3)+"K";
  return s+"$"+Math.round(a);
}
var dlPlugin={id:"dl",afterDatasetsDraw:function(chart){
  var ctx=chart.ctx;
  chart.data.datasets.forEach(function(ds,di){
    chart.getDatasetMeta(di).data.forEach(function(bar,i){
      var val=ds.data[i];if(!val||val<=0)return;
      var p=bar.getCenterPoint(),bH=bar.height||0;
      if(bH<14)return;
      ctx.save();ctx.fillStyle="#fff";ctx.font="bold 9px 'DM Sans',sans-serif";
      ctx.textAlign="center";ctx.textBaseline="middle";
      ctx.fillText(fmt(val),p.x,bar.y+bH*0.5);ctx.restore();
    });
  });
}};
var H=DATA.config;
new Chart(document.getElementById("chart-donut"),{type:"doughnut",data:{labels:["Equipo Joven","Resto"],datasets:[{data:[DATA.ej_fac,DATA.cc_fac-DATA.ej_fac],backgroundColor:["#8b1a27","#e8e2de"],borderColor:["#6b0f1a","#d4c8c4"],borderWidth:2,hoverOffset:4}]},options:{responsive:false,cutout:"72%",plugins:{legend:{display:false},tooltip:{callbacks:{label:function(ctx){return" "+fmt(ctx.raw);}}}}}});
new Chart(document.getElementById("chart-ej"),{type:"bar",plugins:[dlPlugin],data:{labels:["Año ant.","Mes ant.","Actual","Objetivo"],datasets:[{data:[H.ej_ap,H.ej_mar,DATA.ej_fac,H.ej_obj],backgroundColor:["#e8e0dc","#e8e0dc","#8b1a27","rgba(139,26,39,0.15)"],borderColor:["#d4c8c4","#d4c8c4","#6b0f1a","#8b1a27"],borderWidth:[0,0,0,2],borderRadius:5}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:function(ctx){return" "+fmt(ctx.raw);}}}},scales:{x:{ticks:{color:"#8a7570",font:{size:10}},grid:{display:false},border:{display:false}},y:{ticks:{color:"#8a7570",font:{size:9},callback:function(v){return fmt(v);}},grid:{color:"rgba(0,0,0,0.05)"},border:{display:false}}}}});
new Chart(document.getElementById("chart-cc"),{type:"bar",plugins:[dlPlugin],data:{labels:["Año ant.","Mes ant.","Actual","Objetivo"],datasets:[{data:[H.cc_ap,H.cc_mar,DATA.cc_pv,H.cc_obj],backgroundColor:["#e8e0dc","#e8e0dc","#c0394a","rgba(192,57,74,0.15)"],borderColor:["#d4c8c4","#d4c8c4","#8b1a27","#c0394a"],borderWidth:[0,0,0,2],borderRadius:5}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:function(ctx){return" "+fmt(ctx.raw);}}}},scales:{x:{ticks:{color:"#8a7570",font:{size:10}},grid:{display:false},border:{display:false}},y:{ticks:{color:"#8a7570",font:{size:9},callback:function(v){return fmt(v);}},grid:{color:"rgba(0,0,0,0.05)"},border:{display:false}}}}});
"""

if __name__ == "__main__":
    main()
