#!/usr/bin/env python3
"""
generar_dashboard_paig.py
=========================
Script de automatización para Power Automate / Task Scheduler.
Lee el archivo Excel PAIG Fase II desde SharePoint y regenera el HTML
completo con 5 pestañas:

  1. 📊 Proyecciones     — KPIs, tendencia diaria, avance por GRT
  2. 📋 Resumen Ejecutivo — Logística, últimos 3 días, tabla por GRT
  3. 🎯 Reporte Directivo — Vista directiva, entrega vs instalación
  4. 🗺️ Georeferencia    — Mapa interactivo con filtros y leyenda
  5. 📥 BD Control        — Descarga de la base de datos completa en Excel

Fuente de datos: hoja BD_CONTROL (+ CONSOLIDADO, AVANCE_DIARIO, GEOREFERENCIAS)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
INSTALACIÓN (solo una vez):
  pip install pandas openpyxl

USO:
  python generar_dashboard_paig.py
  python generar_dashboard_paig.py --excel "C:\\PAIG\\datos.xlsx" --output "C:\\PAIG\\dashboard.html"

POWER AUTOMATE — FLUJO RECOMENDADO:
  Trigger : Recurrencia (cada X horas) — o "Al modificar archivo en SharePoint"
  Paso 1  : Obtener contenido del archivo → SharePoint (.xlsx)
  Paso 2  : Crear archivo → guarda .xlsx en ruta local temporal
  Paso 3  : Ejecutar script → este archivo Python
  Paso 4  : (Opcional) Actualizar archivo en SharePoint/Teams con el .html generado
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import argparse, json, sys, os
from datetime import date, datetime, timedelta
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("❌ ERROR: ejecuta:  pip install pandas openpyxl")

# ── CONFIGURACIÓN ─────────────────────────────────────────────────────────────
DEFAULT_EXCEL  = "PAIG_FASE_II_Dashboard.xlsx"
DEFAULT_OUTPUT = "PAIG_Dashboard.html"
FECHA_LIMITE_GRT = date(2026, 6, 15)
FECHA_CIERRE     = date(2026, 6, 30)

GRT_ORDER = [
    'GRT Central', 'GRT Occidente', 'GRT Oriente', 'GRT Valle de México',
    'GRT Noroeste', 'GRT Norte', 'GRT Noreste', 'GRT Baja California',
]
GRT_SHORT = {
    'GRT Central': 'Central', 'GRT Occidente': 'Occidente',
    'GRT Oriente': 'Oriente', 'GRT Valle de México': 'Valle de México',
    'GRT Noroeste': 'Noroeste', 'GRT Norte': 'Norte',
    'GRT Noreste': 'Noreste', 'GRT Baja California': 'Baja California',
}
GRT_COLS_DIARIO = {
    'GRT Central': 'GRT Central', 'GRT Occidente': 'GRT Occidente',
    'GRT Oriente': 'GRT Oriente', 'GRT Valle de México': 'GRT V. México',
    'GRT Noroeste': 'GRT Noroeste', 'GRT Norte': 'GRT Norte',
    'GRT Noreste': 'GRT Noreste', 'GRT Baja California': 'GRT Baja Cal.',
}
GRT_ABREV = {
    'GRT Central': 'Cen', 'GRT Occidente': 'Occ', 'GRT Oriente': 'Ori',
    'GRT Valle de México': 'VMex', 'GRT Noroeste': 'Noro',
    'GRT Norte': 'Nte', 'GRT Noreste': 'Nore', 'GRT Baja California': 'BC',
}


# ──────────────────────────────────────────────────────────────────────────────
#  NORMALIZACIÓN
# ──────────────────────────────────────────────────────────────────────────────
def norm_status(s):
    s = str(s).upper().strip()
    if 'OPERANDO' in s:                    return 'Instalado - Operando'
    if 'SIN SE' in s or 'SE\u00d1AL' in s: return 'Instalado - Sin Señal'
    if 'VISITADO' in s:                    return 'Visitado'
    if 'INMUEBLE' in s:                    return 'Sin Inmueble'
    if 'CANCELADO' in s:                   return 'Cancelado'
    return 'Sin Actualizar'

def norm_ambito(s):
    s = str(s).strip().upper()
    if s in ('SI', 'S/I', 'NAN', 'NONE', ''): return 'S/I'
    if s == 'RURAL':  return 'Rural'
    if s == 'URBANO': return 'Urbano'
    return s.capitalize()

def fmt_pct(v, decimals=1):
    return f"{round(float(v)*100, decimals)}" if abs(float(v)) <= 10 else f"{round(float(v), decimals)}"


# ──────────────────────────────────────────────────────────────────────────────
#  LECTURA DE DATOS
# ──────────────────────────────────────────────────────────────────────────────
def leer_excel(excel_path: str) -> dict:
    """Lee todas las hojas necesarias y devuelve un dict con todos los datos."""
    print(f"  📂 Leyendo: {excel_path}")
    xl = pd.ExcelFile(excel_path)
    hoy = date.today()

    # ── BD_CONTROL ──────────────────────────────────────────────────────────
    bd = pd.read_excel(xl, sheet_name='BD_CONTROL', header=4)
    bd.columns = [str(c).strip() for c in bd.columns]
    bd['ESTATUS']     = bd['▼ ESTATUS'].apply(norm_status)
    bd['INSTITUCIÓN'] = bd['INSTITUCIÓN'].fillna('TELEESCUELA').astype(str).str.strip().str.upper()
    bd['ÁMBITO_N']    = bd['ÁMBITO'].apply(norm_ambito)

    total     = len(bd)
    op        = int((bd['ESTATUS'] == 'Instalado - Operando').sum())
    ss        = int((bd['ESTATUS'] == 'Instalado - Sin Señal').sum())
    vis       = int((bd['ESTATUS'] == 'Visitado').sum())
    sin_in    = int((bd['ESTATUS'] == 'Sin Inmueble').sum())
    pendiente = int((bd['ESTATUS'] == 'Sin Actualizar').sum())
    inst_total = op + ss + vis

    # ── CONSOLIDADO (para entregados por GRT) ────────────────────────────────
    con = pd.read_excel(xl, sheet_name='CONSOLIDADO', header=3)
    con.columns = [str(c).replace('\n', ' ') for c in con.columns]
    con_grt = con[~con['GERENCIA (GRT)'].astype(str).str.upper().str.contains('TOTAL', na=False)]
    con_total_row = con[con['GERENCIA (GRT)'].astype(str).str.upper().str.contains('TOTAL', na=False)]

    total_entregados = int(con_total_row['KITS LOG. ALMACÉN'].values[0]) if len(con_total_row) else 3860
    total_almacen    = int(con_total_row['EQ. EN ALMACÉN'].values[0])    if len(con_total_row) else 2571

    # Dict de entregados/almacen por GRT
    entr_by_grt = {}
    alm_by_grt  = {}
    for _, row in con_grt.iterrows():
        g = str(row['GERENCIA (GRT)']).strip()
        entr_by_grt[g] = int(row['KITS LOG. ALMACÉN']) if pd.notna(row.get('KITS LOG. ALMACÉN')) else 0
        alm_by_grt[g]  = int(row['EQ. EN ALMACÉN'])    if pd.notna(row.get('EQ. EN ALMACÉN'))    else 0

    # ── STATS POR GRT ────────────────────────────────────────────────────────
    grt_data = []
    for grt_full in GRT_ORDER:
        g = bd[bd['GERENCIA (GRT)'] == grt_full]
        g_total = len(g)
        g_op    = int((g['ESTATUS'] == 'Instalado - Operando').sum())
        g_ss    = int((g['ESTATUS'] == 'Instalado - Sin Señal').sum())
        g_vis   = int((g['ESTATUS'] == 'Visitado').sum())
        g_sin   = int((g['ESTATUS'] == 'Sin Inmueble').sum())
        g_inst  = g_op + g_ss + g_vis
        g_entr  = entr_by_grt.get(grt_full, 0)
        g_alm   = alm_by_grt.get(grt_full, 0)
        pct_inst = round(g_inst / g_total * 100, 1) if g_total > 0 else 0
        pct_entr = round(g_entr / g_total * 100, 1) if g_total > 0 else 0
        pct_ie   = round(g_inst / g_entr  * 100, 1) if g_entr  > 0 else 0
        grt_data.append({
            'grt': GRT_SHORT[grt_full], 'grt_full': grt_full,
            'total': g_total, 'op': g_op, 'ss': g_ss, 'vis': g_vis,
            'sin': g_sin, 'pend': int((g['ESTATUS'] == 'Sin Actualizar').sum()),
            'instalados': g_inst, 'entregados': g_entr, 'almacen': g_alm,
            'cancelados': 0,
            'pend_entr': g_total - g_entr,
            'pct_inst': pct_inst, 'pct_entr': pct_entr, 'pct_ie': pct_ie,
            'pct_gen': pct_entr,
        })

    # ── AVANCE DIARIO ────────────────────────────────────────────────────────
    diario = pd.read_excel(xl, sheet_name='AVANCE_DIARIO', header=4)
    diario.columns = [str(c).replace('\n', ' ') for c in diario.columns]
    diario = diario.dropna(subset=['FECHA'])
    diario['FECHA'] = pd.to_datetime(diario['FECHA'])
    diario_act = diario[diario['TOTAL DÍA'] > 0].copy()
    diario_act = diario_act[diario_act['FECHA'] >= '2026-04-22'].sort_values('FECHA')

    trend_labels = [r['FECHA'].strftime('%d/%m') for _, r in diario_act.iterrows()]
    trend_vals   = [int(r['TOTAL DÍA']) for _, r in diario_act.iterrows()]
    ma7 = []
    for i in range(len(trend_vals)):
        window = trend_vals[max(0, i-6): i+1]
        ma7.append(round(sum(window) / len(window), 1))

    # ritmo actual (últimos 7 días con actividad)
    last7_vals = trend_vals[-7:] if len(trend_vals) >= 7 else trend_vals
    ritmo_actual = round(sum(last7_vals) / len(last7_vals), 1) if last7_vals else 0

    # últimos 3 días con actividad
    last3_rows = diario_act.tail(3)
    last3_data = []
    for i, (_, row) in enumerate(last3_rows.iterrows()):
        des = []
        for grt_full, col in GRT_COLS_DIARIO.items():
            v = row.get(col, 0)
            if pd.notna(v) and int(v) > 0:
                des.append(f"{GRT_ABREV[grt_full]}:{int(v)}")
        last3_data.append({
            'fecha': row['FECHA'].strftime('%d/%m'),
            'total': int(row['TOTAL DÍA']),
            'acum':  int(row['ACUMULADO']),
            'desglose': des,
            'hoy': (i == len(last3_rows) - 1),
        })

    # ── PROYECCIONES ────────────────────────────────────────────────────────
    dias_limite = max((FECHA_LIMITE_GRT - hoy).days, 1)
    dias_cierre = max((FECHA_CIERRE     - hoy).days, 1)
    pendiente_inst = total - inst_total
    ritmo_req_15  = round(pendiente_inst / dias_limite, 1)
    ritmo_req_30  = round(pendiente_inst / dias_cierre, 1)

    dias_proy = int(pendiente_inst / ritmo_actual) if ritmo_actual > 0 else 999
    fecha_proy = (hoy + timedelta(days=dias_proy)).strftime('%d/%m/%Y')
    en_meta    = dias_proy <= dias_limite

    # promedio semanal real
    prom_semanal = round(sum(trend_vals) / max(len(trend_vals) / 7, 1), 0)

    # ── GEOREFERENCIAS ────────────────────────────────────────────────────────
    geo = pd.read_excel(xl, sheet_name='GEOREFERENCIAS', header=0)
    geo.columns = [str(c).strip() for c in geo.columns]
    merged = bd.merge(geo, on='CLAVE', how='left')
    merged = merged[merged['LATITUD'].notna()].copy()
    estados = sorted(bd['ESTADO'].dropna().unique().tolist())

    geo_records = []
    for _, r in merged.iterrows():
        geo_records.append({
            'c':  str(r['CLAVE']),
            'n':  str(r['NOMBRE DEL SITIO']),
            'e':  str(r['ESTADO']),
            'mu': str(r['MUNICIPIO']),
            'g':  str(r['GERENCIA (GRT)']),
            'i':  str(r['INSTITUCIÓN']),
            'st': str(r['ESTATUS']),
            'am': str(r['ÁMBITO_N']),
            'lat': round(float(r['LATITUD']),  6),
            'lng': round(float(r['LONGITUD']), 6),
        })

    # ── BD_CONTROL para exportar ──────────────────────────────────────────────
    bd_export = []
    for _, r in bd.iterrows():
        try:
            fecha = pd.to_datetime(r['FECHA ACTUALIZACIÓN']).strftime('%d/%m/%Y')
        except:
            fecha = ''
        bd_export.append({
            'CLAVE':              str(r['CLAVE']),
            'ID PUNTO':           str(r['ID PUNTO']) if pd.notna(r['ID PUNTO']) else '',
            'NOMBRE DEL SITIO':   str(r['NOMBRE DEL SITIO']),
            'ESTADO':             str(r['ESTADO']),
            'MUNICIPIO':          str(r['MUNICIPIO']),
            'LOCALIDAD':          str(r['LOCALIDAD'])          if pd.notna(r['LOCALIDAD'])    else '',
            'ÁMBITO':             str(r['ÁMBITO'])             if pd.notna(r['ÁMBITO'])       else '',
            'GERENCIA (GRT)':     str(r['GERENCIA (GRT)']),
            'ALMACÉN':            str(r['ALMACÉN'])            if pd.notna(r['ALMACÉN'])      else '',
            'INSTITUCIÓN':        str(r['INSTITUCIÓN']),
            'ESTATUS':            str(r['ESTATUS']),
            'FECHA ACTUALIZACIÓN': fecha,
            'COMENTARIOS':        str(r['COMENTARIOS'])        if pd.notna(r['COMENTARIOS'])  else '',
        })

    fecha_str = hoy.strftime('%d/%m/%Y')
    print(f"  ✅ Total:{total}  Op:{op}  SS:{ss}  Vis:{vis}  Pend:{pendiente}")
    print(f"  📈 Ritmo actual:{ritmo_actual}/día  Req.15Jun:{ritmo_req_15}/día  Proyección:{fecha_proy}")
    print(f"  🗺️  Geo: {len(geo_records)} sitios  BD export: {len(bd_export)} registros")

    return {
        'fecha_str':       fecha_str,
        'total':           total,
        'op':              op,
        'ss':              ss,
        'vis':             vis,
        'sin_in':          sin_in,
        'pendiente':       pendiente,
        'inst_total':      inst_total,
        'total_entregados': total_entregados,
        'total_almacen':   total_almacen,
        'pct_inst':        round(inst_total / total * 100, 1),
        'pct_entr':        round(total_entregados / total * 100, 1),
        'dias_limite':     dias_limite,
        'dias_cierre':     dias_cierre,
        'pendiente_inst':  pendiente_inst,
        'ritmo_req_15':    ritmo_req_15,
        'ritmo_req_30':    ritmo_req_30,
        'ritmo_actual':    ritmo_actual,
        'fecha_proy':      fecha_proy,
        'en_meta':         en_meta,
        'prom_semanal':    int(prom_semanal),
        'trend_labels':    trend_labels,
        'trend_vals':      trend_vals,
        'ma7':             ma7,
        'last3_data':      last3_data,
        'grt_data':        grt_data,
        'geo_records':     geo_records,
        'estados':         estados,
        'bd_export':       bd_export,
    }


# ──────────────────────────────────────────────────────────────────────────────
#  GENERADOR HTML
# ──────────────────────────────────────────────────────────────────────────────
def generar_html(d: dict, output_path: str):
    geo_js     = json.dumps(d['geo_records'], ensure_ascii=False, separators=(',', ':'))
    grt_js     = json.dumps(d['grt_data'],    ensure_ascii=False)
    last3_js   = json.dumps(d['last3_data'],  ensure_ascii=False)
    bd_js      = json.dumps(d['bd_export'],   ensure_ascii=False, separators=(',', ':'))
    trend_lbl  = json.dumps(d['trend_labels'])
    trend_v    = json.dumps(d['trend_vals'])
    ma7_v      = json.dumps(d['ma7'])
    estado_opts = '\n'.join(
        [f'          <option value="{e}">{e.title()}</option>' for e in d['estados']]
    )
    meta_badge_cls  = 'badge-meta-ok'   if d['en_meta'] else 'badge-meta-warn'
    meta_badge_txt  = '✅ En Meta — cumplimiento ~100%' if d['en_meta'] else '⚠️ Riesgo — ritmo insuficiente'
    proj_color      = '#27ae60' if d['en_meta'] else '#c0392b'
    ritmo_color     = '#27ae60' if d['ritmo_actual'] >= d['ritmo_req_15'] else '#c0392b'
    ritmo_arrow     = '↑' if d['ritmo_actual'] >= d['ritmo_req_15'] else '↓'

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>PAIG Fase II — CFE Subdirección de Transmisión</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<!-- SheetJS para exportar Excel -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{{
  --g:#00703c;--gd:#004d29;--gl:#e8f5ee;--acc:#f5a623;
  --red:#c0392b;--blue:#1a5276;--gray:#566573;
  --bg:#f0f4f3;--card:#fff;--bdr:#d5e8de;--tx:#1a2e23;--tx2:#4a6358;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',Arial,sans-serif;background:var(--bg);color:var(--tx);font-size:13px}}

/* NAV */
.nav{{background:var(--gd);display:flex;align-items:center;padding:0 16px;height:52px;gap:4px;
      position:sticky;top:0;z-index:1000;box-shadow:0 2px 8px rgba(0,0,0,.35)}}
.nav-brand{{display:flex;align-items:center;gap:10px;margin-right:16px}}
.nb-t1{{color:#fff;font-weight:800;font-size:13px;line-height:1.2}}
.nb-t2{{color:rgba(255,255,255,.6);font-size:10px}}
.nav-sep{{width:1px;height:28px;background:rgba(255,255,255,.2);margin:0 6px}}
.tab-btn{{background:transparent;border:none;color:rgba(255,255,255,.7);padding:7px 13px;
          border-radius:6px;cursor:pointer;font-size:12px;font-weight:500;transition:all .2s;white-space:nowrap}}
.tab-btn:hover{{background:rgba(255,255,255,.12);color:#fff}}
.tab-btn.active{{background:var(--g);color:#fff;font-weight:700}}
.nbadge{{background:var(--acc);color:#1a2e23;font-size:10px;padding:1px 5px;border-radius:8px;margin-left:3px;font-weight:700}}

/* PAGES */
.page{{display:none;padding:16px;min-height:calc(100vh - 52px)}}
.page.active{{display:block}}

/* PAGE HEADER */
.ph{{background:linear-gradient(135deg,var(--gd),var(--g));color:#fff;padding:13px 20px;
     border-radius:10px;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center}}
.ph h1{{font-size:16px;font-weight:700}}
.ph .sub{{font-size:11px;opacity:.82;margin-top:2px}}
.ph .dbadge{{background:rgba(255,255,255,.2);padding:5px 12px;border-radius:20px;font-size:11px;
             white-space:nowrap;text-align:right;line-height:1.7}}

/* CARDS */
.cards{{display:grid;gap:12px;margin-bottom:14px}}
.c3{{grid-template-columns:repeat(3,1fr)}}
.c4{{grid-template-columns:repeat(4,1fr)}}
.c5{{grid-template-columns:repeat(5,1fr)}}
.c6{{grid-template-columns:repeat(6,1fr)}}
.card{{background:var(--card);border-radius:10px;padding:13px 15px;border:1px solid var(--bdr);
       box-shadow:0 1px 4px rgba(0,0,0,.06)}}
.clabel{{font-size:10px;color:var(--tx2);text-transform:uppercase;letter-spacing:.5px;font-weight:600;margin-bottom:3px}}
.cval{{font-size:24px;font-weight:800;color:var(--gd);line-height:1}}
.csub{{font-size:11px;color:var(--tx2);margin-top:3px}}
.cgreen .cval{{color:var(--g)}} .camber .cval{{color:#c87f00}} .cblue .cval{{color:var(--blue)}} .cred .cval{{color:var(--red)}}
.pbar{{background:#e8f0eb;border-radius:4px;height:7px;overflow:hidden;margin-top:6px}}
.pbar-fill{{height:100%;border-radius:4px;transition:width .6s}}
.pg{{background:var(--g)}} .pa{{background:var(--acc)}} .pb{{background:var(--blue)}} .pr{{background:var(--red)}}

/* SECTIONS */
.sec{{background:var(--card);border-radius:10px;border:1px solid var(--bdr);overflow:hidden;
      margin-bottom:14px;box-shadow:0 1px 4px rgba(0,0,0,.06)}}
.sec-h{{background:var(--gd);color:#fff;padding:10px 16px;font-weight:700;font-size:12px;
        display:flex;justify-content:space-between;align-items:center}}
.sec-b{{padding:13px 16px}}

/* TABLE */
table{{width:100%;border-collapse:collapse;font-size:12px}}
th{{background:var(--gl);color:var(--gd);padding:8px 10px;text-align:center;font-weight:700;
    font-size:11px;border-bottom:2px solid var(--bdr)}}
td{{padding:7px 10px;text-align:center;border-bottom:1px solid #eef2f0;color:var(--tx)}}
tr:hover td{{background:#f5faf7}}
td.ln{{text-align:left;font-weight:600;color:var(--gd)}}
.pb2{{display:inline-block;padding:2px 8px;border-radius:10px;font-weight:700;font-size:11px}}
.hi{{background:#d5f0e3;color:#1a6640}} .mi{{background:#fff3cd;color:#856404}} .lo{{background:#fde8e8;color:#8b1a1a}}

/* LAYOUT */
.g2{{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}}
.g3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-bottom:14px}}
.chbox{{height:220px;position:relative}}
.chbox-lg{{height:280px;position:relative}}

/* ALERTS */
.alert{{padding:10px 14px;border-radius:8px;margin-bottom:12px;font-size:12px;display:flex;align-items:center;gap:8px}}
.aw{{background:#fff8e1;border-left:4px solid var(--acc);color:#7d5a00}}
.ao{{background:#e8f8ef;border-left:4px solid var(--g);color:#1a6640}}
.ar{{background:#fdecea;border-left:4px solid var(--red);color:#7b1f1f}}

/* PILLS */
.pill{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700}}
.pop{{background:#d5f0e3;color:#1a6640}} .pss{{background:#fff3cd;color:#856404}}

/* PROJ GRID */
.pgrid{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:14px}}
.pcard{{background:var(--card);border-radius:10px;border:1px solid var(--bdr);padding:13px;text-align:center}}
.pcard .big{{font-size:30px;font-weight:900;line-height:1}}
.pcard .lbl{{font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--tx2);margin-bottom:5px;font-weight:600}}
.pcard .note{{font-size:11px;color:var(--tx2);margin-top:5px}}

/* TREND PANEL */
.trend-layout{{display:grid;grid-template-columns:1fr 265px;gap:0}}
.trend-chart{{padding:16px}}
.trend-stat{{background:linear-gradient(160deg,var(--gd) 0%,#006635 100%);color:#fff;
             padding:20px 18px;border-left:1px solid rgba(255,255,255,.1)}}
.trend-stat .proj-dt{{font-size:34px;font-weight:900;letter-spacing:-1px;margin-bottom:3px}}
.trend-stat .proj-sub{{font-size:11px;color:rgba(255,255,255,.65);margin-bottom:14px}}
.badge-meta-ok  {{background:var(--g);border:1.5px solid #4caf85;border-radius:8px;padding:6px 12px;font-size:12px;font-weight:700;display:flex;align-items:center;gap:6px;margin-bottom:16px}}
.badge-meta-warn{{background:#8b4500;border:1.5px solid var(--acc);border-radius:8px;padding:6px 12px;font-size:12px;font-weight:700;display:flex;align-items:center;gap:6px;margin-bottom:16px}}
.trow{{display:flex;justify-content:space-between;align-items:center;padding:6px 0;
       border-bottom:1px solid rgba(255,255,255,.08);font-size:12px}}
.trow:last-child{{border-bottom:none}}
.tlbl{{color:rgba(255,255,255,.7)}}
.tval{{font-weight:700;color:#fff}}
.tval.amber{{color:var(--acc)}} .tval.green{{color:#4cff8a}}

/* LAST 3 DAYS */
.l3grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}}
.l3card{{background:var(--card);border:1.5px solid var(--bdr);border-radius:10px;padding:14px 16px;text-align:center}}
.l3card.today{{border-color:var(--g);background:linear-gradient(135deg,#f0fbf5,#fff)}}
.l3fecha{{font-size:11px;font-weight:700;color:var(--tx2);margin-bottom:4px}}
.l3fecha.today{{color:var(--g)}}
.l3num{{font-size:42px;font-weight:900;color:var(--gd);line-height:1}}
.l3lbl{{font-size:11px;color:var(--tx2);margin-bottom:5px}}
.l3acum{{font-size:12px;font-weight:700;color:var(--blue);margin-bottom:7px}}
.l3des{{font-size:10px;color:var(--tx2);line-height:1.8}}

/* MAP PAGE */
.map-page{{display:none;padding:0;height:calc(100vh - 52px);flex-direction:column}}
.map-page.active{{display:flex}}
.map-topbar{{background:var(--gd);color:#fff;padding:8px 16px;display:flex;align-items:center;
             gap:10px;flex-shrink:0;font-size:11px}}
.map-layout{{display:flex;flex:1;overflow:hidden}}
.map-sidebar{{width:268px;flex-shrink:0;background:#fff;box-shadow:2px 0 12px rgba(0,0,0,.12);
              display:flex;flex-direction:column;overflow:hidden;z-index:100}}
.msb-header{{background:var(--gd);color:#fff;padding:10px 15px;font-size:12px;font-weight:700;flex-shrink:0}}
.msb-body{{padding:13px;flex:1;overflow-y:auto}}
.msb-body::-webkit-scrollbar{{width:4px}}
.msb-body::-webkit-scrollbar-thumb{{background:#b0cfc0;border-radius:4px}}
.fg{{margin-bottom:11px}}
.fl{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--tx2);margin-bottom:4px}}
.fi,.fs{{width:100%;border:1.5px solid var(--bdr);border-radius:7px;padding:6px 9px;
         font-size:12px;color:var(--tx);background:#f8faf9;outline:none;transition:border-color .2s}}
.fi::placeholder{{color:#aac2b5}}
.fi:focus,.fs:focus{{border-color:var(--g);background:#fff}}
.btn-clear{{width:100%;background:var(--g);color:#fff;border:none;border-radius:7px;padding:8px;
            font-size:12px;font-weight:700;cursor:pointer;margin-top:4px;transition:background .2s}}
.btn-clear:hover{{background:#005a30}}
.leg-sec{{margin-top:14px;border-top:1px solid #e8f0eb;padding-top:11px}}
.leg-title{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--tx2);margin-bottom:7px}}
.leg-item{{display:flex;align-items:center;gap:8px;padding:4px 6px;border-radius:6px;cursor:pointer;
           font-size:12px;color:var(--tx);transition:background .15s;margin-bottom:2px;user-select:none}}
.leg-item:hover{{background:#f0f9f4}}
.leg-item.inactive{{opacity:.3;text-decoration:line-through}}
.leg-dot{{width:12px;height:12px;border-radius:50%;border:2px solid rgba(0,0,0,.15);flex-shrink:0}}
.stat-sec{{margin-top:13px;border-top:1px solid #e8f0eb;padding-top:11px}}
.stat-row{{display:flex;justify-content:space-between;padding:3px 0;font-size:12px}}
.stat-lbl{{color:var(--tx2)}} .sv{{font-weight:700;color:var(--tx)}}
.sv.op{{color:#27ae60}} .sv.ss{{color:#2980b9}} .sv.vis{{color:#f39c12}}
.sv.sin{{color:#c0392b}} .sv.pend{{color:#7f8c8d}}
#map{{flex:1}}

/* BD CONTROL PAGE */
.bd-page .search-row{{display:flex;gap:10px;align-items:center;margin-bottom:14px;flex-wrap:wrap}}
.bd-search{{flex:1;min-width:200px;border:1.5px solid var(--bdr);border-radius:8px;padding:8px 12px;
            font-size:12px;outline:none;transition:border-color .2s}}
.bd-search:focus{{border-color:var(--g)}}
.bd-filter-row{{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px;align-items:center}}
.bd-sel{{border:1.5px solid var(--bdr);border-radius:7px;padding:6px 10px;font-size:12px;
         background:#fff;cursor:pointer;outline:none}}
.bd-sel:focus{{border-color:var(--g)}}
.btn-dl{{background:var(--g);color:#fff;border:none;border-radius:8px;padding:8px 16px;
         font-size:12px;font-weight:700;cursor:pointer;display:flex;align-items:center;gap:6px;
         transition:background .2s;white-space:nowrap}}
.btn-dl:hover{{background:#005a30}}
.btn-dl-sec{{background:var(--blue)}}
.btn-dl-sec:hover{{background:#154360}}
.bd-info{{font-size:11px;color:var(--tx2);padding:6px 12px;background:var(--gl);border-radius:8px;
          border:1px solid var(--bdr);margin-bottom:12px;display:flex;align-items:center;gap:6px}}
.bd-table-wrap{{overflow:auto;max-height:calc(100vh - 340px);border:1px solid var(--bdr);border-radius:8px}}
.bd-table-wrap table th{{position:sticky;top:0;z-index:5;font-size:10px;padding:7px 8px}}
.bd-table-wrap table td{{font-size:11px;padding:6px 8px;white-space:nowrap}}
.bd-count{{font-size:11px;color:var(--tx2);margin-bottom:8px;text-align:right}}
.st-op {{color:#27ae60;font-weight:700}} .st-ss{{color:#2980b9;font-weight:700}}
.st-vis{{color:#f39c12;font-weight:700}} .st-sin{{color:#c0392b;font-weight:700}}
.st-pend{{color:#95a5a6}} .st-canc{{color:#8e44ad;font-weight:700}}

.footer{{text-align:center;padding:10px;color:var(--tx2);font-size:10px;margin-top:8px}}

@media(max-width:960px){{
  .c5,.c6{{grid-template-columns:repeat(3,1fr)}}
  .c4{{grid-template-columns:repeat(2,1fr)}}
  .g2,.g3,.l3grid,.pgrid{{grid-template-columns:1fr}}
  .trend-layout{{grid-template-columns:1fr}}
}}
</style>
</head>
<body>

<!-- ░░ NAVBAR ░░ -->
<nav class="nav">
  <div class="nav-brand">
    <svg width="34" height="34" viewBox="0 0 34 34" xmlns="http://www.w3.org/2000/svg">
      <rect width="34" height="34" rx="7" fill="#00703c"/>
      <polygon points="21,3 10,18 17,18 13,31 25,14 18,14" fill="#f5a623"/>
    </svg>
    <div>
      <div class="nb-t1">CFE · SUBDIRECCIÓN DE TRANSMISIÓN</div>
      <div class="nb-t2">PAIG FASE II · {d['total']:,} SITIOS</div>
    </div>
  </div>
  <div class="nav-sep"></div>
  <button class="tab-btn active" onclick="showTab('dash',this)">📊 Proyecciones</button>
  <button class="tab-btn" onclick="showTab('resumen',this)">📋 Resumen Ejecutivo</button>
  <button class="tab-btn" onclick="showTab('directivo',this)">🎯 Reporte Directivo</button>
  <button class="tab-btn" onclick="showTab('geo',this);initMap()">🗺️ Georeferencia<span class="nbadge">{d['total']:,}</span></button>
  <button class="tab-btn" onclick="showTab('bd',this);initBD()">📥 BD Control</button>
</nav>

<!-- ╔══════════════════════════════════════════╗
     ║  TAB 1 — PROYECCIONES                   ║
     ╚══════════════════════════════════════════╝ -->
<div id="tab-dash" class="page active">
  <div class="ph">
    <div>
      <h1>📊 Tablero Ejecutivo — PAIG Fase II</h1>
      <div class="sub">CFE Subdirección de Transmisión · Instalación de Internet en {d['total']:,} Sitios</div>
    </div>
    <div class="dbadge">📅 {d['fecha_str']}<br><span style="font-size:10px;opacity:.8">Límite GRT: 15/Jun · Cierre: 30/Jun/2026</span></div>
  </div>

  <div class="cards c5">
    <div class="card"><div class="clabel">Meta Total</div><div class="cval">{d['total']:,}</div><div class="csub">sitios a instalar</div></div>
    <div class="card cgreen"><div class="clabel">Instalados Operando</div><div class="cval">{d['op']:,}</div>
      <div class="csub">{round(d['op']/d['total']*100,1)}% del total</div>
      <div class="pbar"><div class="pbar-fill pg" style="width:{round(d['op']/d['total']*100,1)}%"></div></div></div>
    <div class="card cgreen"><div class="clabel">Instalados (Op+SS+Vis)</div><div class="cval">{d['inst_total']:,}</div>
      <div class="csub" style="font-weight:700">{d['pct_inst']}% avance</div>
      <div class="pbar"><div class="pbar-fill pg" style="width:{d['pct_inst']}%"></div></div></div>
    <div class="card camber"><div class="clabel">Equipo en Almacén</div><div class="cval">{d['total_almacen']:,}</div>
      <div class="csub">pendiente instalación</div></div>
    <div class="card cblue"><div class="clabel">Eq. Entregado</div><div class="cval">{d['total_entregados']:,}</div>
      <div class="csub">{d['pct_entr']}% de {d['total']:,}</div></div>
  </div>

  <div class="pgrid">
    <div class="pcard"><div class="lbl">Días al Límite GRT</div><div class="big" style="color:var(--red)">{d['dias_limite']}</div><div class="note">Hasta 15/Jun/2026</div></div>
    <div class="pcard"><div class="lbl">Días al Cierre</div><div class="big" style="color:#c87f00">{d['dias_cierre']}</div><div class="note">Hasta 30/Jun/2026</div></div>
    <div class="pcard"><div class="lbl">Pendiente Instalar</div><div class="big" style="color:var(--red)">{d['pendiente_inst']:,}</div><div class="note">sitios restantes</div></div>
    <div class="pcard"><div class="lbl">Ritmo Req. · 15/Jun</div><div class="big" style="color:var(--acc)">{d['ritmo_req_15']}</div><div class="note">instalaciones/día</div></div>
    <div class="pcard"><div class="lbl">Ritmo Actual (7d)</div><div class="big" style="color:{ritmo_color}">{d['ritmo_actual']}</div>
      <div class="note" style="color:{ritmo_color};font-weight:700">{'✅ Sobre la meta' if d['en_meta'] else '⚠️ Bajo la meta'} {ritmo_arrow}</div></div>
    <div class="pcard" style="background:{'linear-gradient(135deg,#f0fbf5,#e8f8ef);border:2px solid var(--g)' if d['en_meta'] else 'linear-gradient(135deg,#fff8e1,#fff3e0);border:2px solid var(--acc)'}">
      <div class="lbl">Proyección de Término</div>
      <div class="big" style="font-size:21px;color:{proj_color}">{d['fecha_proy']}</div>
      <div class="note" style="color:{proj_color};font-weight:700">{'✅ Dentro del límite' if d['en_meta'] else '⚠️ Fuera del límite'}</div></div>
  </div>

  <!-- TENDENCIA -->
  <div class="sec">
    <div class="sec-h">📈 Tendencia Instalaciones Diarias <span style="font-weight:400;opacity:.8;font-size:11px;margin-left:8px">· verde ≥ ritmo requerido 15/Jun</span></div>
    <div class="trend-layout">
      <div class="trend-chart"><div class="chbox-lg"><canvas id="chartTend"></canvas></div></div>
      <div class="trend-stat">
        <div class="proj-dt" style="color:{proj_color}">{d['fecha_proy']}</div>
        <div class="proj-sub">Ritmo actual: {d['ritmo_actual']} inst/día</div>
        <div class="{meta_badge_cls}">{meta_badge_txt}</div>
        <div class="trow"><span class="tlbl">Inicio instalaciones</span><span class="tval">17 Mar 2026</span></div>
        <div class="trow"><span class="tlbl">Días transcurridos</span><span class="tval">{56} días</span></div>
        <div class="trow"><span class="tlbl">Días restantes (15 Jun)</span><span class="tval amber">{d['dias_limite']} días</span></div>
        <div class="trow"><span class="tlbl">Sitios pendientes</span><span class="tval">{d['pendiente_inst']:,}</span></div>
        <div class="trow"><span class="tlbl">Ritmo requerido</span><span class="tval">{d['ritmo_req_15']}/día</span></div>
        <div class="trow"><span class="tlbl">Ritmo actual (7d)</span><span class="tval green">{d['ritmo_actual']}/día {ritmo_arrow}</span></div>
        <div class="trow"><span class="tlbl">% avance acumulado</span><span class="tval">{d['pct_inst']}%</span></div>
      </div>
    </div>
  </div>

  <!-- AVANCE POR GRT -->
  <div class="sec">
    <div class="sec-h">🏢 Avance por Gerencia Regional de Transmisión</div>
    <div class="sec-b">
      <table>
        <thead><tr>
          <th>Gerencia (GRT)</th><th>Total</th><th>Inst. Operando</th><th>Sin Señal</th>
          <th>En Almacén</th><th>% Instalado</th><th>% Entregado</th>
          <th>Ritmo Req./día</th><th>Barra</th>
        </tr></thead>
        <tbody id="grt-tb"></tbody>
        <tfoot><tr style="background:var(--gl);font-weight:700">
          <td class="ln">TOTAL GENERAL</td>
          <td>{d['total']:,}</td><td>{d['op']:,}</td><td>{d['ss']}</td><td>{d['total_almacen']:,}</td>
          <td><span class="pb2 lo">{d['pct_inst']}%</span></td>
          <td><span class="pb2 mi">{d['pct_entr']}%</span></td>
          <td style="color:var(--acc);font-weight:700">{d['ritmo_req_15']}</td><td></td>
        </tr></tfoot>
      </table>
    </div>
  </div>

  <div class="g2">
    <div class="sec" style="margin-bottom:0"><div class="sec-h">🥧 Distribución de Estatus</div>
      <div class="sec-b"><div class="chbox"><canvas id="chartDonut"></canvas></div></div></div>
    <div class="sec" style="margin-bottom:0"><div class="sec-h">🎯 Ritmo Necesario por GRT (inst/día · 15/Jun)</div>
      <div class="sec-b"><div class="chbox"><canvas id="chartRitmo"></canvas></div></div></div>
  </div>
  <div class="footer">CFE Subdirección de Transmisión · PAIG Fase II · Límite GRT: 15/Jun/2026 · Cierre: 30/Jun/2026 · {d['fecha_str']}</div>
</div>

<!-- ╔══════════════════════════════════════════╗
     ║  TAB 2 — RESUMEN EJECUTIVO              ║
     ╚══════════════════════════════════════════╝ -->
<div id="tab-resumen" class="page">
  <div class="ph">
    <div><h1>📋 Resumen Ejecutivo — PAIG Fase II</h1>
      <div class="sub">CFE Subdirección de Transmisión · Límite GRT: 15/Jun · Cierre: 30/Jun/2026</div></div>
    <div class="dbadge">📅 {d['fecha_str']}</div>
  </div>
  <div class="cards c4">
    <div class="card cblue"><div class="clabel">Equipos Asignados</div><div class="cval">{d['total']:,}</div><div class="csub">100% — meta total</div></div>
    <div class="card cgreen"><div class="clabel">Equipos Entregados</div><div class="cval">{d['total_entregados']:,}</div>
      <div class="csub">{d['pct_entr']}% del asignado</div>
      <div class="pbar"><div class="pbar-fill pg" style="width:{d['pct_entr']}%"></div></div></div>
    <div class="card camber"><div class="clabel">Instalados (Op+SS+Vis)</div><div class="cval">{d['inst_total']:,}</div>
      <div class="csub">{d['pct_inst']}% del asignado</div>
      <div class="pbar"><div class="pbar-fill pa" style="width:{d['pct_inst']}%"></div></div></div>
    <div class="card"><div class="clabel">Eq. en Almacén</div><div class="cval">{d['total_almacen']:,}</div>
      <div class="csub">{round(d['total_almacen']/d['total_entregados']*100,1)}% del entregado</div></div>
  </div>
  <div class="sec"><div class="sec-h">📅 Instalaciones — Últimos 3 Días con Actividad</div>
    <div class="sec-b"><div class="l3grid" id="l3-resumen"></div></div></div>
  <div class="sec"><div class="sec-h">📊 Avance por Gerencia Regional de Transmisión</div>
    <div class="sec-b">
      <table>
        <thead><tr><th>Gerencia</th><th>Total</th><th>Cancelados</th><th>Instalados</th>
          <th>% Inst./Asig.</th><th>Entregados</th><th>% Entr./Asig.</th><th>En Almacén</th><th>% Alm./Entr.</th></tr></thead>
        <tbody id="res-tb"></tbody>
      </table>
    </div>
  </div>
  <div class="g2">
    <div class="sec" style="margin-bottom:0"><div class="sec-h">📊 Instalados vs Entregados por GRT</div>
      <div class="sec-b"><div class="chbox"><canvas id="chartResBar"></canvas></div></div></div>
    <div class="sec" style="margin-bottom:0"><div class="sec-h">📈 % de Instalación por GRT</div>
      <div class="sec-b"><div class="chbox"><canvas id="chartResPct"></canvas></div></div></div>
  </div>
  <div class="footer">CFE Subdirección de Transmisión · PAIG Fase II · Resumen Ejecutivo · {d['fecha_str']}</div>
</div>

<!-- ╔══════════════════════════════════════════╗
     ║  TAB 3 — REPORTE DIRECTIVO              ║
     ╚══════════════════════════════════════════╝ -->
<div id="tab-directivo" class="page">
  <div class="ph">
    <div><h1>🎯 Reporte Directivo — PAIG Fase II</h1>
      <div class="sub">CFE Subdirección de Transmisión · {d['total']:,} Sitios</div></div>
    <div class="dbadge">📅 {d['fecha_str']}<br><span style="font-size:10px;opacity:.8">Límite GRT: 15/Jun · Cierre: 30/Jun</span></div>
  </div>
  <div class="cards c3">
    <div class="card cblue"><div class="clabel">Equipos Asignados</div><div class="cval">{d['total']:,}</div><div class="csub">Meta total</div></div>
    <div class="card cgreen"><div class="clabel">Equipos Entregados</div><div class="cval">{d['total_entregados']:,}</div>
      <div class="csub">Pendiente: <strong>{d['total']-d['total_entregados']:,}</strong></div></div>
    <div class="card"><div class="clabel">% Equipos Entregados</div><div class="cval">{d['pct_entr']}%</div><div class="csub">del total asignado</div></div>
  </div>
  <div class="cards c3">
    <div class="card camber"><div class="clabel">Instalados</div><div class="cval">{d['inst_total']:,}</div><div class="csub">Op + Sin Señal + Vis</div></div>
    <div class="card"><div class="clabel">% Inst. / Asignados</div><div class="cval">{d['pct_inst']}%</div>
      <div class="pbar"><div class="pbar-fill pa" style="width:{d['pct_inst']}%"></div></div></div>
    <div class="card"><div class="clabel">% Inst. / Entregados</div>
      <div class="cval">{round(d['inst_total']/d['total_entregados']*100,1)}%</div>
      <div class="pbar"><div class="pbar-fill pg" style="width:{round(d['inst_total']/d['total_entregados']*100,1)}%"></div></div></div>
  </div>
  <div class="sec"><div class="sec-h">📅 Instalaciones — Últimos 3 Días con Actividad</div>
    <div class="sec-b"><div class="l3grid" id="l3-dir"></div></div></div>
  <div class="sec"><div class="sec-h">📋 Desglose por Gerencia Regional de Transmisión</div>
    <div class="sec-b">
      <table>
        <thead><tr><th>Gerencia Regional</th><th>Asignados</th><th>Entregados</th>
          <th>Pend. Entrega</th><th>% Eq. Entregado</th><th>Instalados</th>
          <th>% Inst./Asig.</th><th>% Inst./Entr.</th></tr></thead>
        <tbody id="dir-tb"></tbody>
      </table>
    </div>
  </div>
  <div class="g2">
    <div class="sec" style="margin-bottom:0"><div class="sec-h">📦 Cumplimiento Entrega de Equipo (%)</div>
      <div class="sec-b"><div class="chbox"><canvas id="chartDirEntr"></canvas></div></div></div>
    <div class="sec" style="margin-bottom:0"><div class="sec-h">🔧 % Instalación / Asignados por GRT</div>
      <div class="sec-b"><div class="chbox"><canvas id="chartDirInst"></canvas></div></div></div>
  </div>
  <div class="alert aw" style="margin-top:12px">📌 <strong>CFE Subdirección de Transmisión · PAIG Fase II</strong> · Meta: {d['total']:,} sitios · Límite GRT: 15/Jun/2026 · Cierre: 30/Jun/2026</div>
  <div class="footer">CFE Subdirección de Transmisión · Reporte Directivo · PAIG Fase II · {d['fecha_str']}</div>
</div>

<!-- ╔══════════════════════════════════════════╗
     ║  TAB 4 — GEOREFERENCIA (FULL SCREEN)    ║
     ╚══════════════════════════════════════════╝ -->
<div id="tab-geo" class="map-page">
  <div class="map-topbar">
    <svg width="26" height="26" viewBox="0 0 26 26"><rect width="26" height="26" rx="5" fill="rgba(255,255,255,.15)"/><polygon points="16,2 8,14 13,14 10,24 19,10 14,10" fill="#f5a623"/></svg>
    <strong style="font-size:13px">CFE · PAIG Fase II</strong> &nbsp;·&nbsp; Georeferencia de Sitios &nbsp;·&nbsp;
    <span style="color:var(--acc)"><strong>{d['total']:,}</strong></span> sitios &nbsp;·&nbsp; Límite GRT: 15/Jun/2026 &nbsp;·&nbsp; Cierre: 30/Jun/2026
    <span style="margin-left:auto;opacity:.7">📅 {d['fecha_str']}</span>
  </div>
  <div class="map-layout">
    <div class="map-sidebar">
      <div class="msb-header">🔍 Filtros</div>
      <div class="msb-body">
        <div class="fg"><div class="fl">Buscar Sitio / Clave</div>
          <input class="fi" id="f-search" type="text" placeholder="Nombre o clave..." oninput="applyFilters()"></div>
        <div class="fg"><div class="fl">Gerencia Regional</div>
          <select class="fs" id="f-grt" onchange="applyFilters()">
            <option value="">Todas las GRT</option>
            <option>GRT Baja California</option><option>GRT Central</option>
            <option>GRT Noreste</option><option>GRT Norte</option>
            <option>GRT Noroeste</option><option>GRT Occidente</option>
            <option>GRT Oriente</option><option>GRT Valle de México</option>
          </select></div>
        <div class="fg"><div class="fl">Estado</div>
          <select class="fs" id="f-estado" onchange="applyFilters()">
            <option value="">Todos los estados</option>
{estado_opts}
          </select></div>
        <div class="fg"><div class="fl">Estatus</div>
          <select class="fs" id="f-status" onchange="applyFilters()">
            <option value="">Todos</option>
            <option>Instalado - Operando</option><option>Instalado - Sin Señal</option>
            <option>Visitado</option><option>Sin Inmueble</option><option>Sin Actualizar</option>
          </select></div>
        <div class="fg"><div class="fl">Ámbito</div>
          <select class="fs" id="f-ambito" onchange="applyFilters()">
            <option value="">Todos</option><option>Rural</option><option>Urbano</option><option>S/I</option>
          </select></div>
        <div class="fg"><div class="fl">Tipo de Institución</div>
          <select class="fs" id="f-inst" onchange="applyFilters()">
            <option value="">Todas las instituciones</option>
            <option value="TELEESCUELA">▲ Teleescuela</option>
            <option value="IMSS BIENESTAR">● IMSS Bienestar</option>
          </select></div>
        <button class="btn-clear" onclick="resetFilters()">↺ Limpiar filtros</button>
        <div class="leg-sec">
          <div class="leg-title">Leyenda — Clic Filtra</div>
          <div class="leg-item" id="leg-op"   onclick="togLeg('Instalado - Operando','leg-op')"><div class="leg-dot" style="background:#27ae60"></div>Instalado - Operando</div>
          <div class="leg-item" id="leg-ss"   onclick="togLeg('Instalado - Sin Señal','leg-ss')"><div class="leg-dot" style="background:#2980b9"></div>Instalado - Sin Señal</div>
          <div class="leg-item" id="leg-vis"  onclick="togLeg('Visitado','leg-vis')"><div class="leg-dot" style="background:#f39c12"></div>Visitado</div>
          <div class="leg-item" id="leg-sin"  onclick="togLeg('Sin Inmueble','leg-sin')"><div class="leg-dot" style="background:#c0392b"></div>Sin Inmueble</div>
          <div class="leg-item" id="leg-pend" onclick="togLeg('Sin Actualizar','leg-pend')"><div class="leg-dot" style="background:#95a5a6"></div>Sin Actualizar</div>
          <div style="margin-top:9px;padding-top:8px;border-top:1px dashed #ddd;font-size:10px;color:#7a9a8a;line-height:1.8">
            <strong style="color:var(--tx)">Forma del marcador:</strong><br>▲ Triángulo = Teleescuela<br>● Círculo &nbsp; = IMSS Bienestar</div>
        </div>
        <div class="stat-sec">
          <div class="leg-title">Sitios Visibles</div>
          <div class="stat-row"><span class="stat-lbl">Total</span><span class="sv" id="sv-tot">-</span></div>
          <div class="stat-row"><span class="stat-lbl">Operando</span><span class="sv op" id="sv-op">-</span></div>
          <div class="stat-row"><span class="stat-lbl">Sin señal</span><span class="sv ss" id="sv-ss">-</span></div>
          <div class="stat-row"><span class="stat-lbl">Visitado</span><span class="sv vis" id="sv-vis">-</span></div>
          <div class="stat-row"><span class="stat-lbl">Sin inmueble</span><span class="sv sin" id="sv-sin">-</span></div>
          <div class="stat-row"><span class="stat-lbl">Sin actualizar</span><span class="sv pend" id="sv-pend">-</span></div>
        </div>
      </div>
    </div>
    <div id="map"></div>
  </div>
</div>

<!-- ╔══════════════════════════════════════════╗
     ║  TAB 5 — BD CONTROL                     ║
     ╚══════════════════════════════════════════╝ -->
<div id="tab-bd" class="page bd-page">
  <div class="ph">
    <div><h1>📥 Base de Datos Control — PAIG Fase II</h1>
      <div class="sub">CFE Subdirección de Transmisión · {d['total']:,} registros · Fuente: BD_CONTROL</div></div>
    <div class="dbadge">📅 {d['fecha_str']}</div>
  </div>

  <div class="bd-info">
    ℹ️ Puedes filtrar la tabla y descargar el resultado filtrado como Excel. Los datos provienen directamente de la hoja <strong>BD_CONTROL</strong> del archivo de SharePoint.
  </div>

  <div class="bd-filter-row">
    <input class="bd-search" id="bd-q" type="text" placeholder="🔍  Buscar por nombre, clave, estado, municipio..." oninput="filterBD()">
    <select class="bd-sel" id="bd-grt" onchange="filterBD()">
      <option value="">Todas las GRT</option>
      <option>GRT Baja California</option><option>GRT Central</option>
      <option>GRT Noreste</option><option>GRT Norte</option>
      <option>GRT Noroeste</option><option>GRT Occidente</option>
      <option>GRT Oriente</option><option>GRT Valle de México</option>
    </select>
    <select class="bd-sel" id="bd-st" onchange="filterBD()">
      <option value="">Todos los estatus</option>
      <option>Instalado - Operando</option><option>Instalado - Sin Señal</option>
      <option>Visitado</option><option>Sin Inmueble</option><option>Sin Actualizar</option>
    </select>
    <select class="bd-sel" id="bd-inst" onchange="filterBD()">
      <option value="">Todas las instituciones</option>
      <option value="TELEESCUELA">Teleescuela</option>
      <option value="IMSS BIENESTAR">IMSS Bienestar</option>
    </select>
    <button class="btn-dl" onclick="downloadXLSX()">
      <svg width="14" height="14" fill="none" viewBox="0 0 14 14"><path d="M7 1v8M4 6l3 3 3-3M2 11h10" stroke="white" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>
      Descargar Excel
    </button>
    <button class="btn-dl btn-dl-sec" onclick="downloadCSV()">
      <svg width="14" height="14" fill="none" viewBox="0 0 14 14"><path d="M7 1v8M4 6l3 3 3-3M2 11h10" stroke="white" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>
      Descargar CSV
    </button>
  </div>

  <div class="bd-count" id="bd-count">Cargando...</div>

  <div class="bd-table-wrap">
    <table id="bd-table">
      <thead><tr>
        <th>#</th><th>Clave</th><th>Nombre del Sitio</th><th>Estado</th>
        <th>Municipio</th><th>Localidad</th><th>Ámbito</th><th>GRT</th>
        <th>Almacén</th><th>Institución</th><th>Estatus</th>
        <th>Fecha Actualización</th><th>Comentarios</th>
      </tr></thead>
      <tbody id="bd-tbody"></tbody>
    </table>
  </div>
  <div class="footer">CFE Subdirección de Transmisión · BD Control · PAIG Fase II · {d['fecha_str']}</div>
</div>

<!-- ░░ SCRIPTS ░░ -->
<script>
// ── DATOS ──────────────────────────────────────────────────────────────────
const GRT_DATA  = {grt_js};
const LAST3     = {last3_js};
const TREND_LBL = {trend_lbl};
const TREND_V   = {trend_v};
const MA7_V     = {ma7_v};
const RITMO_REQ = {d['ritmo_req_15']};
const GEO_DATA  = {geo_js};
const BD_DATA   = {bd_js};

// ── NAVEGACIÓN ─────────────────────────────────────────────────────────────
function showTab(name, btn) {{
  document.querySelectorAll('.page,.map-page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  if (btn) btn.classList.add('active');
  if (leafletMap) setTimeout(() => leafletMap.invalidateSize(), 100);
}}

// ── ÚLTIMOS 3 DÍAS ─────────────────────────────────────────────────────────
function buildLast3(cid) {{
  const c = document.getElementById(cid);
  LAST3.forEach(d => {{
    const half = Math.ceil(d.desglose.length / 2);
    const d1 = d.desglose.slice(0, half).join(' · ');
    const d2 = d.desglose.slice(half).join(' · ');
    c.innerHTML += `<div class="l3card ${{d.hoy?'today':''}}">
      <div class="l3fecha ${{d.hoy?'today':''}}">${{d.hoy?'◄ ':''}}${{d.fecha}}</div>
      <div class="l3num">${{d.total}}</div>
      <div class="l3lbl">instalaciones</div>
      <div class="l3acum">Acum: ${{d.acum.toLocaleString()}}</div>
      <div class="l3des">${{d1}}<br>${{d2}}</div>
    </div>`;
  }});
}}

// ── TABLA GRT (DASH) ───────────────────────────────────────────────────────
function buildGRTTable() {{
  const tb = document.getElementById('grt-tb');
  GRT_DATA.forEach(r => {{
    const cls = r.pct_inst>=50?'hi':r.pct_inst>=20?'mi':'lo';
    const gc  = r.pct_entr>=90?'hi':r.pct_entr>=70?'mi':'lo';
    const pend = r.total - r.instalados - r.cancelados;
    const rit  = (pend / {d['dias_limite']}).toFixed(1);
    const rc   = parseFloat(rit) <= {d['ritmo_actual']} ? 'var(--g)' : 'var(--red)';
    tb.innerHTML += `<tr>
      <td class="ln">GRT ${{r.grt}}</td><td>${{r.total.toLocaleString()}}</td>
      <td><span class="pill pop">${{r.op.toLocaleString()}}</span></td>
      <td><span class="pill pss">${{r.ss}}</span></td>
      <td>${{r.almacen.toLocaleString()}}</td>
      <td><span class="pb2 ${{cls}}">${{r.pct_inst}}%</span></td>
      <td><span class="pb2 ${{gc}}">${{r.pct_entr.toFixed(1)}}%</span></td>
      <td style="font-weight:700;color:${{rc}}">${{rit}}</td>
      <td><div class="pbar" style="min-width:70px"><div class="pbar-fill ${{r.pct_inst>=50?'pg':'pa'}}" style="width:${{Math.min(r.pct_inst,100)}}%"></div></div></td>
    </tr>`;
  }});
}}

// ── TABLA RESUMEN ──────────────────────────────────────────────────────────
function buildResumenTable() {{
  const tb = document.getElementById('res-tb');
  GRT_DATA.forEach(r => {{
    const cls = r.pct_inst>=50?'hi':r.pct_inst>=20?'mi':'lo';
    const alm = r.entregados - r.instalados;
    const pA  = r.entregados>0 ? (alm/r.entregados*100).toFixed(1) : '0.0';
    tb.innerHTML += `<tr>
      <td class="ln">GRT ${{r.grt}}</td><td>${{r.total}}</td><td>0</td><td>${{r.instalados}}</td>
      <td><span class="pb2 ${{cls}}">${{r.pct_inst}}%</span></td><td>${{r.entregados}}</td>
      <td><span class="pb2 hi">${{r.pct_entr.toFixed(1)}}%</span></td>
      <td>${{alm}}</td><td>${{pA}}%</td>
    </tr>`;
  }});
  const totalAlm = {d['total_entregados']} - {d['inst_total']};
  tb.innerHTML += `<tr style="background:var(--gl);font-weight:700">
    <td class="ln">TOTAL GENERAL</td><td>{d['total']:,}</td><td>0</td><td>{d['inst_total']:,}</td>
    <td><span class="pb2 lo">{d['pct_inst']}%</span></td><td>{d['total_entregados']:,}</td>
    <td><span class="pb2 hi">{d['pct_entr']}%</span></td>
    <td>${{totalAlm.toLocaleString()}}</td><td>${{(totalAlm/{d['total_entregados']}*100).toFixed(1)}}%</td>
  </tr>`;
}}

// ── TABLA DIRECTIVO ────────────────────────────────────────────────────────
function buildDirectivoTable() {{
  const tb = document.getElementById('dir-tb');
  GRT_DATA.forEach(r => {{
    const ce = r.pct_entr>=95?'hi':r.pct_entr>=80?'mi':'lo';
    const ci = r.pct_inst>=50?'hi':r.pct_inst>=20?'mi':'lo';
    tb.innerHTML += `<tr>
      <td class="ln">GRT ${{r.grt}}</td><td>${{r.total}}</td><td>${{r.entregados}}</td>
      <td>${{r.pend_entr}}</td>
      <td><span class="pb2 ${{ce}}">${{r.pct_entr.toFixed(1)}}%</span></td>
      <td>${{r.instalados}}</td>
      <td><span class="pb2 ${{ci}}">${{r.pct_inst}}%</span></td>
      <td>${{r.pct_ie.toFixed(1)}}%</td>
    </tr>`;
  }});
  tb.innerHTML += `<tr style="background:var(--gl);font-weight:700">
    <td class="ln">TOTAL GENERAL</td><td>{d['total']:,}</td><td>{d['total_entregados']:,}</td>
    <td>{d['total']-d['total_entregados']:,}</td>
    <td><span class="pb2 hi">{d['pct_entr']}%</span></td>
    <td>{d['inst_total']:,}</td>
    <td><span class="pb2 lo">{d['pct_inst']}%</span></td>
    <td>{round(d['inst_total']/d['total_entregados']*100,1)}%</td>
  </tr>`;
}}

// ── CHARTS ─────────────────────────────────────────────────────────────────
const PAL = ['#27ae60','#f39c12','#2980b9','#8e44ad','#e74c3c','#16a085','#d35400','#1abc9c'];
const GN  = GRT_DATA.map(r => r.grt.length>9 ? r.grt.substring(0,9)+'…' : r.grt);

function initCharts() {{
  // Tendencia
  new Chart(document.getElementById('chartTend'), {{
    data: {{
      labels: TREND_LBL,
      datasets: [
        {{ type:'bar', label:'Inst./día', data:TREND_V,
           backgroundColor: TREND_V.map(v => v>=RITMO_REQ?'#27ae60':'#f5a623'),
           borderRadius:4, yAxisID:'y', order:2 }},
        {{ type:'line', label:'Media móvil 7d', data:MA7_V,
           borderColor:'#1a5276', backgroundColor:'rgba(26,82,118,.1)',
           borderWidth:2.5, pointRadius:3, tension:.4, yAxisID:'y', order:1 }},
        {{ type:'line', label:`Ritmo req. (${{RITMO_REQ}})`,
           data:Array(TREND_LBL.length).fill(RITMO_REQ),
           borderColor:'#c0392b', borderDash:[6,4], borderWidth:2,
           pointRadius:0, fill:false, yAxisID:'y', order:0 }}
      ]
    }},
    options:{{ responsive:true, maintainAspectRatio:false,
      plugins:{{ legend:{{ labels:{{ font:{{ size:11 }}, padding:12, usePointStyle:true }} }},
                tooltip:{{ mode:'index', intersect:false }} }},
      scales:{{ y:{{ beginAtZero:true, ticks:{{ font:{{ size:10 }} }}, grid:{{ color:'#eef2f0' }},
                     title:{{ display:true, text:'Inst./día', font:{{ size:10 }} }} }},
                x:{{ ticks:{{ font:{{ size:10 }} }}, grid:{{ display:false }} }} }}
    }}
  }});
  // Donut
  new Chart(document.getElementById('chartDonut'), {{
    type:'doughnut',
    data:{{ labels:['Inst. Operando','Sin Señal','Visitado','Sin Inmueble','Sin Actualizar'],
            datasets:[{{ data:[{d['op']},{d['ss']},{d['vis']},{d['sin_in']},{d['pendiente']}],
                         backgroundColor:['#27ae60','#2980b9','#f39c12','#c0392b','#bdc3c7'], borderWidth:2 }}] }},
    options:{{ responsive:true, maintainAspectRatio:false,
               plugins:{{ legend:{{ position:'bottom', labels:{{ font:{{ size:10 }}, padding:8 }} }} }} }}
  }});
  // Ritmo
  const pends = GRT_DATA.map(r => r.total - r.instalados - r.cancelados);
  const rits  = pends.map(p => (p/{d['dias_limite']}).toFixed(1));
  new Chart(document.getElementById('chartRitmo'), {{
    type:'bar', data:{{ labels:GN, datasets:[{{ label:'Inst./día req.', data:rits,
      backgroundColor:rits.map(r=>parseFloat(r)<={d['ritmo_actual']}?'#27ae60':'#c0392b'), borderRadius:5 }}] }},
    options:{{ responsive:true, maintainAspectRatio:false,
      plugins:{{ legend:{{ display:false }} }},
      scales:{{ y:{{ beginAtZero:true, ticks:{{ font:{{ size:10 }} }},
                     title:{{ display:true, text:'Inst./día al 15/Jun', font:{{ size:10 }} }} }},
                x:{{ ticks:{{ font:{{ size:10 }} }} }} }}
    }}
  }});
  // Res Bar
  new Chart(document.getElementById('chartResBar'), {{
    type:'bar', data:{{ labels:GN, datasets:[
      {{ label:'Entregados', data:GRT_DATA.map(r=>r.entregados), backgroundColor:'#1a5276', borderRadius:3 }},
      {{ label:'Instalados', data:GRT_DATA.map(r=>r.instalados), backgroundColor:'#27ae60', borderRadius:3 }}
    ]}},
    options:{{ responsive:true, maintainAspectRatio:false,
               plugins:{{ legend:{{ labels:{{ font:{{ size:10 }} }} }} }},
               scales:{{ y:{{ beginAtZero:true, ticks:{{ font:{{ size:10 }} }} }}, x:{{ ticks:{{ font:{{ size:10 }} }} }} }} }}
  }});
  // Res Pct
  new Chart(document.getElementById('chartResPct'), {{
    type:'bar', data:{{ labels:GN, datasets:[{{ label:'% Instalado',
      data:GRT_DATA.map(r=>r.pct_inst), backgroundColor:PAL, borderRadius:5 }}] }},
    options:{{ responsive:true, maintainAspectRatio:false, plugins:{{ legend:{{ display:false }} }},
               scales:{{ y:{{ beginAtZero:true, max:110, ticks:{{ font:{{ size:10 }}, callback:v=>v+'%' }} }}, x:{{ ticks:{{ font:{{ size:10 }} }} }} }} }}
  }});
  // Dir Entr
  new Chart(document.getElementById('chartDirEntr'), {{
    type:'bar', data:{{ labels:GN, datasets:[{{ label:'% Entregado',
      data:GRT_DATA.map(r=>r.pct_entr), backgroundColor:PAL, borderRadius:5 }}] }},
    options:{{ responsive:true, maintainAspectRatio:false, plugins:{{ legend:{{ display:false }} }},
               scales:{{ y:{{ beginAtZero:true, max:110, ticks:{{ font:{{ size:10 }}, callback:v=>v+'%' }} }}, x:{{ ticks:{{ font:{{ size:10 }} }} }} }} }}
  }});
  // Dir Inst
  new Chart(document.getElementById('chartDirInst'), {{
    type:'bar', data:{{ labels:GN, datasets:[{{ label:'% Inst./Asignados',
      data:GRT_DATA.map(r=>r.pct_inst), backgroundColor:'#27ae60', borderRadius:5 }}] }},
    options:{{ responsive:true, maintainAspectRatio:false, plugins:{{ legend:{{ display:false }} }},
               scales:{{ y:{{ beginAtZero:true, max:110, ticks:{{ font:{{ size:10 }}, callback:v=>v+'%' }} }}, x:{{ ticks:{{ font:{{ size:10 }} }} }} }} }}
  }});
}}

// ── MAPA ────────────────────────────────────────────────────────────────────
let leafletMap=null, markersLayer=null, mapReady=false;
const SC = {{'Instalado - Operando':'#27ae60','Instalado - Sin Señal':'#2980b9','Visitado':'#f39c12','Sin Inmueble':'#c0392b','Sin Actualizar':'#95a5a6'}};
const SB = {{'Instalado - Operando':'#1a7a44','Instalado - Sin Señal':'#1a5e8a','Visitado':'#c87f00','Sin Inmueble':'#8b1a1a','Sin Actualizar':'#7f8c8d'}};
const BC = {{'Instalado - Operando':'#d5f0e3','Instalado - Sin Señal':'#d6eaf8','Visitado':'#fff3cd','Sin Inmueble':'#fdecea','Sin Actualizar':'#f0f0f0'}};
const FC = {{'Instalado - Operando':'#1a6640','Instalado - Sin Señal':'#1a5276','Visitado':'#856404','Sin Inmueble':'#8b1a1a','Sin Actualizar':'#555'}};
const legA = {{'Instalado - Operando':true,'Instalado - Sin Señal':true,'Visitado':true,'Sin Inmueble':true,'Sin Actualizar':true}};

function makeIcon(s) {{
  const c=SC[s.st]||'#95a5a6', b=SB[s.st]||'#7f8c8d', iT=s.i==='TELEESCUELA';
  if(iT) {{
    return L.divIcon({{className:'',iconSize:[18,18],iconAnchor:[9,16],popupAnchor:[0,-17],
      html:`<svg width="18" height="18" viewBox="0 0 18 18"><polygon points="9,1 17,16 1,16" fill="${{c}}" stroke="${{b}}" stroke-width="1.5"/></svg>`}});
  }}
  return L.divIcon({{className:'',iconSize:[15,15],iconAnchor:[7,7],popupAnchor:[0,-8],
    html:`<svg width="15" height="15" viewBox="0 0 15 15"><circle cx="7.5" cy="7.5" r="6.5" fill="${{c}}" stroke="${{b}}" stroke-width="1.5"/></svg>`}});
}}

function popupHTML(s) {{
  const ib=s.i==='TELEESCUELA'?'#e8f5ee':'#fdeef0', ifc=s.i==='TELEESCUELA'?'#1a6640':'#8b1a1a', sym=s.i==='TELEESCUELA'?'▲':'●';
  return `<div style="font-size:12px;min-width:200px">
    <strong style="color:#004d29;font-size:13px">${{s.n}}</strong><br>
    <span style="background:${{ib}};color:${{ifc}};padding:2px 6px;border-radius:8px;font-size:10px;font-weight:700">${{sym}} ${{s.i}}</span>
    <hr style="margin:5px 0">
    <div style="color:#555;font-size:11px">📍 ${{s.e}} · ${{s.mu}}</div>
    <div style="color:#555;font-size:11px">🏢 ${{s.g}}</div>
    <div style="color:#555;font-size:11px">🔑 ${{s.c}} · 🏘️ ${{s.am}}</div>
    <div style="font-weight:700;margin-top:4px;color:${{FC[s.st]||'#555'}}">● ${{s.st}}</div>
  </div>`;
}}

function initMap() {{
  if(mapReady) return; mapReady=true;
  leafletMap = L.map('map',{{zoomControl:true}}).setView([23.6,-102.5],5);
  L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png',{{attribution:'© OpenStreetMap',maxZoom:19}}).addTo(leafletMap);
  markersLayer = L.layerGroup().addTo(leafletMap);
  renderMap(GEO_DATA);
}}

function renderMap(data) {{
  markersLayer.clearLayers();
  const cnt={{'Instalado - Operando':0,'Instalado - Sin Señal':0,'Visitado':0,'Sin Inmueble':0,'Sin Actualizar':0}};
  data.forEach(s=>{{
    L.marker([s.lat,s.lng],{{icon:makeIcon(s)}}).bindPopup(popupHTML(s),{{maxWidth:240}}).addTo(markersLayer);
    if(cnt[s.st]!==undefined) cnt[s.st]++;
  }});
  document.getElementById('sv-tot').textContent = data.length.toLocaleString();
  document.getElementById('sv-op').textContent  = cnt['Instalado - Operando'].toLocaleString();
  document.getElementById('sv-ss').textContent  = cnt['Instalado - Sin Señal'].toLocaleString();
  document.getElementById('sv-vis').textContent = cnt['Visitado'];
  document.getElementById('sv-sin').textContent = cnt['Sin Inmueble'];
  document.getElementById('sv-pend').textContent= cnt['Sin Actualizar'].toLocaleString();
}}

function getMapFiltered() {{
  const sq=document.getElementById('f-search').value.trim().toLowerCase();
  const g=document.getElementById('f-grt').value, e=document.getElementById('f-estado').value;
  const st=document.getElementById('f-status').value, am=document.getElementById('f-ambito').value;
  const ins=document.getElementById('f-inst').value;
  return GEO_DATA.filter(s=>{{
    if(!legA[s.st]) return false;
    if(sq && !s.n.toLowerCase().includes(sq) && !s.c.toLowerCase().includes(sq)) return false;
    if(g && s.g!==g) return false; if(e && s.e!==e) return false;
    if(st && s.st!==st) return false; if(am && s.am!==am) return false;
    if(ins && s.i!==ins) return false; return true;
  }});
}}

let ft=null;
function applyFilters(){{ clearTimeout(ft); ft=setTimeout(()=>renderMap(getMapFiltered()),120); }}
function resetFilters(){{
  ['f-search','f-grt','f-estado','f-status','f-ambito','f-inst'].forEach(id=>{{
    const el=document.getElementById(id); el.tagName==='INPUT'?el.value='':el.value='';
  }});
  Object.keys(legA).forEach(k=>legA[k]=true);
  document.querySelectorAll('.leg-item').forEach(el=>el.classList.remove('inactive'));
  applyFilters();
}}
function togLeg(st,eid){{ legA[st]=!legA[st]; document.getElementById(eid).classList.toggle('inactive',!legA[st]); applyFilters(); }}

// ── BD CONTROL ──────────────────────────────────────────────────────────────
let bdFiltered = BD_DATA;
let bdInit = false;

function initBD() {{
  if(bdInit) return; bdInit=true;
  filterBD();
}}

function stClass(st) {{
  if(st==='Instalado - Operando')  return 'st-op';
  if(st==='Instalado - Sin Señal') return 'st-ss';
  if(st==='Visitado')              return 'st-vis';
  if(st==='Sin Inmueble')          return 'st-sin';
  if(st==='Cancelado')             return 'st-canc';
  return 'st-pend';
}}

function filterBD() {{
  const q    = document.getElementById('bd-q').value.trim().toLowerCase();
  const grt  = document.getElementById('bd-grt').value;
  const st   = document.getElementById('bd-st').value;
  const inst = document.getElementById('bd-inst').value;
  bdFiltered = BD_DATA.filter(r => {{
    if(q && !r['NOMBRE DEL SITIO'].toLowerCase().includes(q) &&
           !r['CLAVE'].toLowerCase().includes(q) &&
           !r['ESTADO'].toLowerCase().includes(q) &&
           !r['MUNICIPIO'].toLowerCase().includes(q)) return false;
    if(grt  && r['GERENCIA (GRT)'] !== grt)  return false;
    if(st   && r['ESTATUS']         !== st)   return false;
    if(inst && r['INSTITUCIÓN']     !== inst) return false;
    return true;
  }});
  renderBDTable();
}}

function renderBDTable() {{
  const tb = document.getElementById('bd-tbody');
  document.getElementById('bd-count').textContent = `${{bdFiltered.length.toLocaleString()}} registros`;
  const show = bdFiltered.slice(0, 500);  // max 500 rows rendered at once
  tb.innerHTML = show.map((r,i) => `<tr>
    <td>${{i+1}}</td>
    <td style="font-family:monospace;font-size:10px">${{r['CLAVE']}}</td>
    <td style="text-align:left;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r['NOMBRE DEL SITIO']}}">${{r['NOMBRE DEL SITIO']}}</td>
    <td>${{r['ESTADO']}}</td><td>${{r['MUNICIPIO']}}</td>
    <td style="max-width:100px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${{r['LOCALIDAD']}}</td>
    <td>${{r['ÁMBITO']}}</td>
    <td style="white-space:nowrap">${{r['GERENCIA (GRT)'].replace('GRT ','')}}</td>
    <td style="max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r['ALMACÉN']}}">${{r['ALMACÉN']}}</td>
    <td><span style="font-size:10px;font-weight:700;color:${{r['INSTITUCIÓN']==='TELEESCUELA'?'#1a6640':'#8b1a1a'}}">${{r['INSTITUCIÓN']==='TELEESCUELA'?'▲':'●'}} ${{r['INSTITUCIÓN']}}</span></td>
    <td><span class="${{stClass(r['ESTATUS'])}}">${{r['ESTATUS']}}</span></td>
    <td>${{r['FECHA ACTUALIZACIÓN']}}</td>
    <td style="max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r['COMENTARIOS']}}">${{r['COMENTARIOS']}}</td>
  </tr>`).join('');
  if(bdFiltered.length > 500) {{
    tb.innerHTML += `<tr><td colspan="13" style="text-align:center;padding:12px;color:var(--tx2);font-style:italic">
      Mostrando 500 de ${{bdFiltered.length.toLocaleString()}} registros. Descarga el Excel para ver todos.
    </td></tr>`;
  }}
}}

function downloadXLSX() {{
  const cols = ['CLAVE','ID PUNTO','NOMBRE DEL SITIO','ESTADO','MUNICIPIO','LOCALIDAD','ÁMBITO','GERENCIA (GRT)','ALMACÉN','INSTITUCIÓN','ESTATUS','FECHA ACTUALIZACIÓN','COMENTARIOS'];
  const ws = XLSX.utils.json_to_sheet(bdFiltered, {{header: cols}});
  // Column widths
  ws['!cols'] = [{{wch:14}},{{wch:10}},{{wch:40}},{{wch:18}},{{wch:18}},{{wch:20}},{{wch:8}},{{wch:22}},{{wch:25}},{{wch:14}},{{wch:22}},{{wch:18}},{{wch:30}}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'BD Control PAIG');
  const fname = `BD_Control_PAIG_${{new Date().toISOString().slice(0,10)}}.xlsx`;
  XLSX.writeFile(wb, fname);
}}

function downloadCSV() {{
  const cols = Object.keys(bdFiltered[0]);
  const rows = [cols.join(','), ...bdFiltered.map(r => cols.map(c => `"${{String(r[c]).replace(/"/g,'""')}}"`).join(','))];
  const blob = new Blob([rows.join('\\n')], {{type:'text/csv;charset=utf-8;'}});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
  a.download = `BD_Control_PAIG_${{new Date().toISOString().slice(0,10)}}.csv`; a.click();
}}

// ── INIT ────────────────────────────────────────────────────────────────────
buildGRTTable();
buildResumenTable();
buildDirectivoTable();
buildLast3('l3-resumen');
buildLast3('l3-dir');
initCharts();
</script>
</body>
</html>"""

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    size_kb = os.path.getsize(output_path) // 1024
    print(f"  ✅ HTML generado: {output_path}  ({size_kb} KB)")


# ──────────────────────────────────────────────────────────────────────────────
#  MAIN
# ──────────────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description='Genera el Dashboard PAIG Fase II (5 pestañas) desde el Excel de SharePoint'
    )
    parser.add_argument('--excel',  default=DEFAULT_EXCEL,  help='Ruta al archivo Excel (.xlsx)')
    parser.add_argument('--output', default=DEFAULT_OUTPUT, help='Ruta del HTML de salida')
    args = parser.parse_args()

    print(f"\n{'='*65}")
    print(f"  CFE · PAIG Fase II — Generador de Dashboard")
    print(f"  {datetime.now():%d/%m/%Y  %H:%M:%S}")
    print(f"{'='*65}")

    if not Path(args.excel).exists():
        sys.exit(f"\n❌ Archivo no encontrado: {args.excel}\n"
                 f"   Ajusta la ruta con --excel o revisa el paso de descarga en Power Automate.\n")

    datos = leer_excel(args.excel)
    generar_html(datos, args.output)

    print(f"\n  📊 Resumen generado:")
    print(f"     Instalados : {datos['inst_total']:,} / {datos['total']:,}  ({datos['pct_inst']}%)")
    print(f"     Entregados : {datos['total_entregados']:,}  ({datos['pct_entr']}%)")
    print(f"     Ritmo req. : {datos['ritmo_req_15']}/día  (15/Jun)   |   actual: {datos['ritmo_actual']}/día")
    print(f"     Proyección : {datos['fecha_proy']}  ({'✅ dentro del límite' if datos['en_meta'] else '⚠️ fuera del límite'})")
    print(f"\n{'='*65}\n")


if __name__ == '__main__':
    main()
