"""
Honeywell Program Dashboard — Refresh Script
=============================================
HOW TO USE EACH TIME YOU WANT TO UPDATE THE DASHBOARD:
  1. Save and close your Excel file
  2. Open Command Prompt (search "cmd" in Windows Start)
  3. Run this command:
       python "C:\\Users\\zn424f\\OneDrive - The Boeing Company\\Working KPIs\\refresh_dashboard.py"
  4. It generates a fresh index.html in your Working KPIs folder
  5. Upload index.html to your GitHub repo — dashboard is live in ~60 seconds

FIRST TIME SETUP (one time only):
  pip install pandas openpyxl
"""

import pandas as pd
import numpy as np
from datetime import date
import sys

# ─────────────────────────────────────────────────────────────
#  CONFIG — only edit this section if paths or sheet names change
# ─────────────────────────────────────────────────────────────
EXCEL_FILE  = r"C:\Users\zn424f\OneDrive - The Boeing Company\Working KPIs\Working Honeywell KPI Dashboard.xlsx"
SAW_SHEET   = "SAW Report Data for Current Day"
HIST_SHEET  = "Historical Revenue & Past Dues"
OUTPUT_FILE = r"C:\Users\zn424f\OneDrive - The Boeing Company\Working KPIs\index.html"
TODAY       = pd.Timestamp(date.today())
# ─────────────────────────────────────────────────────────────


# ── Helpers ──────────────────────────────────────────────────
def fmt_m(val):
    return f"${val/1e6:.1f}M"

def fmt_k(val):
    return f"${val/1e3:.0f}K"

def fmt_pct(val):
    return f"{val:.1f}%"

def fmt_comma(val):
    return f"{int(val):,}"

def js_arr(lst):
    parts = ["null" if v is None else str(v) for v in lst]
    return "[" + ",".join(parts) + "]"

def js_str_arr(lst):
    escaped = [str(v).replace("'", "\\'").replace("&", "and") for v in lst]
    return "['" + "','".join(escaped) + "']"

def safe(val, fallback=0):
    if val is None:
        return fallback
    if isinstance(val, float) and np.isnan(val):
        return fallback
    return val


# ── Load data ─────────────────────────────────────────────────
def load_data():
    print(f"\n  Loading Excel file...")
    try:
        saw  = pd.read_excel(EXCEL_FILE, sheet_name=SAW_SHEET)
        hist = pd.read_excel(EXCEL_FILE, sheet_name=HIST_SHEET)
    except FileNotFoundError:
        print("\n  ERROR: Excel file not found. Check the path in CONFIG.")
        print(f"  Expected: {EXCEL_FILE}")
        sys.exit(1)
    except Exception as e:
        print(f"\n  ERROR: {e}")
        sys.exit(1)

    hist["Date"] = pd.to_datetime(hist["Date"])
    hist = hist.sort_values("Date")
    print(f"  SAW rows loaded:  {len(saw):,}")
    print(f"  Hist rows loaded: {len(hist):,}")
    return saw, hist


# ── Compute all metrics ───────────────────────────────────────
def compute(saw, hist):
    print("  Computing metrics...")

    past_due = saw[saw["Status"] == "Past Due"].copy()
    po_pd    = past_due[past_due["PO/Bin"] == "PO"]
    bin_pd   = past_due[past_due["PO/Bin"] == "Binstock"]
    backlog  = saw[~saw["Status"].isin(["Past Due", "Today"])]

    cur_year      = TODAY.year
    prev_year     = cur_year - 1
    same_day_prev = TODAY.replace(year=prev_year)

    h_cur  = hist[hist["Date"].dt.year == cur_year].sort_values("Date")
    h_prev = hist[
        (hist["Date"].dt.year == prev_year) &
        (hist["Date"] <= same_day_prev)
    ].sort_values("Date")

    latest = h_cur.iloc[-1] if len(h_cur) else hist.iloc[-1]
    as_of  = latest["Date"].strftime("%B %d, %Y")

    ytd_rev       = safe(latest["Revenue ($)"])
    baseline_rev  = safe(latest.get("Baseline Plan (Revenue)"), None)
    pgm           = safe(latest["PGM (%)"]) * 100
    b_pgm_raw     = latest.get("Baseline Plan PGM")
    baseline_pgm  = (safe(b_pgm_raw, None) * 100) if b_pgm_raw is not None and pd.notna(b_pgm_raw) else None
    pd_dollars    = safe(latest["Past Due Dollars ($)"])
    pd_lines      = int(safe(latest["Past Due Lines"]))
    ytd_rev_prev  = h_prev["Revenue ($)"].max() if len(h_prev) else None

    total_book    = saw["Extended Price"].sum()
    backlog_val   = backlog["Extended Price"].sum()
    backlog_lines = len(backlog)

    po_total_lines  = len(saw[saw["PO/Bin"] == "PO"])
    bin_total_lines = len(saw[saw["PO/Bin"] == "Binstock"])
    po_pd_lines     = len(po_pd)
    po_pd_dollars   = po_pd["Extended Price"].sum()
    bin_pd_lines    = len(bin_pd)
    bin_pd_dollars  = bin_pd["Extended Price"].sum()

    # Trend series — deduplicated daily snapshots for current year
    h_cur_u     = h_cur.drop_duplicates(subset="Date")
    trend_dates = list(h_cur_u["Date"].dt.strftime("%b %d"))

    def to_js(series, divisor=1, mult=1, decimals=2):
        return [
            round(float(v) * mult / divisor, decimals) if pd.notna(v) else None
            for v in series
        ]

    trend_rev  = to_js(h_cur_u["Revenue ($)"],       divisor=1e6)
    trend_pgm  = to_js(h_cur_u["PGM (%)"],           mult=100, decimals=1)

    brev_col = h_cur_u.get("Baseline Plan (Revenue)", pd.Series([None]*len(h_cur_u)))
    bpgm_col = h_cur_u.get("Baseline Plan PGM",       pd.Series([None]*len(h_cur_u)))
    trend_brev = to_js(brev_col, divisor=1e6)
    trend_bpgm = to_js(bpgm_col, mult=100, decimals=1)

    # PO by function
    po_func = (po_pd.groupby("ActionBy - New")
               .agg(Lines=("Extended Price","count"), Dollars=("Extended Price","sum"))
               .sort_values("Dollars", ascending=False).reset_index())

    # PO by site (top 10)
    po_site = (po_pd.groupby("Customer Name")
               .agg(Lines=("Extended Price","count"), Dollars=("Extended Price","sum"))
               .sort_values("Dollars", ascending=False).head(10).reset_index())
    po_site["Label"] = po_site["Customer Name"].str.replace("Honeywell ", "", regex=False)

    # Bin by function
    bin_func = (bin_pd.groupby("ActionBy - New")
                .agg(Lines=("Extended Price","count"), Dollars=("Extended Price","sum"))
                .sort_values("Lines", ascending=False).reset_index())

    # Bin by site (top 10)
    bin_site = (bin_pd.groupby("Customer Name")
                .agg(Lines=("Extended Price","count"), Dollars=("Extended Price","sum"))
                .sort_values("Lines", ascending=False).head(10).reset_index())
    bin_site["Label"] = bin_site["Customer Name"].str.replace("Honeywell ", "", regex=False)

    # Top callout values
    top_func     = po_func.iloc[0]["ActionBy - New"] if len(po_func) else "N/A"
    top_func_l   = int(po_func.iloc[0]["Lines"])     if len(po_func) else 0
    top_func_d   = po_func.iloc[0]["Dollars"]        if len(po_func) else 0
    top_func_pct = round(top_func_d / po_pd_dollars * 100, 1) if po_pd_dollars else 0

    top_site     = po_site.iloc[0]["Label"]      if len(po_site) else "N/A"
    top_site_l   = int(po_site.iloc[0]["Lines"]) if len(po_site) else 0
    top_site_d   = po_site.iloc[0]["Dollars"]    if len(po_site) else 0
    top_site_pct = round(top_site_d / po_pd_dollars * 100, 1) if po_pd_dollars else 0

    return dict(
        as_of=as_of, cur_year=cur_year, prev_year=prev_year,
        ytd_rev=ytd_rev, baseline_rev=baseline_rev, ytd_rev_prev=ytd_rev_prev,
        pgm=pgm, baseline_pgm=baseline_pgm,
        pd_dollars=pd_dollars, pd_lines=pd_lines,
        total_book=total_book, backlog_val=backlog_val, backlog_lines=backlog_lines,
        po_total_lines=po_total_lines, bin_total_lines=bin_total_lines,
        po_pd_lines=po_pd_lines, po_pd_dollars=po_pd_dollars,
        bin_pd_lines=bin_pd_lines, bin_pd_dollars=bin_pd_dollars,
        trend_dates=trend_dates, trend_rev=trend_rev, trend_pgm=trend_pgm,
        trend_brev=trend_brev, trend_bpgm=trend_bpgm,
        po_func=po_func, po_site=po_site,
        bin_func=bin_func, bin_site=bin_site,
        top_func=top_func, top_func_l=top_func_l,
        top_func_d=top_func_d, top_func_pct=top_func_pct,
        top_site=top_site, top_site_l=top_site_l,
        top_site_d=top_site_d, top_site_pct=top_site_pct,
    )


# ── Build HTML ────────────────────────────────────────────────
def build_html(m):
    print("  Building HTML...")

    rev_vs_plan  = (f"+{fmt_m(m['ytd_rev'] - m['baseline_rev'])} vs baseline plan ({fmt_m(m['baseline_rev'])})"
                    if m["baseline_rev"] else "baseline plan N/A")
    rev_yoy_d    = fmt_m(m["ytd_rev"] - m["ytd_rev_prev"]) if m["ytd_rev_prev"] else "N/A"
    rev_yoy_pct  = (f"{(m['ytd_rev']-m['ytd_rev_prev'])/m['ytd_rev_prev']*100:.1f}%"
                    if m["ytd_rev_prev"] else "")
    pgm_vs_plan  = (f"+{(m['pgm']-m['baseline_pgm'])*100:.0f}bps above baseline ({fmt_pct(m['baseline_pgm'])})"
                    if m["baseline_pgm"] else "")
    po_pd_pct     = round(m["po_pd_lines"]  / m["po_total_lines"]  * 100, 1) if m["po_total_lines"]  else 0
    po_dollar_pct = round(m["po_pd_dollars"] / m["pd_dollars"]      * 100, 1) if m["pd_dollars"]      else 0
    bin_pd_pct    = round(m["bin_pd_lines"]  / m["bin_total_lines"] * 100, 1) if m["bin_total_lines"] else 0
    bin_dollar_pct= round(m["bin_pd_dollars"]/ m["pd_dollars"]      * 100, 1) if m["pd_dollars"]      else 0

    po_func_labels  = js_str_arr(m["po_func"]["ActionBy - New"].tolist())
    po_func_d_vals  = js_arr([round(v/1e6,2)  for v in m["po_func"]["Dollars"].tolist()])
    po_func_l_vals  = js_arr(m["po_func"]["Lines"].tolist())

    po_site_labels  = js_str_arr(m["po_site"]["Label"].tolist())
    po_site_d_vals  = js_arr([round(v/1e6,3)  for v in m["po_site"]["Dollars"].tolist()])
    po_site_colors  = js_str_arr([
        "#C0392B" if i==0 else "#E07050" if i<3 else "#185FA5"
        for i in range(len(m["po_site"]))
    ])

    bin_func_labels = js_str_arr(m["bin_func"]["ActionBy - New"].tolist())
    bin_func_d_vals = js_arr([round(v/1e3,1)  for v in m["bin_func"]["Dollars"].tolist()])
    bin_func_l_vals = js_arr(m["bin_func"]["Lines"].tolist())

    bin_site_labels = js_str_arr(m["bin_site"]["Label"].tolist())
    bin_site_l_vals = js_arr(m["bin_site"]["Lines"].tolist())

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Honeywell Global Program &mdash; Executive Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:Arial,sans-serif;background:#f5f5f2;padding:24px;color:#2C2C2A}}
.dash{{max-width:1200px;margin:0 auto}}
.hdr{{background:#0D1B3E;border-radius:10px;padding:20px 24px;margin-bottom:20px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:14px}}
.hdr-title{{font-size:20px;font-weight:500;color:#fff}}
.hdr-sub{{font-size:12px;color:rgba(255,255,255,0.6);margin-top:3px}}
.hdr-badges{{display:flex;gap:10px;flex-wrap:wrap}}
.hdr-badge{{background:rgba(255,255,255,0.12);border-radius:6px;padding:8px 16px;text-align:center;min-width:100px}}
.hdr-badge-val{{font-size:18px;font-weight:500;color:#fff}}
.hdr-badge-lbl{{font-size:10px;color:rgba(255,255,255,0.6);margin-top:2px}}
.sl{{font-size:10px;font-weight:500;color:#5F5E5A;text-transform:uppercase;letter-spacing:.07em;margin:18px 0 10px}}
.kpi-grid{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin-bottom:14px}}
.kpi-grid-2{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;margin-bottom:14px}}
@media(max-width:800px){{.kpi-grid{{grid-template-columns:repeat(2,1fr)}}.kpi-grid-2{{grid-template-columns:1fr}}}}
@media(max-width:480px){{.kpi-grid{{grid-template-columns:1fr}}}}
.kpi{{background:#fff;border-radius:8px;padding:14px 16px;border:.5px solid #D3D1C7}}
.kpi-label{{font-size:11px;color:#888780;margin-bottom:4px}}
.kpi-val{{font-size:21px;font-weight:500;color:#2C2C2A;line-height:1.2}}
.kpi-val.g{{color:#0F6E56}}.kpi-val.r{{color:#993C1D}}.kpi-val.b{{color:#185FA5}}.kpi-val.a{{color:#854F0B}}
.kpi-sub{{font-size:11px;color:#888780;margin-top:4px}}
.kpi-sub.up{{color:#0F6E56}}.kpi-sub.dn{{color:#993C1D}}
.cg{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin-bottom:14px}}
@media(max-width:800px){{.cg{{grid-template-columns:1fr}}}}
.cb{{background:#fff;border:.5px solid #D3D1C7;border-radius:10px;padding:16px 18px}}
.ct{{font-size:13px;font-weight:500;color:#2C2C2A;margin-bottom:3px}}
.cs{{font-size:11px;color:#888780;margin-bottom:10px}}
.leg{{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:8px;font-size:11px;color:#5F5E5A}}
.lsq{{width:10px;height:10px;border-radius:2px;display:inline-block;margin-right:3px;vertical-align:middle}}
.divider{{border:none;border-top:1px solid #D3D1C7;margin:22px 0 6px}}
.fn{{background:#EBF3FC;border-radius:8px;padding:12px 16px;font-size:12px;color:#0C447C;border:.5px solid #B5D4F4;margin-top:20px}}
.rt{{font-size:11px;color:#888780;text-align:right;margin-top:14px}}
</style>
</head>
<body>
<div class="dash">

<div class="hdr">
  <div>
    <div class="hdr-title">Honeywell Global Program &mdash; Executive Dashboard</div>
    <div class="hdr-sub">YTD {m['cur_year']} &middot; As of {m['as_of']}</div>
  </div>
  <div class="hdr-badges">
    <div class="hdr-badge"><div class="hdr-badge-val">{fmt_m(m['ytd_rev'])}</div><div class="hdr-badge-lbl">YTD Revenue</div></div>
    <div class="hdr-badge"><div class="hdr-badge-val">{fmt_pct(m['pgm'])}</div><div class="hdr-badge-lbl">PGM %</div></div>
    <div class="hdr-badge"><div class="hdr-badge-val">{fmt_m(m['pd_dollars'])}</div><div class="hdr-badge-lbl">Past Due $</div></div>
    <div class="hdr-badge"><div class="hdr-badge-val">{fmt_comma(m['pd_lines'])}</div><div class="hdr-badge-lbl">Past Due Lines</div></div>
  </div>
</div>

<div class="sl">Revenue &amp; Margin Performance</div>
<div class="kpi-grid">
  <div class="kpi"><div class="kpi-label">YTD Revenue</div><div class="kpi-val g">{fmt_m(m['ytd_rev'])}</div><div class="kpi-sub up">{rev_vs_plan}</div></div>
  <div class="kpi"><div class="kpi-label">vs {m['prev_year']} same period</div><div class="kpi-val g">+{rev_yoy_d}</div><div class="kpi-sub up">{m['prev_year']} YTD was {fmt_m(m['ytd_rev_prev']) if m['ytd_rev_prev'] else 'N/A'} &middot; +{rev_yoy_pct} YoY</div></div>
  <div class="kpi"><div class="kpi-label">Program Gross Margin</div><div class="kpi-val g">{fmt_pct(m['pgm'])}</div><div class="kpi-sub up">{pgm_vs_plan}</div></div>
  <div class="kpi"><div class="kpi-label">Total Book Value</div><div class="kpi-val b">{fmt_m(m['total_book'])}</div><div class="kpi-sub">Backlog: {fmt_m(m['backlog_val'])} &middot; {fmt_comma(m['backlog_lines'])} lines</div></div>
</div>

<div class="cg">
  <div class="cb">
    <div class="ct">YTD Revenue vs Baseline Plan ($M)</div>
    <div class="cs">Cumulative {m['cur_year']} &middot; {rev_vs_plan}</div>
    <div class="leg"><span><span class="lsq" style="background:#185FA5"></span>Actual revenue</span><span><span class="lsq" style="background:#AAAAAA;border:1px dashed #888"></span>Baseline plan</span></div>
    <div style="position:relative;width:100%;height:240px"><canvas id="revChart" role="img" aria-label="YTD Revenue vs Baseline Plan"></canvas></div>
  </div>
  <div class="cb">
    <div class="ct">PGM % vs Baseline Plan</div>
    <div class="cs">{pgm_vs_plan}</div>
    <div class="leg"><span><span class="lsq" style="background:#0F7A58"></span>Actual PGM %</span><span><span class="lsq" style="background:#AAAAAA;border:1px dashed #888"></span>Baseline plan</span></div>
    <div style="position:relative;width:100%;height:240px"><canvas id="pgmChart" role="img" aria-label="PGM % vs Baseline Plan"></canvas></div>
  </div>
</div>

<hr class="divider">
<div class="sl">Past Due Analysis &mdash; PO Orders (Planned Commitments)</div>
<div class="kpi-grid">
  <div class="kpi"><div class="kpi-label">PO Past Due Lines</div><div class="kpi-val r">{fmt_comma(m['po_pd_lines'])}</div><div class="kpi-sub dn">of {fmt_comma(m['po_total_lines'])} total PO lines &middot; {po_pd_pct}%</div></div>
  <div class="kpi"><div class="kpi-label">PO Past Due $</div><div class="kpi-val r">{fmt_m(m['po_pd_dollars'])}</div><div class="kpi-sub dn">{po_dollar_pct}% of total past due dollars</div></div>
  <div class="kpi"><div class="kpi-label">#1 Function Owner</div><div class="kpi-val r">{m['top_func']}</div><div class="kpi-sub dn">{fmt_comma(m['top_func_l'])} lines &middot; {fmt_m(m['top_func_d'])} &middot; {m['top_func_pct']}% of PO past due $</div></div>
  <div class="kpi"><div class="kpi-label">#1 Customer at Risk</div><div class="kpi-val r">{m['top_site']}</div><div class="kpi-sub dn">{fmt_comma(m['top_site_l'])} lines &middot; {fmt_m(m['top_site_d'])} &middot; {m['top_site_pct']}% of PO past due $</div></div>
</div>

<div class="cg">
  <div class="cb">
    <div class="ct">PO Past Due by Function &mdash; Lines &amp; Dollars</div>
    <div class="cs">ActionBy ownership &middot; who holds the execution miss</div>
    <div class="leg"><span><span class="lsq" style="background:#4A3FBF"></span>$ impact</span><span><span class="lsq" style="background:rgba(61,170,132,0.85)"></span>line count</span></div>
    <div style="position:relative;width:100%;height:280px"><canvas id="poFuncChart" role="img" aria-label="PO Past Due by Function"></canvas></div>
  </div>
  <div class="cb">
    <div class="ct">PO Past Due by Site &mdash; Top 10</div>
    <div class="cs">Financial impact per Honeywell site</div>
    <div style="position:relative;width:100%;height:280px"><canvas id="poSiteChart" role="img" aria-label="PO Past Due by Site"></canvas></div>
  </div>
</div>

<hr class="divider">
<div class="sl">Past Due Analysis &mdash; Binstock Orders (Ad-Hoc Replenishment)</div>
<div class="kpi-grid-2">
  <div class="kpi"><div class="kpi-label">Binstock Past Due Lines</div><div class="kpi-val a">{fmt_comma(m['bin_pd_lines'])}</div><div class="kpi-sub">of {fmt_comma(m['bin_total_lines'])} total Binstock lines &middot; {bin_pd_pct}%</div></div>
  <div class="kpi"><div class="kpi-label">Binstock Past Due $</div><div class="kpi-val a">{fmt_m(m['bin_pd_dollars'])}</div><div class="kpi-sub">{bin_dollar_pct}% of total past due dollars</div></div>
</div>

<div class="cg">
  <div class="cb">
    <div class="ct">Binstock Past Due by Function &mdash; Lines &amp; Dollars</div>
    <div class="cs">ActionBy ownership of bin replenishment misses</div>
    <div class="leg"><span><span class="lsq" style="background:#4A3FBF"></span>$ impact</span><span><span class="lsq" style="background:rgba(61,170,132,0.85)"></span>line count</span></div>
    <div style="position:relative;width:100%;height:260px"><canvas id="binFuncChart" role="img" aria-label="Binstock Past Due by Function"></canvas></div>
  </div>
  <div class="cb">
    <div class="ct">Binstock Past Due by Site &mdash; Top 10</div>
    <div class="cs">Line volume per site &middot; full bin map needed for rate analysis</div>
    <div style="position:relative;width:100%;height:260px"><canvas id="binSiteChart" role="img" aria-label="Binstock Past Due by Site"></canvas></div>
  </div>
</div>

<div class="fn"><strong>Note on Binstock:</strong> Binstock orders are ad-hoc scan-triggered replenishments under a 2-bin system, not planned commitments. Past due lines represent replenishment cycles that have not yet completed. A full bin map is needed to calculate true stockout rates by site. This analysis will be enhanced as that data becomes available.</div>
<div class="rt">Last refreshed: {m['as_of']} &middot; Source: Working Honeywell KPI Dashboard.xlsx</div>

</div>

<script>
Chart.register(ChartDataLabels);
const gc='rgba(0,0,0,0.06)',tc='#5F5E5A';

new Chart(document.getElementById('revChart'),{{type:'line',data:{{labels:{js_str_arr(m['trend_dates'])},datasets:[
  {{label:'Actual',data:{js_arr(m['trend_rev'])},borderColor:'#185FA5',backgroundColor:'rgba(24,95,165,0.07)',fill:true,tension:0.3,pointRadius:3,borderWidth:2.5,
    datalabels:{{display:ctx=>[0,{len(m['trend_rev'])-1}].includes(ctx.dataIndex),anchor:'top',align:'top',color:'#185FA5',font:{{size:10,weight:'500'}},formatter:v=>'$'+v+'M'}}}},
  {{label:'Baseline',data:{js_arr(m['trend_brev'])},borderColor:'#AAAAAA',borderDash:[5,3],tension:0.3,pointRadius:0,borderWidth:1.5,spanGaps:true,datalabels:{{display:false}}}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},datalabels:{{}}}},
  scales:{{x:{{ticks:{{color:tc,font:{{size:9}},maxRotation:45,autoSkip:true,maxTicksLimit:10}},grid:{{color:gc}}}},
           y:{{ticks:{{color:tc,font:{{size:10}},callback:v=>'$'+v+'M'}},grid:{{color:gc}}}}}}  }}}});

new Chart(document.getElementById('pgmChart'),{{type:'line',data:{{labels:{js_str_arr(m['trend_dates'])},datasets:[
  {{label:'PGM',data:{js_arr(m['trend_pgm'])},borderColor:'#0F7A58',backgroundColor:'rgba(15,122,88,0.07)',fill:true,tension:0.3,pointRadius:3,borderWidth:2.5,
    datalabels:{{display:ctx=>[0,1,{len(m['trend_pgm'])-1}].includes(ctx.dataIndex),anchor:'top',align:'top',color:'#0F7A58',font:{{size:10,weight:'500'}},formatter:v=>v+'%'}}}},
  {{label:'Baseline',data:{js_arr(m['trend_bpgm'])},borderColor:'#AAAAAA',borderDash:[5,3],tension:0.3,pointRadius:0,borderWidth:1.5,spanGaps:true,datalabels:{{display:false}}}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},datalabels:{{}}}},
  scales:{{x:{{ticks:{{color:tc,font:{{size:9}},maxRotation:45,autoSkip:true,maxTicksLimit:10}},grid:{{color:gc}}}},
           y:{{min:20,max:35,ticks:{{color:tc,font:{{size:10}},callback:v=>v+'%'}},grid:{{color:gc}}}}  }}}}}});

new Chart(document.getElementById('poFuncChart'),{{type:'bar',data:{{labels:{po_func_labels},datasets:[
  {{label:'$',data:{po_func_d_vals},backgroundColor:'#4A3FBF',xAxisID:'xD',
    datalabels:{{display:ctx=>ctx.dataIndex<4,anchor:'end',align:'right',color:'#4A3FBF',font:{{size:9,weight:'500'}},formatter:v=>'$'+v+'M'}}}},
  {{label:'Lines',data:{po_func_l_vals},backgroundColor:'rgba(61,170,132,0.80)',xAxisID:'xL',
    datalabels:{{display:ctx=>ctx.dataIndex<4,anchor:'end',align:'right',color:'#0F6E56',font:{{size:9}},formatter:v=>v+' lines'}}}}
]}},options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},datalabels:{{}}}},
  scales:{{y:{{ticks:{{color:tc,font:{{size:11}}}},grid:{{color:gc}}}},
           xD:{{position:'bottom',ticks:{{color:'#4A3FBF',font:{{size:9}},callback:v=>'$'+v+'M'}},grid:{{color:gc}}}},
           xL:{{position:'top',ticks:{{color:'#0F6E56',font:{{size:9}},callback:v=>v}},grid:{{display:false}}}}  }}}}}});

const psc={po_site_colors};
new Chart(document.getElementById('poSiteChart'),{{type:'bar',data:{{labels:{po_site_labels},datasets:[
  {{label:'PO Past Due $',data:{po_site_d_vals},backgroundColor:psc,
    datalabels:{{anchor:'end',align:'right',color:ctx=>psc[ctx.dataIndex],font:{{size:9,weight:'500'}},formatter:v=>'$'+Math.round(v*1000)+'K'}}}}
]}},options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},datalabels:{{}}}},
  scales:{{y:{{ticks:{{color:tc,font:{{size:10}}}},grid:{{color:gc}}}},
           x:{{ticks:{{color:tc,font:{{size:9}},callback:v=>'$'+v+'M'}},grid:{{color:gc}}}}  }}}}}});

new Chart(document.getElementById('binFuncChart'),{{type:'bar',data:{{labels:{bin_func_labels},datasets:[
  {{label:'$',data:{bin_func_d_vals},backgroundColor:'#4A3FBF',xAxisID:'xD',
    datalabels:{{anchor:'end',align:'right',color:'#4A3FBF',font:{{size:9,weight:'500'}},formatter:v=>'$'+v+'K'}}}},
  {{label:'Lines',data:{bin_func_l_vals},backgroundColor:'rgba(61,170,132,0.80)',xAxisID:'xL',
    datalabels:{{anchor:'end',align:'right',color:'#0F6E56',font:{{size:9}},formatter:v=>v+' lines'}}}}
]}},options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},datalabels:{{}}}},
  scales:{{y:{{ticks:{{color:tc,font:{{size:11}}}},grid:{{color:gc}}}},
           xD:{{position:'bottom',ticks:{{color:'#4A3FBF',font:{{size:9}},callback:v=>'$'+v+'K'}},grid:{{color:gc}}}},
           xL:{{position:'top',ticks:{{color:'#0F6E56',font:{{size:9}},callback:v=>v}},grid:{{display:false}}}}  }}}}}});

new Chart(document.getElementById('binSiteChart'),{{type:'bar',data:{{labels:{bin_site_labels},datasets:[
  {{label:'Lines',data:{bin_site_l_vals},backgroundColor:'rgba(186,117,23,0.78)',
    datalabels:{{anchor:'end',align:'right',color:'#854F0B',font:{{size:9,weight:'500'}},formatter:v=>v+' lines'}}}}
]}},options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},datalabels:{{}}}},
  scales:{{y:{{ticks:{{color:tc,font:{{size:10}}}},grid:{{color:gc}}}},
           x:{{ticks:{{color:tc,font:{{size:9}}}},grid:{{color:gc}}}}  }}}}}});
</script>
</body>
</html>"""


# ── Entry point ───────────────────────────────────────────────
if __name__ == "__main__":
    print("\n  Honeywell Dashboard Refresh")
    print("  " + "=" * 42)
    saw, hist = load_data()
    metrics   = compute(saw, hist)
    html      = build_html(metrics)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\n  Done! File written to:")
    print(f"  {OUTPUT_FILE}")
    print(f"\n  Next: upload index.html to your GitHub repo.")
    print("  " + "=" * 42 + "\n")