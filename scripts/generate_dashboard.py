"""
Dashboard Generator - Giaco SDC
Reads "Queries Meta ads" and "leadscoring" from Google Sheets,
embeds raw data as JSON in the HTML so the user can filter by any date range
interactively in the browser.
"""

import json, os, sys
from datetime import date, datetime, timedelta
from collections import defaultdict

import gspread
from google.oauth2.service_account import Credentials

SPREADSHEET_ID = "1Uep5K2-fJqBxNX-j8rDa-7V-Q2-iCRlz9S6KzR37VsQ"
CPL_META = 50.0

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# ─── auth ───────────────────────────────────────────────────────────────────

def get_client():
    raw = os.environ.get("GOOGLE_CREDENTIALS")
    if not raw:
        raise ValueError("GOOGLE_CREDENTIALS env var not set")
    creds = Credentials.from_service_account_info(json.loads(raw), scopes=SCOPES)
    return gspread.authorize(creds)

# ─── helpers ────────────────────────────────────────────────────────────────

def parse_brl(v):
    if not v or str(v).strip() in ("", "Sem Dados", "#DIV/0!", "null", "#REF!"):
        return 0.0
    s = str(v).replace("R$", "").replace("\xa0", "").strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def parse_num(v):
    if not v or str(v).strip() in ("", "Sem Dados", "#DIV/0!", "null", "#REF!"):
        return 0.0
    s = str(v).replace(".", "").replace(",", ".").replace("%", "").strip()
    try:
        return float(s)
    except Exception:
        return 0.0

def parse_date(v):
    v = str(v).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(v[:10], fmt).date()
        except Exception:
            pass
    return None

def find_col(header, *names):
    lnames = [n.lower() for n in names]
    for i, h in enumerate(header):
        if h.strip().lower() in lnames:
            return i
    return -1

def open_worksheet(gc, *name_candidates):
    ss = gc.open_by_key(SPREADSHEET_ID)
    available = [ws.title for ws in ss.worksheets()]
    print(f"Available worksheets: {available}", flush=True)
    for name in name_candidates:
        try:
            return ss.worksheet(name)
        except Exception:
            pass
    for name in name_candidates:
        for ws in ss.worksheets():
            if ws.title.lower().strip() == name.lower().strip():
                return ws
    for name in name_candidates:
        for ws in ss.worksheets():
            if name.lower() in ws.title.lower() or ws.title.lower() in name.lower():
                print(f"Partial match: '{ws.title}' for '{name}'", flush=True)
                return ws
    raise RuntimeError(f"Worksheet not found. Tried: {name_candidates}. Available: {available}")

# ─── readers ────────────────────────────────────────────────────────────────

def read_meta_ads(gc, d_from, d_to):
    ws = open_worksheet(gc,
        "Queries Meta ads", "queries meta ads", "Queries Meta Ads",
        "Query Meta Ads", "Meta Ads", "queries_meta_ads")
    rows = ws.get_all_values()
    if len(rows) < 2:
        return []
    hdr = rows[0]
    c = {
        "day":   find_col(hdr, "Day", "Data", "Date"),
        "camp":  find_col(hdr, "Campaign Name", "campaign name"),
        "adset": find_col(hdr, "Ad Set Name", "Ad set name", "Adset Name"),
        "ad":    find_col(hdr, "Ad Name", "Ad name"),
        "spend": find_col(hdr, "Amount Spent", "Gasto", "Spend", "Cost In Local Currency (Spend)"),
        "imp":   find_col(hdr, "Impressions"),
        "clk":   find_col(hdr, "Link Clicks", "Clicks"),
        "leads": find_col(hdr, "Leads"),
        "reach": find_col(hdr, "Reach"),
    }
    def g(row, key):
        idx = c[key]
        return row[idx] if 0 <= idx < len(row) else ""

    out = []
    for row in rows[1:]:
        d = parse_date(g(row, "day"))
        if d is None or not (d_from <= d <= d_to):
            continue
        out.append({
            "date":        str(d),
            "campaign":    g(row, "camp").strip(),
            "adset":       g(row, "adset").strip(),
            "ad":          g(row, "ad").strip(),
            "spend":       parse_brl(g(row, "spend")),
            "impressions": parse_num(g(row, "imp")),
            "clicks":      parse_num(g(row, "clk")),
            "leads":       parse_num(g(row, "leads")),
            "reach":       parse_num(g(row, "reach")),
        })
    return out


def read_leadscoring(gc, d_from, d_to):
    ws = open_worksheet(gc,
        "leadscoring", "Leadscoring", "Lead Scoring", "lead scoring",
        "lead_scoring", "LeadScoring")
    rows = ws.get_all_values()
    if len(rows) < 2:
        return []
    hdr = rows[0]
    c = {
        "nota":   find_col(hdr, "Nota"),
        "date":   find_col(hdr, "DATA", "Data", "created_time", "date"),
        "camp":   find_col(hdr, "campaign_name", "CAMPANHA", "Campaign Name"),
        "adset":  find_col(hdr, "adset_name", "CONJUNTO", "Ad Set Name"),
        "ad":     find_col(hdr, "ad_name", "ANÚNCIO", "Ad Name"),
    }
    def g(row, key):
        idx = c[key]
        return row[idx] if 0 <= idx < len(row) else ""

    out = []
    for row in rows[1:]:
        d = parse_date(g(row, "date"))
        if d is None or not (d_from <= d <= d_to):
            continue
        nota = g(row, "nota").strip().upper()
        if nota not in ("A", "B", "C", "D", "E"):
            nota = ""
        out.append({
            "date":     str(d),
            "nota":     nota,
            "campaign": g(row, "camp").strip(),
            "adset":    g(row, "adset").strip(),
            "ad":       g(row, "ad").strip(),
        })
    return out

# ─── HTML ───────────────────────────────────────────────────────────────────

def generate_html(meta_rows, ls_rows, d_from, d_to, generated_at):
    meta_json = json.dumps(meta_rows, ensure_ascii=False)
    ls_json   = json.dumps(ls_rows,   ensure_ascii=False)
    d_from_str = str(d_from)
    d_to_str   = str(d_to)
    cpl_meta   = CPL_META

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Dashboard SDC — Giaco</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --or:#f97316;--am:#fbbf24;--dark:#0f172a;--card:#1e293b;
  --text:#e2e8f0;--muted:#94a3b8;--border:#334155;
}}
body{{font-family:'Segoe UI',sans-serif;background:var(--dark);color:var(--text);min-height:100vh}}
header{{background:linear-gradient(135deg,#ea580c,#b45309);padding:16px 28px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px}}
header h1{{font-size:1.4rem;font-weight:700;color:#fff}}
.updated{{font-size:.72rem;color:rgba(255,255,255,.65)}}
.container{{max-width:1700px;margin:0 auto;padding:20px 16px}}

/* Date filter */
.date-bar{{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:14px 20px;display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:22px}}
.date-bar label{{font-size:.82rem;color:var(--muted)}}
.date-bar input[type=date]{{background:#0f172a;border:1px solid var(--border);border-radius:7px;color:var(--text);padding:6px 10px;font-size:.85rem}}
.date-bar input[type=date]:focus{{outline:none;border-color:var(--or)}}
.btn{{background:var(--or);color:#fff;border:none;border-radius:7px;padding:7px 18px;font-size:.85rem;font-weight:600;cursor:pointer}}
.btn:hover{{background:#ea580c}}
.btn-outline{{background:transparent;color:var(--muted);border:1px solid var(--border);border-radius:7px;padding:7px 14px;font-size:.82rem;cursor:pointer}}
.btn-outline:hover{{border-color:var(--or);color:var(--or)}}
.period-label{{font-size:.8rem;color:var(--am);font-weight:600}}

.sec-title{{font-size:1rem;font-weight:700;color:var(--am);margin:24px 0 10px;text-transform:uppercase;letter-spacing:.06em;display:flex;align-items:center;gap:8px}}
.sec-title::after{{content:'';flex:1;height:1px;background:var(--border)}}

/* Funnel cards */
.funnel-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:10px}}
.fcard{{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:14px 16px}}
.fcard .lbl{{font-size:.7rem;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px}}
.fcard .val{{font-size:1.55rem;font-weight:700;color:#fff;line-height:1.1}}
.fcard .sub{{font-size:.78rem;color:var(--muted);margin-top:4px}}
.fcard .sub b{{color:var(--am)}}
.fcard.hi{{border-color:var(--or);background:linear-gradient(135deg,#1e1e3a,#2d1b00)}}
.fcard.hi .val{{color:var(--am)}}
.gauge-bar{{height:8px;background:#334155;border-radius:5px;overflow:hidden;margin-top:8px}}
.gauge-fill{{height:100%;border-radius:5px;transition:width .4s ease}}
.gauge-label{{font-size:.7rem;color:var(--muted);margin-top:3px}}

/* Score chips */
.chips{{display:flex;gap:8px;flex-wrap:wrap}}
.chip{{border-radius:8px;padding:9px 16px;font-size:.86rem;font-weight:600}}
.ca{{background:#064e3b;color:#6ee7b7;border:1px solid #065f46}}
.cb{{background:#1e3a8a;color:#93c5fd;border:1px solid #1d4ed8}}
.cc{{background:#78350f;color:#fcd34d;border:1px solid #92400e}}
.cd{{background:#3b0764;color:#d8b4fe;border:1px solid #6b21a8}}

/* Chart */
.chart-card{{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:18px}}

/* Table */
.table-wrap{{overflow-x:auto;border-radius:12px;border:1px solid var(--border)}}
#bt{{width:100%;border-collapse:collapse;font-size:.79rem;white-space:nowrap}}
#bt thead tr{{background:var(--or);color:#fff}}
#bt th{{padding:9px 11px;text-align:right;font-weight:600}}
#bt th:first-child{{text-align:left;min-width:260px}}
#bt td{{padding:7px 11px;text-align:right;border-bottom:1px solid #1e293b}}
#bt td:first-child{{text-align:left}}
.rc{{background:#162032;font-weight:600}} .rc td:first-child{{color:var(--am)}}
.ra{{background:var(--card)}} .ra td:first-child{{color:#cbd5e1}}
.rd{{background:#0d1b2a;font-size:.75rem}} .rd td:first-child{{color:var(--muted)}}
.rc.hide,.ra.hide,.rd.hide{{display:none}}
.i1{{padding-left:22px!important}} .i2{{padding-left:44px!important}}
.nc{{display:flex;align-items:center;gap:5px}}
.tg{{background:transparent;border:1px solid var(--border);border-radius:3px;color:var(--muted);cursor:pointer;font-size:.62rem;padding:2px 4px;flex-shrink:0}}
.tg:hover{{background:var(--border)}}
</style>
</head>
<body>

<header>
  <div><h1>📊 Dashboard SDC — Giaco</h1></div>
  <div class="updated">Gerado: {generated_at}</div>
</header>

<div class="container">

  <!-- Date filter -->
  <div class="date-bar">
    <label>De</label>
    <input type="date" id="d_from" value="{d_from_str}">
    <label>Até</label>
    <input type="date" id="d_to"   value="{d_to_str}">
    <button class="btn" onclick="applyFilter()">Aplicar</button>
    <button class="btn-outline" onclick="setCurrentMonth()">Mês atual</button>
    <button class="btn-outline" onclick="setLast7()">Últimos 7 dias</button>
    <button class="btn-outline" onclick="setLast30()">Últimos 30 dias</button>
    <span class="period-label" id="period-label"></span>
  </div>

  <div class="sec-title">Funil de Captura</div>
  <div class="funnel-grid" id="funnel-grid">
    <div class="fcard hi"><div class="lbl">Valor Gasto</div><div class="val" id="f-spend">—</div></div>
    <div class="fcard"><div class="lbl">Impressões</div><div class="val" id="f-imp">—</div><div class="sub" id="f-cpm-freq">—</div></div>
    <div class="fcard"><div class="lbl">Alcance</div><div class="val" id="f-reach">—</div></div>
    <div class="fcard"><div class="lbl">Link Clicks</div><div class="val" id="f-clicks">—</div><div class="sub" id="f-ctr-cpc">—</div></div>
    <div class="fcard"><div class="lbl">Conv. Formulário</div><div class="val" id="f-conv">—</div><div class="sub">Leads / Clicks</div></div>
    <div class="fcard hi">
      <div class="lbl">Leads Totais</div><div class="val" id="f-leads">—</div>
      <div class="sub" id="f-cpl-sub">—</div>
      <div class="gauge-bar"><div class="gauge-fill" id="gauge-fill" style="width:0%;background:#4ade80"></div></div>
      <div class="gauge-label" id="gauge-label">—</div>
    </div>
    <div class="fcard"><div class="lbl">Tx Lead-A</div><div class="val" id="f-txa">—</div><div class="sub" id="f-cpla-sub">—</div></div>
  </div>

  <div class="sec-title" style="margin-top:22px">Lead Scoring</div>
  <div class="chips">
    <div class="chip ca">🟢 Lead A &nbsp;<strong id="c-a">0</strong></div>
    <div class="chip cb">🔵 Lead B &nbsp;<strong id="c-b">0</strong></div>
    <div class="chip cc">🟡 Lead C &nbsp;<strong id="c-c">0</strong></div>
    <div class="chip cd">🟣 Lead D &nbsp;<strong id="c-d">0</strong></div>
  </div>

  <div class="sec-title" style="margin-top:22px">Evolução Diária</div>
  <div class="chart-card"><canvas id="trendChart" height="75"></canvas></div>

  <div class="sec-title" style="margin-top:22px">Meta Ads — Campanha / Conjunto / Anúncio</div>
  <div class="table-wrap"><table id="bt">
    <thead><tr>
      <th>Campanha / Conjunto / Anúncio</th>
      <th>Gasto</th><th>Leads</th><th>CPL</th>
      <th>CPL-A</th><th>Tx-A</th>
      <th>Leads-A</th><th>Leads-B</th><th>Leads-C</th><th>Leads-D</th>
      <th>CPM</th><th>CTR</th><th>CPC</th>
    </tr></thead>
    <tbody id="tbl-body"></tbody>
  </table></div>

</div><!-- /container -->

<script>
// ── Raw data embedded by Python ──────────────────────────────────────────────
const RAW_META = {meta_json};
const RAW_LS   = {ls_json};
const CPL_META = {cpl_meta};

// ── Formatters ───────────────────────────────────────────────────────────────
const brl  = v => v === 0 ? 'R$ 0,00' : 'R$ ' + v.toFixed(2).replace('.', ',').replace(/\\B(?=(\\d{{3}})+(?!\\d))/g, '.');
const num  = v => Math.round(v).toLocaleString('pt-BR');
const pct  = v => v.toFixed(2).replace('.', ',') + '%';
const sdiv = (a,b) => b ? a/b : 0;

// ── Filter & aggregate ───────────────────────────────────────────────────────
let trendChart = null;

function applyFilter() {{
  const from = document.getElementById('d_from').value;
  const to   = document.getElementById('d_to').value;
  if (!from || !to) return;

  const meta = RAW_META.filter(r => r.date >= from && r.date <= to);
  const ls   = RAW_LS.filter(r => r.date >= from && r.date <= to);

  document.getElementById('period-label').textContent =
    from.split('-').reverse().join('/') + ' — ' + to.split('-').reverse().join('/');

  renderFunnel(meta, ls);
  renderChips(ls);
  renderTrend(meta, ls, from, to);
  renderTable(meta, ls);
}}

function renderFunnel(meta, ls) {{
  const spend = meta.reduce((s,r)=>s+r.spend,0);
  const imp   = meta.reduce((s,r)=>s+r.impressions,0);
  const clk   = meta.reduce((s,r)=>s+r.clicks,0);
  const leads = meta.reduce((s,r)=>s+r.leads,0);
  const reach = meta.reduce((s,r)=>s+r.reach,0);

  const la = ls.filter(r=>r.nota==='A').length;
  const lb = ls.filter(r=>r.nota==='B').length;
  const lc = ls.filter(r=>r.nota==='C').length;
  const ld = ls.filter(r=>r.nota==='D').length;

  const cpl  = sdiv(spend,leads);
  const cpla = sdiv(spend,la);
  const cpm  = sdiv(spend,imp)*1000;
  const cpc  = sdiv(spend,clk);
  const ctr  = sdiv(clk,imp)*100;
  const conv = sdiv(leads,clk)*100;
  const freq = sdiv(imp,reach);
  const txa  = sdiv(la,leads)*100;
  const cplPct = cpl>0 ? Math.min(sdiv(CPL_META,cpl)*100,100) : 0;
  const gaugeColor = cplPct>=80?'#4ade80':cplPct>=50?'#fb923c':'#f87171';

  document.getElementById('f-spend').textContent  = brl(spend);
  document.getElementById('f-imp').textContent    = num(imp);
  document.getElementById('f-cpm-freq').innerHTML = `CPM <b>${{brl(cpm)}}</b> &nbsp;|&nbsp; Freq. <b>${{freq.toFixed(2).replace('.',',')}}</b>`;
  document.getElementById('f-reach').textContent  = num(reach);
  document.getElementById('f-clicks').textContent = num(clk);
  document.getElementById('f-ctr-cpc').innerHTML  = `CTR <b>${{pct(ctr)}}</b> &nbsp;|&nbsp; CPC <b>${{brl(cpc)}}</b>`;
  document.getElementById('f-conv').textContent   = pct(conv);
  document.getElementById('f-leads').textContent  = num(leads);
  document.getElementById('f-cpl-sub').innerHTML  = `CPL <b>${{brl(cpl)}}</b>`;
  document.getElementById('f-txa').textContent    = pct(txa);
  document.getElementById('f-cpla-sub').innerHTML = `CPL-A <b>${{brl(cpla)}}</b>`;
  document.getElementById('gauge-fill').style.width = cplPct.toFixed(1)+'%';
  document.getElementById('gauge-fill').style.background = gaugeColor;
  document.getElementById('gauge-label').textContent = cplPct.toFixed(1)+'% da meta (R$ '+CPL_META+',00)';
}}

function renderChips(ls) {{
  document.getElementById('c-a').textContent = ls.filter(r=>r.nota==='A').length;
  document.getElementById('c-b').textContent = ls.filter(r=>r.nota==='B').length;
  document.getElementById('c-c').textContent = ls.filter(r=>r.nota==='C').length;
  document.getElementById('c-d').textContent = ls.filter(r=>r.nota==='D').length;
}}

function renderTrend(meta, ls, from, to) {{
  // Build day list
  const days = [];
  let d = new Date(from + 'T00:00:00');
  const end = new Date(to + 'T00:00:00');
  while (d <= end) {{
    days.push(d.toISOString().slice(0,10));
    d.setDate(d.getDate()+1);
  }}
  const labels  = days.map(d => d.slice(5).replace('-','/'));
  const spends  = days.map(d => meta.filter(r=>r.date===d).reduce((s,r)=>s+r.spend,0));
  const leadsArr= days.map(d => meta.filter(r=>r.date===d).reduce((s,r)=>s+r.leads,0));
  const laArr   = days.map(d => ls.filter(r=>r.date===d&&r.nota==='A').length);

  if (trendChart) trendChart.destroy();
  trendChart = new Chart(document.getElementById('trendChart'), {{
    type:'bar',
    data:{{
      labels,
      datasets:[
        {{label:'Gasto (R$)',data:spends,backgroundColor:'rgba(249,115,22,.6)',borderColor:'#f97316',borderWidth:1,yAxisID:'y'}},
        {{label:'Leads',data:leadsArr,type:'line',borderColor:'#fbbf24',backgroundColor:'rgba(251,191,36,.12)',tension:.35,fill:true,pointRadius:3,yAxisID:'y1'}},
        {{label:'Leads A',data:laArr,type:'line',borderColor:'#4ade80',tension:.35,fill:false,pointRadius:3,yAxisID:'y1'}},
      ]
    }},
    options:{{
      responsive:true,
      interaction:{{mode:'index',intersect:false}},
      scales:{{
        x:{{ticks:{{color:'#94a3b8',maxRotation:45}},grid:{{color:'#1e293b'}}}},
        y:{{position:'left',ticks:{{color:'#94a3b8',callback:v=>'R$'+v.toLocaleString('pt-BR')}},grid:{{color:'#1e293b'}}}},
        y1:{{position:'right',ticks:{{color:'#94a3b8'}},grid:{{drawOnChartArea:false}}}},
      }},
      plugins:{{legend:{{labels:{{color:'#e2e8f0'}}}},tooltip:{{bodyColor:'#e2e8f0',titleColor:'#fbbf24'}}}},
    }}
  }});
}}

function renderTable(meta, ls) {{
  // Aggregate by campaign→adset→ad
  const camps = {{}};
  meta.forEach(r => {{
    const kc = r.campaign, ka = kc+'|||'+r.adset, kd = ka+'|||'+r.ad;
    [kc,ka,kd].forEach((k,i) => {{
      if (!camps[k]) camps[k] = {{level:i,name:i===0?r.campaign:i===1?r.adset:r.ad,
        parent:i===0?null:i===1?kc:ka,
        spend:0,imp:0,clk:0,leads:0,reach:0,la:0,lb:0,lc:0,ld:0}};
      camps[k].spend+=r.spend; camps[k].imp+=r.impressions;
      camps[k].clk+=r.clicks; camps[k].leads+=r.leads; camps[k].reach+=r.reach;
    }});
  }});
  ls.forEach(r => {{
    const kc = r.campaign, ka = kc+'|||'+r.adset, kd = ka+'|||'+r.ad;
    [kc,ka,kd].forEach(k => {{
      // fuzzy match
      let key = k;
      if (!camps[key]) {{
        key = Object.keys(camps).find(ck => ck.startsWith(r.campaign) || r.campaign.includes(ck.split('|||')[0]));
      }}
      if (!key || !camps[key]) return;
      if (r.nota==='A') camps[key].la++;
      else if (r.nota==='B') camps[key].lb++;
      else if (r.nota==='C') camps[key].lc++;
      else if (r.nota==='D') camps[key].ld++;
    }});
  }});

  // Sort: camps first, then adsets, then ads
  const order = Object.keys(camps).sort((a,b) => {{
    const la = camps[a].level, lb2 = camps[b].level;
    if (la !== lb2) return la - lb2;
    return a.localeCompare(b);
  }});

  const shorten = (s,n) => s.length>n ? s.slice(0,n)+'…' : s;
  const row = (k, m) => {{
    const s=m.spend,l=m.leads,la=m.la;
    const cpl=sdiv(s,l),cpla=sdiv(s,la),txa=sdiv(la,l)*100;
    const cpm=sdiv(s,m.imp)*1000,cpc=sdiv(s,m.clk),ctr=sdiv(m.clk,m.imp)*100;
    const cls = m.level===0?'rc':m.level===1?'ra':'rd';
    const indent = m.level===0?'':m.level===1?' i1':' i2';
    const toggle = m.level<2 ? `<button class="tg" onclick="tog(this,'${{k}}')">▼</button>` : '';
    return `<tr class="${{cls}}" data-key="${{k}}" data-parent="${{m.parent||''}}">
      <td class="nc${{indent}}">${{toggle}}<span title="${{m.name}}">${{shorten(m.name,m.level===0?55:m.level===1?48:42)}}</span></td>
      <td>${{brl(s)}}</td><td>${{num(l)}}</td><td>${{brl(cpl)}}</td>
      <td>${{la>0?brl(cpla):'—'}}</td><td>${{pct(txa)}}</td>
      <td>${{la}}</td><td>${{m.lb}}</td><td>${{lc}}</td><td>${{m.ld}}</td>
      <td>${{brl(cpm)}}</td><td>${{pct(ctr)}}</td><td>${{brl(cpc)}}</td>
    </tr>`;
  }};

  document.getElementById('tbl-body').innerHTML = order.map(k=>row(k,camps[k])).join('');

  // hide adsets & ads by default, show on click
  document.querySelectorAll('#bt .ra,#bt .rd').forEach(r=>r.classList.add('hide'));
}}

function tog(btn, key) {{
  const open = btn.textContent==='▼';
  const rows = document.querySelectorAll(`#bt tr[data-parent="${{key}}"]`);
  rows.forEach(r => {{
    r.classList.toggle('hide', open);
    // collapse grandchildren when closing
    if (open) {{
      const k2 = r.dataset.key;
      if (k2) document.querySelectorAll(`#bt tr[data-parent="${{k2}}"]`).forEach(rr=>rr.classList.add('hide'));
      const b2 = r.querySelector('.tg');
      if (b2) b2.textContent='▼';
    }}
  }});
  btn.textContent = open ? '▶' : '▼';
}}

// ── Date shortcuts ────────────────────────────────────────────────────────────
function setCurrentMonth() {{
  const t = new Date();
  document.getElementById('d_from').value = `${{t.getFullYear()}}-${{String(t.getMonth()+1).padStart(2,'0')}}-01`;
  document.getElementById('d_to').value   = t.toISOString().slice(0,10);
  applyFilter();
}}
function setLast7() {{
  const t=new Date(), f=new Date(t); f.setDate(f.getDate()-6);
  document.getElementById('d_from').value=f.toISOString().slice(0,10);
  document.getElementById('d_to').value=t.toISOString().slice(0,10);
  applyFilter();
}}
function setLast30() {{
  const t=new Date(), f=new Date(t); f.setDate(f.getDate()-29);
  document.getElementById('d_from').value=f.toISOString().slice(0,10);
  document.getElementById('d_to').value=t.toISOString().slice(0,10);
  applyFilter();
}}

// ── Init ─────────────────────────────────────────────────────────────────────
applyFilter();
</script>
</body>
</html>"""


# ─── main ────────────────────────────────────────────────────────────────────

def main():
    today = date.today()
    # Load a full month (or custom range) — user filters interactively in browser
    if len(sys.argv) >= 3:
        d_from = datetime.strptime(sys.argv[1], "%Y-%m-%d").date()
        d_to   = datetime.strptime(sys.argv[2], "%Y-%m-%d").date()
    else:
        d_from = date(today.year, today.month, 1)
        d_to   = today

    print(f"Period: {d_from} → {d_to}", flush=True)
    gc = get_client()

    print("Reading Queries Meta Ads…", flush=True)
    meta_rows = read_meta_ads(gc, d_from, d_to)
    print(f"  {len(meta_rows)} rows", flush=True)

    print("Reading Lead Scoring…", flush=True)
    ls_rows = read_leadscoring(gc, d_from, d_to)
    print(f"  {len(ls_rows)} rows", flush=True)

    generated_at = datetime.utcnow().strftime("%d/%m/%Y %H:%M UTC")
    html = generate_html(meta_rows, ls_rows, d_from, d_to, generated_at)

    out = os.path.join(os.path.dirname(__file__), "..", "docs", "index.html")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Saved: {os.path.abspath(out)}", flush=True)


if __name__ == "__main__":
    main()
