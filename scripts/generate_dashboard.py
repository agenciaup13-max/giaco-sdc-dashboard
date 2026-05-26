"""
Dashboard Generator - Giaco SDC
Reads "Queries Meta ads" and "leadscoring" from Google Sheets,
cross-references both, and generates a static HTML dashboard.
"""

import json, os, sys
from datetime import date, datetime, timedelta
from collections import defaultdict, OrderedDict

import gspread
from google.oauth2.service_account import Credentials

SPREADSHEET_ID = "1Uep5K2-fJqBxNX-j8rDa-7V-Q2-iCRlz9S6KzR37VsQ"
CPL_META = 50.0

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# ─── helpers ────────────────────────────────────────────────────────────────

def get_client():
    raw = os.environ.get("GOOGLE_CREDENTIALS")
    if not raw:
        raise ValueError("GOOGLE_CREDENTIALS env var not set")
    creds = Credentials.from_service_account_info(json.loads(raw), scopes=SCOPES)
    return gspread.authorize(creds)


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


def fmt_brl(v):
    if v == 0:
        return "R$ 0,00"
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return "R$ " + s


def fmt_int(v):
    return f"{int(v):,}".replace(",", ".")


def fmt_pct(v):
    return f"{v:.2f}".replace(".", ",") + "%"


def fmt_float2(v):
    return f"{v:.2f}".replace(".", ",")


def safe_div(a, b):
    return a / b if b else 0.0


def find_col(header, *names):
    """Return index of first matching column name (case-insensitive)."""
    lnames = [n.lower() for n in names]
    for i, h in enumerate(header):
        if h.strip().lower() in lnames:
            return i
    return -1


# ─── Google Sheets readers ───────────────────────────────────────────────────

def open_worksheet(gc, *name_candidates):
    ss = gc.open_by_key(SPREADSHEET_ID)
    for name in name_candidates:
        try:
            return ss.worksheet(name)
        except Exception:
            pass
    raise RuntimeError(f"Worksheet not found. Tried: {name_candidates}")


def read_meta_ads(gc, d_from, d_to):
    ws = open_worksheet(gc,
        "Queries Meta ads", "queries meta ads", "Queries Meta Ads",
        "Query Meta Ads", "Meta Ads")
    rows = ws.get_all_values()
    if len(rows) < 2:
        return []

    hdr = rows[0]
    c = {
        "day":   find_col(hdr, "Day", "Data", "Date"),
        "camp":  find_col(hdr, "Campaign Name", "campaign name"),
        "adset": find_col(hdr, "Ad Set Name", "Ad set name", "Adset Name"),
        "ad":    find_col(hdr, "Ad Name", "Ad name"),
        "spend": find_col(hdr, "Amount Spent", "Gasto", "Spend",
                          "Cost In Local Currency (Spend)"),
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
            "date":        d,
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
        "leadscoring", "Leadscoring", "Lead Scoring", "lead scoring")
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
        "compra": find_col(hdr, "Compra", "compra"),
        "fat":    find_col(hdr, "FAT", "fat"),
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
            "date":     d,
            "nota":     nota,
            "campaign": g(row, "camp").strip(),
            "adset":    g(row, "adset").strip(),
            "ad":       g(row, "ad").strip(),
            "compra":   g(row, "compra").strip().upper() == "SIM",
            "fat":      parse_brl(g(row, "fat")),
        })
    return out


# ─── aggregation ────────────────────────────────────────────────────────────

def make_row():
    return dict(spend=0.0, impressions=0.0, clicks=0.0,
                leads=0.0, reach=0.0,
                leads_a=0, leads_b=0, leads_c=0, leads_d=0,
                compras=0, fat=0.0)


def calc(m):
    s = m["spend"]; l = m["leads"]; imp = m["impressions"]
    clk = m["clicks"]; la = m["leads_a"]
    return {
        **m,
        "cpl":       safe_div(s, l),
        "cpl_a":     safe_div(s, la),
        "tx_a":      safe_div(la, l) * 100,
        "cpm":       safe_div(s, imp) * 1000,
        "cpc":       safe_div(s, clk),
        "ctr":       safe_div(clk, imp) * 100,
        "conv_form": safe_div(l, clk) * 100,
        "roas":      safe_div(m["fat"], s),
    }


def aggregate(meta_rows, ls_rows):
    # ── global
    g = make_row()
    for r in meta_rows:
        for k in ("spend", "impressions", "clicks", "leads", "reach"):
            g[k] += r[k]

    # Frequency via daily reach sum
    daily_reach = defaultdict(float)
    for r in meta_rows:
        daily_reach[r["date"]] += r["reach"]
    g["reach"] = sum(daily_reach.values())
    g["frequency"] = safe_div(g["impressions"], g["reach"]) if g["reach"] else 0

    for r in ls_rows:
        n = r["nota"]
        if n == "A": g["leads_a"] += 1
        elif n == "B": g["leads_b"] += 1
        elif n == "C": g["leads_c"] += 1
        elif n == "D": g["leads_d"] += 1
        if r["compra"]:
            g["compras"] += 1
            g["fat"] += r["fat"]

    global_metrics = calc(g)

    # ── campaign / adset / ad breakdown
    camps  = defaultdict(make_row)
    adsets = defaultdict(make_row)   # key: (camp, adset)
    ads    = defaultdict(make_row)   # key: (camp, adset, ad)

    for r in meta_rows:
        kc = r["campaign"]
        ka = (r["campaign"], r["adset"])
        kd = (r["campaign"], r["adset"], r["ad"])
        for key, store in [(kc, camps), (ka, adsets), (kd, ads)]:
            for k in ("spend", "impressions", "clicks", "leads", "reach"):
                store[key][k] += r[k]

    # Match leadscoring to meta rows by (campaign, adset, ad)
    for r in ls_rows:
        n = r["nota"]
        kc = r["campaign"]
        ka = (r["campaign"], r["adset"])
        kd = (r["campaign"], r["adset"], r["ad"])

        # Fuzzy match campaign if exact key missing
        if kc not in camps:
            match = next((k for k in camps if r["campaign"] in k or k in r["campaign"]), None)
            if not match:
                continue
            kc = match
            ka = (kc, r["adset"])
            kd = (kc, r["adset"], r["ad"])

        def bump(key, store):
            if key not in store:
                return
            if n == "A": store[key]["leads_a"] += 1
            elif n == "B": store[key]["leads_b"] += 1
            elif n == "C": store[key]["leads_c"] += 1
            elif n == "D": store[key]["leads_d"] += 1
            if r["compra"]:
                store[key]["compras"] += 1
                store[key]["fat"] += r["fat"]

        bump(kc, camps)
        bump(ka, adsets)
        bump(kd, ads)

    # Build tree dict
    tree = {}
    for c_name in sorted(camps):
        cm = calc(camps[c_name])
        cm["name"] = c_name
        cm["adsets"] = {}
        for (cc, a_name) in sorted(k for k in adsets if k[0] == c_name):
            am = calc(adsets[(cc, a_name)])
            am["name"] = a_name
            am["ads"] = {}
            for (cc2, aa, ad_name) in sorted(k for k in ads if k[0] == c_name and k[1] == a_name):
                dm = calc(ads[(cc2, aa, ad_name)])
                dm["name"] = ad_name
                am["ads"][ad_name] = dm
            cm["adsets"][a_name] = am
        tree[c_name] = cm

    return global_metrics, tree


def daily_trend(meta_rows, ls_rows, d_from, d_to):
    days = OrderedDict()
    d = d_from
    while d <= d_to:
        days[d] = {"spend": 0.0, "leads": 0.0, "leads_a": 0}
        d += timedelta(days=1)
    for r in meta_rows:
        if r["date"] in days:
            days[r["date"]]["spend"] += r["spend"]
            days[r["date"]]["leads"] += r["leads"]
    for r in ls_rows:
        if r["date"] in days and r["nota"] == "A":
            days[r["date"]]["leads_a"] += 1
    return [
        {"label": d.strftime("%d/%m"), "spend": round(days[d]["spend"], 2),
         "leads": int(days[d]["leads"]), "leads_a": days[d]["leads_a"]}
        for d in days
    ]


# ─── HTML ───────────────────────────────────────────────────────────────────

COLS = [
    ("Campanha / Conjunto / Anúncio", None, None),
    ("Gasto",   "spend",   fmt_brl),
    ("Leads",   "leads",   lambda v: fmt_int(int(v))),
    ("CPL",     "cpl",     fmt_brl),
    ("CPL-A",   "cpl_a",   lambda v: fmt_brl(v) if v > 0 else "—"),
    ("Tx-A",    "tx_a",    fmt_pct),
    ("Leads-A", "leads_a", lambda v: fmt_int(int(v))),
    ("Leads-B", "leads_b", lambda v: fmt_int(int(v))),
    ("Leads-C", "leads_c", lambda v: fmt_int(int(v))),
    ("Leads-D", "leads_d", lambda v: fmt_int(int(v))),
    ("CPM",     "cpm",     fmt_brl),
    ("CTR",     "ctr",     fmt_pct),
    ("CPC",     "cpc",     fmt_brl),
]


def row_tds(m):
    return "".join(
        f"<td>{fn(m.get(k, 0))}</td>"
        for _, k, fn in COLS[1:]
    )


def shorten(s, n):
    return (s[:n] + "…") if len(s) > n else s


def render_tree(tree):
    heads = "".join(f"<th>{label}</th>" for label, _, _ in COLS)
    rows = []
    camp_id = 0
    for c_name, cm in tree.items():
        camp_id += 1
        cid = f"c{camp_id}"
        rows.append(
            f'<tr class="row-camp" data-id="{cid}">'
            f'<td class="name-cell">'
            f'<button class="toggle" onclick="toggleLevel(this,\'camp\',\'{cid}\')">▼</button>'
            f'<span title="{c_name}">{shorten(c_name, 55)}</span></td>'
            f"{row_tds(cm)}</tr>"
        )
        adset_id = 0
        for a_name, am in cm["adsets"].items():
            adset_id += 1
            aid = f"{cid}_a{adset_id}"
            rows.append(
                f'<tr class="row-adset" data-id="{aid}" data-parent="{cid}">'
                f'<td class="name-cell indent-1">'
                f'<button class="toggle" onclick="toggleLevel(this,\'adset\',\'{aid}\')">▼</button>'
                f'<span title="{a_name}">{shorten(a_name, 50)}</span></td>'
                f"{row_tds(am)}</tr>"
            )
            for ad_name, dm in am["ads"].items():
                rows.append(
                    f'<tr class="row-ad" data-parent="{aid}">'
                    f'<td class="name-cell indent-2">'
                    f'<span title="{ad_name}">{shorten(ad_name, 45)}</span></td>'
                    f"{row_tds(dm)}</tr>"
                )
    return (
        '<table id="breakdown-table"><thead><tr>'
        + heads
        + "</tr></thead><tbody>"
        + "".join(rows)
        + "</tbody></table>"
    )


def gauge_bar(pct, label):
    clamped = min(pct, 100)
    color = "#4ade80" if pct >= 80 else "#fb923c" if pct >= 50 else "#f87171"
    return (
        f'<div class="gauge-wrap">'
        f'<div class="gauge-bar"><div class="gauge-fill" style="width:{clamped:.1f}%;background:{color}"></div></div>'
        f'<div class="gauge-label">{label}</div></div>'
    )


def generate_html(gm, tree, trend, d_from, d_to, generated_at):
    period = f"{d_from.strftime('%d/%m/%Y')} — {d_to.strftime('%d/%m/%Y')}"
    cpl_pct = safe_div(CPL_META, gm["cpl"]) * 100 if gm["cpl"] > 0 else 0
    gauge_html = gauge_bar(cpl_pct,
        f'{fmt_pct(cpl_pct)} da meta (R$ {int(CPL_META)},00)')
    tree_html = render_tree(tree)
    trend_json = json.dumps(trend)

    # Pre-format all values to avoid complex f-string nesting
    v = {
        "spend":     fmt_brl(gm["spend"]),
        "imp":       fmt_int(gm["impressions"]),
        "cpm":       fmt_brl(gm["cpm"]),
        "freq":      fmt_float2(gm.get("frequency", 0)),
        "reach":     fmt_int(gm["reach"]),
        "clicks":    fmt_int(gm["clicks"]),
        "ctr":       fmt_pct(gm["ctr"]),
        "cpc":       fmt_brl(gm["cpc"]),
        "conv_form": fmt_pct(gm["conv_form"]),
        "leads":     fmt_int(int(gm["leads"])),
        "cpl":       fmt_brl(gm["cpl"]),
        "tx_a":      fmt_pct(gm["tx_a"]),
        "cpl_a":     fmt_brl(gm["cpl_a"]),
        "leads_a":   str(gm["leads_a"]),
        "leads_b":   str(gm["leads_b"]),
        "leads_c":   str(gm["leads_c"]),
        "leads_d":   str(gm["leads_d"]),
    }

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
  --or:#f97316;--am:#fbbf24;--dark:#0f172a;--card:#1e293b;--card2:#162032;
  --text:#e2e8f0;--muted:#94a3b8;--green:#4ade80;--red:#f87171;--border:#334155;
}}
body{{font-family:'Segoe UI',sans-serif;background:var(--dark);color:var(--text);min-height:100vh}}
header{{background:linear-gradient(135deg,#ea580c 0%,#b45309 100%);padding:18px 32px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px}}
header h1{{font-size:1.5rem;font-weight:700;color:#fff}}
header .period{{font-size:.82rem;color:rgba(255,255,255,.85);margin-top:3px}}
.updated{{font-size:.72rem;color:rgba(255,255,255,.65)}}
.container{{max-width:1700px;margin:0 auto;padding:24px 18px}}
.sec-title{{font-size:1.05rem;font-weight:700;color:var(--am);margin:30px 0 12px;text-transform:uppercase;letter-spacing:.06em;display:flex;align-items:center;gap:8px}}
.sec-title::after{{content:'';flex:1;height:1px;background:var(--border)}}

/* Funnel cards */
.funnel-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:12px}}
.fcard{{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:16px 18px}}
.fcard .lbl{{font-size:.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin-bottom:5px}}
.fcard .val{{font-size:1.65rem;font-weight:700;color:#fff;line-height:1.1}}
.fcard .sub{{font-size:.8rem;color:var(--muted);margin-top:5px}}
.fcard .sub b{{color:var(--am)}}
.fcard.hi{{border-color:var(--or);background:linear-gradient(135deg,#1e1e3a,#2d1b00)}}
.fcard.hi .val{{color:var(--am)}}

/* Gauge */
.gauge-wrap{{margin-top:10px}}
.gauge-bar{{height:9px;background:#334155;border-radius:5px;overflow:hidden}}
.gauge-fill{{height:100%;border-radius:5px;transition:width .5s ease}}
.gauge-label{{font-size:.72rem;color:var(--muted);margin-top:4px}}

/* Score chips */
.chips{{display:flex;gap:10px;flex-wrap:wrap}}
.chip{{border-radius:8px;padding:10px 18px;font-size:.88rem;font-weight:600}}
.ca{{background:#064e3b;color:#6ee7b7;border:1px solid #065f46}}
.cb{{background:#1e3a8a;color:#93c5fd;border:1px solid #1d4ed8}}
.cc{{background:#78350f;color:#fcd34d;border:1px solid #92400e}}
.cd{{background:#3b0764;color:#d8b4fe;border:1px solid #6b21a8}}

/* Chart */
.chart-card{{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:20px}}

/* Table */
.table-wrap{{overflow-x:auto;border-radius:12px;border:1px solid var(--border)}}
#breakdown-table{{width:100%;border-collapse:collapse;font-size:.8rem;white-space:nowrap}}
#breakdown-table thead tr{{background:var(--or);color:#fff}}
#breakdown-table th{{padding:10px 12px;text-align:right;font-weight:600}}
#breakdown-table th:first-child{{text-align:left;min-width:270px}}
#breakdown-table td{{padding:8px 12px;text-align:right;border-bottom:1px solid #1e293b}}
#breakdown-table td:first-child{{text-align:left}}
.row-camp{{background:#162032;font-weight:600}}
.row-camp td:first-child{{color:var(--am)}}
.row-adset{{background:var(--card)}}
.row-adset td:first-child{{color:#cbd5e1}}
.row-ad{{background:#0d1b2a;font-size:.75rem}}
.row-ad td:first-child{{color:var(--muted)}}
.row-adset.collapsed,.row-ad.collapsed{{display:none}}
.indent-1{{padding-left:24px!important}}
.indent-2{{padding-left:48px!important}}
.name-cell{{display:flex;align-items:center;gap:6px}}
.toggle{{background:transparent;border:1px solid var(--border);border-radius:4px;color:var(--muted);cursor:pointer;font-size:.65rem;padding:2px 5px;flex-shrink:0}}
.toggle:hover{{background:var(--border)}}
@media(max-width:600px){{header{{flex-direction:column}}.fcard .val{{font-size:1.3rem}}}}
</style>
</head>
<body>

<header>
  <div>
    <h1>📊 Dashboard SDC — Giaco</h1>
    <div class="period">📅 {period}</div>
  </div>
  <div class="updated">Atualizado: {generated_at}</div>
</header>

<div class="container">

  <div class="sec-title">Funil de Captura</div>
  <div class="funnel-grid">

    <div class="fcard hi">
      <div class="lbl">Valor Gasto</div>
      <div class="val">{v["spend"]}</div>
    </div>

    <div class="fcard">
      <div class="lbl">Impressões</div>
      <div class="val">{v["imp"]}</div>
      <div class="sub">CPM <b>{v["cpm"]}</b> &nbsp;|&nbsp; Freq. <b>{v["freq"]}</b></div>
    </div>

    <div class="fcard">
      <div class="lbl">Alcance</div>
      <div class="val">{v["reach"]}</div>
    </div>

    <div class="fcard">
      <div class="lbl">Link Clicks</div>
      <div class="val">{v["clicks"]}</div>
      <div class="sub">CTR <b>{v["ctr"]}</b> &nbsp;|&nbsp; CPC <b>{v["cpc"]}</b></div>
    </div>

    <div class="fcard">
      <div class="lbl">Conversão Formulário</div>
      <div class="val">{v["conv_form"]}</div>
      <div class="sub">Leads / Clicks</div>
    </div>

    <div class="fcard hi">
      <div class="lbl">Leads Totais</div>
      <div class="val">{v["leads"]}</div>
      <div class="sub">CPL <b>{v["cpl"]}</b></div>
      {gauge_html}
    </div>

    <div class="fcard">
      <div class="lbl">Tx Lead-A</div>
      <div class="val">{v["tx_a"]}</div>
      <div class="sub">CPL-A <b>{v["cpl_a"]}</b></div>
    </div>

  </div>

  <div class="sec-title">Lead Scoring</div>
  <div class="chips">
    <div class="chip ca">🟢 Lead A &nbsp;<strong>{v["leads_a"]}</strong></div>
    <div class="chip cb">🔵 Lead B &nbsp;<strong>{v["leads_b"]}</strong></div>
    <div class="chip cc">🟡 Lead C &nbsp;<strong>{v["leads_c"]}</strong></div>
    <div class="chip cd">🟣 Lead D &nbsp;<strong>{v["leads_d"]}</strong></div>
  </div>

  <div class="sec-title" style="margin-top:30px">Evolução Diária</div>
  <div class="chart-card">
    <canvas id="trendChart" height="80"></canvas>
  </div>

  <div class="sec-title">Otimizações Meta Ads — Campanha / Conjunto / Anúncio</div>
  <div class="table-wrap">
    {tree_html}
  </div>

</div>

<script>
// ── Chart ──────────────────────────────────────────────────────────────────
const TREND = {trend_json};
(function(){{
  const labels = TREND.map(r=>r.label);
  new Chart(document.getElementById("trendChart"),{{
    type:"bar",
    data:{{
      labels,
      datasets:[
        {{label:"Gasto (R$)",data:TREND.map(r=>r.spend),backgroundColor:"rgba(249,115,22,.6)",borderColor:"#f97316",borderWidth:1,yAxisID:"y"}},
        {{label:"Leads",data:TREND.map(r=>r.leads),type:"line",borderColor:"#fbbf24",backgroundColor:"rgba(251,191,36,.12)",tension:.35,fill:true,pointRadius:3,yAxisID:"y1"}},
        {{label:"Leads A",data:TREND.map(r=>r.leads_a),type:"line",borderColor:"#4ade80",tension:.35,fill:false,pointRadius:3,yAxisID:"y1"}},
      ]
    }},
    options:{{
      responsive:true,
      interaction:{{mode:"index",intersect:false}},
      scales:{{
        x:{{ticks:{{color:"#94a3b8",maxRotation:45}},grid:{{color:"#1e293b"}}}},
        y:{{position:"left",ticks:{{color:"#94a3b8",callback:v=>"R$"+v.toLocaleString("pt-BR")}},grid:{{color:"#1e293b"}}}},
        y1:{{position:"right",ticks:{{color:"#94a3b8"}},grid:{{drawOnChartArea:false}}}},
      }},
      plugins:{{legend:{{labels:{{color:"#e2e8f0"}}}},tooltip:{{bodyColor:"#e2e8f0",titleColor:"#fbbf24"}}}},
    }}
  }});
}})();

// ── Table expand / collapse ─────────────────────────────────────────────────
function toggleLevel(btn, level, id) {{
  const isOpen = btn.textContent === "▼";
  const tbody = document.querySelector("#breakdown-table tbody");
  const rows = Array.from(tbody.rows);

  if (level === "camp") {{
    rows.filter(r => r.dataset.parent === id).forEach(r => {{
      const opening = !isOpen;
      r.classList.toggle("collapsed", !opening);
      // always collapse grandchildren when toggling camp
      if (!opening) {{
        const aid = r.dataset.id;
        if (aid) {{
          rows.filter(rr => rr.dataset.parent === aid).forEach(rr => rr.classList.add("collapsed"));
          const adBtn = r.querySelector(".toggle");
          if (adBtn) adBtn.textContent = "▶";
        }}
      }}
    }});
  }} else {{
    // adset: toggle direct ad children
    rows.filter(r => r.dataset.parent === id).forEach(r => {{
      r.classList.toggle("collapsed", isOpen);
    }});
  }}
  btn.textContent = isOpen ? "▶" : "▼";
}}
</script>
</body>
</html>"""


# ─── main ────────────────────────────────────────────────────────────────────

def main():
    today = date.today()
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

    print("Aggregating…", flush=True)
    gm, tree = aggregate(meta_rows, ls_rows)
    trend = daily_trend(meta_rows, ls_rows, d_from, d_to)

    generated_at = datetime.utcnow().strftime("%d/%m/%Y %H:%M UTC")
    html = generate_html(gm, tree, trend, d_from, d_to, generated_at)

    out = os.path.join(os.path.dirname(__file__), "..", "docs", "index.html")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Saved: {os.path.abspath(out)}", flush=True)


if __name__ == "__main__":
    main()
