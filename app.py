import streamlit as st
import requests
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from datetime import date
import json
import io
import statistics

st.set_page_config(
    page_title="Aviation Lessee Credit Analysis",
    page_icon="\u2708",
    layout="wide",
)

AIRLINES = {
    "Delta Air Lines": {"ticker": "DAL", "cik": "0000027904"},
    "United Airlines Holdings": {"ticker": "UAL", "cik": "0000100517"},
    "American Airlines Group": {"ticker": "AAL", "cik": "0001549922"},
    "Southwest Airlines": {"ticker": "LUV", "cik": "0000092380"},
    "Alaska Air Group": {"ticker": "ALK", "cik": "0000766421"},
    "JetBlue Airways": {"ticker": "JBLU", "cik": "0001158463"},
    "Allegiant Travel Company": {"ticker": "ALGT", "cik": "0001362988"},
    "Frontier Group Holdings": {"ticker": "ULCC", "cik": "0001836035"},
    "Spirit Airlines": {"ticker": "SAVE", "cik": "0001418121"},
    "Sun Country Airlines": {"ticker": "SNCY", "cik": "0001549802"},
    "Hawaiian Holdings": {"ticker": "HA", "cik": "0000046619"},
    "SkyWest Inc": {"ticker": "SKYW", "cik": "0000070858"},
    "Air Transport Services Group": {"ticker": "ATSG", "cik": "0000894871"},
    "Mesa Air Group": {"ticker": "MESA", "cik": "0000810332"},
}

SEC_HEADERS = {
    "User-Agent": "Aviation Credit Tool aviationcredit@gmail.com",
    "Accept": "application/json",
}

# RAG hex colors
GREEN_BG, GREEN_TEXT = "#f0fdf4", "#15803d"
AMBER_BG, AMBER_TEXT = "#fffbeb", "#b45309"
RED_BG, RED_TEXT = "#fef2f2", "#b91c1c"
NM_BG, NM_TEXT = "#f9fafb", "#6b7280"

COLOR_MAP = {
    "green": (GREEN_BG, GREEN_TEXT),
    "amber": (AMBER_BG, AMBER_TEXT),
    "red": (RED_BG, RED_TEXT),
    "nm": (NM_BG, NM_TEXT),
}

INSTANT = "instant"
DURATION = "duration"

# concept_key: (list_of_xbrl_concepts_in_priority_order, kind)
CONCEPTS = {
    "revenue": (["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax", "SalesRevenueNet"], DURATION),
    "assets_current": (["AssetsCurrent"], INSTANT),
    "liabilities_current": (["LiabilitiesCurrent"], INSTANT),
    "cash": (["CashAndCashEquivalentsAtCarryingValue"], INSTANT),
    "short_term_investments": (["ShortTermInvestments", "AvailableForSaleSecuritiesCurrent"], INSTANT),
    "assets_total": (["Assets"], INSTANT),
    "debt_current": (["DebtCurrent", "LongTermDebtCurrent"], INSTANT),
    "debt_noncurrent": (["LongTermDebtNoncurrent", "LongTermDebt"], INSTANT),
    "op_lease_current": (["OperatingLeaseLiabilityCurrent"], INSTANT),
    "op_lease_noncurrent": (["OperatingLeaseLiabilityNoncurrent"], INSTANT),
    "equity": (["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"], INSTANT),
    "operating_income": (["OperatingIncomeLoss"], DURATION),
    "interest_expense": (["InterestExpense", "InterestAndDebtExpense"], DURATION),
    "da": (["DepreciationDepletionAndAmortization", "DepreciationAndAmortization"], DURATION),
    "operating_lease_cost": (["OperatingLeaseCost", "LeaseAndRentalExpense"], DURATION),
}

DEFAULT_ZERO = {"short_term_investments", "debt_current", "op_lease_current", "op_lease_noncurrent"}

# id, label, key, unit_symbol, dimension
RATIOS = [
    (1, "Current Ratio", "current_ratio", "x", "LIQUIDITY"),
    (2, "Quick Ratio", "quick_ratio", "x", "LIQUIDITY"),
    (3, "Cash % Revenue", "cash_pct_revenue", "%", "LIQUIDITY"),
    (4, "Net Debt / EBITDAR", "net_debt_ebitdar", "x", "LEVERAGE"),
    (5, "Total Debt / Assets", "debt_assets", "x", "LEVERAGE"),
    (6, "Total Debt / Equity", "debt_equity", "x", "LEVERAGE"),
    (7, "Interest Coverage", "interest_coverage", "x", "COVERAGE"),
    (8, "EBITDAR Coverage", "ebitdar_coverage", "x", "COVERAGE"),
]
RATIO_BY_ID = {r[0]: r for r in RATIOS}
HIGHER_BETTER = {1, 2, 3, 7, 8}
LEVERAGE_IDS = {4, 5, 6}

ABS_THRESHOLDS = {
    # id: (green_cut, amber_cut). For higher-better: v>=green green, v>=amber amber, else red.
    # For lower-better: v<=green green, v<=amber amber, else red.
    1: (0.90, 0.60),
    2: (0.80, 0.50),
    3: (15.0, 10.0),
    4: (4.0, 6.0),
    5: (0.70, 0.85),
    6: (3.0, 5.0),
    7: (2.0, 1.0),
    8: (3.0, 2.0),
}
# threshold used for chart reference lines (the green cutoff)
GREEN_LINE = {rid: ABS_THRESHOLDS[rid][0] for rid in ABS_THRESHOLDS}


# ----------------------------------------------------------------------------
# Data fetching & extraction
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def fetch_companyfacts(cik):
    url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
    resp = requests.get(url, headers=SEC_HEADERS, timeout=30)
    resp.raise_for_status()
    return resp.json()


def concept_annual(facts_root, concepts, kind):
    """Return ({fiscal_year_int: value}, matched_concept_or_None).

    Period is derived from the fact's END DATE, not the SEC 'fy' field, because
    'fy'/'fp' describe the filing the fact came from, not the period it covers.
    Comparative-period figures inside a later 10-K are therefore mapped to the
    correct fiscal year. Duration concepts additionally require a ~annual window.
    """
    usgaap = facts_root.get("facts", {}).get("us-gaap", {})
    if not usgaap:
        usgaap = facts_root.get("us-gaap", {})  # tolerate spec's flatter shape
    for concept in concepts:
        node = usgaap.get(concept)
        if not node:
            continue
        units = node.get("units", {})
        series = units.get("USD")
        if not series and units:
            series = next(iter(units.values()))
        if not series:
            continue
        ann = {}
        for e in series:
            if e.get("form") != "10-K":
                continue
            end = e.get("end")
            if not end:
                continue
            try:
                end_d = date.fromisoformat(end)
            except (ValueError, TypeError):
                continue
            if kind == DURATION:
                start = e.get("start")
                if not start:
                    continue
                try:
                    start_d = date.fromisoformat(start)
                except (ValueError, TypeError):
                    continue
                length = (end_d - start_d).days
                if length < 350 or length > 380:
                    continue
            val = e.get("val")
            if val is None:
                continue
            fy = end_d.year
            filed = e.get("filed", "")
            prev = ann.get(fy)
            if prev is None or filed >= prev[1]:
                ann[fy] = (val, filed)
        if ann:
            return {fy: v[0] for fy, v in ann.items()}, concept
    return {}, None


def safe_div(a, b):
    if a is None or b is None or b == 0:
        return None
    return a / b


def compute_records(raw, lease_cost_found):
    """Build ascending year records with raw $ values, derived metrics, ratios."""
    anchor = raw.get("assets_total") or raw.get("revenue") or {}
    years = sorted(anchor.keys(), reverse=True)[:3]
    years = sorted(years)
    records = []
    for y in years:
        rec = {"fy": y}
        for key in CONCEPTS:
            v = raw.get(key, {}).get(y)
            if v is None and key in DEFAULT_ZERO:
                v = 0
            rec[key] = v

        dc = rec["debt_current"] or 0
        dnc = rec["debt_noncurrent"]
        olc = rec["op_lease_current"] or 0
        olnc = rec["op_lease_noncurrent"] or 0
        total_debt = None if dnc is None else (dc + dnc + olc + olnc)

        cash = rec["cash"]
        sti = rec["short_term_investments"] or 0
        net_debt = None if total_debt is None else total_debt - (cash or 0) - sti

        oi = rec["operating_income"]
        da = rec["da"]
        ebitda = (oi + da) if (oi is not None and da is not None) else None
        olcost = rec["operating_lease_cost"]
        if ebitda is not None and olcost is not None:
            ebitdar = ebitda + olcost
        else:
            ebitdar = ebitda

        ac = rec["assets_current"]
        lc = rec["liabilities_current"]
        at = rec["assets_total"]
        eq = rec["equity"]
        ie = rec["interest_expense"]
        rev = rec["revenue"]

        ratios = {}
        nm = {4: False, 6: False}

        ratios[1] = safe_div(ac, lc)
        num_quick = None if cash is None else (cash + sti)
        ratios[2] = safe_div(num_quick, lc)
        r3 = safe_div(cash, rev)
        ratios[3] = None if r3 is None else r3 * 100.0

        if ebitdar is None:
            ratios[4] = None
        elif ebitdar <= 0:
            ratios[4] = None
            nm[4] = True
        else:
            ratios[4] = safe_div(net_debt, ebitdar)

        ratios[5] = safe_div(total_debt, at)

        if eq is None:
            ratios[6] = None
        elif eq <= 0:
            ratios[6] = None
            nm[6] = True
        else:
            ratios[6] = safe_div(total_debt, eq)

        ratios[7] = safe_div(oi, ie)
        ratios[8] = safe_div(ebitdar, ie)

        rec.update({
            "total_debt": total_debt,
            "net_debt": net_debt,
            "ebitda": ebitda,
            "ebitdar": ebitdar,
            "ebitdar_label": "EBITDAR" if lease_cost_found else "EBITDA",
            "ratios": ratios,
            "nm": nm,
        })
        records.append(rec)
    return records


def build_airline(name, info):
    facts = fetch_companyfacts(info["cik"])
    raw = {}
    matched = {}
    for key, (concepts, kind) in CONCEPTS.items():
        vals, mc = concept_annual(facts, concepts, kind)
        raw[key] = vals
        matched[key] = mc
    lease_cost_found = matched.get("operating_lease_cost") is not None
    records = compute_records(raw, lease_cost_found)
    return {
        "ticker": info["ticker"],
        "label": "EBITDAR" if lease_cost_found else "EBITDA",
        "records": records,
        "years": [r["fy"] for r in records],
    }


# ----------------------------------------------------------------------------
# RAG colour resolution
# ----------------------------------------------------------------------------
def absolute_rag(rid, v):
    if v is None:
        return None
    g, a = ABS_THRESHOLDS[rid]
    if rid in HIGHER_BETTER:
        if v >= g:
            return "green"
        if v >= a:
            return "amber"
        return "red"
    if v <= g:
        return "green"
    if v <= a:
        return "amber"
    return "red"


def peer_rag(rid, v, peers):
    vals = [x for x in peers if x is not None]
    if v is None or len(vals) < 2:
        return absolute_rag(rid, v)
    med = statistics.median(vals)
    try:
        qs = statistics.quantiles(vals, n=4)
    except statistics.StatisticsError:
        qs = [med, med, med]
    q1, q3 = qs[0], qs[2]
    if rid in LEVERAGE_IDS:
        if v <= med:
            return "green"
        if v >= q3:
            return "red"
        return "amber"
    if v >= med:
        return "green"
    if v <= q1:
        return "red"
    return "amber"


def base_color(rid, v, peers, peer_mode):
    if peer_mode:
        return peer_rag(rid, v, peers)
    return absolute_rag(rid, v)


def resolve_color(rid, records, peers, peer_mode):
    """Latest-year colour name for a ratio, including N/M trend rules."""
    if not records:
        return "nm"
    latest = records[-1]
    v = latest["ratios"].get(rid)
    if v is not None:
        return base_color(rid, v, peers, peer_mode) or "nm"

    # Latest is None. Apply N/M trend logic for ratios 4 and 6, else neutral.
    prior = records[-2] if len(records) >= 2 else None
    if rid == 4 and latest["nm"].get(4):
        cur_e = latest["ebitdar"]
        prior_e = prior["ebitdar"] if prior else None
        if prior_e is None or prior_e <= 0:
            return "red"
        if cur_e is not None and cur_e > prior_e:
            return "amber"
        return "red"
    if rid == 6 and latest["nm"].get(6):
        cur_q = latest["equity"]
        prior_q = prior["equity"] if prior else None
        if prior_q is None:
            return "red"
        if cur_q is not None and cur_q > prior_q:
            return "amber"
        return "red"
    return "nm"


# ----------------------------------------------------------------------------
# Formatting
# ----------------------------------------------------------------------------
def fmt_ratio(rid, v):
    if v is None:
        return "N/M"
    sym = RATIO_BY_ID[rid][3]
    if sym == "%":
        return f"{v:.1f}%"
    return f"{v:.2f}x"


def to_millions(v):
    if v is None:
        return None
    return v / 1_000_000.0


# ----------------------------------------------------------------------------
# UI: sidebar
# ----------------------------------------------------------------------------
st.sidebar.title("\u2708 Airline Selection")
selected = st.sidebar.multiselect("Select airlines", list(AIRLINES.keys()))
load = st.sidebar.button("Load Data")

if "data" not in st.session_state:
    st.session_state.data = {}
    st.session_state.peer_mode = False

if load:
    data = {}
    for name in selected:
        with st.spinner(f"Fetching {name} ..."):
            try:
                data[name] = build_airline(name, AIRLINES[name])
            except requests.exceptions.RequestException as exc:
                st.error(f"Request failed for {name}: {exc}")
            except (ValueError, KeyError, json.JSONDecodeError) as exc:
                st.error(f"Could not parse SEC data for {name}: {exc}")
    st.session_state.data = data
    st.session_state.peer_mode = len(data) >= 4

data = st.session_state.data
peer_mode = st.session_state.peer_mode

if data:
    mode_label = "peer comparison" if peer_mode else "absolute thresholds"
    st.sidebar.success(f"{len(data)} airline(s) loaded \u2014 using {mode_label}.")
    if not peer_mode:
        st.sidebar.warning("Peer comparison requires 4+ airlines. Using absolute thresholds.")


# ----------------------------------------------------------------------------
# Header
# ----------------------------------------------------------------------------
st.title("Aviation Lessee Credit Analysis Dashboard")
st.caption("Aircraft Lessor Credit Tool | Data: SEC EDGAR XBRL 10-K")

if not data:
    st.info("Select airlines in the sidebar and click **Load Data** to begin.")
    st.stop()

# ASC 842 banner if any pre-2020 fiscal year appears
pre2020 = [n for n, d in data.items() if any(y < 2020 for y in d["years"])]
if pre2020:
    st.warning(
        "ASC 842 (operating lease balance-sheet recognition) took effect for fiscal years "
        "beginning after 15 Dec 2018. Pre-2020 figures for "
        + ", ".join(pre2020)
        + " may not include capitalised operating-lease liabilities and are not strictly comparable."
    )

airlines = list(data.keys())
color_seq = px.colors.qualitative.Plotly
airline_colors = {a: color_seq[i % len(color_seq)] for i, a in enumerate(airlines)}

# precompute latest-year peer arrays per ratio
peers_latest = {
    rid: [data[a]["records"][-1]["ratios"].get(rid) if data[a]["records"] else None for a in airlines]
    for rid, *_ in RATIOS
}


def latest_fy_label():
    yrs = [d["years"][-1] for d in data.values() if d["years"]]
    return f"FY{max(yrs)}" if yrs else "Latest FY"


# ----------------------------------------------------------------------------
# Excel export
# ----------------------------------------------------------------------------
def openpyxl_fill(color_name):
    bg, _ = COLOR_MAP[color_name]
    hexv = bg.replace("#", "")
    return PatternFill(start_color=hexv, end_color=hexv, fill_type="solid")


def build_excel():
    wb = Workbook()
    bold = Font(bold=True)

    # Sheet 1: Credit Summary (latest FY) with RAG fills
    ws = wb.active
    ws.title = "Credit Summary"
    ws.cell(row=1, column=1, value="Ratio").font = bold
    for j, a in enumerate(airlines, start=2):
        ws.cell(row=1, column=j, value=a).font = bold
    r = 2
    last_dim = None
    for rid, label, key, sym, dim in RATIOS:
        if dim != last_dim:
            ws.cell(row=r, column=1, value=dim).font = bold
            r += 1
            last_dim = dim
        ws.cell(row=r, column=1, value=label)
        for j, a in enumerate(airlines, start=2):
            recs = data[a]["records"]
            v = recs[-1]["ratios"].get(rid) if recs else None
            cell = ws.cell(row=r, column=j, value=fmt_ratio(rid, v))
            cname = resolve_color(rid, recs, peers_latest[rid], peer_mode)
            cell.fill = openpyxl_fill(cname)
            cell.alignment = Alignment(horizontal="center")
        r += 1
    ws.column_dimensions["A"].width = 24
    for j in range(2, 2 + len(airlines)):
        ws.column_dimensions[ws.cell(row=1, column=j).column_letter].width = 18

    # Sheet 2: 3-Year Trends
    ws2 = wb.create_sheet("3-Year Trends")
    headers = ["Airline", "Ticker", "FY"] + [r[1] for r in RATIOS]
    for j, h in enumerate(headers, start=1):
        ws2.cell(row=1, column=j, value=h).font = bold
    row = 2
    for a in airlines:
        for rec in data[a]["records"]:
            ws2.cell(row=row, column=1, value=a)
            ws2.cell(row=row, column=2, value=data[a]["ticker"])
            ws2.cell(row=row, column=3, value=f"FY{rec['fy']}")
            for k, (rid, *_rest) in enumerate(RATIOS, start=4):
                ws2.cell(row=row, column=k, value=fmt_ratio(rid, rec["ratios"].get(rid)))
            row += 1

    # Sheet 3: Raw Financials ($M)
    ws3 = wb.create_sheet("Raw Financials")
    rawhead = ["Airline", "Ticker", "FY", "Revenue ($M)", "EBITDA(R)", "EBITDA(R) ($M)",
               "Total Debt ($M)", "Net Debt ($M)", "Cash ($M)", "Total Assets ($M)",
               "Equity ($M)", "Interest Expense ($M)"]
    for j, h in enumerate(rawhead, start=1):
        ws3.cell(row=1, column=j, value=h).font = bold
    row = 2
    for a in airlines:
        for rec in data[a]["records"]:
            vals = [
                a, data[a]["ticker"], f"FY{rec['fy']}",
                to_millions(rec["revenue"]), rec["ebitdar_label"], to_millions(rec["ebitdar"]),
                to_millions(rec["total_debt"]), to_millions(rec["net_debt"]),
                to_millions(rec["cash"]), to_millions(rec["assets_total"]),
                to_millions(rec["equity"]), to_millions(rec["interest_expense"]),
            ]
            for j, v in enumerate(vals, start=1):
                if isinstance(v, float):
                    ws3.cell(row=row, column=j, value=round(v, 1))
                else:
                    ws3.cell(row=row, column=j, value=v)
            row += 1

    # Sheet 4: Methodology
    ws4 = wb.create_sheet("Methodology")
    ws4.column_dimensions["A"].width = 110
    lines = [
        "AVIATION LESSEE CREDIT ANALYSIS \u2014 METHODOLOGY",
        "",
        "DATA SOURCE",
        "All figures sourced from SEC EDGAR XBRL companyfacts (10-K filings).",
        "Fiscal periods are assigned by each fact's period END DATE (not the SEC 'fy' tag), so",
        "comparative-period figures inside later filings are mapped to the correct fiscal year.",
        "",
        "DERIVED METRICS",
        "Total Debt = current debt + non-current debt + current op-lease liab + non-current op-lease liab",
        "Net Debt = Total Debt - Cash - Short-term Investments",
        "EBITDA = Operating Income + Depreciation & Amortisation",
        "EBITDAR = EBITDA + Operating Lease Cost (when reported; otherwise EBITDA is used)",
        "",
        "RATIOS, LESSOR PURPOSE, AND THRESHOLD RATIONALE",
        "1. Current Ratio = Current Assets / Current Liabilities",
        "   Purpose: near-term ability to meet obligations incl. lease rentals. Green >=0.90, Amber 0.60-0.90, Red <0.60.",
        "2. Quick Ratio = (Cash + Short-term Investments) / Current Liabilities",
        "   Purpose: liquidity excluding less-liquid current assets. Green >=0.80, Amber 0.50-0.80, Red <0.50.",
        "3. Cash % Revenue = Cash / Revenue x 100",
        "   Purpose: liquidity buffer relative to operating scale. Green >=15, Amber 10-15, Red <10.",
        "4. Net Debt / EBITDAR = Net Debt / EBITDAR",
        "   Purpose: lease-adjusted leverage \u2014 the core lessor metric. Green <=4.0, Amber 4.0-6.0, Red >6.0.",
        "5. Total Debt / Assets = Total Debt / Total Assets",
        "   Purpose: balance-sheet leverage. Green <=0.70, Amber 0.70-0.85, Red >0.85.",
        "6. Total Debt / Equity = Total Debt / Stockholders' Equity",
        "   Purpose: leverage vs equity cushion. Green <=3.0, Amber 3.0-5.0, Red >5.0.",
        "7. Interest Coverage = Operating Income / Interest Expense",
        "   Purpose: ability to service interest. Green >=2.0, Amber 1.0-2.0, Red <1.0.",
        "8. EBITDAR Coverage = EBITDAR / Interest Expense",
        "   Purpose: lease-adjusted earnings vs interest. Green >=3.0, Amber 2.0-3.0, Red <2.0.",
        "",
        "EBITDAR vs EBITDA",
        "EBITDAR adds back operating lease cost to allow comparison across airlines with different",
        "owned/leased fleet mixes. When operating lease cost is not separately reported, EBITDA is used",
        "and the metric is labelled accordingly per airline.",
        "",
        "N/M TREATMENT",
        "Net Debt / EBITDAR is Not Meaningful when EBITDAR <= 0 (a leverage multiple over negative",
        "earnings is uninformative). Total Debt / Equity is Not Meaningful when equity <= 0 (deficit).",
        "For the latest year, N/M cells are coloured by trend: deteriorating/persistent N/M = Red;",
        "improving toward positive = Amber. N/M only in prior years is coloured normally on the valid latest year.",
        "",
        "ASC 842",
        "Operating lease liabilities appear on the balance sheet only for fiscal years beginning after",
        "15 Dec 2018. Pre-2020 figures may understate Total Debt and are not strictly comparable.",
        "",
        "PEER COMPARISON RULE",
        "With 4+ airlines selected, RAG bands use peer median/quartile logic (leverage: below median = green,",
        "top quartile = red; others: above median = green, bottom quartile = red). With fewer than 4 airlines,",
        "fixed absolute thresholds are used.",
        "",
        "Verify all figures against primary filings before any investment or credit decision.",
    ]
    for i, ln in enumerate(lines, start=1):
        c = ws4.cell(row=i, column=1, value=ln)
        if ln and ln.isupper() and len(ln) < 60:
            c.font = bold

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


export_bytes = build_excel()
st.download_button(
    "\u2b07 Export to Excel",
    data=export_bytes,
    file_name=f"Aviation_Credit_Analysis_{date.today().isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

tab1, tab2, tab3, tab4 = st.tabs(["Credit Summary", "Trends", "Peer Ranking", "Raw Data"])

# ----------------------------------------------------------------------------
# TAB 1 \u2014 Credit Summary
# ----------------------------------------------------------------------------
with tab1:
    st.subheader(f"Credit Summary \u2014 {latest_fy_label()}")

    # Signature element: RAG proportion bar per airline
    st.markdown("**Metric health (latest FY)**")
    for a in airlines:
        recs = data[a]["records"]
        counts = {"green": 0, "amber": 0, "red": 0, "nm": 0}
        for rid, *_ in RATIOS:
            counts[resolve_color(rid, recs, peers_latest[rid], peer_mode)] += 1
        total = sum(counts.values()) or 1
        seg = ""
        for cname in ["green", "amber", "red", "nm"]:
            if counts[cname] == 0:
                continue
            w = counts[cname] / total * 100
            bg, txt = COLOR_MAP[cname]
            border = txt
            seg += (
                f"<div style='width:{w:.1f}%;background:{bg};color:{txt};"
                f"border-right:1px solid #fff;text-align:center;font-size:12px;"
                f"line-height:22px;border:1px solid {border}33;'>{counts[cname]}</div>"
            )
        bar = (
            f"<div style='display:flex;width:100%;height:22px;border-radius:4px;"
            f"overflow:hidden;margin-bottom:6px;'>{seg}</div>"
        )
        st.markdown(
            f"<div style='font-size:13px;margin-bottom:2px;'>{a} "
            f"<span style='color:#6b7280;'>({data[a]['ticker']})</span></div>{bar}",
            unsafe_allow_html=True,
        )

    st.markdown("---")

    cols = list(airlines)
    show_median = peer_mode
    if show_median:
        cols = cols + ["Peer Median"]

    index_rows = []
    text_grid = {}
    color_grid = {}
    last_dim = None
    for rid, label, key, sym, dim in RATIOS:
        if dim != last_dim:
            index_rows.append(dim)
            text_grid[dim] = {c: "" for c in cols}
            color_grid[dim] = {c: f"font-weight:700;background:#eef2ff;color:#1e3a8a" for c in cols}
            last_dim = dim
        index_rows.append(label)
        text_grid[label] = {}
        color_grid[label] = {}
        for a in airlines:
            recs = data[a]["records"]
            v = recs[-1]["ratios"].get(rid) if recs else None
            text_grid[label][a] = fmt_ratio(rid, v)
            cname = resolve_color(rid, recs, peers_latest[rid], peer_mode)
            bg, txt = COLOR_MAP[cname]
            color_grid[label][a] = f"background:{bg};color:{txt};text-align:center"
        if show_median:
            vals = [v for v in peers_latest[rid] if v is not None]
            med = statistics.median(vals) if vals else None
            text_grid[label]["Peer Median"] = fmt_ratio(rid, med)
            color_grid[label]["Peer Median"] = "background:#f1f5f9;color:#334155;text-align:center;font-style:italic"

    df_disp = pd.DataFrame(
        [[text_grid[r].get(c, "") for c in cols] for r in index_rows],
        index=index_rows, columns=cols,
    )
    css_df = pd.DataFrame(
        [[color_grid[r].get(c, "") for c in cols] for r in index_rows],
        index=index_rows, columns=cols,
    )
    styler = df_disp.style.apply(lambda _: css_df, axis=None)
    st.dataframe(styler, use_container_width=True)
    st.caption("Ratios shown as Nx; Cash % Revenue as %. N/M = not meaningful. "
               "Colour bands: " + ("peer median/quartile." if peer_mode else "absolute thresholds."))

# ----------------------------------------------------------------------------
# TAB 2 \u2014 Trends
# ----------------------------------------------------------------------------
with tab2:
    st.subheader("3-Year Trends")
    grid = st.columns(2)
    for i, (rid, label, key, sym, dim) in enumerate(RATIOS):
        rows = []
        for a in airlines:
            for rec in data[a]["records"]:
                rows.append({"Airline": a, "FY": f"FY{rec['fy']}", "Value": rec["ratios"].get(rid)})
        dfp = pd.DataFrame(rows)
        with grid[i % 2]:
            if dfp.empty or dfp["Value"].dropna().empty:
                st.markdown(f"**{label}**")
                st.info("No data.")
                continue
            order = sorted(dfp["FY"].unique())
            fig = px.line(
                dfp, x="FY", y="Value", color="Airline", markers=True,
                category_orders={"FY": order}, color_discrete_map=airline_colors,
                title=label,
            )
            fig.add_hline(
                y=GREEN_LINE[rid], line_dash="dash", line_color="#15803d",
                annotation_text="green threshold", annotation_position="top left",
            )
            fig.update_layout(height=320, margin=dict(l=10, r=10, t=40, b=10),
                              yaxis_title=("%" if sym == "%" else "x"))
            st.plotly_chart(fig, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 3 \u2014 Peer Ranking
# ----------------------------------------------------------------------------
with tab3:
    st.subheader(f"Peer Ranking \u2014 {latest_fy_label()}")
    if len(airlines) < 2:
        st.info("Peer ranking needs at least 2 airlines. Load more to compare.")
    else:
        grid = st.columns(2)
        for i, (rid, label, key, sym, dim) in enumerate(RATIOS):
            rows = []
            for a in airlines:
                recs = data[a]["records"]
                v = recs[-1]["ratios"].get(rid) if recs else None
                if v is None:
                    continue
                cname = resolve_color(rid, recs, peers_latest[rid], peer_mode)
                rows.append({"Airline": a, "Value": v, "Color": COLOR_MAP[cname][1]})
            with grid[i % 2]:
                st.markdown(f"**{label}**")
                if not rows:
                    st.info("No comparable data.")
                    continue
                dfp = pd.DataFrame(rows)
                asc = rid in LEVERAGE_IDS  # leverage: lower better -> best at top
                dfp = dfp.sort_values("Value", ascending=not asc)
                fig = go.Figure(go.Bar(
                    x=dfp["Value"], y=dfp["Airline"], orientation="h",
                    marker_color=list(dfp["Color"]),
                ))
                vals = [r["Value"] for r in rows]
                med = statistics.median(vals)
                fig.add_vline(x=med, line_dash="dash", line_color="#6b7280",
                              annotation_text="median", annotation_position="top")
                fig.update_layout(height=max(220, 60 + 32 * len(dfp)),
                                  margin=dict(l=10, r=10, t=10, b=10),
                                  xaxis_title=("%" if sym == "%" else "x"))
                st.plotly_chart(fig, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 4 \u2014 Raw Data
# ----------------------------------------------------------------------------
with tab4:
    st.subheader("Raw Financials ($M)")
    rows = []
    for a in airlines:
        for rec in data[a]["records"]:
            rows.append({
                "Airline": a,
                "Ticker": data[a]["ticker"],
                "FY": f"FY{rec['fy']}",
                "Revenue ($M)": to_millions(rec["revenue"]),
                "EBITDA(R)": rec["ebitdar_label"],
                "EBITDA(R) ($M)": to_millions(rec["ebitdar"]),
                "Total Debt ($M)": to_millions(rec["total_debt"]),
                "Net Debt ($M)": to_millions(rec["net_debt"]),
                "Cash ($M)": to_millions(rec["cash"]),
                "Total Assets ($M)": to_millions(rec["assets_total"]),
                "Equity ($M)": to_millions(rec["equity"]),
                "Interest Expense ($M)": to_millions(rec["interest_expense"]),
            })
    raw_df = pd.DataFrame(rows)
    num_cols = [c for c in raw_df.columns if c.endswith("($M)")]

    def red_negative(v):
        if isinstance(v, (int, float)) and v < 0:
            return f"color:{RED_TEXT}"
        return ""

    styler = (raw_df.style
              .format({c: "{:,.1f}" for c in num_cols}, na_rep="\u2014")
              .map(red_negative, subset=num_cols))
    st.dataframe(styler, use_container_width=True)
    st.caption("Data sourced from SEC EDGAR XBRL 10-K filings. "
               "Verify against primary filings for investment decisions.")

# To run: streamlit run app.py