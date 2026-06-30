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
import yfinance as yf

st.set_page_config(
    page_title="Aviation Finance Dashboard",
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


@st.cache_data(show_spinner=False, ttl=900)
def fetch_stock_info(ticker):
    """Price + market cap via yfinance. Degrades to None on any failure."""
    try:
        info = yf.Ticker(ticker).info or {}
    except Exception:
        return {"price": None, "prev": None, "mcap": None}
    price = info.get("currentPrice") or info.get("regularMarketPrice")
    return {"price": price, "prev": info.get("previousClose"), "mcap": info.get("marketCap")}


def fmt_mcap(m):
    if m is None:
        return "\u2014"
    if m >= 1_000_000_000:
        return f"${m / 1e9:.1f}B"
    if m >= 1_000_000:
        return f"${m / 1e6:.1f}M"
    return f"${m:,.0f}"


def price_line_html(stock):
    price = stock.get("price")
    prev = stock.get("prev")
    mcap = stock.get("mcap")
    if price is None:
        return "<span style='color:#9ca3af;'>price unavailable</span>"
    chg_html = ""
    if prev:
        chg = price - prev
        pct = chg / prev * 100 if prev else 0.0
        arrow = "\u25b2" if chg >= 0 else "\u25bc"
        col = GREEN_TEXT if chg >= 0 else RED_TEXT
        chg_html = f" <span style='color:{col};'>{arrow} {pct:+.1f}%</span>"
    mcap_html = f" | Mkt Cap: {fmt_mcap(mcap)}" if mcap else ""
    return f"${price:,.2f}{chg_html}{mcap_html}"


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
                continue
            except (ValueError, KeyError, json.JSONDecodeError) as exc:
                st.error(f"Could not parse SEC data for {name}: {exc}")
                continue
            data[name]["stock"] = fetch_stock_info(AIRLINES[name]["ticker"])
    st.session_state.data = data

data = st.session_state.data
peer_mode = len(data) >= 4

if data:
    mode_label = "peer comparison" if peer_mode else "absolute thresholds"
    st.sidebar.success(f"{len(data)} airline(s) loaded \u2014 using {mode_label}.")
    if not peer_mode:
        st.sidebar.warning("Peer comparison requires 4+ airlines. Using absolute thresholds.")

    st.sidebar.markdown("---")
    remove = None
    for name in list(data.keys()):
        stock = data[name].get("stock") or {}
        ticker = data[name]["ticker"]
        c1, c2 = st.sidebar.columns([6, 1])
        c1.markdown(
            f"<div style='line-height:1.3;'>"
            f"<div style='font-weight:600;font-size:13px;'>{name} "
            f"<span style='color:#6b7280;'>({ticker})</span></div>"
            f"<div style='font-size:11px;'>{price_line_html(stock)}</div></div>",
            unsafe_allow_html=True,
        )
        if c2.button("\u00d7", key=f"rm_{name}", help=f"Remove {name}"):
            remove = name
    if remove is not None:
        del st.session_state.data[remove]
        st.rerun()

    st.sidebar.markdown(
        "<div style='font-size:10px;color:#9ca3af;font-style:italic;margin-top:6px;'>"
        "Prices delayed ~15 min. Source: Yahoo Finance</div>",
        unsafe_allow_html=True,
    )


# ----------------------------------------------------------------------------
# Header
# ----------------------------------------------------------------------------
st.title("Aviation Finance Dashboard")
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
        "AVIATION FINANCE DASHBOARD \u2014 METHODOLOGY",
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

# ----------------------------------------------------------------------------
# Financial Statements engine
# ----------------------------------------------------------------------------
# Row spec tuple: (label, concepts|None, kind, style, neg, direction, calc_key)
#   style:     line / sub / total / eps
#   neg:       True -> display as -abs(value)
#   direction: "up" (higher better), "down" (higher worse), None (neutral)
#   calc_key:  None for pulled rows; a key handled by the calc dispatcher otherwise
# NOTE: Many airline line items (fuel, maintenance, air-traffic liability, etc.)
# are reported as company extension tags rather than standard us-gaap concepts,
# so those rows will frequently be empty and are hidden when all years are None.

INCOME_SPEC = [
    ("Total Revenue", ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax"], DURATION, "line", False, "up", None),
    ("Cost of Revenue", ["CostOfRevenue", "CostOfGoodsAndServicesSold"], DURATION, "line", True, "down", None),
    ("Gross Profit", None, DURATION, "sub", False, "up", "gross_profit"),
    ("Salaries and Benefits", ["LaborAndRelatedExpense", "EmployeeBenefitsAndShareBasedCompensation"], DURATION, "line", False, "down", None),
    ("Aircraft Fuel", ["AirlineFuelCosts", "FuelCostsAirline"], DURATION, "line", False, "down", None),
    ("Depreciation and Amortization", ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization"], DURATION, "line", False, "down", None),
    ("Maintenance", ["AirlineCapacityPurchaseArrangements", "MaintenanceCostsCivil", "AircraftMaintenanceMaterialsAndRepairs"], DURATION, "line", False, "down", None),
    ("Other Operating Expenses", ["OtherOperatingIncomeExpenseNet", "OtherCostAndExpenseOperating"], DURATION, "line", False, "down", None),
    ("Total Operating Expenses", None, DURATION, "sub", False, "down", "total_opex"),
    ("Operating Income (EBIT)", ["OperatingIncomeLoss"], DURATION, "total", False, "up", None),
    ("Interest Income", ["InterestIncomeOther", "InvestmentIncomeInterest"], DURATION, "line", False, "up", None),
    ("Interest Expense", ["InterestExpense", "InterestAndDebtExpense"], DURATION, "line", True, "down", None),
    ("Other Income/Expense", ["NonoperatingIncomeExpense", "OtherNonoperatingIncomeExpense"], DURATION, "line", False, None, None),
    ("Income Before Tax", ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"], DURATION, "sub", False, "up", None),
    ("Income Tax", ["IncomeTaxExpenseBenefit"], DURATION, "line", True, None, None),
    ("Net Income", ["NetIncomeLoss"], DURATION, "total", False, "up", None),
    ("EPS Basic", ["EarningsPerShareBasic"], DURATION, "eps", False, "up", None),
    ("EPS Diluted", ["EarningsPerShareDiluted"], DURATION, "eps", False, "up", None),
    ("EBITDA (calculated)", None, DURATION, "sub", False, "up", "ebitda"),
]

BALANCE_SPEC = [
    ("CURRENT ASSETS", None, INSTANT, "header", False, None, None),
    ("Cash and Equivalents", ["CashAndCashEquivalentsAtCarryingValue"], INSTANT, "line", False, "up", None),
    ("Short-term Investments", ["ShortTermInvestments", "AvailableForSaleSecuritiesCurrent"], INSTANT, "line", False, "up", None),
    ("Accounts Receivable", ["AccountsReceivableNetCurrent", "ReceivablesNetCurrent"], INSTANT, "line", False, None, None),
    ("Inventories", ["InventoryNet", "AirlineRelatedInventoryNet"], INSTANT, "line", False, None, None),
    ("Prepaid and Other", ["PrepaidExpenseAndOtherAssetsCurrent"], INSTANT, "line", False, None, None),
    ("Total Current Assets", ["AssetsCurrent"], INSTANT, "sub", False, "up", None),
    ("NON-CURRENT ASSETS", None, INSTANT, "header", False, None, None),
    ("Property Plant and Equipment (net)", ["PropertyPlantAndEquipmentNet"], INSTANT, "line", False, None, None),
    ("Operating Lease ROU Assets", ["OperatingLeaseRightOfUseAsset"], INSTANT, "line", False, None, None),
    ("Goodwill", ["Goodwill"], INSTANT, "line", False, None, None),
    ("Other Non-current Assets", ["OtherAssetsNoncurrent"], INSTANT, "line", False, None, None),
    ("Total Assets", ["Assets"], INSTANT, "total", False, None, None),
    ("CURRENT LIABILITIES", None, INSTANT, "header", False, None, None),
    ("Accounts Payable", ["AccountsPayableCurrent"], INSTANT, "line", False, None, None),
    ("Accrued Liabilities", ["AccruedLiabilitiesCurrent"], INSTANT, "line", False, None, None),
    ("Air Traffic Liability", ["AirTrafficLiability", "DeferredRevenueCurrent"], INSTANT, "line", False, None, None),
    ("Current Debt", ["DebtCurrent", "LongTermDebtCurrent"], INSTANT, "line", False, "down", None),
    ("Current Operating Lease", ["OperatingLeaseLiabilityCurrent"], INSTANT, "line", False, "down", None),
    ("Total Current Liabilities", ["LiabilitiesCurrent"], INSTANT, "sub", False, "down", None),
    ("NON-CURRENT LIABILITIES", None, INSTANT, "header", False, None, None),
    ("Long-term Debt", ["LongTermDebtNoncurrent", "LongTermDebt"], INSTANT, "line", False, "down", None),
    ("Non-current Operating Lease", ["OperatingLeaseLiabilityNoncurrent"], INSTANT, "line", False, "down", None),
    ("Deferred Tax", ["DeferredIncomeTaxLiabilitiesNet", "DeferredTaxLiabilitiesNoncurrent"], INSTANT, "line", False, None, None),
    ("Other Non-current Liabilities", ["OtherLiabilitiesNoncurrent"], INSTANT, "line", False, None, None),
    ("Total Liabilities", ["Liabilities"], INSTANT, "total", False, "down", None),
    ("SHAREHOLDERS EQUITY", None, INSTANT, "header", False, None, None),
    ("Common Stock", ["CommonStockValue"], INSTANT, "line", False, None, None),
    ("Additional Paid-in Capital", ["AdditionalPaidInCapital"], INSTANT, "line", False, None, None),
    ("Retained Earnings", ["RetainedEarningsAccumulatedDeficit"], INSTANT, "line", False, "up", None),
    ("Treasury Stock", ["TreasuryStockValue"], INSTANT, "line", False, None, None),
    ("Total Equity", ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"], INSTANT, "total", False, "up", None),
    ("Total Liabilities and Equity", None, INSTANT, "total", False, None, "total_le"),
]

CASHFLOW_SPEC = [
    ("OPERATING ACTIVITIES", None, DURATION, "header", False, None, None),
    ("Net Income", ["NetIncomeLoss"], DURATION, "line", False, "up", None),
    ("Depreciation and Amortization", ["DepreciationDepletionAndAmortization"], DURATION, "line", False, None, None),
    ("Stock-based Compensation", ["ShareBasedCompensation"], DURATION, "line", False, None, None),
    ("Changes in Working Capital", ["IncreaseDecreaseInOperatingCapital", "IncreaseDecreaseInOperatingLiabilities"], DURATION, "line", False, None, None),
    ("Other Operating", ["OtherOperatingActivitiesCashFlowStatement"], DURATION, "line", False, None, None),
    ("Net Cash from Operations", ["NetCashProvidedByUsedInOperatingActivities"], DURATION, "sub", False, "up", None),
    ("INVESTING ACTIVITIES", None, DURATION, "header", False, None, None),
    ("Capital Expenditures", ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsForFlightEquipment"], DURATION, "line", True, None, None),
    ("Purchases of Investments", ["PaymentsToAcquireInvestments", "PaymentsToAcquireAvailableForSaleSecurities"], DURATION, "line", True, None, None),
    ("Proceeds from Asset Sales", ["ProceedsFromSaleOfPropertyPlantAndEquipment"], DURATION, "line", False, None, None),
    ("Other Investing", ["PaymentsForProceedsFromOtherInvestingActivities"], DURATION, "line", False, None, None),
    ("Net Cash from Investing", ["NetCashProvidedByUsedInInvestingActivities"], DURATION, "sub", False, None, None),
    ("FINANCING ACTIVITIES", None, DURATION, "header", False, None, None),
    ("Debt Proceeds", ["ProceedsFromIssuanceOfLongTermDebt", "ProceedsFromDebtMaturingInMoreThanThreeMonths"], DURATION, "line", False, None, None),
    ("Debt Repayments", ["RepaymentsOfLongTermDebt", "RepaymentsOfDebtMaturingInMoreThanThreeMonths"], DURATION, "line", True, None, None),
    ("Dividends Paid", ["PaymentsOfDividends", "PaymentsOfDividendsCommonStock"], DURATION, "line", True, None, None),
    ("Share Repurchases", ["PaymentsForRepurchaseOfCommonStock"], DURATION, "line", True, None, None),
    ("Other Financing", ["ProceedsFromRepaymentsOfOtherDebt"], DURATION, "line", False, None, None),
    ("Net Cash from Financing", ["NetCashProvidedByUsedInFinancingActivities"], DURATION, "sub", False, None, None),
    ("Net Change in Cash", None, DURATION, "total", False, "up", "net_change"),
    ("Beginning Cash", None, INSTANT, "line", False, None, "beginning_cash"),
    ("Ending Cash", ["CashAndCashEquivalentsAtCarryingValue"], INSTANT, "total", False, "up", None),
]


def _pull(facts, concepts, kind):
    vals, _ = concept_annual(facts, concepts, kind)
    return vals


def _get(vmap, label, fy):
    d = vmap.get(label)
    return d.get(fy) if d else None


def _calc(stmt, key, vmap, fy, years):
    if key == "gross_profit":
        r, c = _get(vmap, "Total Revenue", fy), _get(vmap, "Cost of Revenue", fy)
        return (r - c) if (r is not None and c is not None) else None
    if key == "total_opex":
        comps = ["Salaries and Benefits", "Aircraft Fuel", "Depreciation and Amortization",
                 "Maintenance", "Other Operating Expenses"]
        found = [_get(vmap, l, fy) for l in comps]
        found = [x for x in found if x is not None]
        return sum(found) if found else None
    if key == "ebitda":
        oi, da = _get(vmap, "Operating Income (EBIT)", fy), _get(vmap, "Depreciation and Amortization", fy)
        return (oi + da) if (oi is not None and da is not None) else None
    if key == "total_le":
        tl, te = _get(vmap, "Total Liabilities", fy), _get(vmap, "Total Equity", fy)
        return (tl + te) if (tl is not None and te is not None) else None
    if key == "net_change":
        parts = [_get(vmap, "Net Cash from Operations", fy),
                 _get(vmap, "Net Cash from Investing", fy),
                 _get(vmap, "Net Cash from Financing", fy)]
        parts = [x for x in parts if x is not None]
        return sum(parts) if parts else None
    if key == "beginning_cash":
        return _get(vmap, "Ending Cash", fy - 1)
    return None


def build_statement(facts, years, spec, stmt):
    vmap = {}
    for (label, concepts, kind, style, neg, direction, calc_key) in spec:
        if concepts is not None:
            vmap[label] = _pull(facts, concepts, kind)
    for (label, concepts, kind, style, neg, direction, calc_key) in spec:
        if calc_key is not None:
            vmap[label] = {fy: _calc(stmt, calc_key, vmap, fy, years) for fy in years}
    rows = []
    for (label, concepts, kind, style, neg, direction, calc_key) in spec:
        if style == "header":
            rows.append({"label": label, "style": "header", "neg": False,
                         "dir": None, "vals": {fy: None for fy in years}, "eps": False})
            continue
        valdict = vmap.get(label, {})
        vals = {fy: valdict.get(fy) for fy in years}
        if all(v is None for v in vals.values()):
            continue
        rows.append({"label": label, "style": style, "neg": neg, "dir": direction,
                     "vals": vals, "eps": style == "eps"})
    # drop section headers immediately followed by another header or end of list
    cleaned = []
    for i, r in enumerate(rows):
        if r["style"] == "header":
            nxt = rows[i + 1] if i + 1 < len(rows) else None
            if nxt is None or nxt["style"] == "header":
                continue
        cleaned.append(r)
    return cleaned, vmap


def _pct_change(vals, years):
    if len(years) < 2:
        return None
    a, b = vals.get(years[-2]), vals.get(years[-1])
    if a is None or a == 0 or b is None:
        return None
    return (b - a) / abs(a)


def _fmt_value(v, neg, eps):
    if v is None:
        return "\u2014", False
    disp = -abs(v) if neg else v
    if eps:
        return f"${disp:,.2f}", disp < 0
    return f"{disp / 1e6:,.2f}", disp < 0


def _pct_text_color(p, direction):
    if p is None:
        return "\u2014", NM_TEXT
    txt = f"{p * 100:+.1f}%"
    if direction == "up":
        col = GREEN_TEXT if p > 0 else (RED_TEXT if p < 0 else NM_TEXT)
    elif direction == "down":
        col = RED_TEXT if p > 0 else (GREEN_TEXT if p < 0 else NM_TEXT)
    else:
        col = NM_TEXT
    return txt, col


ROW_BG = {"sub": "#f3f4f6", "total": "#e5e7eb", "header": "#1e293b"}


def render_statement_html(title, rows, years, vmap, is_balance):
    yr_labels = [f"FY{y}" for y in years]
    pct_label = f"% Change (FY{years[-2]}\u2192FY{years[-1]})" if len(years) >= 2 else "% Change"
    head = "".join(f"<th style='text-align:right;padding:6px 10px;'>{h}</th>" for h in yr_labels)
    html = [
        f"<div style='font-weight:700;font-size:16px;margin:14px 0 4px;'>{title}</div>",
        "<table style='width:100%;border-collapse:collapse;font-size:13px;'>",
        f"<tr style='border-bottom:2px solid #cbd5e1;'>"
        f"<th style='text-align:left;padding:6px 10px;'>Line Item</th>{head}"
        f"<th style='text-align:right;padding:6px 10px;'>{pct_label}</th></tr>",
    ]
    for r in rows:
        if r["style"] == "header":
            html.append(
                f"<tr><td colspan='{len(years) + 2}' style='background:{ROW_BG['header']};"
                f"color:#fff;font-weight:700;padding:5px 10px;letter-spacing:.04em;'>"
                f"{r['label']}</td></tr>"
            )
            continue
        bg = ROW_BG.get(r["style"], "transparent")
        weight = "700" if r["style"] in ("sub", "total") else "400"
        cells = ""
        for y in years:
            txt, isneg = _fmt_value(r["vals"].get(y), r["neg"], r["eps"])
            color = RED_TEXT if isneg else "#111827"
            cells += f"<td style='text-align:right;padding:5px 10px;color:{color};'>{txt}</td>"
        ptxt, pcol = _pct_text_color(_pct_change(r["vals"], years), r["dir"])
        html.append(
            f"<tr style='background:{bg};border-bottom:1px solid #f1f5f9;'>"
            f"<td style='padding:5px 10px;font-weight:{weight};'>{r['label']}</td>"
            f"{cells}"
            f"<td style='text-align:right;padding:5px 10px;color:{pcol};font-weight:{weight};'>{ptxt}</td></tr>"
        )
    if is_balance:
        assets = vmap.get("Total Assets", {})
        tle = vmap.get("Total Liabilities and Equity", {})
        chk_cells = ""
        for y in years:
            a, le = assets.get(y), tle.get(y)
            if a is None or le is None or a == 0:
                chk_cells += "<td style='text-align:right;padding:5px 10px;color:#6b7280;'>\u2014</td>"
            elif abs(a - le) <= 0.01 * abs(a):
                chk_cells += f"<td style='text-align:right;padding:5px 10px;color:{GREEN_TEXT};font-weight:700;'>\u2713 Balanced</td>"
            else:
                chk_cells += f"<td style='text-align:right;padding:5px 10px;color:{AMBER_TEXT};font-weight:700;'>\u26a0 Check</td>"
        html.append(
            f"<tr style='background:{ROW_BG['sub']};'>"
            f"<td style='padding:5px 10px;font-weight:700;'>Balance Check</td>{chk_cells}"
            f"<td style='padding:5px 10px;'></td></tr>"
        )
    html.append("</table>")
    return "".join(html)


def build_airline_statements_excel(name, ticker, facts, years):
    wb = Workbook()
    bold = Font(bold=True)
    red = Font(color="b91c1c")
    red_bold = Font(bold=True, color="b91c1c")
    sub_fill = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
    total_fill = PatternFill(start_color="E5E7EB", end_color="E5E7EB", fill_type="solid")

    sheets = [
        ("Income Statement", INCOME_SPEC, "income", False),
        ("Balance Sheet", BALANCE_SPEC, "balance", True),
        ("Cash Flow", CASHFLOW_SPEC, "cashflow", False),
    ]
    first = True
    for sheet_name, spec, stmt, is_bal in sheets:
        rows, vmap = build_statement(facts, years, spec, stmt)
        ws = wb.active if first else wb.create_sheet(sheet_name)
        if first:
            ws.title = sheet_name
            first = False
        header = ["Line Item"] + [f"FY{y}" for y in years]
        header += [f"% Change FY{years[-2]}->FY{years[-1]}" if len(years) >= 2 else "% Change"]
        for j, h in enumerate(header, start=1):
            ws.cell(row=1, column=j, value=h).font = bold
        rownum = 2
        for r in rows:
            if r["style"] == "header":
                c = ws.cell(row=rownum, column=1, value=r["label"])
                c.font = bold
                rownum += 1
                continue
            is_tot = r["style"] in ("sub", "total")
            ws.cell(row=rownum, column=1, value=r["label"]).font = bold if is_tot else None
            for k, y in enumerate(years, start=2):
                v = r["vals"].get(y)
                if v is None:
                    cell = ws.cell(row=rownum, column=k, value=None)
                else:
                    disp = -abs(v) if r["neg"] else v
                    out = round(disp, 2) if r["eps"] else round(disp / 1e6, 2)
                    cell = ws.cell(row=rownum, column=k, value=out)
                    if out < 0:
                        cell.font = red_bold if is_tot else red
                    elif is_tot:
                        cell.font = bold
            p = _pct_change(r["vals"], years)
            ws.cell(row=rownum, column=len(years) + 2,
                    value=(f"{p * 100:+.1f}%" if p is not None else None))
            if is_tot:
                fill = total_fill if r["style"] == "total" else sub_fill
                for col in range(1, len(years) + 3):
                    ws.cell(row=rownum, column=col).fill = fill
            rownum += 1
        if is_bal:
            assets = vmap.get("Total Assets", {})
            tle = vmap.get("Total Liabilities and Equity", {})
            ws.cell(row=rownum, column=1, value="Balance Check").font = bold
            for k, y in enumerate(years, start=2):
                a, le = assets.get(y), tle.get(y)
                if a is None or le is None or a == 0:
                    txt = "\u2014"
                elif abs(a - le) <= 0.01 * abs(a):
                    txt = "Balanced"
                else:
                    txt = "Check"
                ws.cell(row=rownum, column=k, value=txt).font = bold
        ws.column_dimensions["A"].width = 34
        for k in range(2, len(years) + 3):
            ws.column_dimensions[ws.cell(row=1, column=k).column_letter].width = 16

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# yfinance financial-statements engine (Financial Statements tab)
# ----------------------------------------------------------------------------
# Row spec: (yf_label, display_label, style, special, direction)
#   style:     "line" | "bold" (#f8fafc) | "sub" (#f1f5f9)
#   special:   None (=> $M, 2dp) | "eps" ($/share) | "pct" (%)
#   direction: "up" (green if rising) | "down" (red if rising) | None (neutral)
# yfinance row labels drift across versions; matching is case/space-insensitive
# with a small alias map, and missing rows are skipped silently per spec.

YF_ALIASES = {
    "operating expense": ["operating expenses", "total operating expenses"],
    "operating income": ["operating income loss", "ebit"],
    "tax provision": ["income tax expense", "tax provision benefit"],
    "pretax income": ["income before tax", "pre tax income"],
    "net ppe": ["net property plant and equipment", "property plant and equipment net"],
    "total liabilities net minority interest": ["total liabilities"],
    "total equity gross minority interest": ["total equity"],
    "stockholders equity": ["common stock equity", "total stockholders equity"],
    "operating cash flow": ["cash flow from operations", "total cash from operating activities"],
    "investing cash flow": ["total cash from investing activities"],
    "financing cash flow": ["total cash from financing activities"],
    "end cash position": ["ending cash position", "end cash position"],
    "total revenue": ["total revenues", "revenue"],
}

INCOME_YF = [
    ("Total Revenue", "Total Revenue", "line", None, "up"),
    ("Cost Of Revenue", "Cost of Revenue", "line", None, "down"),
    ("Gross Profit", "Gross Profit", "sub", None, "up"),
    ("Operating Expense", "Total Operating Expenses", "bold", None, "down"),
    ("Operating Income", "Operating Income / EBIT", "bold", None, "up"),
    ("EBITDA", "EBITDA", "bold", None, None),
    ("Interest Expense", "Interest Expense", "line", None, "down"),
    ("Interest Income", "Interest Income", "line", None, "up"),
    ("Pretax Income", "Income Before Tax", "line", None, "up"),
    ("Tax Provision", "Income Tax Expense", "line", None, "down"),
    ("Net Income", "Net Income", "bold", None, "up"),
    ("__NET_MARGIN__", "Net Margin %", "line", "pct", "up"),
    ("Basic EPS", "EPS Basic ($)", "eps", "eps", None),
    ("Diluted EPS", "EPS Diluted ($)", "eps", "eps", None),
    ("Normalized EBITDA", "Normalized EBITDA", "line", None, None),
]

BALANCE_YF = [
    ("Cash And Cash Equivalents", "Cash & Equivalents", "line", None, None),
    ("Short Term Investments", "Short-term Investments", "line", None, None),
    ("Accounts Receivable", "Accounts Receivable", "line", None, None),
    ("Inventory", "Inventory", "line", None, None),
    ("Current Assets", "Total Current Assets", "sub", None, None),
    ("Net PPE", "Property, Plant & Equipment (net)", "line", None, None),
    ("Goodwill", "Goodwill", "line", None, None),
    ("Total Assets", "Total Assets", "sub", None, None),
    ("Accounts Payable", "Accounts Payable", "line", None, None),
    ("Current Debt", "Current Portion of Debt", "line", None, "down"),
    ("Current Liabilities", "Total Current Liabilities", "sub", None, "down"),
    ("Long Term Debt", "Long-term Debt", "line", None, "down"),
    ("Total Liabilities Net Minority Interest", "Total Liabilities", "sub", None, "down"),
    ("Stockholders Equity", "Total Stockholders Equity", "bold", None, "up"),
    ("Total Equity Gross Minority Interest", "Total Equity", "sub", None, "up"),
    ("Total Capitalization", "Total Capitalization", "line", None, None),
]

CASHFLOW_YF = [
    ("Net Income", "Net Income", "line", None, "up"),
    ("Depreciation And Amortization", "Depreciation & Amortization", "line", None, None),
    ("Change In Working Capital", "Changes in Working Capital", "line", None, None),
    ("Operating Cash Flow", "Cash from Operations", "sub", None, None),
    ("Capital Expenditure", "Capital Expenditures", "line", None, None),
    ("Purchase Of Investment", "Purchases of Investments", "line", None, None),
    ("Sale Of Investment", "Proceeds from Investment Sales", "line", None, None),
    ("Investing Cash Flow", "Cash from Investing", "sub", None, None),
    ("Issuance Of Debt", "Debt Proceeds", "line", None, None),
    ("Repayment Of Debt", "Debt Repayments", "line", None, "down"),
    ("Repurchase Of Capital Stock", "Share Repurchases", "line", None, None),
    ("Payment Of Dividends", "Dividends Paid", "line", None, None),
    ("Financing Cash Flow", "Cash from Financing", "sub", None, None),
    ("End Cash Position", "Ending Cash Balance", "bold", None, None),
    ("Free Cash Flow", "Free Cash Flow", "bold", None, "up"),
    ("__FCF_MARGIN__", "FCF Margin %", "line", "pct", "up"),
]

YF_BG = {"bold": "#f8fafc", "sub": "#f1f5f9", "line": "transparent"}


@st.cache_data(show_spinner=False, ttl=3600)
def fetch_yf_statements(ticker):
    out = {}
    try:
        t = yf.Ticker(ticker)
        getters = {
            "fin": lambda: t.financials, "bs": lambda: t.balance_sheet, "cf": lambda: t.cashflow,
            "qfin": lambda: t.quarterly_financials, "qbs": lambda: t.quarterly_balance_sheet,
            "qcf": lambda: t.quarterly_cashflow,
        }
        for k, g in getters.items():
            try:
                df = g()
                out[k] = df if isinstance(df, pd.DataFrame) else pd.DataFrame()
            except Exception:
                out[k] = pd.DataFrame()
    except Exception:
        pass
    for k in ["fin", "bs", "cf", "qfin", "qbs", "qcf"]:
        out.setdefault(k, pd.DataFrame())
    return out


def _yf_find(df, yf_label):
    """Resolve a display/yf label to an actual DataFrame index label, or None."""
    if df is None or df.empty:
        return None
    norm = {str(idx).strip().lower(): idx for idx in df.index}
    candidates = [yf_label.strip().lower()]
    candidates += YF_ALIASES.get(yf_label.strip().lower(), [])
    for c in candidates:
        if c in norm:
            return norm[c]
    return None


def _yf_series(df, yf_label, cols):
    idx = _yf_find(df, yf_label)
    if idx is None:
        return {c: None for c in cols}
    row = df.loc[idx]
    out = {}
    for c in cols:
        try:
            v = row[c]
            out[c] = None if pd.isna(v) else float(v)
        except Exception:
            out[c] = None
    return out


def _col_label(ts, period):
    try:
        if period == "Annual":
            return f"FY{ts.year}"
        return f"Q{(ts.month - 1) // 3 + 1} {ts.year}"
    except Exception:
        return str(ts)


def prepare_yf_statement(df, spec, period, n, revenue_df=None, fcf_df=None):
    """Return (col_labels_ascending, columns, rows). Rows include calc rows."""
    if df is None or df.empty:
        return [], [], []
    cols = list(df.columns)[:n]            # yfinance: most recent first
    cols = sorted(cols)                    # ascending for display
    labels = [_col_label(c, period) for c in cols]

    # revenue per column for margins (match by column label)
    rev_by_label = {}
    if revenue_df is not None and not revenue_df.empty:
        rcols = sorted(list(revenue_df.columns)[:n])
        rser = _yf_series(revenue_df, "Total Revenue", rcols)
        rev_by_label = {_col_label(c, period): rser.get(c) for c in rcols}

    rows = []
    for yf_label, disp, style, special, direction in spec:
        if yf_label == "__NET_MARGIN__":
            ni = next((r for r in rows if r["disp"] == "Net Income"), None)
            vals = {}
            for c in cols:
                lab = _col_label(c, period)
                rev = rev_by_label.get(lab)
                niv = ni["raw"].get(c) if ni else None
                vals[c] = (niv / rev * 100) if (ni and niv is not None and rev not in (None, 0)) else None
            rows.append({"disp": disp, "style": style, "special": special,
                         "dir": direction, "raw": vals})
            continue
        if yf_label == "__FCF_MARGIN__":
            fcf = next((r for r in rows if r["disp"] == "Free Cash Flow"), None)
            vals = {}
            for c in cols:
                lab = _col_label(c, period)
                rev = rev_by_label.get(lab)
                fv = fcf["raw"].get(c) if fcf else None
                vals[c] = (fv / rev * 100) if (fcf and fv is not None and rev not in (None, 0)) else None
            rows.append({"disp": disp, "style": style, "special": special,
                         "dir": direction, "raw": vals})
            continue
        series = _yf_series(df, yf_label, cols)
        if all(v is None for v in series.values()):
            continue
        rows.append({"disp": disp, "style": style, "special": special,
                     "dir": direction, "raw": series})
    return labels, cols, rows


def _yf_fmt(v, special):
    if v is None:
        return "\u2014", False
    if special == "eps":
        return f"${v:,.2f}", v < 0
    if special == "pct":
        return f"{v:.1f}%", v < 0
    return f"{v / 1e6:,.2f}", v < 0


def _yf_pct(raw, cols, direction):
    if len(cols) < 2:
        return "\u2014", NM_TEXT
    a, b = raw.get(cols[-2]), raw.get(cols[-1])
    if a is None or a == 0 or b is None:
        return "\u2014", NM_TEXT
    p = (b - a) / abs(a)
    txt = f"{p * 100:+.1f}%"
    if direction == "up":
        col = GREEN_TEXT if p > 0 else (RED_TEXT if p < 0 else NM_TEXT)
    elif direction == "down":
        col = RED_TEXT if p > 0 else (GREEN_TEXT if p < 0 else NM_TEXT)
    else:
        col = NM_TEXT
    return txt, col


def render_yf_html(title, labels, cols, rows, period, balance_extra=None):
    head = "".join(f"<th style='text-align:right;padding:6px 10px;'>{h}</th>" for h in labels)
    pct_hdr = f"% \u0394 ({labels[-2]}\u2192{labels[-1]})" if len(labels) >= 2 else "% \u0394"
    html = [
        f"<div style='font-weight:700;font-size:16px;margin:14px 0 4px;'>{title}</div>",
        "<table style='width:100%;border-collapse:collapse;font-size:13px;'>",
        f"<tr style='border-bottom:2px solid #cbd5e1;'>"
        f"<th style='text-align:left;padding:6px 10px;'>Line Item</th>{head}"
        f"<th style='text-align:right;padding:6px 10px;'>{pct_hdr}</th></tr>",
    ]
    for r in rows:
        bg = YF_BG.get(r["style"], "transparent")
        weight = "700" if r["style"] in ("bold", "sub") else "400"
        cells = ""
        for c in cols:
            txt, isneg = _yf_fmt(r["raw"].get(c), r["special"])
            color = RED_TEXT if isneg else "#111827"
            cells += f"<td style='text-align:right;padding:5px 10px;color:{color};'>{txt}</td>"
        ptxt, pcol = _yf_pct(r["raw"], cols, r["dir"])
        html.append(
            f"<tr style='background:{bg};border-bottom:1px solid #f1f5f9;'>"
            f"<td style='padding:5px 10px;font-weight:{weight};'>{r['disp']}</td>{cells}"
            f"<td style='text-align:right;padding:5px 10px;color:{pcol};font-weight:{weight};'>{ptxt}</td></tr>"
        )
    if balance_extra is not None:
        assets, tl, te = balance_extra
        chk = ""
        for c in cols:
            a = assets.get(c)
            le = (tl.get(c) or 0) + (te.get(c) or 0) if (tl.get(c) is not None or te.get(c) is not None) else None
            if a is None or le is None or a == 0:
                chk += "<td style='text-align:right;padding:5px 10px;color:#6b7280;'>\u2014</td>"
            elif abs(a - le) < 0.01 * abs(a):
                chk += f"<td style='text-align:right;padding:5px 10px;color:{GREEN_TEXT};font-weight:700;'>\u2713 Balanced</td>"
            else:
                chk += f"<td style='text-align:right;padding:5px 10px;color:{AMBER_TEXT};font-weight:700;'>\u26a0 Review</td>"
        html.append(f"<tr style='background:{YF_BG['sub']};'>"
                    f"<td style='padding:5px 10px;font-weight:700;'>Balance Check</td>{chk}"
                    f"<td style='padding:5px 10px;'></td></tr>")
    html.append("</table>")
    return "".join(html)


def _xl_write_yf(ws, labels, cols, rows, balance_extra=None):
    bold = Font(bold=True)
    red = Font(color="b91c1c")
    red_bold = Font(bold=True, color="b91c1c")
    sub_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    bold_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    header = ["Line Item"] + labels + ([f"% Change {labels[-2]}->{labels[-1]}"] if len(labels) >= 2 else ["% Change"])
    for j, h in enumerate(header, start=1):
        ws.cell(row=1, column=j, value=h).font = bold
    rn = 2
    for r in rows:
        is_b = r["style"] in ("bold", "sub")
        ws.cell(row=rn, column=1, value=r["disp"]).font = bold if is_b else None
        for k, c in enumerate(cols, start=2):
            v = r["raw"].get(c)
            if v is None:
                ws.cell(row=rn, column=k, value=None)
                continue
            if r["special"] == "eps":
                out = round(v, 2)
            elif r["special"] == "pct":
                out = round(v, 1)
            else:
                out = round(v / 1e6, 2)
            cell = ws.cell(row=rn, column=k, value=out)
            if out < 0:
                cell.font = red_bold if is_b else red
            elif is_b:
                cell.font = bold
        ptxt, _ = _yf_pct(r["raw"], cols, r["dir"])
        ws.cell(row=rn, column=len(cols) + 2, value=(ptxt if ptxt != "\u2014" else None))
        if is_b:
            fill = sub_fill if r["style"] == "sub" else bold_fill
            for col in range(1, len(cols) + 3):
                ws.cell(row=rn, column=col).fill = fill
        rn += 1
    if balance_extra is not None:
        assets, tl, te = balance_extra
        ws.cell(row=rn, column=1, value="Balance Check").font = bold
        for k, c in enumerate(cols, start=2):
            a = assets.get(c)
            le = (tl.get(c) or 0) + (te.get(c) or 0) if (tl.get(c) is not None or te.get(c) is not None) else None
            txt = "\u2014" if (a is None or le is None or a == 0) else ("Balanced" if abs(a - le) < 0.01 * abs(a) else "Review")
            ws.cell(row=rn, column=k, value=txt).font = bold
    ws.column_dimensions["A"].width = 34
    for k in range(2, len(cols) + 3):
        ws.column_dimensions[ws.cell(row=1, column=k).column_letter].width = 15


def build_yf_excel(ticker, stmts):
    wb = Workbook()
    plans = [
        ("Income Statement - Annual", stmts["fin"], INCOME_YF, "Annual", 3),
        ("Income Statement - Quarterly", stmts["qfin"], INCOME_YF, "Quarterly", 4),
        ("Balance Sheet - Annual", stmts["bs"], BALANCE_YF, "Annual", 3),
        ("Balance Sheet - Quarterly", stmts["qbs"], BALANCE_YF, "Quarterly", 4),
        ("Cash Flow - Annual", stmts["cf"], CASHFLOW_YF, "Annual", 3),
        ("Cash Flow - Quarterly", stmts["qcf"], CASHFLOW_YF, "Quarterly", 4),
    ]
    first = True
    for sheet_name, df, spec, period, n in plans:
        ws = wb.active if first else wb.create_sheet(sheet_name)
        if first:
            ws.title = sheet_name
            first = False
        rev_df = stmts["fin"] if period == "Annual" else stmts["qfin"]
        labels, cols, rows = prepare_yf_statement(df, spec, period, n, revenue_df=rev_df)
        if not labels:
            ws.cell(row=1, column=1, value="No data available from Yahoo Finance.")
            continue
        extra = None
        if "Balance" in sheet_name:
            extra = (_yf_series(df, "Total Assets", cols),
                     _yf_series(df, "Total Liabilities Net Minority Interest", cols),
                     _yf_series(df, "Total Equity Gross Minority Interest", cols))
        _xl_write_yf(ws, labels, cols, rows, balance_extra=extra)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Credit Summary", "Trends", "Peer Ranking", "Raw Data", "Financial Statements"
])

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

# ----------------------------------------------------------------------------
# TAB 5 \u2014 Financial Statements
# ----------------------------------------------------------------------------
with tab5:
    st.subheader("Financial Statements")
    fs_airline = st.selectbox("Select Airline", airlines, key="fs_airline")
    fs_period = st.radio("Period", ["Annual", "Quarterly"], horizontal=True, key="fs_period")
    fs_ticker = data[fs_airline]["ticker"]
    n_periods = 3 if fs_period == "Annual" else 4

    stmts = fetch_yf_statements(fs_ticker)

    inc_df = stmts["fin"] if fs_period == "Annual" else stmts["qfin"]
    bal_df = stmts["bs"] if fs_period == "Annual" else stmts["qbs"]
    cf_df = stmts["cf"] if fs_period == "Annual" else stmts["qcf"]

    if inc_df.empty and bal_df.empty and cf_df.empty:
        st.error(
            f"No financial-statement data returned from Yahoo Finance for {fs_ticker}. "
            "This is common for delisted or recently acquired carriers, or when Yahoo "
            "rate-limits the request. Try again shortly or pick another airline."
        )
    else:
        xlsx_bytes = build_yf_excel(fs_ticker, stmts)
        st.download_button(
            f"\u2b07 Download {fs_airline} Financials to Excel",
            data=xlsx_bytes,
            file_name=f"{fs_ticker}_Financials_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="fs_download",
        )
        st.caption(
            "Source: Yahoo Finance (yfinance). Values in $millions (2dp) unless marked "
            "EPS ($/share) or % (margins). Rows absent from Yahoo's data are skipped. "
            "Excel export includes both Annual and Quarterly sheets. "
            + ("Note: quarterly %\u0394 is period-over-period (seasonal), not year-on-year."
               if fs_period == "Quarterly" else "")
        )

        i_labels, i_cols, i_rows = prepare_yf_statement(inc_df, INCOME_YF, fs_period, n_periods, revenue_df=inc_df)
        if i_labels:
            st.markdown(render_yf_html("Income Statement", i_labels, i_cols, i_rows, fs_period),
                        unsafe_allow_html=True)
        else:
            st.info("No income statement data available.")

        b_labels, b_cols, b_rows = prepare_yf_statement(bal_df, BALANCE_YF, fs_period, n_periods)
        if b_labels:
            extra = (_yf_series(bal_df, "Total Assets", b_cols),
                     _yf_series(bal_df, "Total Liabilities Net Minority Interest", b_cols),
                     _yf_series(bal_df, "Total Equity Gross Minority Interest", b_cols))
            st.markdown(render_yf_html("Balance Sheet", b_labels, b_cols, b_rows, fs_period,
                                       balance_extra=extra), unsafe_allow_html=True)
        else:
            st.info("No balance sheet data available.")

        c_labels, c_cols, c_rows = prepare_yf_statement(cf_df, CASHFLOW_YF, fs_period, n_periods, revenue_df=inc_df)
        if c_labels:
            st.markdown(render_yf_html("Cash Flow Statement", c_labels, c_cols, c_rows, fs_period),
                        unsafe_allow_html=True)
        else:
            st.info("No cash flow data available.")

# To run: streamlit run app.py
