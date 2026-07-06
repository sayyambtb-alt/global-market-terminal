import streamlit as st
import pandas as pd
import yfinance as yf
import altair as alt
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
import json
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Global Market Terminal", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
.metric-card{background:#1e272e;padding:15px;border-radius:8px;border-left:5px solid #0abde3;}
.terminal-header{font-family:'Courier New',monospace;color:#10ac84;}
.news-card{background:#2f3640;padding:15px;border-radius:5px;margin-bottom:10px;border-left:3px solid #e1b12c;}
.risk-card{background:#2d3436;padding:20px;border-radius:8px;border-top:4px solid #e84393;text-align:center;min-height:180px;display:flex;flex-direction:column;justify-content:center;}
.nlp-bull{border-left:4px solid #2ed573;background:#1a3a2a;padding:10px;margin-bottom:8px;border-radius:4px;color:#fff;}
.nlp-bear{border-left:4px solid #ff4757;background:#3a1a1a;padding:10px;margin-bottom:8px;border-radius:4px;color:#fff;}
.nlp-neut{border-left:4px solid #7f8fa6;background:#2d3436;padding:10px;margin-bottom:8px;border-radius:4px;color:#fff;}
.sig-card{border-radius:8px;padding:14px 16px;text-align:center;border-top:3px solid;}
</style>""", unsafe_allow_html=True)

# ── CACHING ──────────────────────────────────────────────────────────────────
@st.cache_data(ttl=3600)
def cached_history(ticker, period="1y"):
    d = yf.Ticker(ticker).history(period=period)
    return d.dropna(subset=["Close"]) if not d.empty else d

@st.cache_data(ttl=3600)
def cached_info(ticker):
    return yf.Ticker(ticker).info

@st.cache_data(ttl=3600)
def cached_income_stmt(ticker):
    obj = yf.Ticker(ticker)
    return obj.income_stmt if (hasattr(obj,"income_stmt") and not obj.income_stmt.empty) else obj.financials

@st.cache_data(ttl=3600)
def cached_balance_sheet(ticker):
    obj = yf.Ticker(ticker)
    return obj.balance_sheet if (hasattr(obj,"balance_sheet") and not obj.balance_sheet.empty) else None

@st.cache_data(ttl=3600)
def cached_cashflow(ticker):
    obj = yf.Ticker(ticker)
    return obj.cashflow if (hasattr(obj,"cashflow") and not obj.cashflow.empty) else None

@st.cache_data(ttl=3600)
def cached_upgrades(ticker):
    return yf.Ticker(ticker).upgrades_downgrades

@st.cache_data(ttl=3600)
def cached_recs(ticker):
    return yf.Ticker(ticker).recommendations

@st.cache_data(ttl=3600)
def fetch_macro_yield(ticker, fallback):
    if ticker is None: return fallback
    try:
        d = yf.Ticker(ticker).history(period="1d")
        if not d.empty: return float(d["Close"].iloc[-1])
    except: pass
    return fallback

@st.cache_data(ttl=3600)
def get_peer_multiples(peers):
    rows = []
    for sym in peers:
        try:
            info = cached_info(sym)
            name = info.get("shortName") or info.get("longName") or sym
            price = info.get("currentPrice") or info.get("regularMarketPrice")
            pe    = info.get("trailingPE")
            eveb  = info.get("enterpriseToEbitda")
            pb    = info.get("priceToBook")
            mc    = info.get("marketCap")
            hi52  = info.get("fiftyTwoWeekHigh")
            lo52  = info.get("fiftyTwoWeekLow")
            rows.append({"Ticker":sym,"Company":name[:22] if name else sym,
                         "Price":price,"P/E":pe,"EV/EBITDA":eveb,
                         "P/B":pb,"Mkt Cap":mc,"52W High":hi52,"52W Low":lo52})
        except: pass
    return pd.DataFrame(rows)

# ── SEARCH ────────────────────────────────────────────────────────────────────
def search_ticker_candidates(query, max_results=8):
    if not query: return []
    url = (f"https://query2.finance.yahoo.com/v1/finance/search"
           f"?q={urllib.parse.quote(query.strip())}&quotesCount={max_results}&newsCount=0&enableFuzzyQuery=true")
    try:
        req = urllib.request.Request(url, headers={"User-Agent":"Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=5) as r:
            data = json.loads(r.read().decode())
        out = []
        for q in data.get("quotes",[]):
            sym = q.get("symbol","")
            if sym:
                out.append({"symbol":sym,
                    "name":   q.get("longname") or q.get("shortname") or sym,
                    "exchange":q.get("exchDisp") or q.get("exchange") or "Unknown",
                    "type":   q.get("typeDisp") or q.get("quoteType") or "Equity",
                    "label":  f"{sym}  |  {q.get('longname') or q.get('shortname') or sym}  |  {q.get('exchDisp') or q.get('exchange') or 'Unknown'}  ({q.get('typeDisp') or 'Equity'})"})
        return out
    except: return []

def ticker_search_widget(default_query="Reliance"):
    ck,lk,sk,qk = "gc","gl","gs","gq"
    if qk not in st.session_state: st.session_state[qk] = default_query
    query = st.text_input("Search Company or Ticker  (e.g. Reliance, Apple, TCS)", key=qk)
    q = query.strip()
    if not q: return None
    if q != st.session_state.get(lk,""):
        with st.spinner("Searching…"):
            cands = search_ticker_candidates(q)
        st.session_state[ck]=cands; st.session_state[lk]=q
        if cands: st.session_state[sk]=cands[0]["label"]
    cands = st.session_state.get(ck,[])
    if not cands:
        st.warning(f"No results for **{q}**."); return None
    labels=[c["label"] for c in cands]
    if st.session_state.get(sk) not in labels: st.session_state[sk]=labels[0]
    st.selectbox(f"Found **{len(cands)}** matches — select security:", options=labels, key=sk)
    chosen = cands[labels.index(st.session_state[sk])]
    st.success(f"✅ **{chosen['symbol']}** — {chosen['name']} ({chosen['exchange']})")
    return chosen["symbol"]

# ── SHARED HELPERS ────────────────────────────────────────────────────────────
EXCHANGE_MAP = {".NS":"NSE",".BO":"BOM",".L":"LON",".T":"TYO",".HK":"HKG"}
NLP_POS = ["beat","beats","record","growth","profit","surge","rally","upgrade","buy","outperform","strong",
           "raises","raised","dividend","acquisition","expansion","wins","bullish","upside","robust",
           "target raised","revenue beat","earnings beat","recovery"]
NLP_NEG = ["miss","misses","loss","decline","cut","downgrade","sell","underperform","weak","falls","drops",
           "concern","warning","debt","bearish","downturn","penalty","fine","lawsuit","lowers guidance",
           "margin pressure","disappoints","below expectations"]

def ticker_meta(sym):
    is_in = sym.endswith(".NS") or sym.endswith(".BO")
    clean = sym.split(".")[0] if "." in sym else sym
    gex   = next((v for k,v in EXCHANGE_MAP.items() if sym.endswith(k)), None)
    return (is_in, "₹" if is_in else "$", clean,
            "^NSEI" if is_in else "^GSPC",
            "Nifty 50" if is_in else "S&P 500", gex)

def safe_get(df, *labels, col=0):
    if df is None or df.empty: return None
    for l in labels:
        if l in df.index:
            v = df.loc[l].iloc[col]
            if pd.notna(v): return float(v)
    return None

def fmt_money(val, is_indian=True):
    if val is None: return "N/A"
    if is_indian:
        cr = val/10_000_000
        return f"₹{cr/100_000:.2f}L Cr" if abs(cr)>=100_000 else f"₹{cr:,.0f} Cr"
    b = val/1_000_000_000
    return f"${b:.2f}B" if abs(b)>=1 else f"${val/1_000_000:.0f}M"

def metric_card(label, value, tooltip=None, color="#0abde3"):
    tip = f' <span title="{tooltip}" style="cursor:help;color:#7f8fa6;">ⓘ</span>' if tooltip else ""
    return (f'<div style="background:#1e272e;border-radius:8px;padding:18px 20px;text-align:center;'
            f'border-top:3px solid {color};height:100%;">'
            f'<div style="color:#7f8fa6;font-size:13px;margin-bottom:8px;">{label}{tip}</div>'
            f'<div style="color:#fff;font-size:26px;font-weight:700;">{value}</div></div>')

def format_statement(df, t):
    if df is None or df.empty: return pd.DataFrame()
    H = {
        "IS": ["Total Revenue","Operating Revenue","Cost Of Revenue","Gross Profit",
               "Research And Development","Selling General And Administration",
               "Operating Expense","Operating Income","Normalized EBITDA",
               "Interest Expense","Pretax Income","Tax Provision",
               "Net Income","Net Income Common Stockholders","Diluted EPS"],
        "BS": [
            # ASSETS
            "Total Assets",
            "Current Assets","Cash And Cash Equivalents","Short Term Investments",
            "Accounts Receivable","Inventory","Other Current Assets",
            "Total Non Current Assets","Net PPE","Goodwill",
            "Other Intangible Assets","Long Term Investments",
            # LIABILITIES
            "Total Liabilities Net Minority Interest",
            "Current Liabilities","Accounts Payable","Current Debt",
            "Other Current Liabilities",
            "Total Non Current Liabilities Net Minority Interest",
            "Long Term Debt","Other Non Current Liabilities",
            # EQUITY
            "Stockholders Equity","Common Stock","Retained Earnings",
            # SUMMARY
            "Total Debt","Net Debt","Working Capital"],
        "CF": [
            # OPERATING
            "Operating Cash Flow","Depreciation And Amortization",
            "Change In Working Capital","Change In Receivables",
            "Change In Inventory","Change In Payable",
            # FREE CASH FLOW
            "Capital Expenditure","Free Cash Flow",
            # INVESTING
            "Investing Cash Flow","Purchase Of Investment",
            "Sale Of Investment","Net PPE Purchase And Sale",
            # FINANCING
            "Financing Cash Flow","Repayment Of Debt","Issuance Of Debt",
            "Repurchase Of Capital Stock","Cash Dividends Paid",
            # SUMMARY
            "End Cash Position"]
    }
    order = H[t]
    rows  = [x for x in order if x in df.index]   # curated only — no "others" dump
    df2   = df.reindex(rows).iloc[:,:3].copy()
    df2.columns = [c.strftime('%Y') if hasattr(c,'strftime') else str(c)[:4] for c in df2.columns]
    return df2.dropna(how='all')

# ── TECHNICAL INDICATOR FUNCTIONS ─────────────────────────────────────────────
def compute_rsi(prices, period=14):
    delta    = prices.diff()
    gain     = delta.clip(lower=0).rolling(period).mean()
    loss     = (-delta.clip(upper=0)).rolling(period).mean()
    rs       = gain / loss.replace(0, float("nan"))
    return 100 - (100 / (1 + rs))

def compute_macd(prices, fast=12, slow=26, sig=9):
    macd   = prices.ewm(span=fast,adjust=False).mean() - prices.ewm(span=slow,adjust=False).mean()
    signal = macd.ewm(span=sig,adjust=False).mean()
    return macd, signal, macd - signal

def compute_bollinger(prices, period=20, sd=2):
    sma = prices.rolling(period).mean()
    std = prices.rolling(period).std()
    return sma+sd*std, sma, sma-sd*std

def dcf_iv(base_fcf, fcf_growth, wacc, tg, cash, debt, shares):
    if wacc<=tg or wacc<=0 or shares<=0: return None
    pv, fcfs = 0, []
    for yr in range(1,6):
        pf = base_fcf*((1+fcf_growth)**yr)
        pv += pf/((1+wacc)**yr); fcfs.append(pf)
    tv = (fcfs[-1]*(1+tg))/(wacc-tg)
    return (pv + tv/((1+wacc)**5) + cash - debt) / shares

# ── SECTOR PEER BASKETS ───────────────────────────────────────────────────────
PEERS_IN = {
    "Automobile Manufacturers":["MARUTI.NS","TATAMOTORS.NS","M&M.NS","BAJAJ-AUTO.NS","HEROMOTOCO.NS"],
    "Auto Components":["MOTHERSON.NS","BOSCHLTD.NS","BALKRISIND.NS","MRF.NS","APOLLOTYRE.NS"],
    "Information Technology":["TCS.NS","INFY.NS","WIPRO.NS","HCLTECH.NS","TECHM.NS"],
    "Banks":["HDFCBANK.NS","ICICIBANK.NS","KOTAKBANK.NS","AXISBANK.NS","SBIN.NS"],
    "Financial Services":["BAJFINANCE.NS","BAJAJFINSV.NS","MUTHOOTFIN.NS","CHOLAFIN.NS","HDFCAMC.NS"],
    "Pharmaceuticals":["SUNPHARMA.NS","CIPLA.NS","DRREDDY.NS","DIVISLAB.NS","AUROPHARMA.NS"],
    "Consumer Defensive":["HINDUNILVR.NS","NESTLEIND.NS","BRITANNIA.NS","DABUR.NS","MARICO.NS"],
    "Energy":["RELIANCE.NS","ONGC.NS","BPCL.NS","IOC.NS","GAIL.NS"],
    "Basic Materials":["TATASTEEL.NS","JSWSTEEL.NS","HINDALCO.NS","VEDL.NS","SAIL.NS"],
    "Industrials":["LT.NS","SIEMENS.NS","ABB.NS","BHEL.NS","CUMMINSIND.NS"],
    "Telecom":["BHARTIARTL.NS","TATACOMM.NS","IDEA.NS","MTNL.NS","RAILTEL.NS"],
    "Entertainment":["TIPSMUSIC.NS","SAREGAMA.NS","ZEEL.NS","PVRINOX.NS","SUNTV.NS"],
    "Media":["ZEEL.NS","SUNTV.NS","NDTV.NS","DISHTV.NS","TVSM.NS"],
    "Broadcasting":["ZEEL.NS","SUNTV.NS","NDTV.NS","DISHTV.NS","TIPSMUSIC.NS"],
    "Internet Content":["INDIAMART.NS","JUSTDIAL.NS","NAUKRI.NS","POLICYBZR.NS","ZOMATO.NS"],
    "Communication Services":["BHARTIARTL.NS","TATACOMM.NS","INDIAMART.NS"],
    "Real Estate":["DLF.NS","GODREJPROP.NS","OBEROIRLTY.NS","PRESTIGE.NS","PHOENIXLTD.NS"],
    "Utilities":["NTPC.NS","POWERGRID.NS","TATAPOWER.NS","CESC.NS","TORNTPOWER.NS"],
    "Consumer Cyclical":["TITAN.NS","TRENT.NS","DMART.NS","PAGEIND.NS","VOLTAS.NS"],
    "Insurance":["SBILIFE.NS","HDFCLIFE.NS","ICICIPRULI.NS","LICI.NS"],
    "Chemicals":["PIDILITIND.NS","AARTI.NS","DEEPAKNTR.NS","NAVINFLUOR.NS","ALKYLAMINE.NS"],
    "Textiles":["PAGEIND.NS","VARDHMAN.NS","WELSPUNIND.NS","RAYMOND.NS","ARVIND.NS"],
    "Hotels":["INDHOTEL.NS","EIHOTEL.NS","LEMERIDIEN.NS","CHALET.NS","MAHINDRAHOLIDAYS.NS"],
}
PEERS_US = {
    "Technology":["AAPL","MSFT","GOOGL","META","NVDA"],
    "Financials":["JPM","BAC","WFC","GS","MS"],
    "Healthcare":["JNJ","PFE","UNH","ABBV","MRK"],
    "Consumer Cyclical":["AMZN","TSLA","HD","MCD","NKE"],
    "Energy":["XOM","CVX","COP","SLB","EOG"],
    "Industrials":["HON","GE","CAT","RTX","LMT"],
    "Communication Services":["GOOGL","META","NFLX","DIS","CMCSA"],
    "Basic Materials":["LIN","APD","ECL","SHW","FCX"],
    "Real Estate":["AMT","PLD","CCI","EQIX","PSA"],
    "Utilities":["NEE","DUK","SO","D","AEP"],
    "Consumer Defensive":["PG","KO","PEP","WMT","COST"],
}

# Explicit industry → basket mapping (prevents wrong word-overlap matches)
INDUSTRY_EXACT = {
    "auto manufacturers":                       "Automobile Manufacturers",
    "auto parts":                               "Auto Components",
    "truck manufacturing":                      "Automobile Manufacturers",
    "farm & heavy construction machinery":      "Industrials",
    "electrical equipment & parts":            "Industrials",
    "engineering & construction":              "Industrials",
    "banks—regional":                           "Banks",
    "banks—diversified":                        "Banks",
    "diversified banks":                        "Banks",
    "insurance—life":                           "Insurance",
    "insurance—diversified":                    "Insurance",
    "software—application":                     "Information Technology",
    "software—infrastructure":                  "Information Technology",
    "information technology services":          "Information Technology",
    "drug manufacturers—general":               "Pharmaceuticals",
    "drug manufacturers—specialty & generic":   "Pharmaceuticals",
    "diagnostics & research":                   "Pharmaceuticals",
    "biotechnology":                            "Pharmaceuticals",
    "packaged foods":                           "Consumer Defensive",
    "household & personal products":            "Consumer Defensive",
    "grocery stores":                           "Consumer Defensive",
    "oil & gas integrated":                     "Energy",
    "oil & gas e&p":                            "Energy",
    "oil & gas refining & marketing":           "Energy",
    "oil & gas midstream":                      "Energy",
    "steel":                                    "Basic Materials",
    "aluminum":                                 "Basic Materials",
    "specialty chemicals":                      "Chemicals",
    "basic chemicals":                          "Chemicals",
    "entertainment":                            "Entertainment",
    "broadcasting":                             "Broadcasting",
    "publishing":                               "Media",
    "telecom services":                         "Telecom",
    "internet content & information":           "Internet Content",
    "real estate—development":                  "Real Estate",
    "real estate—diversified":                  "Real Estate",
    "utilities—regulated electric":             "Utilities",
    "utilities—regulated gas":                  "Utilities",
    "apparel manufacturing":                    "Textiles",
    "textiles":                                 "Textiles",
    "hotels & motels":                          "Hotels",
    "lodging":                                  "Hotels",
    "specialty retail":                         "Consumer Cyclical",
    "apparel retail":                           "Consumer Cyclical",
    "consumer electronics":                     "Consumer Cyclical",
    "leisure":                                  "Consumer Cyclical",
}

def get_peers(ticker_sym, sector, industry, is_indian):
    pool = PEERS_IN if is_indian else PEERS_US

    # Step 1: exact industry → basket lookup (most reliable)
    basket_name = INDUSTRY_EXACT.get(industry.lower())
    if basket_name and basket_name in pool:
        peers = pool[basket_name]
        return [p for p in peers if p.split(".")[0] != ticker_sym.split(".")[0]][:5]

    # Step 2: word-score match against industry
    def score(bname, query):
        stop = {"and","the","of","in","services","manufacturing","manufacturers","&"}
        qw = {w for w in query.lower().replace("&"," ").split() if w not in stop and len(w)>3}
        bw = {w for w in bname.lower().replace("&"," ").split()  if w not in stop and len(w)>3}
        return len(qw & bw)

    best, best_score = [], 0
    for bname, lst in pool.items():
        s = score(bname, industry)
        if s > best_score: best_score = s; best = lst

    # Step 3: fall back to sector word-score
    if best_score == 0:
        for bname, lst in pool.items():
            s = score(bname, sector)
            if s > best_score: best_score = s; best = lst

    if not best and pool:
        best = list(pool.values())[0]

    return [p for p in best if p.split(".")[0] != ticker_sym.split(".")[0]][:5]

# ── EXCEL STYLE CONSTANTS ─────────────────────────────────────────────────────
_H  = PatternFill("solid", fgColor="1A3A5C")
_SH = PatternFill("solid", fgColor="2E5E8E")
_IN = PatternFill("solid", fgColor="FFF2CC")
_CA = PatternFill("solid", fgColor="DEEAF1")
_PO = PatternFill("solid", fgColor="E2EFDA")
_NE = PatternFill("solid", fgColor="FCE4D6")
_AL = PatternFill("solid", fgColor="F5F5F5")
_WH = PatternFill("solid", fgColor="FFFFFF")
_FH = Font(color="FFFFFF",bold=True,size=10)
_FB = Font(bold=True)
_FI = Font(color="000080",bold=True)
_FR = Font(color="C00000",bold=True)
_FG = Font(color="375623",bold=True)
_BD = Border(left=Side(style="thin"),right=Side(style="thin"),top=Side(style="thin"),bottom=Side(style="thin"))
_CT = Alignment(horizontal="center",vertical="center",wrap_text=True)
_LT = Alignment(horizontal="left",  vertical="center")
_RT = Alignment(horizontal="right", vertical="center")

def _c(ws,row,col,val=None,fill=None,font=None,align=None,fmt=None,bdr=_BD):
    cell = ws.cell(row=row,column=col,value=val)
    if fill:  cell.fill=fill
    if font:  cell.font=font
    if align: cell.alignment=align
    if fmt:   cell.number_format=fmt
    if bdr:   cell.border=bdr
    return cell

def _hrow(ws,row,text,ncols=8):
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=ncols)
    cell = ws.cell(row=row,column=1,value=text)
    cell.fill=_H; cell.font=Font(color="FFFFFF",bold=True,size=11)
    cell.alignment=_CT; ws.row_dimensions[row].height=22
    return cell

def _shrow(ws,row,text,ncols=8):
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=ncols)
    cell = ws.cell(row=row,column=1,value=text)
    cell.fill=_SH; cell.font=Font(color="FFFFFF",bold=True,size=10)
    cell.alignment=_LT; ws.row_dimensions[row].height=18
    return cell

# ── HEADER + NAV ──────────────────────────────────────────────────────────────
st.markdown("<h1 class='terminal-header'>🖥️ GLOBAL MARKET TERMINAL</h1>", unsafe_allow_html=True)
st.caption("Institutional-Grade Financial Intelligence — One-Stop Analysis Platform")
st.sidebar.header("📁 Navigation Desk")
menu = st.sidebar.radio("Select Module", [
    "📋 Fundamental Deep Dive",
    "🎯 Valuation Modeler",
    "📈 Technical Analysis",
    "📊 Macro & CAPM",
])

# ═══════════════════════════════════════════════════════════════════════════
# MODULE 1 — FUNDAMENTAL DEEP DIVE (COMPREHENSIVE)
# ═══════════════════════════════════════════════════════════════════════════
if menu == "📋 Fundamental Deep Dive":
    st.markdown("## 📋 Fundamental Deep Dive")
    st.markdown("*Corporate health, alpha generation, analyst tracking, and working capital intelligence.*")
    ticker_sym = ticker_search_widget()
    if not ticker_sym: st.stop()
    is_indian, currency, clean_sym, bench_sym, bench_name, goog_exc = ticker_meta(ticker_sym)
    st.markdown("---")

    # ── Corporate Overview ─────────────────────────────────────────────────
    with st.spinner("Loading corporate data…"):
        try:
            info         = cached_info(ticker_sym)
            company_name = info.get("longName", ticker_sym)
            sector       = info.get("sector","Unknown")
            industry     = info.get("industry","Unknown")
            summary      = info.get("longBusinessSummary","No summary available.")
            curr_px      = info.get("currentPrice") or info.get("regularMarketPrice") or 0
        except: info={}; company_name=ticker_sym; sector=industry=summary=""; curr_px=0

    st.markdown(f"### 🏢 {company_name}")
    st.markdown(f"**{sector}** — *{industry}*")
    with st.expander("Read Full Company Brief"):
        st.write(summary)
    st.markdown("---")

    # ── Institutional Telemetry ────────────────────────────────────────────
    st.markdown("### 📈 Institutional Telemetry")
    st.caption("Idiosyncratic Alpha = today's stock return minus today's benchmark return (live computation). Volume footprint = today's volume vs 20-day average. All real, no estimates.")
    with st.spinner("Calculating alpha and volume footprint…"):
        try:
            sd = cached_history(ticker_sym,"1mo")
            bd = cached_history(bench_sym,"5d")
            if len(sd)>=2 and len(bd)>=2:
                sc,sp = float(sd["Close"].iloc[-1]),float(sd["Close"].iloc[-2])
                bc,bp = float(bd["Close"].iloc[-1]),float(bd["Close"].iloc[-2])
                pct   = ((sc-sp)/sp)*100
                bpct  = ((bc-bp)/bp)*100
                alpha = pct - bpct
                avgv  = sd["Volume"].iloc[-21:-1].mean() if len(sd)>=20 else sd["Volume"].mean()
                vsurge= (float(sd["Volume"].iloc[-1])/avgv)*100 if avgv>0 else 100
                sc_c  = "#2ed573" if pct>=0 else "#ff4757"
                al_c  = "#2ed573" if alpha>=0 else "#ff4757"
                vl_c  = "#0abde3" if vsurge>150 else "#e1b12c" if vsurge>80 else "#7f8fa6"
                vs_lbl= ("MASSIVE ACCUMULATION" if vsurge>200 else "ELEVATED INSTITUTIONAL FLOW" if vsurge>120 else "NORMAL RETAIL FLOW")
                c1,c2,c3 = st.columns(3)
                with c1:
                    st.markdown(f"<div class='metric-card' style='border-left-color:{sc_c};'><h4 style='margin:0;color:#7f8fa6;'>Live Price</h4><p style='font-size:26px;font-weight:bold;margin:5px 0;color:#ffffff;'>{currency}{sc:,.2f}</p><span style='color:{sc_c};font-weight:bold;'>{pct:+.2f}% Today</span></div>",unsafe_allow_html=True)
                with c2:
                    st.markdown(f"<div class='metric-card' style='border-left-color:{al_c};'><h4 style='margin:0;color:#7f8fa6;'>Idiosyncratic Alpha</h4><p style='font-size:26px;font-weight:bold;margin:5px 0;color:{al_c};'>{alpha:+.2f}%</p><span style='color:#b2bec3;font-size:12px;'>vs {bench_name} ({bpct:+.2f}%)</span></div>",unsafe_allow_html=True)
                with c3:
                    st.markdown(f"<div class='metric-card' style='border-left-color:{vl_c};'><h4 style='margin:0;color:#7f8fa6;'>Volume Footprint</h4><p style='font-size:26px;font-weight:bold;margin:5px 0;color:{vl_c};'>{vsurge:.0f}%</p><span style='color:#b2bec3;font-size:12px;'>of 20D avg → {vs_lbl}</span></div>",unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Telemetry error: {e}")
    st.markdown("---")

    # ── NLP Catalyst Feed + Analyst Wall ──────────────────────────────────
    st.markdown("### 📰 NLP Catalyst Feed & Analyst Intelligence")
    col_nlp, col_wall = st.columns([1.2,1])
    with col_nlp:
        st.markdown("#### Live Headline Parsing")
        # Use company name for news search — ticker strings like "TIPSMUSIC" return nothing
        search_name = " ".join(company_name.split()[:3]) if company_name and company_name != ticker_sym else clean_sym
        encoded_name = urllib.parse.quote(search_name)
        rss = (f"https://news.google.com/rss/search?q={encoded_name}+stock+NSE&hl=en-IN&gl=IN&ceid=IN:en"
               if goog_exc else
               f"https://news.google.com/rss/search?q={encoded_name}+stock+earnings&hl=en&gl=US&ceid=US:en")
        try:
            req = urllib.request.Request(rss, headers={"User-Agent":"Mozilla/5.0"})
            with urllib.request.urlopen(req,timeout=6) as r: xml=r.read()
            items = ET.fromstring(xml).findall("./channel/item")
            # Fallback to Yahoo Finance RSS if Google News returned nothing
            if not items:
                yahoo_rss = f"https://feeds.finance.yahoo.com/rss/2.0/headline?s={ticker_sym}&region=IN&lang=en-IN"
                try:
                    req2 = urllib.request.Request(yahoo_rss, headers={"User-Agent":"Mozilla/5.0"})
                    with urllib.request.urlopen(req2,timeout=6) as r2: xml=r2.read()
                    items = ET.fromstring(xml).findall("./channel/item")
                except: pass
            items = items[:7]
            bull=bear=neut=0
            for item in items:
                t_el = item.find("title")
                if not t_el: continue
                t = t_el.text or ""
                tl = t.lower()
                if any(w in tl for w in NLP_POS):   css,ic,tag="nlp-bull","🟢","BULLISH CATALYST"; bull+=1
                elif any(w in tl for w in NLP_NEG): css,ic,tag="nlp-bear","🔴","BEARISH CATALYST"; bear+=1
                else:                                css,ic,tag="nlp-neut","⚪","NEUTRAL FLOW";    neut+=1
                st.markdown(f"<div class='{css}'><strong>{ic} {tag}</strong><br><span style='font-size:13px;'>{t}</span></div>",unsafe_allow_html=True)
            tot = bull+bear+neut
            if tot:
                bp = (bull/tot)*100
                sc = "#2ed573" if bp>60 else "#ff4757" if bp<30 else "#e1b12c"
                st.markdown(f"<div style='background:#1e272e;border-radius:6px;padding:10px 14px;margin-top:8px;display:flex;justify-content:space-between;'><span style='color:#7f8fa6;font-size:12px;'>Sentiment ({tot} headlines)</span><span style='color:{sc};font-weight:bold;font-size:12px;'>🟢 {bull} &nbsp; ⚪ {neut} &nbsp; 🔴 {bear}</span></div>",unsafe_allow_html=True)
        except:
            st.warning("Headline feed temporarily unavailable.")

    with col_wall:
        st.markdown("#### 🎯 Analyst Accountability Wall")
        st.caption("Real sell-side rating changes via Yahoo Finance (Bloomberg/Refinitiv aggregated). Consensus buy/hold/sell counts are live.")
        try:
            raw = cached_upgrades(ticker_sym)
            if raw is not None and not raw.empty:
                df_up = raw.reset_index().tail(8).iloc[::-1]
                df_up.columns=[str(c).strip() for c in df_up.columns]
                dc = next((c for c in df_up.columns if "date" in c.lower()),df_up.columns[0])
                fc = next((c for c in df_up.columns if "firm" in c.lower()),None)
                ac = next((c for c in df_up.columns if "action" in c.lower()),None)
                gc = next((c for c in df_up.columns if "tograde" in c.lower() or "to grade" in c.lower()),None)
                if fc and gc:
                    st.dataframe(pd.DataFrame({
                        "Firm": df_up[fc].astype(str).str[:22],
                        "Date": pd.to_datetime(df_up[dc]).dt.strftime("%d %b %Y"),
                        "Action": df_up[ac].str.title() if ac else "N/A",
                        "To Grade": df_up[gc],
                    }), use_container_width=True, hide_index=True)
                else: st.info("Rating data unavailable.")
            else:
                recs = cached_recs(ticker_sym)
                if recs is not None and not recs.empty:
                    latest=recs.iloc[0]
                    cdf=pd.DataFrame({"Rating":["Strong Buy","Buy","Hold","Sell","Strong Sell"],
                        "Count":[int(latest.get(k,0)) for k in ["strongBuy","buy","hold","sell","strongSell"]]})
                    cdf=cdf[cdf["Count"]>0]
                    if not cdf.empty:
                        st.altair_chart(alt.Chart(cdf).mark_bar(cornerRadiusTopRight=4,cornerRadiusBottomRight=4).encode(
                            x="Count:Q",y=alt.Y("Rating:N",sort=["Strong Buy","Buy","Hold","Sell","Strong Sell"]),
                            color=alt.Color("Rating:N",scale=alt.Scale(
                                domain=["Strong Buy","Buy","Hold","Sell","Strong Sell"],
                                range=["#2ed573","#10ac84","#e1b12c","#ff6b81","#ff4757"]),legend=None)
                        ).properties(height=200),use_container_width=True)
                else: st.info("No analyst data available.")
        except Exception as e: st.caption(f"Analyst data: {e}")

    # ── Consensus Price Target ─────────────────────────────────────────────
    st.caption("Consensus price targets sourced from Yahoo Finance (aggregated from sell-side research firms). FCF = Operating Cash Flow – Capex from trailing twelve months.")
    try:
        tm=info.get("targetMeanPrice"); th=info.get("targetHighPrice"); tl=info.get("targetLowPrice")
        na=info.get("numberOfAnalystOpinions","N/A"); rk=info.get("recommendationKey","").replace("_"," ").title()
        if curr_px and tm:
            up=((tm-curr_px)/curr_px)*100
            p1,p2,p3,p4,p5=st.columns(5)
            p1.metric("Current",f"{currency}{curr_px:,.2f}")
            p2.metric("Target Low",f"{currency}{tl:,.2f}" if tl else "N/A")
            p3.metric("Consensus",f"{currency}{tm:,.2f}",f"{up:+.1f}% upside")
            p4.metric("Target High",f"{currency}{th:,.2f}" if th else "N/A")
            p5.metric("# Analysts",str(na))
            if tl and th and th>tl:
                pir=max(0,min(1,(curr_px-tl)/(th-tl)))*100
                nc="#2ed573" if curr_px<tm else "#e1b12c" if curr_px<th else "#ff4757"
                st.markdown(f'<div style="margin:10px 0;"><div style="display:flex;justify-content:space-between;font-size:12px;color:#7f8fa6;margin-bottom:4px;"><span>Low {currency}{tl:,.0f}</span><span style="color:{nc};font-weight:bold;">▲ {currency}{curr_px:,.0f}</span><span>High {currency}{th:,.0f}</span></div><div style="background:#2f3640;border-radius:6px;height:10px;"><div style="background:{nc};border-radius:6px;height:10px;width:{pir:.1f}%;max-width:100%;"></div></div><div style="text-align:center;margin-top:6px;font-size:12px;color:#7f8fa6;">Consensus: <strong style="color:#0abde3;">{currency}{tm:,.0f}</strong> ({up:+.1f}%) | Wall St: <strong style="color:#fff;">{rk}</strong></div></div>',unsafe_allow_html=True)
    except: pass
    st.markdown("---")

    # ── Liquidity & Earnings Quality ───────────────────────────────────────
    st.markdown("### 💧 Liquidity & Earnings Quality")
    try:
        bs=cached_balance_sheet(ticker_sym); inc=cached_income_stmt(ticker_sym); cf=cached_cashflow(ticker_sym)
        cr=info.get("currentRatio"); qr=info.get("quickRatio")
        cfo = safe_get(cf,"Operating Cash Flow","Cash From Operations") or info.get("operatingCashflow") or 0
        capex = abs(safe_get(cf,"Capital Expenditure","Purchase Of Property Plant And Equipment") or
                    abs(info.get("capitalExpenditures") or 0))
        fcff  = (cfo-capex) if cfo else None
        ni    = safe_get(inc,"Net Income","Net Income Common Stockholders")
        cni   = (cfo/ni) if (cfo and ni and ni>0) else None
        c1,c2,c3,c4=st.columns(4)
        with c1: st.markdown(metric_card("Current Ratio",f"{cr:.2f}x" if cr else "N/A",">2x strong · 1-2x adequate · <1x risky"),unsafe_allow_html=True)
        with c2: st.markdown(metric_card("Quick Ratio",f"{qr:.2f}x" if qr else "N/A","Excludes inventory. >1x = obligations covered."),unsafe_allow_html=True)
        with c3: st.markdown(metric_card("FCFF",fmt_money(fcff,is_indian) if fcff else "N/A","Operating CF minus Capex.",color="#2ed573" if (fcff and fcff>0) else "#ff4757"),unsafe_allow_html=True)
        with c4: st.markdown(metric_card("CFO / Net Income",f"{cni:.2f}x" if cni else "N/A",">1.0 cash-backed · <1.0 accrual-driven"),unsafe_allow_html=True)
        st.markdown("<div style='margin-bottom:8px'></div>",unsafe_allow_html=True)

        # ── Working Capital ────────────────────────────────────────────────
        st.markdown("### ⚙️ Working Capital Efficiency")
        dso=doh=dpo=ccc=None
        if bs is not None and inc is not None:
            rec=safe_get(bs,"Accounts Receivable","Net Receivables")
            inv=safe_get(bs,"Inventory")
            pay=safe_get(bs,"Accounts Payable")
            rev=safe_get(inc,"Total Revenue","Operating Revenue")
            cogs=safe_get(inc,"Cost Of Revenue","Cost of Goods Sold")
            if cogs: cogs=abs(cogs)
            if rec and rev and rev>0:   dso=(rec/rev)*365
            if inv and cogs and cogs>0: doh=(inv/cogs)*365
            if pay and cogs and cogs>0: dpo=(pay/cogs)*365
            if all(v is not None for v in [dso,doh,dpo]): ccc=dso+doh-dpo
        if ccc is not None:
            cc_c="#2ed573" if ccc<45 else "#ff4757"
            w1,w2,w3,w4=st.columns(4)
            with w1: st.markdown(metric_card("DSO",f"{dso:.1f} days","Days to collect from customers"),unsafe_allow_html=True)
            with w2: st.markdown(metric_card("DOH",f"{doh:.1f} days","Days inventory held before sale"),unsafe_allow_html=True)
            with w3: st.markdown(metric_card("DPO",f"{dpo:.1f} days","Days to pay suppliers. Higher = better."),unsafe_allow_html=True)
            with w4: st.markdown(f'<div style="background:#1e272e;border-radius:8px;padding:18px 20px;text-align:center;border-top:3px solid {cc_c};height:100%;"><div style="color:#7f8fa6;font-size:13px;margin-bottom:8px;">Cash Conversion Cycle <span title="DSO+DOH-DPO. Negative = supplier float." style="cursor:help;color:#7f8fa6;">ⓘ</span></div><div style="color:{cc_c};font-size:26px;font-weight:700;">{ccc:.1f} days</div></div>',unsafe_allow_html=True)
            st.markdown("<div style='margin-bottom:8px'></div>",unsafe_allow_html=True)
        else:
            st.info("Working capital unavailable — common for banks and financial institutions.")
    except Exception as e:
        st.error(f"Liquidity data error: {e}")
        cr=qr=fcff=ni=cni=dso=doh=dpo=ccc=None; bs=inc=cf=None
    st.markdown("---")

    # ── Historical Margin Trends ───────────────────────────────────────────
    st.markdown("### 📊 Historical Margin & Return Trends (3-Year)")
    try:
        if inc is not None and not inc.empty and bs is not None and not bs.empty:
            years = [c.strftime("%Y") if hasattr(c,"strftime") else str(c)[:4] for c in inc.columns[:3]]
            rev_row= "Total Revenue" if "Total Revenue" in inc.index else "Operating Revenue"
            rev_s  = [safe_get(inc,rev_row,col=i) for i in range(3)]
            gp_s   = [safe_get(inc,"Gross Profit",col=i) for i in range(3)]
            ebitda_s=[safe_get(inc,"Normalized EBITDA",col=i) for i in range(3)]
            ni_s   = [safe_get(inc,"Net Income","Net Income Common Stockholders",col=i) for i in range(3)]
            eq_s   = [safe_get(bs,"Stockholders Equity",col=i) for i in range(3)]
            rows=[]
            for i,yr in enumerate(years):
                r=rev_s[i]; g=gp_s[i]; e=ebitda_s[i]; n=ni_s[i]; eq=eq_s[i]
                if r and r>0:
                    rows.append({"Year":yr,
                        "Gross Margin":round((g/r)*100,2) if g else None,
                        "EBITDA Margin":round((e/r)*100,2) if e else None,
                        "Net Margin":round((n/r)*100,2) if n else None,
                        "ROE":round((n/eq)*100,2) if (n and eq and eq>0) else None})
            if rows:
                mdf = pd.DataFrame(rows).sort_values("Year")
                mc1,mc2 = st.columns(2)
                with mc1:
                    melt1 = mdf.melt("Year",["Gross Margin","EBITDA Margin","Net Margin"],"Metric","Value").dropna()
                    if not melt1.empty:
                        st.altair_chart(alt.Chart(melt1).mark_line(point=True).encode(
                            x=alt.X("Year:O"),y=alt.Y("Value:Q",title="Margin (%)"),
                            color=alt.Color("Metric:N",scale=alt.Scale(
                                domain=["Gross Margin","EBITDA Margin","Net Margin"],
                                range=["#10ac84","#0abde3","#e1b12c"])),
                            tooltip=["Year","Metric","Value"]
                        ).properties(height=220,title="Margin Trends"),use_container_width=True)
                with mc2:
                    melt2 = mdf.melt("Year",["ROE"],"Metric","Value").dropna()
                    if not melt2.empty:
                        st.altair_chart(alt.Chart(melt2).mark_bar(color="#e84393",cornerRadiusTopLeft=4,cornerRadiusTopRight=4).encode(
                            x=alt.X("Year:O"),y=alt.Y("Value:Q",title="ROE (%)"),tooltip=["Year","Value"]
                        ).properties(height=220,title="Return on Equity (%)"),use_container_width=True)
    except Exception as e:
        st.info(f"Margin trend data unavailable: {e}")
    st.markdown("---")

    # ── Ratio Interpreter ─────────────────────────────────────────────────
    st.markdown("### 🔎 Ratio Interpreter")
    interps=[]
    try:
        cr=info.get("currentRatio"); qr=info.get("quickRatio")
        if cr is not None:
            if cr>=2:   interps.append(("✅","Liquidity",f"Current ratio of **{cr:.2f}x** is strong. {company_name} can comfortably cover short-term obligations."))
            elif cr>=1: interps.append(("⚠️","Liquidity",f"Current ratio of **{cr:.2f}x** is adequate. Short-term covered, limited buffer if revenue slows."))
            else:       interps.append(("🚨","Liquidity",f"Current ratio of **{cr:.2f}x** is a liquidity risk. Current liabilities exceed current assets."))
        if qr is not None:
            if qr>=1:   interps.append(("✅","Quick Liquidity",f"Quick ratio of **{qr:.2f}x** — obligations met without selling inventory."))
            else:       interps.append(("⚠️","Quick Liquidity",f"Quick ratio of **{qr:.2f}x** — leaning on inventory. Sector downturn could pressure liquidity."))
        if fcff is not None:
            fm=fmt_money(fcff,is_indian)
            if fcff>0: interps.append(("✅","Free Cash Flow",f"Positive FCFF of **{fm}** — surplus cash after capex. Engine for dividends and debt paydown."))
            else:      interps.append(("🚨","Free Cash Flow",f"Negative FCFF of **{fm}** — capex exceeds operating cash flow. Watch external funding dependency."))
        if cni is not None:
            if cni>1.2:   interps.append(("✅","Earnings Quality",f"CFO/Net Income of **{cni:.2f}x** — every rupee of profit backed by more than a rupee of cash. Low accrual risk."))
            elif cni>=0.8: interps.append(("⚠️","Earnings Quality",f"CFO/Net Income of **{cni:.2f}x** — cash tracks income. Standard quality; watch for divergence."))
            else:          interps.append(("🚨","Earnings Quality",f"CFO/Net Income of **{cni:.2f}x** — significant accrual-based earnings. Risk of future write-downs."))
        if ccc is not None:
            if ccc<0:  interps.append(("✅","Working Capital",f"Negative CCC of **{ccc:.1f} days** — {company_name} runs on supplier float like DMart or Amazon."))
            elif ccc<45:interps.append(("✅","Working Capital",f"CCC of **{ccc:.1f} days** is healthy. Fast capital cycle."))
            else:       interps.append(("⚠️","Working Capital",f"CCC of **{ccc:.1f} days** is elevated. Focus on tightening collections and inventory."))
    except: pass
    for ic,lb,tx in interps:
        bg="#1a3a2a" if ic=="✅" else "#3a1a1a" if ic=="🚨" else "#2d2a1a"
        bd="#2ed573" if ic=="✅" else "#ff4757" if ic=="🚨" else "#e1b12c"
        st.markdown(f'<div style="background:{bg};border-left:4px solid {bd};border-radius:0 8px 8px 0;padding:14px 18px;margin-bottom:10px;"><span style="font-weight:700;color:{bd};">{ic} {lb}</span><p style="margin:6px 0 0 0;color:#dfe6e9;font-size:14px;line-height:1.6;">{tx}</p></div>',unsafe_allow_html=True)
    if not interps: st.info("Insufficient data for ratio interpretation.")
    st.markdown("---")

    # ── News Feed ──────────────────────────────────────────────────────────
    st.markdown("### 📰 Latest News")
    nf=False
    try:
        news_search = urllib.parse.quote(" ".join(company_name.split()[:3]) if company_name and company_name != ticker_sym else clean_sym)
        news_rss = (f"https://news.google.com/rss/search?q={news_search}+stock+NSE&hl=en-IN&gl=IN&ceid=IN:en"
                    if goog_exc else
                    f"https://news.google.com/rss/search?q={news_search}+stock&hl=en&gl=US&ceid=US:en")
        nr=urllib.request.Request(news_rss,headers={"User-Agent":"Mozilla/5.0"})
        with urllib.request.urlopen(nr,timeout=6) as r: xd=r.read()
        for item in ET.fromstring(xd).findall("./channel/item")[:8]:
            t_el=item.find("title"); l_el=item.find("link"); d_el=item.find("pubDate"); s_el=item.find("source")
            if t_el is None: continue
            nf=True
            t=t_el.text or ""; lnk=l_el.text if l_el is not None else "#"
            dt=(d_el.text or "Recent").replace(" +0000","").replace(" GMT","")
            src=s_el.text if s_el is not None else "Google News"
            st.markdown(f'<div class="news-card"><p style="color:#7f8fa6;font-size:12px;margin-bottom:3px;">{src} • {dt}</p><a href="{lnk}" target="_blank" style="color:#0abde3;text-decoration:none;font-weight:bold;font-size:15px;">{t}</a></div>',unsafe_allow_html=True)
    except: pass
    if not nf: st.info("News feed temporarily unavailable.")
    st.markdown("---")

    # ── Export ─────────────────────────────────────────────────────────────
    st.markdown("### 📥 Export Analysis")
    try:
        exp=[("Company",company_name),("Ticker",ticker_sym),("Sector",sector),("Industry",industry),
             ("Current Ratio",f"{cr:.2f}x" if cr else "N/A"),("Quick Ratio",f"{qr:.2f}x" if qr else "N/A"),
             ("FCFF",fmt_money(fcff,is_indian) if fcff else "N/A"),
             ("CFO/Net Income",f"{cni:.2f}x" if cni else "N/A"),
             ("DSO",f"{dso:.1f} days" if dso else "N/A"),("DOH",f"{doh:.1f} days" if doh else "N/A"),
             ("DPO",f"{dpo:.1f} days" if dpo else "N/A"),("CCC",f"{ccc:.1f} days" if ccc else "N/A")]
        edf=pd.DataFrame(exp,columns=["Metric","Value"])
        ec1,ec2=st.columns(2)
        with ec1:
            st.download_button("⬇️ Download CSV",edf.to_csv(index=False).encode(),f"{clean_sym}_fundamental.csv","text/csv",use_container_width=True)
        with ec2:
            tr="".join(f"<tr><td>{m}</td><td><strong>{v}</strong></td></tr>" for m,v in exp)
            ih="".join(f"<div class='i'><strong>{ic} {lb}</strong><p>{tx.replace('**','')}</p></div>" for ic,lb,tx in interps)
            html=f'<!DOCTYPE html><html><head><meta charset="utf-8"><title>{company_name}</title><style>body{{font-family:Arial,sans-serif;max-width:800px;margin:40px auto;font-size:14px;}}h1{{border-bottom:2px solid #0abde3;padding-bottom:8px;}}h2{{color:#0abde3;}}table{{width:100%;border-collapse:collapse;}}td{{padding:8px 12px;border-bottom:1px solid #ddd;}}tr:nth-child(even) td{{background:#f4f6f8;}}.i{{background:#f0f8ff;border-left:4px solid #0abde3;padding:10px;margin:8px 0;border-radius:0 6px 6px 0;}}.i p{{margin:4px 0 0 0;}}.foot{{margin-top:40px;font-size:11px;color:#999;border-top:1px solid #ddd;padding-top:10px;}}</style></head><body><h1>{company_name}</h1><p><strong>Ticker:</strong> {ticker_sym} | <strong>Sector:</strong> {sector} | <strong>Industry:</strong> {industry}</p><h2>Key Metrics</h2><table>{tr}</table><h2>Interpretation</h2>{ih}<div class="foot">Global Market Terminal | Not investment advice</div></body></html>'
            st.download_button("⬇️ Download PDF Report",html.encode(),f"{clean_sym}_fundamental.html","text/html",use_container_width=True)
            st.caption("💡 Ctrl+P → Save as PDF for a clean one-pager.")
    except Exception as e:
        st.error(f"Export error: {e}")

# ═══════════════════════════════════════════════════════════════════════════
# MODULE 2 — VALUATION MODELER
# ═══════════════════════════════════════════════════════════════════════════
elif menu == "🎯 Valuation Modeler":
    st.markdown("## 🎯 Institutional Valuation Desk")
    st.markdown("*Peer comparables, DCF engine, sensitivity analysis, and a ready-to-use Excel financial model.*")
    ticker_sym = ticker_search_widget()
    if not ticker_sym: st.stop()
    is_indian,currency,clean_sym,_,_,_ = ticker_meta(ticker_sym)

    with st.spinner(f"Pulling financials for {ticker_sym}…"):
        try:
            info        = cached_info(ticker_sym)
            curr_px     = info.get("currentPrice",info.get("regularMarketPrice",100))
            mkt_cap     = info.get("marketCap",1_000_000_000)
            pe          = info.get("trailingPE","N/A")
            eveb        = info.get("enterpriseToEbitda","N/A")
            pb          = info.get("priceToBook","N/A")
            ps          = info.get("priceToSalesTrailing12Months","N/A")
            shares      = info.get("sharesOutstanding",1_000_000)
            sector      = info.get("sector","")
            total_cash  = info.get("totalCash",  mkt_cap*0.05)
            total_debt  = info.get("totalDebt",  mkt_cap*0.15)
            op_cf       = info.get("operatingCashflow", mkt_cap*0.08)
            capex_raw   = info.get("capitalExpenditures", mkt_cap*0.02)
            capex       = abs(capex_raw) if capex_raw else mkt_cap*0.02
            base_fcf    = max(op_cf-capex, mkt_cap*0.03)
        except Exception as e:
            st.error(f"Data load error: {e}"); st.stop()

    # ── Live Multiples ─────────────────────────────────────────────────────
    st.markdown("### 📊 Live Valuation Multiples")
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Price",      f"{currency}{curr_px:,.2f}")
    c2.metric("P/E",        f"{pe:.2f}x"   if isinstance(pe,(int,float))   else pe)
    c3.metric("EV/EBITDA",  f"{eveb:.2f}x" if isinstance(eveb,(int,float)) else eveb)
    c4.metric("P/B",        f"{pb:.2f}x"   if isinstance(pb,(int,float))   else pb)
    c5.metric("P/S",        f"{ps:.2f}x"   if isinstance(ps,(int,float))   else ps)
    st.markdown("---")

    # ── Peer Comparables ───────────────────────────────────────────────────
    st.markdown("### 👥 Peer Comparables")
    industry = info.get("industry","")
    st.caption(f"Auto-detected peers — {industry} ({sector})")
    manual_override = st.text_input(
        "🔧 Override peers (optional — paste comma-separated tickers, e.g. TATAMOTORS.NS, M&M.NS)",
        value="", key="manual_peers", placeholder="Leave blank to use auto-detection")
    if manual_override.strip():
        peers = [p.strip() for p in manual_override.split(",") if p.strip()]
    else:
        peers = get_peers(ticker_sym, sector, industry, is_indian)
    if peers:
        with st.spinner("Fetching peer multiples…"):
            peer_df = get_peer_multiples(peers)
        if not peer_df.empty:
            target_row = {"Ticker":ticker_sym,
                          "Company":info.get("longName", info.get("shortName", clean_sym)),
                          "Price":curr_px,
                          "P/E":pe if isinstance(pe,(int,float)) else None,
                          "EV/EBITDA":eveb if isinstance(eveb,(int,float)) else None,
                          "P/B":pb if isinstance(pb,(int,float)) else None,
                          "Mkt Cap":mkt_cap,
                          "52W High":info.get("fiftyTwoWeekHigh"),
                          "52W Low":info.get("fiftyTwoWeekLow")}
            full_df = pd.concat([pd.DataFrame([target_row]),peer_df],ignore_index=True)

            def fmt_peer(df, cur):
                out = df.copy()
                # Truncate company name for display
                out["Company"]   = out["Company"].apply(lambda x: str(x)[:28] if pd.notna(x) else "N/A")
                out["Price"]     = out["Price"].apply(lambda x: f"{cur}{x:,.2f}" if pd.notna(x) else "N/A")
                out["P/E"]       = out["P/E"].apply(lambda x: f"{x:.1f}x"   if pd.notna(x) else "N/A")
                out["EV/EBITDA"] = out["EV/EBITDA"].apply(lambda x: f"{x:.1f}x" if pd.notna(x) else "N/A")
                out["P/B"]       = out["P/B"].apply(lambda x: f"{x:.1f}x"   if pd.notna(x) else "N/A")
                out["Mkt Cap"]   = out["Mkt Cap"].apply(lambda x: fmt_money(x,is_indian) if pd.notna(x) else "N/A")
                out["52W High"]  = out["52W High"].apply(lambda x: f"{cur}{x:,.0f}" if pd.notna(x) else "N/A")
                out["52W Low"]   = out["52W Low"].apply(lambda x: f"{cur}{x:,.0f}" if pd.notna(x) else "N/A")
                return out

            disp = fmt_peer(full_df, currency)
            # Highlight target row
            def highlight_target(row):
                if row.name == 0:
                    return ["background-color:#1a3a5c;color:white;font-weight:bold"]*len(row)
                return [""]*len(row)
            st.dataframe(disp.style.apply(highlight_target,axis=1), use_container_width=True, hide_index=True)
            st.caption("🔵 Highlighted row = target company")
    else:
        st.info("Peer basket not configured for this sector yet.")
    st.markdown("---")

    # ── Historical P/E Normalization ───────────────────────────────────────
    st.markdown("### 📈 Historical P/E Normalization Bands")
    try:
        inc_stmt = cached_income_stmt(ticker_sym)
        if inc_stmt is not None and not inc_stmt.empty:
            avail  = inc_stmt.index.tolist()
            eps_l  = next((l for l in ["Diluted EPS","Basic EPS","Normalized EPS"] if l in avail),None)
            ni_l   = next((l for l in ["Net Income","Net Income Common Stockholders"] if l in avail),None)
            nd=[]
            for date in inc_stmt.columns:
                try:
                    hd  = yf.Ticker(ticker_sym).history(
                          start=(date-pd.Timedelta(days=5)).strftime("%Y-%m-%d"),
                          end=(date+pd.Timedelta(days=5)).strftime("%Y-%m-%d"))
                    hp  = float(hd["Close"].median()) if not hd.empty else curr_px
                    hpe = None
                    if eps_l:
                        ev = inc_stmt.loc[eps_l,date]
                        ev = ev.iloc[0] if isinstance(ev,pd.Series) else ev
                        if pd.notna(ev) and ev>0: hpe=hp/ev
                    elif ni_l:
                        nv = inc_stmt.loc[ni_l,date]
                        nv = nv.iloc[0] if isinstance(nv,pd.Series) else nv
                        if pd.notna(nv) and nv>0: hpe=(shares*hp)/nv
                    if hpe and 0<hpe<400:
                        nd.append({"Fiscal Year":str(date.year),"Historical P/E":round(hpe,2)})
                except: pass
            if len(nd)>1:
                ndf=pd.DataFrame(nd).sort_values("Fiscal Year")
                hmed=ndf["Historical P/E"].mean()
                nh1,nh2=st.columns([1.5,1])
                with nh1:
                    st.altair_chart(alt.Chart(ndf).mark_line(point=True,color="#10ac84").encode(
                        x=alt.X("Fiscal Year:O"),y=alt.Y("Historical P/E:Q",scale=alt.Scale(zero=False)),
                        tooltip=["Fiscal Year","Historical P/E"]).properties(height=230),use_container_width=True)
                with nh2:
                    st.metric("Historical Median P/E",f"{hmed:.2f}x")
                    if isinstance(pe,(int,float)):
                        dev=((pe-hmed)/hmed)*100
                        if dev>15:   st.warning(f"⚠️ Trading at **{dev:.1f}% premium** to history.")
                        elif dev<-15:st.success(f"🔥 Trading at **{abs(dev):.1f}% discount** to history.")
                        else:        st.info("📊 Fairly valued vs historical range.")
    except Exception as e:
        st.info(f"Historical P/E unavailable: {e}")
    st.markdown("---")

    # ── DCF Engine ─────────────────────────────────────────────────────────
    st.markdown("<h3 style='text-align:center;'>🧮 DCF Engine</h3>",unsafe_allow_html=True)
    di,do = st.columns([1,1.5])
    with di:
        st.markdown("#### Assumptions")
        wacc         = st.slider("WACC %",5.0,20.0,11.5,0.5)/100
        fcf_growth   = st.slider("FCF Growth % (Yr 1–5)",0.0,30.0,12.0,0.5)/100
        terminal_g   = st.slider("Terminal Growth %",1.0,6.0,3.0,0.5)/100
    with do:
        pv_fcf=0; fcfs=[]
        for yr in range(1,6):
            pf=base_fcf*((1+fcf_growth)**yr)
            pv_fcf+=pf/((1+wacc)**yr); fcfs.append(pf)
        if wacc>terminal_g:
            tv     = (fcfs[-1]*(1+terminal_g))/(wacc-terminal_g)
            pv_tv  = tv/((1+wacc)**5)
            ev     = pv_fcf+pv_tv
            eq_v   = ev+total_cash-total_debt
            iv     = eq_v/shares if shares>0 else 0
            upside = ((iv-curr_px)/curr_px)*100
            st.markdown("#### Intrinsic Value Output")
            st.info(f"**Calculated Intrinsic Value:**\n### {currency}{iv:,.2f}")
            if upside>0: st.success(f"✅ Undervalued by **{upside:.1f}%** vs current {currency}{curr_px:,.2f}")
            else:        st.error(  f"🚨 Overvalued by **{abs(upside):.1f}%** vs current {currency}{curr_px:,.2f}")
            st.bar_chart(pd.DataFrame({
                "Component":["PV FCFs (5Y)","PV Terminal Value","Net Cash/Debt","Equity Value"],
                "Value":[pv_fcf,pv_tv,total_cash-total_debt,eq_v]
            }).set_index("Component"),color="#10ac84")
        else:
            st.error("WACC must be greater than terminal growth rate.")
            pv_tv=ev=eq_v=iv=upside=0
    st.markdown("---")

    # ── Sensitivity Analysis ───────────────────────────────────────────────
    st.markdown("### 🗺️ Sensitivity Analysis — WACC vs Terminal Growth")
    st.caption("Intrinsic value per share across WACC and terminal growth combinations. 🟢 = upside · 🔴 = downside vs current price.")
    wacc_steps = [wacc+(i-4)*0.005 for i in range(9)]
    tg_steps   = [terminal_g+(j-3)*0.005 for j in range(7)]
    header_cols = f'<th style="background:#1e272e;color:#7f8fa6;padding:8px;font-size:11px;">WACC \\ g</th>' + \
                  "".join(f'<th style="background:#1e272e;color:#7f8fa6;padding:8px;font-size:11px;">{g*100:.1f}%</th>' for g in tg_steps)
    body_rows=""
    for w in wacc_steps:
        cells=""
        for g in tg_steps:
            val=dcf_iv(base_fcf,fcf_growth,w,g,total_cash,total_debt,shares)
            if val is None:
                cells+='<td style="background:#2d3436;color:#7f8fa6;text-align:center;padding:7px 5px;font-size:11px;">N/A</td>'
            else:
                up=((val-curr_px)/curr_px)*100
                bg=("#0d2b1a" if up>30 else "#1a3a2a" if up>10 else "#2a3a1a" if up>0 else "#3a1a1a" if up>-10 else "#4a0a0a")
                fc="#2ed573" if up>0 else "#ff4757"
                cells+=f'<td style="background:{bg};color:{fc};text-align:center;padding:7px 5px;font-size:11px;">{currency}{val:,.0f}<br><span style="font-size:9px;">{up:+.1f}%</span></td>'
        body_rows+=f'<tr><td style="background:#2f3640;color:#e1b12c;text-align:center;padding:7px 5px;font-size:11px;font-weight:bold;">{w*100:.1f}%</td>{cells}</tr>'
    st.markdown(f'<div style="overflow-x:auto;"><table style="border-collapse:collapse;width:100%;"><thead><tr>{header_cols}</tr></thead><tbody>{body_rows}</tbody></table></div>',unsafe_allow_html=True)
    st.markdown("---")

    # ── Excel Financial Model Builder ──────────────────────────────────────
    st.markdown("### 📥 Download Institutional Excel Financial Model")
    st.caption("6-sheet ready-to-use model: Cover · DCF · Sensitivity · Peer Comps · 3-Statement Financials")

    df_is = format_statement(cached_income_stmt(ticker_sym),"IS")
    df_bs = format_statement(cached_balance_sheet(ticker_sym),"BS")
    df_cf = format_statement(cached_cashflow(ticker_sym),"CF")

    wb  = Workbook()
    wb.remove(wb.active)

    # ── SHEET 1: COVER ────────────────────────────────────────────────────
    ws = wb.create_sheet("Cover")
    ws.column_dimensions["A"].width=28; ws.column_dimensions["B"].width=24
    ws.column_dimensions["C"].width=24; ws.column_dimensions["D"].width=24

    ws.merge_cells("A1:D1")
    t=ws["A1"]; t.value="GLOBAL MARKET TERMINAL — INSTITUTIONAL EQUITY RESEARCH"
    t.fill=_H; t.font=Font(color="FFFFFF",bold=True,size=14); t.alignment=_CT
    ws.row_dimensions[1].height=30

    ws.merge_cells("A2:D2")
    t2=ws["A2"]; t2.value=info.get("longName",ticker_sym)
    t2.font=Font(bold=True,size=16,color="1A3A5C"); t2.alignment=_CT
    ws.row_dimensions[2].height=28

    ws.merge_cells("A3:D3")
    t3=ws["A3"]; t3.value=f"{sector}  |  {info.get('industry','')}  |  {ticker_sym}"
    t3.font=Font(color="555555",size=11); t3.alignment=_CT; ws.row_dimensions[3].height=18

    ws.row_dimensions[4].height=8

    for col,h in enumerate(["METRIC","VALUE","METRIC","VALUE"],1):
        _c(ws,5,col,h,fill=_H,font=_FH,align=_CT)

    kv=[("Current Price",f"{currency}{curr_px:,.2f}"),("Market Cap",fmt_money(mkt_cap,is_indian)),
        ("Trailing P/E",f"{pe:.2f}x" if isinstance(pe,(int,float)) else str(pe)),
        ("EV/EBITDA",f"{eveb:.2f}x" if isinstance(eveb,(int,float)) else str(eveb)),
        ("Price/Book",f"{pb:.2f}x" if isinstance(pb,(int,float)) else str(pb)),
        ("52W High",f"{currency}{info.get('fiftyTwoWeekHigh',0):,.2f}"),
        ("52W Low",f"{currency}{info.get('fiftyTwoWeekLow',0):,.2f}"),
        ("Intrinsic Value (DCF)",f"{currency}{iv:,.2f}"),
        ("Implied Upside/Downside",f"{upside:+.2f}%"),
        ("Analyst Consensus",info.get("recommendationKey","N/A").replace("_"," ").title()),
        ("# Analysts",str(info.get("numberOfAnalystOpinions","N/A"))),
        ("Beta",str(round(info.get("beta",0),2))),]

    for i in range(0,len(kv),2):
        row=6+i//2
        _c(ws,row,1,kv[i][0],fill=_CA,font=_FB,align=_LT)
        val_cell=_c(ws,row,2,kv[i][1],fill=_WH,align=_RT)
        if i+1<len(kv):
            _c(ws,row,3,kv[i+1][0],fill=_CA,font=_FB,align=_LT)
            _c(ws,row,4,kv[i+1][1],fill=_WH,align=_RT)

    sumrow=6+len(kv)//2+2
    ws.merge_cells(start_row=sumrow,start_column=1,end_row=sumrow,end_column=4)
    ws.cell(row=sumrow,column=1,value="Business Summary").font=Font(bold=True,size=10,color="1A3A5C")
    sumrow+=1
    ws.merge_cells(start_row=sumrow,start_column=1,end_row=sumrow+4,end_column=4)
    full_summary = info.get("longBusinessSummary","")
    sc=ws.cell(row=sumrow,column=1,value=full_summary)
    sc.alignment=Alignment(wrap_text=True,vertical="top"); ws.row_dimensions[sumrow].height=80

    # ── SHEET 2: DCF VALUATION ────────────────────────────────────────────
    ws2=wb.create_sheet("DCF Valuation")
    for col,w in enumerate([30,20,20,20,20,20,20],1):
        ws2.column_dimensions[get_column_letter(col)].width=w

    _hrow(ws2,1,"DCF VALUATION MODEL — DISCOUNTED CASH FLOW ANALYSIS",7)
    _shrow(ws2,3,"ASSUMPTIONS (Yellow = Modifiable Inputs)",7)
    assumptions=[
        ("WACC (Cost of Capital)",wacc*100,"%"),
        ("FCF Growth Rate (Years 1–5)",fcf_growth*100,"%"),
        ("Terminal Growth Rate",terminal_g*100,"%"),
        ("Base Free Cash Flow",base_fcf,"abs"),
        ("Total Cash & Equivalents",total_cash,"abs"),
        ("Total Debt",total_debt,"abs"),
        ("Diluted Shares Outstanding",shares,"abs"),
    ]
    for i,(lbl,val,typ) in enumerate(assumptions):
        r=4+i
        _c(ws2,r,1,lbl,fill=_CA,font=_FB,align=_LT)
        fmt="#,##0.00%" if typ=="%" else "#,##0"
        dval=val/100 if typ=="%" else val
        _c(ws2,r,2,dval,fill=_IN,font=_FI,align=_RT,fmt=fmt)

    _shrow(ws2,12,"PROJECTION TABLE (Years 1–5)",7)
    yr_hdrs=["Metric"]+[f"Year {y}" for y in range(1,6)]
    for c,h in enumerate(yr_hdrs,1):
        _c(ws2,13,c,h,fill=_H,font=_FH,align=_CT)

    proj_rows=[
        ("Projected FCF",          fcfs,                                          "#,##0"),
        ("Discount Factor",        [1/((1+wacc)**y) for y in range(1,6)],         "0.0000"),
        ("Present Value of FCF",   [fcfs[y]/((1+wacc)**(y+1)) for y in range(5)], "#,##0"),
    ]
    for ri,(lbl,vals,fmt_str) in enumerate(proj_rows):
        r=14+ri
        _c(ws2,r,1,lbl,fill=_CA,font=_FB,align=_LT)
        for ci,v in enumerate(vals):
            _c(ws2,r,ci+2,round(v,6),fill=_WH,align=_RT,fmt=fmt_str)

    _shrow(ws2,18,"VALUATION BRIDGE",7)
    bridge=[
        ("Sum of PV FCFs (Years 1–5)",      pv_fcf),
        ("PV of Terminal Value",             pv_tv),
        ("(+) Cash & Equivalents",           total_cash),
        ("(–) Total Debt",                   -total_debt),
        ("Implied Enterprise Value",         ev),
        ("Implied Equity Value",             eq_v),
        ("Diluted Shares Outstanding",       shares),
        ("INTRINSIC VALUE PER SHARE",        iv),
        ("Current Market Price",             curr_px),
        ("Implied Upside / (Downside)",      upside/100),
    ]
    for i,(lbl,val) in enumerate(bridge):
        r=19+i
        is_total=lbl.startswith("=") or lbl.startswith("INTRINSIC")
        is_upside="Upside" in lbl
        fill=_PO if (is_total and val>0) else _NE if (is_total and val<0) else _CA if is_total else _WH
        font=_FG if (is_total and val>0) else _FR if (is_total and val<0) else _FB if is_total else None
        fmt="0.00%" if is_upside else "#,##0"
        _c(ws2,r,1,lbl,fill=_CA if not is_total else fill,font=_FB,align=_LT)
        _c(ws2,r,2,val,fill=fill,font=font,align=_RT,fmt=fmt)

    # ── SHEET 3: SENSITIVITY ──────────────────────────────────────────────
    ws3=wb.create_sheet("Sensitivity Matrix")
    for col,w in enumerate([12]+[14]*7,1):
        ws3.column_dimensions[get_column_letter(col)].width=w
    _hrow(ws3,1,"SENSITIVITY ANALYSIS — Intrinsic Value per Share",8)
    _shrow(ws3,2,"Rows = WACC | Columns = Terminal Growth Rate",8)
    ws3.cell(row=3,column=1,value="WACC \\ g").fill=_H
    ws3.cell(row=3,column=1).font=_FH; ws3.cell(row=3,column=1).alignment=_CT
    for j,g in enumerate(tg_steps):
        _c(ws3,3,j+2,f"{g*100:.1f}%",fill=_H,font=_FH,align=_CT)
    for i,w in enumerate(wacc_steps):
        _c(ws3,4+i,1,f"{w*100:.1f}%",fill=_SH,font=Font(color="FFFFFF",bold=True),align=_CT)
        for j,g in enumerate(tg_steps):
            val=dcf_iv(base_fcf,fcf_growth,w,g,total_cash,total_debt,shares)
            if val is None: cell=_c(ws3,4+i,j+2,"N/A",fill=_AL,align=_CT)
            else:
                cell=_c(ws3,4+i,j+2,round(val,2),fill=_PO if val>curr_px else _NE,
                         font=_FG if val>curr_px else _FR,align=_CT,fmt="#,##0.00")

    # ── SHEET 4: PEER COMPARABLES ─────────────────────────────────────────
    ws4=wb.create_sheet("Peer Comparables")
    for col,w in enumerate([28,14,14,12,14,12,16,14,14],1):
        ws4.column_dimensions[get_column_letter(col)].width=w
    _hrow(ws4,1,"PEER COMPARABLES — RELATIVE VALUATION ANALYSIS",9)
    hdrs=["Company","Ticker","Price","P/E","EV/EBITDA","P/B","Mkt Cap (₹ Cr / $B)","52W High","52W Low"]
    for c,h in enumerate(hdrs,1):
        _c(ws4,2,c,h,fill=_H,font=_FH,align=_CT)
    if not peer_df.empty:
        trow={"Company":info.get("longName",clean_sym),"Ticker":ticker_sym,
              "Price":curr_px,"P/E":pe if isinstance(pe,(int,float)) else None,
              "EV/EBITDA":eveb if isinstance(eveb,(int,float)) else None,
              "P/B":pb if isinstance(pb,(int,float)) else None,
              "Mkt Cap":mkt_cap,
              "52W High":info.get("fiftyTwoWeekHigh"),
              "52W Low":info.get("fiftyTwoWeekLow")}
        all_peers=[trow]+peer_df.to_dict("records")
        for ri,row in enumerate(all_peers):
            r=3+ri; is_t=ri==0
            fill=PatternFill("solid",fgColor="1A3A5C") if is_t else (_AL if ri%2==0 else _WH)
            fn=Font(color="FFFFFF",bold=True) if is_t else None
            # Convert raw market cap to readable units (Cr for Indian, B for US)
            raw_mc = row.get("Mkt Cap")
            if raw_mc:
                mc_display = raw_mc/10_000_000 if is_indian else raw_mc/1_000_000_000
            else:
                mc_display = None
            _c(ws4,r,1,row.get("Company","")[:30],fill=fill,font=fn,align=_LT)
            _c(ws4,r,2,row.get("Ticker",""),fill=fill,font=fn,align=_CT)
            _c(ws4,r,3,row.get("Price"),fill=fill,font=fn,align=_RT,fmt="#,##0.00")
            _c(ws4,r,4,row.get("P/E"),fill=fill,font=fn,align=_RT,fmt="#,##0.1")
            _c(ws4,r,5,row.get("EV/EBITDA"),fill=fill,font=fn,align=_RT,fmt="#,##0.1")
            _c(ws4,r,6,row.get("P/B"),fill=fill,font=fn,align=_RT,fmt="#,##0.1")
            _c(ws4,r,7,mc_display,fill=fill,font=fn,align=_RT,fmt="#,##0.0")
            _c(ws4,r,8,row.get("52W High"),fill=fill,font=fn,align=_RT,fmt="#,##0.0")
            _c(ws4,r,9,row.get("52W Low"),fill=fill,font=fn,align=_RT,fmt="#,##0.0")

    # ── SHEETS 5–7: FINANCIAL STATEMENTS ──────────────────────────────────
    def write_fs_sheet(wb, name, df, divisor=1_000_000):
        if df is None or df.empty: return
        ws=wb.create_sheet(name)
        ws.column_dimensions["A"].width=38
        _hrow(ws,1,name.upper(),len(df.columns)+1)
        _c(ws,2,1,"Line Item",fill=_H,font=_FH,align=_LT)
        for ci,col in enumerate(df.columns,2):
            _c(ws,2,ci,col,fill=_H,font=_FH,align=_CT)
            ws.column_dimensions[get_column_letter(ci)].width=18
        for ri,(idx,row) in enumerate(df.iterrows()):
            r=3+ri
            fill=_CA if any(k in idx for k in ["Revenue","Income","EBITDA","Profit","Total","Equity","Cash","Flow"]) else (_AL if ri%2==0 else _WH)
            fb=_FB if fill==_CA else None
            _c(ws,r,1,idx,fill=fill,font=fb,align=_LT)
            for ci,val in enumerate(row,2):
                try:
                    nval=val/divisor if pd.notna(val) else None
                    cell=_c(ws,r,ci,nval,fill=fill,align=_RT,fmt="#,##0.0")
                    if nval and nval<0: cell.font=_FR
                except: _c(ws,r,ci,None,fill=fill,align=_RT)

    write_fs_sheet(wb,"Income Statement (₹M)",df_is)
    write_fs_sheet(wb,"Balance Sheet (₹M)",df_bs)
    write_fs_sheet(wb,"Cash Flow (₹M)",df_cf)

    buf=io.BytesIO()
    wb.save(buf)
    st.download_button(
        label=f"🚀 Download Institutional Excel Model — {clean_sym}",
        data=buf.getvalue(),
        file_name=f"{clean_sym}_Institutional_Financial_Model.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)
    st.markdown("---")

    # ── 3-Year Financial Statements (Preview) ─────────────────────────────
    st.markdown("### 📚 3-Year Financial Statements (₹M / $M)")
    st.info("📌 **Sanity note:** Section totals (Total Assets, Total Liabilities, Operating CF etc.) are directly from company filings — accurate. Sub-items shown are significant ones only; not all components are returned by the data provider, so sub-items may not sum to totals. Balance sheet equation (Assets = Liabilities + Equity) is verified and holds.")
    t_is,t_bs,t_cf=st.tabs(["Income Statement","Balance Sheet","Cash Flow"])
    with t_is:
        if not df_is.empty: st.dataframe((df_is/1e6).style.format("{:,.2f}"),use_container_width=True)
        else: st.info("Unavailable.")
    with t_bs:
        if not df_bs.empty: st.dataframe((df_bs/1e6).style.format("{:,.2f}"),use_container_width=True)
        else: st.info("Unavailable.")
    with t_cf:
        if not df_cf.empty: st.dataframe((df_cf/1e6).style.format("{:,.2f}"),use_container_width=True)
        else: st.info("Unavailable.")

# ═══════════════════════════════════════════════════════════════════════════
# MODULE 3 — TECHNICAL ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════
elif menu == "📈 Technical Analysis":
    st.markdown("## 📈 Technical Analysis Engine")
    st.markdown("*RSI, MACD, Bollinger Bands, 50/200 DMA crossover, and volume intelligence.*")
    ticker_sym = ticker_search_widget()
    if not ticker_sym: st.stop()
    is_indian,currency,clean_sym,_,_,_ = ticker_meta(ticker_sym)

    period_map={"1 Month":"1mo","3 Months":"3mo","6 Months":"6mo","1 Year":"1y","2 Years":"2y"}
    sel_period=st.selectbox("Chart Period",list(period_map.keys()),index=3)
    period_str=period_map[sel_period]

    with st.spinner(f"Computing technical indicators for {ticker_sym}…"):
        try:
            raw = cached_history(ticker_sym, period_str).reset_index()
            date_col = "Date" if "Date" in raw.columns else raw.columns[0]
            raw = raw.rename(columns={date_col:"Date"})
            raw["Date"] = pd.to_datetime(raw["Date"]).dt.tz_localize(None)
            close = raw["Close"]

            raw["RSI"]          = compute_rsi(close)
            raw["MACD"],raw["MACD_Signal"],raw["MACD_Hist"] = compute_macd(close)
            raw["BB_Upper"],raw["BB_Mid"],raw["BB_Lower"]   = compute_bollinger(close)
            raw["DMA50"]        = close.rolling(50).mean()
            raw["DMA200"]       = close.rolling(200).mean()
            raw["vol_up"]       = (close >= close.shift(1)).astype(int)

            # Current indicator values
            cur_rsi   = float(raw["RSI"].dropna().iloc[-1])  if not raw["RSI"].dropna().empty  else None
            cur_macd  = float(raw["MACD"].dropna().iloc[-1]) if not raw["MACD"].dropna().empty else None
            cur_sig   = float(raw["MACD_Signal"].dropna().iloc[-1]) if not raw["MACD_Signal"].dropna().empty else None
            cur_close = float(close.iloc[-1])
            cur_bbu   = float(raw["BB_Upper"].dropna().iloc[-1]) if not raw["BB_Upper"].dropna().empty else None
            cur_bbl   = float(raw["BB_Lower"].dropna().iloc[-1]) if not raw["BB_Lower"].dropna().empty else None
            cur_d50   = float(raw["DMA50"].dropna().iloc[-1])  if not raw["DMA50"].dropna().empty  else None
            cur_d200  = float(raw["DMA200"].dropna().iloc[-1]) if not raw["DMA200"].dropna().empty else None

            # ── Signal Summary Cards ───────────────────────────────────────
            st.markdown("### 🚦 Signal Dashboard")
            s1,s2,s3,s4 = st.columns(4)

            def sig_card(col_widget, title, signal, value, color):
                col_widget.markdown(
                    f'<div style="background:#1e272e;border-radius:8px;padding:14px 16px;'
                    f'text-align:center;border-top:3px solid {color};">'
                    f'<div style="color:#7f8fa6;font-size:12px;margin-bottom:6px;">{title}</div>'
                    f'<div style="color:{color};font-size:15px;font-weight:700;margin-bottom:4px;">{signal}</div>'
                    f'<div style="color:#dfe6e9;font-size:12px;">{value}</div></div>',
                    unsafe_allow_html=True)

            if cur_rsi is not None:
                if cur_rsi>70:   rsi_sig,rsi_col="OVERBOUGHT","#ff4757"
                elif cur_rsi<30: rsi_sig,rsi_col="OVERSOLD","#2ed573"
                else:            rsi_sig,rsi_col="NEUTRAL","#7f8fa6"
                sig_card(s1,"RSI (14)",rsi_sig,f"{cur_rsi:.1f}",rsi_col)

            if cur_macd is not None and cur_sig is not None:
                mc_sig = "BULLISH" if cur_macd>cur_sig else "BEARISH"
                mc_col = "#2ed573" if cur_macd>cur_sig else "#ff4757"
                sig_card(s2,"MACD (12,26,9)",mc_sig,f"MACD {cur_macd:.3f}",mc_col)

            if cur_d50 is not None and cur_d200 is not None:
                if cur_d50>cur_d200: ma_sig,ma_col="GOLDEN CROSS ✨","#2ed573"
                else:                ma_sig,ma_col="DEATH CROSS 💀","#ff4757"
                sig_card(s3,"50/200 DMA",ma_sig,f"50D {cur_d50:,.1f}",ma_col)

            if cur_bbu is not None and cur_bbl is not None:
                if cur_close>cur_bbu:   bb_sig,bb_col="ABOVE UPPER BAND","#ff4757"
                elif cur_close<cur_bbl: bb_sig,bb_col="BELOW LOWER BAND","#2ed573"
                else:                   bb_sig,bb_col="WITHIN BANDS","#7f8fa6"
                sig_card(s4,"Bollinger Bands",bb_sig,f"Upper {cur_bbu:,.1f}",bb_col)

            st.markdown("---")

            # ── Price + Overlays Chart ─────────────────────────────────────
            st.markdown("### 📉 Price Chart — Moving Averages & Bollinger Bands")
            plot = raw.dropna(subset=["Close"])
            base = alt.Chart(plot).encode(x=alt.X("Date:T",axis=alt.Axis(format="%b %Y",labelAngle=-30,title="")))

            price_l  = base.mark_line(strokeWidth=2,color="#0abde3").encode(y=alt.Y("Close:Q",title=f"Price ({currency})"),tooltip=["Date:T","Close:Q"])
            bbu_l    = base.mark_line(strokeWidth=1,opacity=0.5,color="#7f8fa6").encode(y="BB_Upper:Q")
            bbm_l    = base.mark_line(strokeWidth=1,opacity=0.3,color="#7f8fa6",strokeDash=[3,3]).encode(y="BB_Mid:Q")
            bbl_l    = base.mark_line(strokeWidth=1,opacity=0.5,color="#7f8fa6").encode(y="BB_Lower:Q")
            d50_l    = base.mark_line(strokeWidth=2,strokeDash=[6,2],color="#e1b12c").encode(y="DMA50:Q")
            d200_l   = base.mark_line(strokeWidth=2,strokeDash=[6,2],color="#ff4757").encode(y="DMA200:Q")

            legend_data = pd.DataFrame({"Series":["Price","50 DMA","200 DMA","Bollinger Bands"],
                                         "Color":["#0abde3","#e1b12c","#ff4757","#7f8fa6"]})
            price_chart = (bbu_l+bbl_l+bbm_l+d50_l+d200_l+price_l).properties(height=340,
                title=f"{ticker_sym} — Price with 50/200 DMA & Bollinger Bands (20, ±2σ)")
            st.altair_chart(price_chart,use_container_width=True)

            # ── RSI Chart ─────────────────────────────────────────────────
            st.markdown("### 📊 RSI (14)")
            rsi_plot = raw.dropna(subset=["RSI"])
            rsi_line = alt.Chart(rsi_plot).mark_line(color="#10ac84",strokeWidth=1.5).encode(
                x=alt.X("Date:T",axis=alt.Axis(format="%b %Y",labelAngle=-30,title="")),
                y=alt.Y("RSI:Q",scale=alt.Scale(domain=[0,100]),title="RSI"),
                tooltip=["Date:T","RSI:Q"])
            ob  = alt.Chart(pd.DataFrame({"y":[70]})).mark_rule(color="#ff4757",strokeDash=[5,5],opacity=0.8).encode(y="y:Q")
            os  = alt.Chart(pd.DataFrame({"y":[30]})).mark_rule(color="#2ed573",strokeDash=[5,5],opacity=0.8).encode(y="y:Q")
            mid = alt.Chart(pd.DataFrame({"y":[50]})).mark_rule(color="#7f8fa6",strokeDash=[3,3],opacity=0.4).encode(y="y:Q")
            st.altair_chart((rsi_line+ob+os+mid).properties(height=180),use_container_width=True)

            # ── MACD Chart ────────────────────────────────────────────────
            st.markdown("### 📊 MACD (12, 26, 9)")
            macd_plot = raw.dropna(subset=["MACD"])
            hist_bars = alt.Chart(macd_plot).mark_bar(opacity=0.7).encode(
                x=alt.X("Date:T",axis=alt.Axis(format="%b %Y",labelAngle=-30,title="")),
                y=alt.Y("MACD_Hist:Q",title="MACD"),
                color=alt.condition(alt.datum.MACD_Hist>0,alt.value("#2ed573"),alt.value("#ff4757")),
                tooltip=["Date:T","MACD:Q","MACD_Signal:Q","MACD_Hist:Q"])
            macd_l  = alt.Chart(macd_plot).mark_line(color="#0abde3",strokeWidth=1.5).encode(x="Date:T",y="MACD:Q")
            signal_l= alt.Chart(macd_plot).mark_line(color="#ff4757",strokeWidth=1.5,strokeDash=[4,2]).encode(x="Date:T",y="MACD_Signal:Q")
            st.altair_chart((hist_bars+macd_l+signal_l).properties(height=180),use_container_width=True)

            # ── Volume Chart ──────────────────────────────────────────────
            st.markdown("### 📊 Volume")
            vol_chart = alt.Chart(raw).mark_bar(opacity=0.6).encode(
                x=alt.X("Date:T",axis=alt.Axis(format="%b %Y",labelAngle=-30,title="")),
                y=alt.Y("Volume:Q",title="Volume"),
                color=alt.condition(alt.datum.vol_up==1,alt.value("#2ed573"),alt.value("#ff4757")),
                tooltip=["Date:T","Volume:Q"])
            st.altair_chart(vol_chart.properties(height=140),use_container_width=True)

            st.markdown("""
            <div style="background:#1e272e;border-radius:6px;padding:12px 16px;font-size:12px;color:#7f8fa6;margin-top:8px;">
            📖 <strong>Legend:</strong>
            <span style="color:#0abde3;">━━ Price</span> &nbsp;
            <span style="color:#e1b12c;">╌╌ 50 DMA</span> &nbsp;
            <span style="color:#ff4757;">╌╌ 200 DMA</span> &nbsp;
            <span style="color:#7f8fa6;">━━ Bollinger Bands (±2σ)</span> &nbsp; | &nbsp;
            RSI: >70 Overbought · <30 Oversold &nbsp; | &nbsp;
            Golden Cross: 50D above 200D = Bullish
            </div>""",unsafe_allow_html=True)

        except Exception as e:
            st.error(f"Technical analysis error: {e}")

# ═══════════════════════════════════════════════════════════════════════════
# MODULE 4 — MACRO & CAPM
# ═══════════════════════════════════════════════════════════════════════════
elif menu == "📊 Macro & CAPM":
    st.markdown("## 📊 Macro Factors & CAPM Engine")
    st.markdown("*Live global liquidity, commodity prices, yield curves, and dynamic beta regression.*")

    selected_region = st.selectbox("Select Region",["India","United States","United Kingdom","Eurozone","Japan"])
    MACRO_CONFIG={
        "India":         {"fx_label":"USD/INR","fx_ticker":"USDINR=X","bond_label":"India 10Y G-Sec","bond_ticker":"IN10YT=RR","fallback_yield":6.77,"erp":5.5,"bench_ticker":"^NSEI","bench_name":"Nifty 50","vix_ticker":"^INDIAVIX","vix_name":"India VIX","short_ticker":None,"fallback_short":6.85,"policy_rate":6.50},
        "United States": {"fx_label":"US Dollar Index","fx_ticker":"DX-Y.NYB","bond_label":"US 10Y Treasury","bond_ticker":"^TNX","fallback_yield":4.37,"erp":4.5,"bench_ticker":"^GSPC","bench_name":"S&P 500","vix_ticker":"^VIX","vix_name":"CBOE VIX","short_ticker":"^IRX","fallback_short":3.64,"policy_rate":5.25},
        "United Kingdom":{"fx_label":"GBP/USD","fx_ticker":"GBPUSD=X","bond_label":"UK 10Y Gilt","bond_ticker":"GB10YT=RR","fallback_yield":4.74,"erp":5.0,"bench_ticker":"^FTSE","bench_name":"FTSE 100","vix_ticker":"^VIX","vix_name":"Volatility Proxy","short_ticker":None,"fallback_short":5.15,"policy_rate":5.25},
        "Eurozone":      {"fx_label":"EUR/USD","fx_ticker":"EURUSD=X","bond_label":"EU 10Y Bund","bond_ticker":"EU10YT=RR","fallback_yield":2.55,"erp":5.0,"bench_ticker":"^STOXX50E","bench_name":"Euro Stoxx 50","vix_ticker":"^V2TX","vix_name":"Euro VSTOXX","short_ticker":None,"fallback_short":3.75,"policy_rate":4.00},
        "Japan":         {"fx_label":"USD/JPY","fx_ticker":"JPY=X","bond_label":"Japan 10Y JGB","bond_ticker":"JP10YT=RR","fallback_yield":2.66,"erp":6.0,"bench_ticker":"^N225","bench_name":"Nikkei 225","vix_ticker":"^VIX","vix_name":"Volatility Proxy","short_ticker":None,"fallback_short":0.15,"policy_rate":0.10},
    }
    cfg=MACRO_CONFIG[selected_region]

    live_10y   = fetch_macro_yield(cfg["bond_ticker"],  cfg["fallback_yield"])
    live_short = fetch_macro_yield(cfg["short_ticker"], cfg["fallback_short"])
    vix_level  = fetch_macro_yield(cfg["vix_ticker"],   14.8)

    # ── Systemic Risk Cards ─────────────────────────────────────────────────
    st.markdown("### 🏛️ Systemic Risk & Policy Desk")
    r1,r2,r3=st.columns(3)
    curve=(live_10y-live_short)*100
    curve_lbl = "NORMAL SLOPE" if curve>0 else "⚠️ INVERTED — RECESSION SIGNAL"
    vix_lbl   = ("LOW FEAR — RISK ON" if vix_level<15 else "ELEVATED VOLATILITY" if vix_level<25 else "HIGH FEAR — RISK OFF")
    vix_col   = "#2ed573" if vix_level<15 else "#e1b12c" if vix_level<25 else "#ff4757"
    with r1:
        st.markdown(f"<div class='risk-card' style='border-top-color:#0abde3;'><h4 style='margin:0 0 8px;color:#7f8fa6;font-size:13px;'>Yield Curve (10Y – 3M)</h4><h2 style='margin:0 0 8px;color:#ffffff;font-size:28px;'>{curve:+.0f} bps</h2><strong style='color:#0abde3;font-size:13px;'>{curve_lbl}</strong></div>",unsafe_allow_html=True)
    with r2:
        st.markdown(f"<div class='risk-card' style='border-top-color:#10ac84;'><h4 style='margin:0 0 8px;color:#7f8fa6;font-size:13px;'>Implied Short Rate</h4><h2 style='margin:0 0 8px;color:#ffffff;font-size:28px;'>{live_short:.2f}%</h2><strong style='color:#7f8fa6;font-size:13px;'>Policy Target: {cfg['policy_rate']:.2f}%</strong></div>",unsafe_allow_html=True)
    with r3:
        st.markdown(f"<div class='risk-card' style='border-top-color:{vix_col};'><h4 style='margin:0 0 8px;color:#7f8fa6;font-size:13px;'>{cfg['vix_name']}</h4><h2 style='margin:0 0 8px;color:#ffffff;font-size:28px;'>{vix_level:.2f}</h2><strong style='color:{vix_col};font-size:13px;'>{vix_lbl}</strong></div>",unsafe_allow_html=True)

    st.markdown("---")

    # ── Live Macro Telemetry Tiles ──────────────────────────────────────────
    st.markdown("### 🌐 Live Macro Telemetry")
    macro_tickers={cfg["fx_label"]:cfg["fx_ticker"],"Gold (XAU/USD)":"GC=F","Crude Oil (WTI)":"CL=F"}
    m_data=[{"Indicator":cfg["bond_label"],"Value":f"{live_10y:.2f}%","Change":0.0}]
    with st.spinner("Fetching macro data…"):
        for name,t in macro_tickers.items():
            try:
                h=cached_history(t,"5d")
                if len(h)>=2:
                    c=float(h["Close"].iloc[-1]); p=float(h["Close"].iloc[-2])
                    m_data.append({"Indicator":name,"Value":f"{c:,.2f}","Change":((c-p)/p)*100})
            except: pass
    if m_data:
        cols=st.columns(len(m_data))
        for i,row in enumerate(m_data):
            with cols[i]:
                c="#2ed573" if row["Change"]>0 else "#ff4757" if row["Change"]<0 else "#7f8fa6"
                st.markdown(f"<div class='metric-card' style='border-left-color:{c};'><h4 style='margin:0;font-size:13px;color:#7f8fa6;'>{row['Indicator']}</h4><p style='font-size:22px;font-weight:bold;margin:5px 0;'>{row['Value']}</p><span style='color:{c};font-weight:bold;font-size:13px;'>{row['Change']:+.2f}%</span></div>",unsafe_allow_html=True)

    st.markdown("---")

    # ── CAPM Beta Regression ────────────────────────────────────────────────
    st.markdown(f"### 🧮 Dynamic CAPM Engine — {selected_region}")
    st.caption(f"3-year rolling beta regression against the {cfg['bench_name']}.")
    ticker_sym=ticker_search_widget(default_query="Reliance")
    if not ticker_sym: st.stop()

    with st.spinner(f"Running regression for {ticker_sym}…"):
        try:
            ad=cached_history(ticker_sym,"3y"); bd=cached_history(cfg["bench_ticker"],"3y")
            if not ad.empty and not bd.empty:
                ar=ad["Close"].pct_change().dropna(); br=bd["Close"].pct_change().dropna()
                aln=pd.concat([ar,br],axis=1).dropna(); aln.columns=["Asset","Benchmark"]
                beta   = aln.cov().iloc[0,1]/aln["Benchmark"].var()
                ke     = live_10y + beta*cfg["erp"]
                c1,c2,c3,c4=st.columns(4)
                c1.metric("3Y Beta",         f"{beta:.2f}x",  help=f"vs {cfg['bench_name']}")
                c2.metric("Risk-Free Rate",  f"{live_10y:.2f}%")
                c3.metric("Equity Risk Prem",f"{cfg['erp']:.1f}%",help="Damodaran country ERP")
                c4.metric("Cost of Equity",  f"{ke:.2f}%",    help="Rf + β × ERP")
                st.info(f"**CAPM:** Investors require **{ke:.2f}%** annual return for **{ticker_sym}**. Beta of {beta:.2f}x implies {'higher' if beta>1 else 'lower'} systematic risk than {cfg['bench_name']}.")
                cd=(aln*100).reset_index(drop=True)
                scatter=alt.Chart(cd).mark_circle(size=40,opacity=0.4,color="#0abde3").encode(
                    x=alt.X("Benchmark:Q",title=f"{cfg['bench_name']} Daily Return (%)"),
                    y=alt.Y("Asset:Q",title="Asset Daily Return (%)"),tooltip=["Benchmark","Asset"])
                reg=scatter.transform_regression("Benchmark","Asset").mark_line(color="#ff4757",size=2)
                st.altair_chart((scatter+reg).properties(height=300,title=f"Beta Regression — {ticker_sym} vs {cfg['bench_name']}"),use_container_width=True)
            else:
                st.warning("Insufficient data for regression.")
        except Exception as e:
            st.error(f"CAPM error: {e}")