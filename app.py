import io
import re
from datetime import datetime
from html import escape
from difflib import SequenceMatcher
from typing import Optional, List, Dict, Tuple

import pandas as pd
import streamlit as st

try:
    from rapidfuzz import fuzz
except Exception:
    fuzz = None

try:
    from pypdf import PdfReader
except Exception:
    try:
        from PyPDF2 import PdfReader
    except Exception:
        PdfReader = None

try:
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except Exception:
    PatternFill = Font = Alignment = Border = Side = get_column_letter = None

PAS_YELLOW = "#FFD400"
PAS_BLACK = "#0A0A0A"
GREEN = "#C6EFCE"
RED = "#FFC7CE"
ORANGE = "#FCE4D6"
GREY = "#E7E6E6"

st.set_page_config(page_title="PAS Vendor Hire Report Checker", page_icon="pas_logo.png", layout="wide")

st.markdown(f"""
<style>
.stApp {{ background:#f7f8fa; color:#0A0A0A; font-family:Inter, "Segoe UI", Arial, sans-serif; }}
.block-container {{ max-width:1580px; padding:1.05rem 2rem 2rem 2rem; }}
section[data-testid="stSidebar"] {{ background:linear-gradient(180deg,#050606 0%,#0b1015 100%); border-right:1px solid #161b22; }}
section[data-testid="stSidebar"] * {{ color:white !important; }}
section[data-testid="stSidebar"] img {{ border-radius:14px; box-shadow:0 10px 24px rgba(0,0,0,.26); }}
.pas-sidebar-title {{ color:#fff; font-size:18px; font-weight:950; line-height:1.15; text-align:center; margin:20px 0 8px; }}
.pas-yellow-line {{ width:72px; height:4px; background:{PAS_YELLOW}; border-radius:99px; margin:0 auto 22px; }}
.pas-sidebar-copy {{ color:#fff; font-size:14px; line-height:1.52; font-weight:650; margin-bottom:24px; }}
.pas-sidebar-rule {{ border-top:1px solid rgba(255,255,255,.22); margin:22px 0; }}
.pas-sidebar-heading {{ color:{PAS_YELLOW} !important; font-size:19px; font-weight:950; margin:0 0 16px; }}
.pas-nav-row {{ display:grid; grid-template-columns:26px 1fr; gap:10px; align-items:start; margin:15px 0; color:#fff; font-weight:750; line-height:1.25; font-size:14px; }}
.pas-nav-icon svg {{ width:21px; height:21px; stroke:{PAS_YELLOW}; stroke-width:2.4; fill:none; stroke-linecap:round; stroke-linejoin:round; }}
.pas-sidebar-footer {{ color:#fff; font-size:12px; font-weight:800; margin-top:28px; }}
.pas-hero {{ display:flex; align-items:center; gap:16px; background:linear-gradient(100deg,#08090b 0%,#151718 70%,#c9aa00 130%); border-radius:16px; padding:12px 22px; margin:0 0 18px 0; box-shadow:0 9px 25px rgba(0,0,0,.13); min-height:60px; }}
.pas-hero-logo {{ width:37px; height:37px; border-radius:7px; background:{PAS_YELLOW}; color:#000; display:inline-flex; align-items:center; justify-content:center; font-weight:950; font-size:14px; letter-spacing:-1px; }}
.pas-hero-text {{ color:#fff; font-size:18px; font-weight:950; letter-spacing:-.02em; }}
.pas-hero-dot {{ color:#fff; opacity:.8; margin:0 7px; }}
.pas-hero-version {{ color:{PAS_YELLOW}; font-weight:950; }}
.pas-upload-card {{ background:#fff; border:1px solid #e5e7eb; border-radius:18px; box-shadow:0 5px 18px rgba(15,23,42,.08); padding:16px 18px 14px; margin-bottom:14px; }}
.pas-upload-title {{ color:#0A0A0A; font-size:16px; font-weight:950; margin-bottom:10px; }}
div[data-testid="stFileUploader"] label {{ display:none !important; }}
div[data-testid="stFileUploaderDropzoneInstructions"] {{ display:none !important; }}
div[data-testid="stFileUploader"] section {{ background:transparent !important; border:0 !important; min-height:0 !important; padding:0 !important; }}
div[data-testid="stFileUploader"] button {{ background:#fff !important; color:#0A0A0A !important; border:1px solid #d7dce3 !important; border-radius:10px !important; font-weight:900 !important; min-height:44px !important; box-shadow:0 2px 8px rgba(0,0,0,.06) !important; }}
div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] {{ display:none !important; }}
.pas-file-card {{ display:flex; align-items:center; gap:14px; background:#f4f6f8; border:1px solid #dfe4ea; border-radius:12px; padding:11px 14px; min-height:54px; margin:4px 0 12px; }}
.pas-file-icon {{ width:32px; height:32px; border-radius:8px; display:flex; align-items:center; justify-content:center; color:#fff; font-weight:950; font-size:11px; box-shadow:0 2px 8px rgba(0,0,0,.12); flex:none; }}
.pas-file-icon.excel {{ background:#118a3b; }} .pas-file-icon.pdf {{ background:#df1f2d; }}
.pas-file-main {{ flex:1; min-width:0; }} .pas-file-name {{ color:#0A0A0A; font-weight:950; font-size:15px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }} .pas-file-size {{ color:#4b5563; font-weight:650; font-size:13px; margin-top:2px; }} .pas-file-check {{ width:24px; height:24px; border-radius:50%; background:#108a37; color:white; display:flex; align-items:center; justify-content:center; font-size:15px; font-weight:950; flex:none; }}
.stButton > button, .stDownloadButton > button {{ background:{PAS_YELLOW} !important; color:{PAS_BLACK} !important; border:1px solid {PAS_BLACK} !important; border-radius:12px !important; font-weight:900 !important; min-height:52px !important; font-size:16px !important; box-shadow:0 6px 18px rgba(255,212,0,.25) !important; }}
.stDownloadButton > button {{ min-height:62px !important; font-size:20px !important; }}
.kpi-card {{ background:#fff; border-radius:18px; border:1px solid #e4e7eb; box-shadow:0 5px 20px rgba(15,23,42,.08); min-height:118px; padding:18px 22px; display:flex; align-items:center; gap:18px; }}
.kpi-icon {{ width:64px; height:64px; border-radius:50%; background:#fff5bd; display:flex; align-items:center; justify-content:center; flex:none; }}
.kpi-icon svg {{ width:35px; height:35px; stroke:#0A0A0A; stroke-width:2.5; fill:none; stroke-linecap:round; stroke-linejoin:round; }}
.kpi-label {{ color:#111; font-size:15px; font-weight:950; margin:0 0 3px; }}
.kpi-value {{ color:#e9b900; font-size:42px; line-height:.98; font-weight:950; }}
.kpi-sub {{ color:#374151; font-size:14px; margin-top:6px; }}
.kpi-red .kpi-value {{ color:#e12626; }} .kpi-green .kpi-value {{ color:#16a34a; }} .kpi-orange .kpi-value {{ color:#f59e0b; }}
.pas-results-title {{ color:#0A0A0A; font-size:28px; font-weight:950; margin:22px 0 8px; }}
.pas-pill {{ display:inline-flex; background:{PAS_YELLOW}; color:#0A0A0A; border-radius:14px 14px 0 0; padding:13px 20px; font-size:18px; font-weight:950; box-shadow:0 4px 14px rgba(0,0,0,.09); }}
.pas-table-wrap {{ background:#fff; border:1px solid #e0e4e9; border-radius:0 16px 16px 16px; max-height:430px; overflow:auto; box-shadow:0 7px 25px rgba(15,23,42,.10); }}
table.pas-table {{ width:100%; border-collapse:collapse; font-size:14px; color:#0A0A0A; }}
table.pas-table thead th {{ background:{PAS_YELLOW}; color:#0A0A0A; border:1px solid #e2ba00; padding:12px 14px; font-weight:950; text-align:left; white-space:nowrap; position:sticky; top:0; z-index:5; }}
table.pas-table tbody td {{ background:#fff; color:#0A0A0A; border:1px solid #e1e5eb; padding:10px 14px; vertical-align:top; }}
table.pas-table tbody tr.green td {{ background:{GREEN}; }} table.pas-table tbody tr.red td {{ background:{RED}; }} table.pas-table tbody tr.orange td {{ background:{ORANGE}; }}
div[data-testid="stAlert"], div[data-testid="stAlert"] * {{ color:#0A0A0A !important; }}
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.image("pas_logo.png", use_column_width=True)
    st.markdown("""
    <div class="pas-sidebar-title">PAS Vendor<br>On-Hire Checker</div>
    <div class="pas-yellow-line"></div>
    <div class="pas-sidebar-copy">Upload a vendor hire report and the PAS Material & Plant Orders spreadsheet, then check whether each chargeable item is still live on hire.</div>
    <div class="pas-sidebar-rule"></div>
    <div class="pas-sidebar-heading">Instructions</div>
    <div class="pas-nav-row"><span class="pas-nav-icon"><svg viewBox="0 0 24 24"><path d="M16 16l-4-4-4 4"/><path d="M12 12v9"/><path d="M20 16.6A5 5 0 0 0 18 7h-1.3A8 8 0 1 0 4 15.3"/></svg></span><span>Upload Vendor Hire Report</span></div>
    <div class="pas-nav-row"><span class="pas-nav-icon"><svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M9 13h6"/><path d="M9 17h6"/></svg></span><span>Upload Materials & Plant Report</span></div>
    <div class="pas-nav-row"><span class="pas-nav-icon"><svg viewBox="0 0 24 24"><path d="M5 3l14 9-14 9V3z"/></svg></span><span>Run Reconciliation</span></div>
    <div class="pas-nav-row"><span class="pas-nav-icon"><svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="M7 10l5 5 5-5"/><path d="M12 15V3"/></svg></span><span>Download Checked Excel</span></div>
    <div class="pas-sidebar-rule"></div>
    <div class="pas-sidebar-footer">PAS NW Ltd • v1.0 Prototype Build</div>
    """, unsafe_allow_html=True)

st.markdown("""
<div class="pas-hero"><div class="pas-hero-logo">PAS</div><div class="pas-hero-text">PAS Vendor On-Hire Checker<span class="pas-hero-dot">•</span><span class="pas-hero-version">v1.0 Prototype Build</span></div></div>
""", unsafe_allow_html=True)

# ---------- helpers ----------
def clean_cell(value) -> str:
    if value is None or pd.isna(value): return ""
    text = str(value).strip()
    return "" if text.lower() in {"nan", "none", "nat"} else text

def normalise_text(text: str) -> str:
    text = clean_cell(text).lower()
    repl = {
        "tonne":"t", "ton":"t", "tons":"t", "3 ton":"3t", "3 tonne":"3t", "1 ton":"1t", "1 tonne":"1t",
        "mini digger":"excavator", "digger":"excavator", "cutquick":"stihl saw", "disc cutter":"stihl saw",
        "cabins":"cabin", "site unit":"cabin", "welfare unit":"welfare", "telehandler":"telehandler",
        "fork lift":"forklift", "fork-lift":"forklift", "genie":"lift", "generator":"genny"
    }
    for k, v in repl.items(): text = text.replace(k, v)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    stop = {"the","and","with","c/w","cw","hire","plant","ltd","limited","weekly","week","per","for","to","of","x"}
    words = [w for w in text.split() if w not in stop]
    return " ".join(words)

def money_to_float(value) -> Optional[float]:
    if value is None or pd.isna(value): return None
    if isinstance(value, (int, float)): return round(float(value), 2)
    text = str(value).replace(",", "")
    m = re.findall(r"-?£?\s*([0-9]+(?:\.[0-9]+)?)", text)
    if not m: return None
    try: return round(float(m[-1]), 2)
    except Exception: return None

def normalise_po(value) -> str:
    text = clean_cell(value).upper()
    # Vendor order often looks like P114/H4767. PAS has Job No P114 and Order Number H4767.
    parts = re.findall(r"[A-Z]+\d+[A-Z0-9&]*", text)
    return " ".join(parts) if parts else re.sub(r"[^A-Z0-9]+", " ", text).strip()

def extract_job(value) -> str:
    text = clean_cell(value).upper()
    m = re.search(r"\bP\d+[A-Z0-9&]*\b", text)
    return m.group(0) if m else ""

def extract_hire_no(value) -> str:
    text = clean_cell(value).upper()
    m = re.search(r"\bH\d+\b", text)
    return m.group(0) if m else ""

def job_base(job: str) -> str:
    m = re.match(r"^(P\d+)", clean_cell(job).upper())
    return m.group(1) if m else clean_cell(job).upper()

def find_col(columns: List[str], keywords: List[str]) -> Optional[str]:
    norm = {c: re.sub(r"[^a-z0-9]+", "", str(c).lower()) for c in columns}
    for key in keywords:
        nkey = re.sub(r"[^a-z0-9]+", "", key.lower())
        for col, ncol in norm.items():
            if nkey == ncol: return col
    for key in keywords:
        nkey = re.sub(r"[^a-z0-9]+", "", key.lower())
        for col, ncol in norm.items():
            if nkey in ncol or ncol in nkey: return col
    return None

def similarity(a: str, b: str) -> float:
    a2, b2 = normalise_text(a), normalise_text(b)
    if not a2 or not b2: return 0.0
    if fuzz:
        return max(fuzz.token_set_ratio(a2, b2), fuzz.partial_ratio(a2, b2)) / 100
    return SequenceMatcher(None, a2, b2).ratio()

def format_date(value):
    if value is None or pd.isna(value): return ""
    try:
        d = pd.to_datetime(value, dayfirst=True, errors="coerce")
        if pd.isna(d): return clean_cell(value)
        return d.strftime("%d/%m/%Y")
    except Exception:
        return clean_cell(value)

def render_selected_file_card(uploaded_file, file_kind="excel"):
    size = getattr(uploaded_file, "size", 0) or 0
    size_text = f"{size/(1024*1024):.1f} MB" if size >= 1024*1024 else f"{size/1024:.0f} KB"
    icon = "PDF" if file_kind == "pdf" else "XLS"
    icon_class = "pdf" if file_kind == "pdf" else "excel"
    st.markdown(f'''<div class="pas-file-card"><div class="pas-file-icon {icon_class}">{icon}</div><div class="pas-file-main"><div class="pas-file-name">{escape(getattr(uploaded_file,'name','Uploaded file'))}</div><div class="pas-file-size">{size_text}</div></div><div class="pas-file-check">✓</div></div>''', unsafe_allow_html=True)

# ---------- file loading ----------
def read_excel_any(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".xls"):
        return pd.read_excel(uploaded_file, engine="xlrd")
    return pd.read_excel(uploaded_file)

def load_vendor_report(uploaded_file) -> Tuple[pd.DataFrame, pd.DataFrame]:
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        if PdfReader is None:
            raise RuntimeError("PDF reader unavailable. Add pypdf to requirements.txt.")
        pdf_bytes = uploaded_file.read(); uploaded_file.seek(0)
        reader = PdfReader(io.BytesIO(pdf_bytes))
        rows = []
        for page_no, page in enumerate(reader.pages, start=1):
            text = page.extract_text() or ""
            for line in text.splitlines():
                line = re.sub(r"\s+", " ", line).strip()
                if not line: continue
                rate = money_to_float(line)
                if rate is None: continue
                rows.append({"Description": line, "Hire Rate": rate, "Source Page": page_no})
        raw = pd.DataFrame(rows)
    else:
        raw = read_excel_any(uploaded_file)
    raw = raw.dropna(how="all").copy()
    cols = list(raw.columns)
    mapping = {
        "Vendor Site": find_col(cols, ["Site", "Location", "Project", "Job", "Site Name"]),
        "Vendor Fleet No": find_col(cols, ["Item No", "Fleet No", "Fleet", "Asset", "Serial", "Plant No"]),
        "Vendor Description": find_col(cols, ["Description", "Item Description", "Product", "Equipment", "Plant Description"]),
        "Vendor Qty": find_col(cols, ["Quantity", "Qty"]),
        "Vendor On Hire Date": find_col(cols, ["Date", "On Hire Date", "Start Date", "Delivery Date", "Hired Date"]),
        "Vendor Contract No": find_col(cols, ["Syrinx Contract No", "Contract No", "Contract", "Hire Contract"]),
        "Vendor Order No": find_col(cols, ["Order No", "PO", "Purchase Order", "Order Number", "Customer Order"]),
        "Vendor Rate": find_col(cols, ["Hire Rate", "Rate", "Weekly Rate", "Cost", "Value", "Charge"]),
    }
    out = pd.DataFrame()
    for new_col, old_col in mapping.items():
        out[new_col] = raw[old_col] if old_col in raw.columns else ""
    out["Vendor Rate Value"] = out["Vendor Rate"].apply(money_to_float)
    out["Vendor Job"] = out["Vendor Order No"].apply(extract_job)
    out["Vendor Hire No"] = out["Vendor Order No"].apply(extract_hire_no)
    out["Vendor Description Clean"] = out["Vendor Description"].apply(normalise_text)
    out["Vendor Row No"] = range(2, len(out) + 2)
    return raw, out

def load_pas_plant(uploaded_file) -> pd.DataFrame:
    xls = pd.ExcelFile(uploaded_file)
    sheet = "Plant" if "Plant" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet).dropna(how="all").copy()
    cols = list(df.columns)
    out = pd.DataFrame()
    out["PAS Description"] = df[find_col(cols, ["Description", "Item Description"])] if find_col(cols, ["Description", "Item Description"]) else ""
    out["PAS Fleet No"] = df[find_col(cols, ["Fleet No.", "Fleet No", "Fleet", "Item No"])] if find_col(cols, ["Fleet No.", "Fleet No", "Fleet", "Item No"]) else ""
    out["PAS Supplier"] = df[find_col(cols, ["Supplier", "Vendor"])] if find_col(cols, ["Supplier", "Vendor"]) else ""
    out["PAS Qty"] = df[find_col(cols, ["Qty", "Quantity"])] if find_col(cols, ["Qty", "Quantity"]) else ""
    out["PAS Weekly Cost"] = df[find_col(cols, ["Cost", "Weekly Rate", "Hire Rate"])] if find_col(cols, ["Cost", "Weekly Rate", "Hire Rate"]) else ""
    out["PAS Day Rate"] = df[find_col(cols, ["Day Rate", "Daily Rate"])] if find_col(cols, ["Day Rate", "Daily Rate"]) else ""
    out["PAS On Hire Date"] = df[find_col(cols, ["On Hire / Delivery Date", "On Hire Date", "Delivery Date"])] if find_col(cols, ["On Hire / Delivery Date", "On Hire Date", "Delivery Date"]) else ""
    out["PAS Expected Off Hire Date"] = df[find_col(cols, ["Expected Off-Hire Date", "Expected Off Hire Date"])] if find_col(cols, ["Expected Off-Hire Date", "Expected Off Hire Date"]) else ""
    out["PAS Off Hire Date"] = df[find_col(cols, ["Off Hire Date", "Off-Hire Date"])] if find_col(cols, ["Off Hire Date", "Off-Hire Date"]) else ""
    out["PAS Status"] = df[find_col(cols, ["Status"])] if find_col(cols, ["Status"]) else ""
    out["PAS Job No"] = df[find_col(cols, ["Job No", "Job Number", "Job"])] if find_col(cols, ["Job No", "Job Number", "Job"]) else ""
    out["PAS Order Number"] = df[find_col(cols, ["Order Number", "Sage Order No", "PO"])] if find_col(cols, ["Order Number", "Sage Order No", "PO"]) else ""
    out["PAS Site Name"] = df[find_col(cols, ["Site Name", "Site"])] if find_col(cols, ["Site Name", "Site"]) else ""
    out["PAS Description Clean"] = out["PAS Description"].apply(normalise_text)
    out["PAS Job Base"] = out["PAS Job No"].apply(job_base)
    out["PAS Row No"] = range(2, len(out) + 2)
    return out

# ---------- matching ----------
def pas_is_live(status: str) -> bool:
    status_text = clean_cell(status).lower().strip()
    return status_text in {"on hire", "missing"}

def pas_is_off_hired(status: str) -> bool:
    status_text = clean_cell(status).lower().strip()
    return "off" in status_text and "hire" in status_text

def score_candidate(vrow, prow) -> Tuple[float, List[str]]:
    """Score likely PAS matches. Price is deliberately ignored.

    The goal is only to decide whether the vendor line is still live on the
    PAS Materials & Plant report. Fleet helps confidence, but a fleet mismatch
    does not fail the line.
    """
    reasons = []
    score = 0.0

    desc_score = similarity(vrow.get("Vendor Description", ""), prow.get("PAS Description", ""))
    score += desc_score * 65
    if desc_score >= 0.88:
        reasons.append("Strong description match")
    elif desc_score >= 0.72:
        reasons.append("Possible description match")

    vjob = job_base(vrow.get("Vendor Job", ""))
    pjob = job_base(prow.get("PAS Job No", ""))
    if vjob and pjob and vjob == pjob:
        score += 22
        reasons.append("Job matches")

    # Vendor reports commonly show P114/H4767. The H number is the PAS order.
    vhire = clean_cell(vrow.get("Vendor Hire No", "")).upper()
    phire = clean_cell(prow.get("PAS Order Number", "")).upper()
    if vhire and phire and vhire == phire:
        score += 32
        reasons.append("Order number matches")

    vfleet = clean_cell(vrow.get("Vendor Fleet No", "")).upper()
    pfleet = clean_cell(prow.get("PAS Fleet No", "")).upper()
    if vfleet and pfleet and vfleet == pfleet:
        score += 8
        reasons.append("Fleet matches")

    return score, reasons

def reconcile(vendor_df: pd.DataFrame, pas_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    results = []
    ignored = []

    for _, vrow in vendor_df.iterrows():
        rate = vrow.get("Vendor Rate Value")
        if rate is not None and abs(float(rate)) < 0.005:
            row = vrow.to_dict()
            row["Status"] = "Ignored"
            row["Reason"] = "£0 cost/rate ignored"
            row["Colour"] = "Grey"
            ignored.append(row)
            continue

        candidates = []
        for _, prow in pas_df.iterrows():
            s, rs = score_candidate(vrow, prow)
            # Keep this fairly low because descriptions vary supplier-to-supplier,
            # but PO/job/order number will push real matches to the top.
            if s >= 35:
                candidates.append((s, rs, prow))
        candidates.sort(key=lambda x: x[0], reverse=True)
        best = candidates[0] if candidates else None

        status = "Query"
        colour = "Red"
        reason = "No matching item found on PAS report"
        match_score = 0
        prow = None

        if best:
            match_score, reason_bits, prow = best
            desc_score = similarity(vrow.get("Vendor Description", ""), prow.get("PAS Description", ""))
            reason = "; ".join(reason_bits) or "Best match found"

            v_fleet = clean_cell(vrow.get("Vendor Fleet No", "")).upper()
            p_fleet = clean_cell(prow.get("PAS Fleet No", "")).upper()
            fleet_mismatch = bool(v_fleet and p_fleet and v_fleet != p_fleet)
            status_text = clean_cell(prow.get("PAS Status", ""))
            status_low = status_text.lower()
            pas_desc_low = clean_cell(prow.get("PAS Description", "")).lower()

            if "operated plant" in status_low or "operated plant" in pas_desc_low:
                status, colour, reason = "Query", "Red", "Operated plant item - manual review required"
            elif pas_is_off_hired(status_text):
                status, colour, reason = "Query", "Red", "PAS item is off-hired but still appears on vendor report"
            elif not pas_is_live(status_text):
                status, colour, reason = "Query", "Red", f"PAS status is '{status_text or 'blank'}' - not live on hire"
            elif match_score >= 72 and desc_score >= 0.72:
                if fleet_mismatch:
                    status, colour = "Warning", "Orange"
                    reason = reason + "; Fleet mismatch only - item still appears live on PAS"
                else:
                    status, colour = "Matched", "Green"
                    reason = reason + "; Item still live on PAS"
            elif match_score >= 55 and desc_score >= 0.62:
                status, colour = "Warning", "Orange"
                reason = reason + "; Possible fuzzy match - PAS item appears live, review wording"
            else:
                status, colour = "Query", "Red"
                reason = "Weak match only - manual review required"

        out = vrow.to_dict()
        out.update({
            "Status": status,
            "Colour": colour,
            "Reason": reason,
            "Match Score": round(float(match_score), 1),
            "PAS Description": clean_cell(prow.get("PAS Description", "")) if prow is not None else "",
            "PAS Fleet No": clean_cell(prow.get("PAS Fleet No", "")) if prow is not None else "",
            "PAS Supplier": clean_cell(prow.get("PAS Supplier", "")) if prow is not None else "",
            "PAS Qty": clean_cell(prow.get("PAS Qty", "")) if prow is not None else "",
            "PAS On Hire Date": format_date(prow.get("PAS On Hire Date", "")) if prow is not None else "",
            "PAS Off Hire Date": format_date(prow.get("PAS Off Hire Date", "")) if prow is not None else "",
            "PAS Status": clean_cell(prow.get("PAS Status", "")) if prow is not None else "",
            "PAS Job No": clean_cell(prow.get("PAS Job No", "")) if prow is not None else "",
            "PAS Order Number": clean_cell(prow.get("PAS Order Number", "")) if prow is not None else "",
        })
        results.append(out)

    return pd.DataFrame(results), pd.DataFrame(ignored)

# ---------- excel output ----------
def style_workbook(writer, sheet_dfs: Dict[str, pd.DataFrame], colour_col_map: Dict[str, str] = None):
    if not all([PatternFill, Font, Alignment, Border, Side, get_column_letter]): return
    colour_col_map = colour_col_map or {}
    fills = {
        "header": PatternFill("solid", fgColor="FFD400"),
        "Green": PatternFill("solid", fgColor="C6EFCE"),
        "Red": PatternFill("solid", fgColor="FFC7CE"),
        "Orange": PatternFill("solid", fgColor="FCE4D6"),
        "Grey": PatternFill("solid", fgColor="E7E6E6"),
    }
    header_font = Font(name="Calibri", size=10, bold=True, color="000000")
    body_font = Font(name="Calibri", size=10, color="000000")
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
    for sheet, df in sheet_dfs.items():
        ws = writer.book[sheet]
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 25
        colour_col = colour_col_map.get(sheet)
        colour_idx = list(df.columns).index(colour_col) + 1 if colour_col and colour_col in df.columns else None
        for row in ws.iter_rows():
            row_colour = None
            if row[0].row == 1:
                row_colour = "header"
            elif colour_idx:
                row_colour = clean_cell(ws.cell(row=row[0].row, column=colour_idx).value)
            for cell in row:
                cell.font = header_font if cell.row == 1 else body_font
                cell.alignment = align
                cell.border = border
                if row_colour in fills:
                    cell.fill = fills[row_colour]
        ws.auto_filter.ref = ws.dimensions
        for idx, col in enumerate(df.columns, start=1):
            values = [str(col)] + [str(v) for v in df[col].fillna("").astype(str).tolist()[:200]]
            width = min(max(len(v) for v in values) + 2, 50)
            ws.column_dimensions[get_column_letter(idx)].width = width

def make_excel(results_df: pd.DataFrame, ignored_df: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    cols = [
        "Status", "Reason", "Vendor Site", "Vendor Fleet No", "Vendor Description", "Vendor Qty", "Vendor On Hire Date",
        "Vendor Contract No", "Vendor Order No", "Vendor Rate", "Vendor Rate Value", "Match Score",
        "PAS Description", "PAS Fleet No", "PAS Supplier", "PAS Qty", "PAS On Hire Date", "PAS Off Hire Date", "PAS Status", "PAS Job No", "PAS Order Number", "Colour"
    ]
    export = results_df.copy()
    for c in cols:
        if c not in export.columns: export[c] = ""
    export = export[cols]
    queries = export[export["Status"] == "Query"].copy()
    warnings = export[export["Status"] == "Warning"].copy()
    matched = export[export["Status"] == "Matched"].copy()
    ignored = ignored_df.copy()
    if not ignored.empty:
        for c in ["Status", "Reason"]:
            if c not in ignored.columns: ignored[c] = ""
    sheet_dfs = {
        "Summary": summary_df,
        "Vendor Report Checked": export,
        "Queries Only": queries,
        "Warnings": warnings,
        "Matched Only": matched,
        "Ignored £0 Items": ignored,
    }
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet, df in sheet_dfs.items():
            df.to_excel(writer, index=False, sheet_name=sheet[:31])
        style_workbook(writer, sheet_dfs, {"Vendor Report Checked": "Colour", "Queries Only": "Colour", "Warnings": "Colour", "Matched Only": "Colour"})
    return output.getvalue()

# ---------- UI ----------
up_col1, up_col2 = st.columns(2)
with up_col1:
    st.markdown('<div class="pas-upload-card"><div class="pas-upload-title">Upload Vendor Hire Report</div>', unsafe_allow_html=True)
    vendor_file = st.file_uploader("Upload Vendor Hire Report", type=["xlsx", "xls", "pdf"], label_visibility="collapsed", key="vendor_upload")
    if vendor_file: render_selected_file_card(vendor_file, "pdf" if vendor_file.name.lower().endswith(".pdf") else "excel")
    st.markdown('</div>', unsafe_allow_html=True)
with up_col2:
    st.markdown('<div class="pas-upload-card"><div class="pas-upload-title">Upload Materials & Plant Orders</div>', unsafe_allow_html=True)
    pas_file = st.file_uploader("Upload Materials & Plant Orders", type=["xlsx", "xls"], label_visibility="collapsed", key="pas_upload")
    if pas_file: render_selected_file_card(pas_file, "excel")
    st.markdown('</div>', unsafe_allow_html=True)

run = st.button("▶  Run reconciliation", use_container_width=True)
if "vendor_checker_results" not in st.session_state:
    st.session_state["vendor_checker_results"] = None

if run:
    if not vendor_file or not pas_file:
        st.warning("Please upload both the vendor hire report and the Materials & Plant Orders spreadsheet.")
        st.stop()
    try:
        with st.spinner("Reading vendor hire report..."):
            raw_vendor, vendor_df = load_vendor_report(vendor_file)
        with st.spinner("Reading PAS Materials & Plant Orders..."):
            pas_df = load_pas_plant(pas_file)
        with st.spinner("Checking whether vendor items are still live on PAS report..."):
            results_df, ignored_df = reconcile(vendor_df, pas_df)
        total_checked = len(results_df)
        matched = int((results_df["Status"] == "Matched").sum()) if not results_df.empty else 0
        warnings = int((results_df["Status"] == "Warning").sum()) if not results_df.empty else 0
        queries = int((results_df["Status"] == "Query").sum()) if not results_df.empty else 0
        ignored = len(ignored_df)
        match_pct = round((matched / total_checked) * 100, 1) if total_checked else 0.0
        summary_df = pd.DataFrame({
            "Metric": ["Total chargeable lines checked", "Matched lines", "Warnings", "Queries", "Ignored £0 lines", "Match percentage", "Run date/time"],
            "Value": [total_checked, matched, warnings, queries, ignored, f"{match_pct}%", datetime.now().strftime("%d/%m/%Y %H:%M")]
        })
        excel_bytes = make_excel(results_df, ignored_df, summary_df)
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        st.session_state["vendor_checker_results"] = {
            "results_df": results_df, "ignored_df": ignored_df, "summary_df": summary_df, "excel_bytes": excel_bytes,
            "total": total_checked, "matched": matched, "warnings": warnings, "queries": queries, "ignored": ignored,
            "match_pct": match_pct, "filename": f"PAS_Vendor_On_Hire_Checked_{stamp}.xlsx"
        }
    except Exception as e:
        st.error(f"Something went wrong: {e}")
        st.exception(e)

res = st.session_state.get("vendor_checker_results")
if res:
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: st.markdown(f'<div class="kpi-card"><div class="kpi-icon"><svg viewBox="0 0 24 24"><path d="M8 7V3h8l4 4v14H6V7z"/><path d="M16 3v5h5"/><path d="M9 13h6"/><path d="M9 17h4"/></svg></div><div><div class="kpi-label">Checked</div><div class="kpi-value">{res["total"]}</div><div class="kpi-sub">Chargeable lines</div></div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="kpi-card kpi-green"><div class="kpi-icon"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M8 12.5l2.7 2.7L16.5 9"/></svg></div><div><div class="kpi-label">Matched</div><div class="kpi-value">{res["matched"]}</div><div class="kpi-sub">Green lines</div></div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="kpi-card kpi-orange"><div class="kpi-icon"><svg viewBox="0 0 24 24"><path d="M12 3l10 18H2L12 3z"/><path d="M12 9v5"/><path d="M12 18h.01"/></svg></div><div><div class="kpi-label">Warnings</div><div class="kpi-value">{res["warnings"]}</div><div class="kpi-sub">Orange lines</div></div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="kpi-card kpi-red"><div class="kpi-icon"><svg viewBox="0 0 24 24"><path d="M12 3l10 18H2L12 3z"/><path d="M12 9v5"/><path d="M12 18h.01"/></svg></div><div><div class="kpi-label">Queries</div><div class="kpi-value">{res["queries"]}</div><div class="kpi-sub">Red lines</div></div></div>', unsafe_allow_html=True)
    with c5: st.markdown(f'<div class="kpi-card"><div class="kpi-icon"><svg viewBox="0 0 24 24"><path d="M3 20h18"/><path d="M6 16v-4"/><path d="M11 16V8"/><path d="M16 16v-6"/><path d="M19 6l-5 5-3-3-5 5"/></svg></div><div><div class="kpi-label">Match %</div><div class="kpi-value">{res["match_pct"]}%</div><div class="kpi-sub">Core KPI</div></div></div>', unsafe_allow_html=True)

    st.markdown('<div class="pas-results-title">Results Preview</div>', unsafe_allow_html=True)
    preview_cols = ["Status", "Reason", "Vendor Site", "Vendor Fleet No", "Vendor Description", "Vendor Order No", "Vendor Rate", "PAS Description", "PAS Fleet No", "PAS Status", "PAS Job No", "Colour"]
    df = res["results_df"].copy()
    for c in preview_cols:
        if c not in df.columns: df[c] = ""
    rows = []
    for _, row in df[preview_cols].head(100).iterrows():
        cls = clean_cell(row.get("Colour", "")).lower()
        cells = "".join(f"<td>{escape(clean_cell(row.get(col, '')))}</td>" for col in preview_cols[:-1])
        rows.append(f"<tr class='{cls}'>{cells}</tr>")
    header = "".join(f"<th>{escape(c)}</th>" for c in preview_cols[:-1])
    st.markdown('<div class="pas-pill">Vendor Report Checked</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="pas-table-wrap"><table class="pas-table"><thead><tr>{header}</tr></thead><tbody>{"".join(rows)}</tbody></table></div>', unsafe_allow_html=True)
    st.caption(f"Ignored £0 lines: {res['ignored']} | Price is ignored apart from £0 lines, which are excluded.")
    dl_left, dl_mid, dl_right = st.columns([1.3, 1, 1.3])
    with dl_mid:
        st.download_button("⬇  Download checked Excel", data=res["excel_bytes"], file_name=res["filename"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
else:
    st.info("Upload a vendor hire report and Materials & Plant Orders spreadsheet, then click Run reconciliation.")
