# ba_rate_pages_app.py  —  streamlit run app.py

import io
import streamlit as st
from pathlib import Path

import subprocess, sys

st.set_page_config(
    page_title="RatePage Builder · Nationwide",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Session state ────────────────────────────────────────────────────────────
REQUIRED = ["NGIC", "MM", "NACO", "NAFF", "NICOF", "HICNJ", "CCMIC", "NWAG"]
OPTIONAL = ["CW"]
ALL_KEYS = REQUIRED + OPTIONAL

for k in ALL_KEYS:
    st.session_state.setdefault(f"file_{k}", None)
st.session_state.setdefault("save_dir",    "")
st.session_state.setdefault("sched_mod",   0)
st.session_state.setdefault("run_status",  "idle")
st.session_state.setdefault("run_msg",     "")
st.session_state.setdefault("xlsx_path",   "")
st.session_state.setdefault("pdf_path",    "")
st.session_state.setdefault("pdf_status",  "idle")
st.session_state.setdefault("lob",          "Business Auto")
st.session_state.setdefault("confirm_step", "idle")
st.session_state.setdefault("upload_reset", 0)
st.session_state.setdefault("mode",          "individual")
st.session_state.setdefault("multi_src_dir", "")
st.session_state.setdefault("multi_save_dir","")
st.session_state.setdefault("multi_gen_pdf", False)
st.session_state.setdefault("multi_step",    "idle")
st.session_state.setdefault("multi_results", [])
st.session_state.setdefault("multi_sched",        0)    # kept for compat
st.session_state.setdefault("multi_sched_map",    {})   # {state_name: int} per-state mods
st.session_state.setdefault("multi_sched_mode",   "upload")  # "upload" | "manual"
st.session_state.setdefault("multi_sched_excel",  {})   # raw parsed excel {state: rating}

LOB_NAV = [("Business Auto","🚗"),
    ("General Liability", "⚖️"),
    ("Farm Auto",         "🚜"),
    ("Property",          "🏠"),
]
LOB_NAMES   = [l for l, _ in LOB_NAV]
LOB_OPTIONS = [f"{ic}  {nm}" for nm, ic in LOB_NAV]

# ─── Helpers ──────────────────────────────────────────────────────────────────
def _valid(k):
    v = st.session_state.get(f"file_{k}")
    return v is not None and "error" not in v

def n_req():    return sum(1 for k in REQUIRED if _valid(k))
def all_req():  return n_req() == len(REQUIRED)
def any_req():  return n_req() > 0
def has_ngic(): return _valid("NGIC")

def chip(f):
    if f:
        return f'<span class="chip-ok">&#10003; {Path(f["name"]).stem[:17]}</span>'
    return '<span class="chip-none">&#8212; none</span>'

def browse_folder():
    # Run tkinter in a separate process — calling tk.Tk() from a Streamlit
    # background thread on Windows crashes the server process entirely.
    try:
        result = subprocess.run(
            [sys.executable, "-c",
             "import tkinter as tk; from tkinter import filedialog; "
             "root=tk.Tk(); root.withdraw(); root.wm_attributes('-topmost',True); "
             "print(filedialog.askdirectory(title='Select Save Location') or '', end='')"],
            capture_output=True, text=True, timeout=120,
        )
        folder = result.stdout.strip()
        return folder or None
    except Exception:
        return None

def spacer(px=20):
    st.markdown(f"<div style='height:{px}px'></div>", unsafe_allow_html=True)


# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Libre+Baskerville:wght@700&display=swap');

/* ── tokens ── */
:root {
  --nw-blue:     #1A5DAB;
  --nw-deep:     #0D3F7A;
  --nw-lt:       #EBF2FB;
  --gold:        #C8A951;
  --gold-lt:     #EDD97A;
  --off:         #F4F7FB;
  --surface:     #FFFFFF;
  --border:      #D4DFEF;
  --text:        #0C1A35;
  --muted:       #6B7A9E;
  --ok-bg:       #EAF5EE;
  --ok-fg:       #196B38;
  --radius:      10px;
  --shadow:      0 2px 12px rgba(13,63,122,0.08);
  --content-pad: 72px;   /* ← single knob: push main content right/left */
}

/* ── kill white top gap ── */
html, body { margin:0 !important; padding:0 !important; background:var(--off) !important; }
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
section.main { background:var(--off) !important; padding:0 !important; margin:0 !important; }
.block-container {
  background:var(--off) !important;
  padding: 0 var(--content-pad) 64px var(--content-pad) !important;
  margin:0 !important; max-width:100% !important;
}
.stMainBlockContainer { padding-top:0 !important; }
#MainMenu, footer,
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stHeader"],
header { display:none !important; height:0 !important; min-height:0 !important; }
[data-testid="stVerticalBlock"] > div:first-child { padding-top:0 !important; }

/* ── sidebar shell ── */
[data-testid="stSidebar"] {
  background: #F7F9FC !important;
  border-right: 1px solid #DDE5F0 !important;
  padding: 0 !important;
}
[data-testid="stSidebar"] > div:first-child {
  padding: 0 !important;
  background: #F7F9FC !important;
}


/* ══════════════════════════════════════════════════════════════════
   SIDEBAR NAV  —  st.radio disguised as a nav menu
   Strategy: hide every Streamlit chrome element, style only the
   <label> elements.  The CHECKED state is read via aria-checked="true"
   on the <label> itself — this is set by the browser's native radio
   behaviour and is 100% reliable across Chrome / Edge / Firefox.
══════════════════════════════════════════════════════════════════ */

/* Hide the radio group label */
[data-testid="stSidebar"] [data-testid="stRadio"] > label { display:none !important; }

/* Zero out EVERY wrapper between the sidebar edge and the label */
[data-testid="stSidebar"] [data-testid="stRadio"] > div { gap:0 !important; padding:0 !important; margin:0 !important; }
[data-testid="stSidebar"] [data-testid="stRadio"] [data-baseweb="radio"] { margin:0 !important; padding:0 !important; width:100% !important; }
[data-testid="stSidebar"] [data-testid="stRadio"] [data-baseweb="radio"] > div { margin:0 !important; padding:0 !important; width:100% !important; }
[data-testid="stSidebar"] [data-testid="stRadio"] > div > div { margin:0 !important; padding:0 !important; width:100% !important; }

/* Hide the actual circle dot */
[data-testid="stSidebar"] [data-testid="stRadio"] label > div:first-child { display:none !important; }

/* ── base style for every nav label ── */
[data-testid="stSidebar"] [data-testid="stRadio"] label {
  display: flex !important;
  align-items: center !important;
  gap: 10px !important;
  width: 100% !important;
  min-width: 100% !important;
  margin: 0 !important;
  padding: 13px 20px !important;
  border-radius: 0 !important;
  cursor: pointer !important;
  font-size: 13px !important;
  font-family: 'Inter', sans-serif !important;
  font-weight: 500 !important;
  color: #566278 !important;
  background: transparent !important;
  border: none !important;
  border-left: 3px solid transparent !important;
  transition: background 0.14s ease, color 0.14s ease !important;
  user-select: none !important;
  box-sizing: border-box !important;
}

/* ── hover ── */
[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
  background: #EBF2FB !important;
  color: #0D3F7A !important;
  border-left: 3px solid #A8C4E8 !important;
}

/* ── SELECTED item — full-width blue row, every selector variant ── */

/* 1. aria-checked on label (Streamlit >= 1.28) */
[data-testid="stSidebar"] [data-testid="stRadio"] label[aria-checked="true"] {
  background: linear-gradient(90deg, #1A5DAB 0%, #0D3F7A 100%) !important;
  color: #FFFFFF !important;
  font-weight: 700 !important;
  border-left: 3px solid #EDD97A !important;
  box-shadow: none !important;
}
[data-testid="stSidebar"] [data-testid="stRadio"] label[aria-checked="true"]:hover {
  background: linear-gradient(90deg, #1A5DAB 0%, #0D3F7A 100%) !important;
  color: #FFFFFF !important;
}
            
/* Container div — zero all spacing */
[data-testid="stSidebar"] [data-testid="stImage"]        { margin:0; padding:0; line-height:0; }
[data-testid="stSidebar"] [data-testid="stImage"] > div  { margin:0; padding:0; }

/* The img itself gets the breathing room via padding instead */
[data-testid="stSidebar"] [data-testid="stImage"] img    { display:block; padding:16px 20px 12px; margin:0; }

/* 2. checked input sibling */
[data-testid="stSidebar"] [data-testid="stRadio"] input[type="radio"]:checked ~ div,
[data-testid="stSidebar"] [data-testid="stRadio"] input[type="radio"]:checked + div {
  background: linear-gradient(90deg, #1A5DAB 0%, #0D3F7A 100%) !important;
  color: #FFFFFF !important;
  font-weight: 700 !important;
}

/* 3. data-checked attribute (BaseWeb / older Streamlit) */
[data-testid="stSidebar"] [data-testid="stRadio"] label[data-checked="true"] {
  background: linear-gradient(90deg, #1A5DAB 0%, #0D3F7A 100%) !important;
  color: #FFFFFF !important;
  font-weight: 700 !important;
  border-left: 3px solid #EDD97A !important;
}

/* 4. BaseWeb checked state on the radio wrapper */
[data-testid="stSidebar"] [data-testid="stRadio"] [data-baseweb="radio"][aria-checked="true"] label,
[data-testid="stSidebar"] [data-testid="stRadio"] [data-checked="true"] label {
  background: linear-gradient(90deg, #1A5DAB 0%, #0D3F7A 100%) !important;
  color: #FFFFFF !important;
  font-weight: 700 !important;
  border-left: 3px solid #EDD97A !important;
}

/* ── collapse button in sidebar ── */
[data-testid="stSidebar"] [data-testid="stButton"] button {
  background: #ECF0F7 !important;
  border: 1px solid #D0D9E8 !important;
  color: #3A4D6B !important;
  border-radius: 7px !important;
  font-size: 11px !important;
  font-weight: 600 !important;
  letter-spacing: 0.3px !important;
  padding: 7px 14px !important;
  transition: background 0.13s, color 0.13s !important;
}
[data-testid="stSidebar"] [data-testid="stButton"] button:hover {
  background: #DAE2EF !important;
  color: #0D3F7A !important;
}

/* ── HEADER ── */
.nw-header {
  background: var(--nw-blue);
  margin: 0 calc(-1 * var(--content-pad));
  padding: 0 var(--content-pad);
  display: flex; align-items: center; justify-content: space-between;
  height: 64px;
  box-shadow: 0 2px 16px rgba(13,63,122,0.22);
  position: relative; z-index: 100;
}
.nw-header-left { display:flex; align-items:center; gap:14px; }
.nw-eagle {
  width:32px; height:32px;
  background:rgba(255,255,255,0.15);
  border:1px solid rgba(255,255,255,0.22);
  border-radius:7px;
  display:flex; align-items:center; justify-content:center; font-size:17px;
}
.nw-brand      { font-family:'Libre Baskerville',serif; font-size:16px; color:#fff; }
.nw-brand span { color:var(--gold-lt); }
.nw-sep        { width:1px; height:18px; background:rgba(255,255,255,0.18); }
.nw-pgname     { font-size:12px; font-weight:500; color:rgba(255,255,255,0.68); }
.nw-right      { font-size:10px; color:rgba(255,255,255,0.32); letter-spacing:1.2px; text-transform:uppercase; }
.gold-line {
  height:3px;
  background:linear-gradient(90deg,var(--gold) 0%,var(--gold-lt) 45%,transparent 100%);
  margin:0 calc(-1 * var(--content-pad)) 32px;
}

/* ── section label ── */
.sec-label {
  font-size:10px; font-weight:700; letter-spacing:2.2px; text-transform:uppercase;
  color:var(--nw-blue); display:flex; align-items:center; gap:10px; margin:0 0 16px;
}
.sec-label::after { content:''; flex:1; height:1px; background:var(--border); }

/* ── field helpers ── */
.f-label { font-size:10px; font-weight:700; letter-spacing:1.8px; text-transform:uppercase; color:var(--nw-blue); margin:0 0 8px; }
.f-hint  { font-size:10px; color:var(--muted); margin:5px 0 0; }
.f-ok    { font-size:10px; color:var(--ok-fg); margin:5px 0 0; }

/* ── chips ── */
.chip-ok   { display:inline-flex; align-items:center; gap:4px; background:var(--ok-bg); border:1px solid #9ECDB0; border-radius:5px; padding:2px 8px; font-size:10px; color:var(--ok-fg); font-weight:600; max-width:100%; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.chip-none { display:inline-flex; align-items:center; gap:4px; border:1px dashed var(--border); border-radius:5px; padding:2px 8px; font-size:10px; color:var(--muted); }

/* ── file uploader (multi-file drop zone) ── */
[data-testid="stFileUploader"] { background:var(--surface) !important; border:2px dashed var(--border) !important; border-radius:10px !important; transition:border-color 0.18s !important; }
[data-testid="stFileUploader"]:hover { border-color:var(--nw-blue) !important; }
[data-testid="stFileUploader"] section { padding:20px 16px !important; }
[data-testid="stFileUploaderDropzoneInstructions"] > div > span { font-size:12px !important; color:var(--muted) !important; }
[data-testid="stFileUploaderDropzoneInstructions"] > div > small { font-size:10px !important; color:var(--border) !important; }
[data-testid="stFileUploaderDropzone"] { padding:0 !important; }
[data-testid="stFileUploaderDropzone"] > div { flex-direction:column !important; align-items:center !important; gap:8px !important; }
[data-testid="stFileUploaderDropzone"] button { font-size:11px !important; font-weight:600 !important; padding:7px 18px !important; border-radius:6px !important; min-height:unset !important; background:var(--nw-blue) !important; color:#fff !important; border:none !important; }
[data-testid="stFileUploaderDropzone"] button:hover { background:var(--nw-deep) !important; }
[data-testid="stFileUploadedFile"] { padding:3px 8px !important; font-size:10px !important; }

/* ── assignment table ── */
.assign-wrap { border:1px solid var(--border); border-radius:var(--radius); background:var(--surface); margin-top:14px; overflow:hidden; box-shadow:var(--shadow); }
.assign-hdr  { display:flex; justify-content:space-between; align-items:center; padding:9px 16px; background:var(--off); border-bottom:1px solid var(--border); font-size:9px; font-weight:700; letter-spacing:2px; text-transform:uppercase; color:var(--muted); }
.arow        { display:grid; grid-template-columns:120px 1fr 28px; align-items:center; padding:8px 16px; border-bottom:1px solid var(--border); gap:12px; }
.arow:last-child { border-bottom:none; }
.arow-ok    { background:var(--surface); }
.arow-empty { background:var(--surface); }
.arow-error { background:#FFF5F5; }
.aco        { font-size:11px; font-weight:700; color:var(--text); display:flex; align-items:center; gap:5px; flex-wrap:wrap; }
.afile      { font-size:11px; color:var(--muted); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.afile-assigned { color:var(--text) !important; }
.afile-err  { color:#C8102E !important; font-size:10px; }
.astat      { font-size:12px; font-weight:700; text-align:right; }
.astat-ok   { color:var(--ok-fg); }
.astat-empty{ color:var(--border); font-weight:400; }
.astat-err  { color:#C8102E; }
.ab-req     { display:inline-block; background:#FEE8E8; color:#C8102E; border:1px solid #F5C0C0; border-radius:3px; font-size:8px; font-weight:700; letter-spacing:0.5px; text-transform:uppercase; padding:1px 5px; line-height:1.5; }
.ab-opt     { display:inline-block; background:var(--nw-lt); color:var(--nw-blue); border:1px solid var(--border); border-radius:3px; font-size:8px; font-weight:600; letter-spacing:0.5px; text-transform:uppercase; padding:1px 5px; line-height:1.5; }

/* ── widget labels ── */
label[data-testid="stWidgetLabel"] p { font-size:10px !important; font-weight:700 !important; letter-spacing:1.5px !important; text-transform:uppercase !important; color:var(--nw-blue) !important; margin-bottom:4px !important; }

/* ── inputs ── */
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input { border:1.5px solid var(--border) !important; border-radius:7px !important; background:var(--surface) !important; font-family:'Inter',sans-serif !important; font-size:13px !important; color:var(--text) !important; padding:9px 12px !important; }
[data-testid="stTextInput"] input:focus,
[data-testid="stNumberInput"] input:focus { border-color:var(--nw-blue) !important; box-shadow:0 0 0 3px rgba(26,93,171,0.10) !important; outline:none !important; }

/* ── slider ── */
[data-testid="stSlider"] { padding:4px 0 !important; }
[data-baseweb="slider"] div[role="slider"] { background:var(--nw-blue) !important; border:2px solid #fff !important; box-shadow:0 0 0 2px var(--nw-blue) !important; width:18px !important; height:18px !important; }
[data-baseweb="slider"] div[role="progressbar"] { background:var(--nw-blue) !important; }

/* ── expander ── */
[data-testid="stExpander"] { border:1px solid var(--border) !important; border-radius:var(--radius) !important; background:var(--surface) !important; box-shadow:var(--shadow) !important; margin-bottom:0 !important; }
details summary { font-family:'Inter',sans-serif !important; font-size:10px !important; font-weight:700 !important; letter-spacing:2px !important; text-transform:uppercase !important; color:var(--nw-blue) !important; padding:13px 20px !important; }
details[open] summary { border-bottom:1px solid var(--border); }

/* ── readiness card ── */
.rdy-card  { border:1px solid var(--border); border-radius:var(--radius); background:var(--surface); padding:8px 16px 4px; margin-bottom:22px; }
.rdy-row   { display:flex; align-items:flex-start; gap:11px; padding:11px 0; border-bottom:1px solid var(--border); }
.rdy-row:last-child { border-bottom:none; padding-bottom:6px; }
.rdy-dot   { width:22px; height:22px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:11px; font-weight:700; flex-shrink:0; margin-top:1px; }
.dot-ok    { background:var(--ok-bg); color:var(--ok-fg); }
.dot-wait  { background:var(--nw-lt); color:var(--nw-blue); }
.rdy-title { font-size:12px; font-weight:600; color:var(--text); line-height:1.4; }
.rdy-sub   { font-size:10px; color:var(--muted); margin-top:2px; }

/* ── run button ── */
div.btn-ready > div > button { background:linear-gradient(135deg,var(--nw-deep) 0%,var(--nw-blue) 100%) !important; color:#fff !important; border:none !important; border-radius:9px !important; font-weight:700 !important; font-size:13px !important; letter-spacing:0.5px !important; padding:13px 28px !important; width:100% !important; box-shadow:0 4px 18px rgba(13,63,122,0.28) !important; transition:all 0.18s !important; }
div.btn-ready > div > button:hover { background:linear-gradient(135deg,#0a2f5e 0%,var(--nw-deep) 100%) !important; box-shadow:0 7px 26px rgba(13,63,122,0.36) !important; transform:translateY(-1px) !important; }
div.btn-wait  > div > button { background:var(--border) !important; color:var(--muted) !important; border:none !important; border-radius:9px !important; font-weight:600 !important; font-size:13px !important; padding:13px 28px !important; width:100% !important; box-shadow:none !important; }

/* ── secondary button ── */
.stButton > button[kind="secondary"] { background:transparent !important; color:var(--nw-blue) !important; border:1.5px solid var(--border) !important; border-radius:7px !important; font-size:11px !important; font-weight:500 !important; }
.stButton > button[kind="secondary"]:hover { border-color:var(--nw-blue) !important; background:var(--nw-lt) !important; }

[data-testid="stAlert"] { border-radius:8px !important; font-size:12px !important; }
[data-testid="column"]  { padding:0 6px !important; }

/* ── inline warning box ── */
.warn-box {
  border:1.5px solid #EDD97A; border-radius:10px; background:#FFFDF5;
  padding:16px 20px; margin-top:12px;
  box-shadow:0 2px 10px rgba(200,169,81,0.10);
}
.warn-box .wb-head { display:flex; align-items:center; gap:8px; margin-bottom:8px; }
.warn-box .wb-icon { font-size:18px; }
.warn-box .wb-title { font-size:13px; font-weight:700; color:#0C1A35; }
.warn-box .wb-body { font-size:11px; color:#6B7A9E; line-height:1.6; margin:0 0 4px; }

/* ── inline spinner ── */
.inline-loader { display:flex; align-items:center; gap:14px; padding:18px 20px; margin-top:12px;
  border:1.5px solid var(--border); border-radius:10px; background:var(--surface); }
.spin-ring {
  width:28px; height:28px; border:3px solid rgba(26,93,171,0.15);
  border-top:3px solid #1A5DAB; border-radius:50%; flex-shrink:0;
  animation: spin 0.8s linear infinite;
}
@keyframes spin { to { transform: rotate(360deg); } }
.loader-label { font-family:'Inter',sans-serif; font-size:12px; font-weight:600; color:var(--nw-blue); }
.loader-sub   { font-family:'Inter',sans-serif; font-size:10px; color:var(--muted); margin-top:2px; }

/* ── coming soon ── */
.coming-soon { display:flex; flex-direction:column; align-items:center; justify-content:center; padding:80px 32px; background:var(--surface); border:1px solid var(--border); border-radius:var(--radius); text-align:center; margin-top:24px; }
.coming-soon .cs-icon  { font-size:52px; margin-bottom:16px; }
.coming-soon .cs-title { font-family:'Libre Baskerville',serif; font-size:22px; font-weight:700; color:var(--nw-deep); margin-bottom:8px; }
.coming-soon .cs-sub   { font-size:13px; color:var(--muted); max-width:380px; line-height:1.6; }
.coming-soon .cs-tag   { margin-top:20px; display:inline-block; background:var(--nw-lt); color:var(--nw-blue); border:1px solid var(--border); border-radius:20px; font-size:10px; font-weight:700; letter-spacing:1.5px; text-transform:uppercase; padding:5px 14px; }

/* ── inline radio (sched mode toggle) ── */
[data-testid="stRadio"][aria-label="sched_mode_pick"] > div,
div:has(> [data-testid="stRadio"] > div > [data-baseweb="radio"]:first-child) {
  gap:0 !important;
}
/* Per-state schedule rating grid container */
.sched-state-grid { border:1px solid var(--border); border-top:none; border-radius:0 0 8px 8px; overflow:hidden; }

/* ── multi-state progress card ── */
.multi-progress-card { display:flex; align-items:center; gap:14px; padding:14px 18px; border:1.5px solid var(--border); border-radius:10px; background:var(--surface); margin-top:10px; }
.mp-label { font-size:9px; font-weight:700; letter-spacing:1.8px; text-transform:uppercase; color:var(--muted); flex-shrink:0; }
.mp-state { font-size:15px; font-weight:700; color:var(--nw-deep); flex:1; }
.mp-count { font-size:11px; color:var(--muted); flex-shrink:0; }
</style>
""", unsafe_allow_html=True)


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:

    st.image("Nationwide-logo.png",  width=200)

    # # Logo
    st.markdown("""
    <div style="padding:16px 20px 6px; font-size:8px; font-weight:700;
                letter-spacing:2.8px; text-transform:uppercase; color:#9BAABF;">
      Line of Business
    </div>
    """, unsafe_allow_html=True)

    # ── LOB navigation ──
    # st.radio is the correct widget: it has persistent checked state that the
    # browser maintains natively via aria-checked on the label element.
    # CSS selector  label[aria-checked="true"]  always works — no class names,
    # no JS, no Streamlit internals.  This is the only reliable approach.
    current_idx = LOB_NAMES.index(st.session_state.lob) if st.session_state.lob in LOB_NAMES else 0

    chosen = st.radio(
        "lob_selector",
        options=LOB_OPTIONS,
        index=current_idx,
        key="lob_radio",
        label_visibility="collapsed",
    )

    # Derive the plain LOB name from the chosen option label
    chosen_lob = LOB_NAMES[LOB_OPTIONS.index(chosen)]
    if chosen_lob != st.session_state.lob:
        st.session_state.lob        = chosen_lob
        st.session_state.run_status = "idle"
        st.rerun()

    # JS: style active label AND fix all parent wrapper widths/margins
    _active = st.session_state.lob
    st.markdown(f"""
    <script>
    (function highlight() {{
      var active = {repr(_active)};
      var sidebar = window.parent.document.querySelector('[data-testid="stSidebar"]');
      if (!sidebar) return;

      // Zero out every wrapper div inside the radio group so nothing clips the label
      var radio = sidebar.querySelector('[data-testid="stRadio"]');
      if (radio) {{
        radio.querySelectorAll('div').forEach(function(d) {{
          d.style.setProperty('padding', '0', 'important');
          d.style.setProperty('margin', '0', 'important');
          d.style.setProperty('width', '100%', 'important');
          d.style.setProperty('max-width', '100%', 'important');
          d.style.setProperty('box-sizing', 'border-box', 'important');
        }});
      }}

      var labels = sidebar.querySelectorAll('[data-testid="stRadio"] label');
      labels.forEach(function(lbl) {{
        // Also fix the label's own parent chain
        var el = lbl.parentElement;
        while (el && el !== radio) {{
          el.style.setProperty('padding', '0', 'important');
          el.style.setProperty('margin', '0', 'important');
          el.style.setProperty('width', '100%', 'important');
          el = el.parentElement;
        }}

        var txt = lbl.innerText || lbl.textContent || '';
        if (txt.indexOf(active) !== -1) {{
          lbl.style.setProperty('background', 'linear-gradient(90deg,#1A5DAB 0%,#0D3F7A 100%)', 'important');
          lbl.style.setProperty('color', '#FFFFFF', 'important');
          lbl.style.setProperty('font-weight', '700', 'important');
          lbl.style.setProperty('border-left', '3px solid #EDD97A', 'important');
          lbl.style.setProperty('width', '100%', 'important');
          lbl.style.setProperty('min-width', '100%', 'important');
          lbl.style.setProperty('margin', '0', 'important');
          lbl.style.setProperty('border-radius', '0', 'important');
          lbl.style.setProperty('box-sizing', 'border-box', 'important');
          lbl.style.setProperty('padding', '13px 20px', 'important');
        }} else {{
          lbl.style.removeProperty('background');
          lbl.style.removeProperty('color');
          lbl.style.removeProperty('font-weight');
          lbl.style.removeProperty('border-left');
        }}
      }});
    }})();
    setTimeout(highlight, 150);
    setTimeout(highlight, 400);

    function highlight() {{
      var active = {repr(_active)};
      var sidebar = window.parent.document.querySelector('[data-testid="stSidebar"]');
      if (!sidebar) return;
      var radio = sidebar.querySelector('[data-testid="stRadio"]');
      if (radio) {{
        radio.querySelectorAll('div').forEach(function(d) {{
          d.style.setProperty('padding', '0', 'important');
          d.style.setProperty('margin', '0', 'important');
          d.style.setProperty('width', '100%', 'important');
          d.style.setProperty('max-width', '100%', 'important');
          d.style.setProperty('box-sizing', 'border-box', 'important');
        }});
      }}
      var labels = sidebar ? sidebar.querySelectorAll('[data-testid="stRadio"] label') : [];
      labels.forEach(function(lbl) {{
        var el = lbl.parentElement;
        while (el && el !== radio) {{
          el.style.setProperty('padding', '0', 'important');
          el.style.setProperty('margin', '0', 'important');
          el.style.setProperty('width', '100%', 'important');
          el = el.parentElement;
        }}
        var txt = lbl.innerText || lbl.textContent || '';
        if (txt.indexOf(active) !== -1) {{
          lbl.style.setProperty('background', 'linear-gradient(90deg,#1A5DAB 0%,#0D3F7A 100%)', 'important');
          lbl.style.setProperty('color', '#FFFFFF', 'important');
          lbl.style.setProperty('font-weight', '700', 'important');
          lbl.style.setProperty('border-left', '3px solid #EDD97A', 'important');
          lbl.style.setProperty('width', '100%', 'important');
          lbl.style.setProperty('min-width', '100%', 'important');
          lbl.style.setProperty('margin', '0', 'important');
          lbl.style.setProperty('border-radius', '0', 'important');
          lbl.style.setProperty('box-sizing', 'border-box', 'important');
          lbl.style.setProperty('padding', '13px 20px', 'important');
        }} else {{
          lbl.style.removeProperty('background');
          lbl.style.removeProperty('color');
          lbl.style.removeProperty('font-weight');
          lbl.style.removeProperty('border-left');
        }}
      }});
    }}
    </script>
    """, unsafe_allow_html=True)


# ─── HEADER ───────────────────────────────────────────────────────────────────
active_lob = st.session_state.lob

LOB_ICONS = {"Business Auto":"🚗","General Liability":"⚖️","Farm Auto":"🚜","Property":"🏠"}
LOB_SUBS  = {
    "Business Auto":     "Upload proposed ratebooks &nbsp;&middot;&nbsp; Configure options &nbsp;&middot;&nbsp; Generate output",
    "General Liability": "General Liability rate page configuration",
    "Farm Auto":         "Farm Auto rate page configuration",
    "Property":          "Property rate page configuration",
}

st.markdown(f"""
<div class="nw-header">
  <div class="nw-header-left">
    <div class="nw-eagle">📋</div>
    <div class="nw-brand">Nationwide <span>Insurance</span></div>
    <div class="nw-sep"></div>
    <div class="nw-pgname">{LOB_ICONS[active_lob]} &nbsp;{active_lob} &middot; Rate Page Builder</div>
  </div>
  <div class="nw-right">BA &nbsp;&middot;&nbsp; Analytics &nbsp;&middot;&nbsp; Internal Tool</div>
</div>
<div class="gold-line"></div>
""", unsafe_allow_html=True)

# Page heading
h1, h2 = st.columns([3, 1])
with h1:
    st.markdown(f"""
    <p style="font-family:'Libre Baskerville',serif;font-size:25px;font-weight:700;
              color:#0D3F7A;margin:0 0 5px;">Build {active_lob} Rate Pages</p>
    <p style="font-size:13px;color:#6B7A9E;margin:0;">{LOB_SUBS[active_lob]}</p>
    """, unsafe_allow_html=True)
with h2:
    if active_lob == "Business Auto":
        nr = n_req(); tot = len(REQUIRED); pct = int(nr/tot*100)
        fg = "#196B38" if nr == tot else "#1A5DAB"
        st.markdown(f"""
        <div style="text-align:right;padding-top:4px;">
          <div style="font-size:10px;color:#6B7A9E;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:5px;">Uploaded Files</div>
          <div style="font-size:28px;font-weight:700;color:{fg};line-height:1;font-family:'Libre Baskerville',serif;">
            {nr}<span style="font-size:13px;font-weight:400;color:#6B7A9E;">/{tot}</span>
          </div>
          <div style="background:#D4DFEF;border-radius:3px;height:3px;margin-top:8px;overflow:hidden;">
            <div style="width:{pct}%;height:100%;background:linear-gradient(90deg,#0D3F7A,#1A5DAB);border-radius:3px;"></div>
          </div>
        </div>""", unsafe_allow_html=True)

spacer(28)


# ─── BUSINESS AUTO ────────────────────────────────────────────────────────────
if active_lob == "Business Auto":

    DETECT_ORDER = ["HICNJ", "CCMIC", "NICOF", "NWAG", "NACO", "NAFF", "NGIC", "MM", "CW"]

    def scan_states(src_path):
        states = []
        try:
            for item in sorted(Path(src_path).iterdir()):
                if item.is_dir() and not item.name.startswith(('.', '_')):
                    books = {}
                    for f in item.iterdir():
                        if f.is_file() and f.suffix.lower() in ('.xlsx', '.xlsm', '.xls'):
                            name_up = f.name.upper()
                            matched = next((k for k in DETECT_ORDER if k in name_up), None)
                            if matched:
                                books.setdefault(matched, []).append(f)
                    has_ngic  = "NGIC" in books and len(books["NGIC"]) == 1
                    conflicts = [k for k, v in books.items() if len(v) > 1]
                    states.append({"name": item.name, "path": item, "books": books,
                                   "ready": has_ngic and not conflicts,
                                   "conflicts": conflicts, "has_ngic": has_ngic})
        except Exception:
            pass
        return states

    # ── Mode toggle ────────────────────────────────────────────────────────────
    mode = st.session_state.mode
    tc1, tc2, _ = st.columns([2, 2, 8])
    with tc1:
        if st.button("Individual State", key="btn_ind", use_container_width=True,
                     type="primary" if mode == "individual" else "secondary"):
            if mode != "individual":
                st.session_state.mode = "individual"
                st.session_state.confirm_step = "idle"
                st.rerun()
    with tc2:
        if st.button("Multiple States", key="btn_mul", use_container_width=True,
                     type="primary" if mode == "multiple" else "secondary"):
            if mode != "multiple":
                st.session_state.mode = "multiple"
                st.session_state.multi_step = "idle"
                st.rerun()

    spacer(16)

    # ══════════════════════════════════════════════════════════════════════════
    # INDIVIDUAL STATE MODE
    # ══════════════════════════════════════════════════════════════════════════
    if mode == "individual":
        L, R = st.columns([13, 7], gap="large")

        with L:
            st.markdown('<div class="sec-label">&#128194; &nbsp;Proposed Ratebooks</div>', unsafe_allow_html=True)

            uploaded = st.file_uploader(
            "Select all ratebook files at once — filenames must contain the company code (NGIC, MM, NACO, …)",
            type=["xlsx", "xlsm", "xls"],
            accept_multiple_files=True,
            key=f"multi_up_{st.session_state.upload_reset}",
        )

            # ── Auto-detect & assign ───────────────────────────────────────
            if uploaded:
                grouped = {}
                for f in uploaded:
                    name_up = f.name.upper()
                    matched = next((k for k in DETECT_ORDER if k in name_up), None)
                    grouped.setdefault(matched or "__unknown__", []).append(f)
                for key in ALL_KEYS:
                    files = grouped.get(key, [])
                    if len(files) == 1:
                        st.session_state[f"file_{key}"] = {"name": files[0].name, "bytes": files[0].read()}
                    elif len(files) > 1:
                        st.session_state[f"file_{key}"] = {"error": "multiple", "names": [f.name for f in files]}

            # ── Assignment table ───────────────────────────────────────────
            LABELS = {"NGIC": "Required", "CW": "Optional"}
            rows_html = ""; n_ok = n_err = 0
            for key in ALL_KEYS:
                val = st.session_state.get(f"file_{key}")
                bh = '<span class="ab-req">Required</span>' if LABELS.get(key) == "Required" else ('<span class="ab-opt">Optional</span>' if LABELS.get(key) == "Optional" else "")
                if val is None:
                    rows_html += f'<div class="arow arow-empty"><span class="aco">{key} {bh}</span><span class="afile">Not uploaded</span><span class="astat astat-empty">—</span></div>'
                elif "error" in val:
                    n_err += 1
                    rows_html += f'<div class="arow arow-error"><span class="aco">{key} {bh}</span><span class="afile afile-err">&#9888;&nbsp; Multiple files: {", ".join(val["names"])}</span><span class="astat astat-err">&#10005;</span></div>'
                else:
                    n_ok += 1
                    rows_html += f'<div class="arow arow-ok"><span class="aco">{key} {bh}</span><span class="afile afile-assigned">{val["name"]}</span><span class="astat astat-ok">&#10003;</span></div>'
            summary = f'{n_ok} assigned' + (f' &nbsp;&middot;&nbsp; <span style="color:#C8102E;font-weight:700;">{n_err} conflict{"s" if n_err>1 else ""}</span>' if n_err else '')
            st.markdown(f'<div class="assign-wrap"><div class="assign-hdr"><span>File Assignment</span><span>{summary}</span></div>{rows_html}</div>', unsafe_allow_html=True)

            if any(st.session_state[f"file_{k}"] for k in ALL_KEYS):
                spacer(8)
                _, clr = st.columns([5, 1])
                with clr:
                    if st.button("Clear all", type="secondary"):
                        for k in ALL_KEYS: st.session_state[f"file_{k}"] = None
                        st.session_state.upload_reset += 1
                        st.session_state.run_status = "idle"
                        st.rerun()

        with R:
            st.markdown('<div class="sec-label">&#9881; &nbsp;Configuration</div>', unsafe_allow_html=True)
            st.markdown('<p class="f-label">&#128193; &nbsp;Save Location</p>', unsafe_allow_html=True)
            typed = st.text_input("save_path", value=st.session_state.save_dir, placeholder="Paste path or click Browse", label_visibility="collapsed")
            if typed != st.session_state.save_dir: st.session_state.save_dir = typed
            if st.button("Browse", key="browse_btn"):
                folder = browse_folder()
                if folder: st.session_state.save_dir = folder; st.rerun()
            if st.session_state.save_dir:
                p = st.session_state.save_dir
                st.markdown(f'<p class="f-ok">&#10003; &nbsp;{("…"+p[-38:]) if len(p)>40 else p}</p>', unsafe_allow_html=True)
            else:
                st.markdown('<p class="f-hint">Browse your device or paste the full folder path</p>', unsafe_allow_html=True)

            spacer(6)
            st.markdown('<p class="f-label">&#128202; &nbsp;Schedule Rating Mod</p>', unsafe_allow_html=True)
            nc, pc = st.columns([3, 1])
            with nc:
                tm = st.number_input("mod_num", min_value=0, max_value=100, value=st.session_state.sched_mod, step=1, label_visibility="collapsed")
                if tm != st.session_state.sched_mod: st.session_state.sched_mod = int(tm)
            with pc:
                st.markdown(f'<div style="display:flex;align-items:center;height:42px;padding-left:4px;"><span style="font-size:22px;font-weight:700;color:#1A5DAB;line-height:1;">{st.session_state.sched_mod}<span style="font-size:13px;font-weight:400;color:#6B7A9E;">%</span></span></div>', unsafe_allow_html=True)
            spacer(6)
            sv = st.slider("mod_slider", 0, 100, value=st.session_state.sched_mod, step=1, format="%d%%", label_visibility="collapsed")
            if sv != st.session_state.sched_mod: st.session_state.sched_mod = sv; st.rerun()
            st.markdown('<p class="f-hint">Rule 417 &middot; State Schedule Rating Maximum Modification Threshold</p>', unsafe_allow_html=True)

            spacer(6)
            st.markdown('<div class="sec-label">&#128203; &nbsp;Readiness</div>', unsafe_allow_html=True)
            has_files = any_req(); save_ok = bool(st.session_state.save_dir)
            nr_now = n_req(); sdv = st.session_state.save_dir; mv = st.session_state.sched_mod
            req_sub  = f"All {len(REQUIRED)} ratebooks uploaded" if all_req() else f"{nr_now} of {len(REQUIRED)} ratebooks uploaded"
            save_sub = (("…"+sdv[-36:]) if len(sdv)>38 else sdv) if save_ok else "Not yet selected"

            def rdy(ok, title, sub):
                d = "dot-ok" if ok else "dot-wait"; i = "&#10003;" if ok else "&#9675;"
                return f'<div class="rdy-row"><div class="rdy-dot {d}">{i}</div><div><div class="rdy-title">{title}</div><div class="rdy-sub">{sub}</div></div></div>'

            ngic_uploaded = has_ngic(); ngic_sub = "Uploaded" if ngic_uploaded else "Required — please upload NGIC"
            st.markdown('<div class="rdy-card">'
                + rdy(ngic_uploaded, 'NGIC Ratebook <span style="font-size:10px;color:#C8102E;font-weight:600;">REQUIRED</span>', ngic_sub)
                + rdy(has_files, f'Other Ratebooks &nbsp;<span style="font-size:10px;color:#6B7A9E;font-weight:400;">{nr_now}/{len(REQUIRED)}</span>', req_sub)
                + rdy(save_ok, "Save location", save_sub)
                + rdy(True, f'Schedule Mod &nbsp;<span style="font-size:10px;color:#6B7A9E;font-weight:400;">{mv}%</span>', "Rule 417 threshold")
                + '</div>', unsafe_allow_html=True)

            ngic_ok = has_ngic(); ready = ngic_ok and save_ok

            if st.session_state.confirm_step == "idle":
                if ready:
                    st.markdown('<div class="btn-ready">', unsafe_allow_html=True)
                    if st.button("&#129413;  Create Rate Pages", key="run_btn", use_container_width=True):
                        st.session_state.confirm_step = "confirm"; st.session_state.run_status = "idle"; st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)
                else:
                    missing = (["NGIC ratebook"] if not ngic_ok else []) + (["save location"] if not save_ok else [])
                    st.markdown('<div class="btn-wait">', unsafe_allow_html=True)
                    st.button(f"Waiting \u2014 {', '.join(missing)}", key="run_btn_dis", use_container_width=True, disabled=True)
                    st.markdown('</div>', unsafe_allow_html=True)

            elif st.session_state.confirm_step == "confirm":
                st.markdown('<div class="btn-ready">', unsafe_allow_html=True)
                st.button("&#129413;  Create Rate Pages", key="run_btn_cfm", use_container_width=True, disabled=True)
                st.markdown('</div>', unsafe_allow_html=True)
                st.markdown('<div class="warn-box"><div class="wb-head"><span class="wb-icon">⚠️</span><span class="wb-title">Close &amp; save all open Excel files</span></div><p class="wb-body">The builder needs exclusive access to the workbooks. Please save and close any open <code>.xlsx</code> / <code>.xlsm</code> files before proceeding.</p></div>', unsafe_allow_html=True)
                spacer(8)
                bc1, bc2 = st.columns(2)
                with bc1:
                    if st.button("Cancel", key="cancel_btn", use_container_width=True, type="secondary"):
                        st.session_state.confirm_step = "idle"; st.rerun()
                with bc2:
                    if st.button("Proceed", key="proceed_btn", use_container_width=True, type="primary"):
                        st.session_state.confirm_step = "processing"; st.rerun()

            elif st.session_state.confirm_step == "processing":
                st.markdown('<div class="btn-wait">', unsafe_allow_html=True)
                st.button("Processing Excel...", key="run_btn_proc", use_container_width=True, disabled=True)
                st.markdown('</div>', unsafe_allow_html=True)
                loader_ph = st.empty()
                def update_progress(msg):
                    loader_ph.markdown(f'<div class="inline-loader"><div class="spin-ring"></div><div><div class="loader-label">Creating Excel rate pages…</div><div class="loader-sub">{msg}</div></div></div>', unsafe_allow_html=True)
                update_progress("Please wait while the workbooks are processed.")
                from BARatePages import run as run_rate_pages
                try:
                    def _rb(k):
                        f = st.session_state.get(f"file_{k}")
                        return io.BytesIO(f["bytes"]) if f and "error" not in f else None
                    xlsx_out, pdf_out = run_rate_pages(
                        NGICRatebook=_rb("NGIC"), MMRatebook=_rb("MM"), NACORatebook=_rb("NACO"),
                        NAFFRatebook=_rb("NAFF"), NICOFRatebook=_rb("NICOF"), HICNJRatebook=_rb("HICNJ"),
                        CCMICRatebook=_rb("CCMIC"), NWAGRatebook=_rb("NWAG"), CWRatebook=_rb("CW"),
                        folder_selected=st.session_state.save_dir,
                        SchedRatingMod=int(st.session_state.sched_mod) or None,
                        progress_callback=update_progress, skip_pdf=True)
                    st.session_state.xlsx_path = xlsx_out; st.session_state.pdf_path = pdf_out
                    st.session_state.run_status = "success"; st.session_state.pdf_status = "idle"
                except Exception as e:
                    import traceback; traceback.print_exc()
                    st.session_state.run_status = "error"; st.session_state.run_msg = str(e)
                st.session_state.confirm_step = "idle"; st.rerun()

            elif st.session_state.confirm_step == "pdf_processing":
                st.markdown('<div class="btn-wait">', unsafe_allow_html=True)
                st.button("Generating PDF...", key="pdf_btn_proc", use_container_width=True, disabled=True)
                st.markdown('</div>', unsafe_allow_html=True)
                loader_ph2 = st.empty()
                def update_pdf_progress(msg):
                    loader_ph2.markdown(f'<div class="inline-loader"><div class="spin-ring"></div><div><div class="loader-label">Converting to PDF…</div><div class="loader-sub">{msg}</div></div></div>', unsafe_allow_html=True)
                from BARatePages import generate_pdf_only
                try:
                    generate_pdf_only(st.session_state.xlsx_path, st.session_state.pdf_path, progress_callback=update_pdf_progress)
                    st.session_state.pdf_status = "success"
                except Exception as e:
                    import traceback; traceback.print_exc()
                    st.session_state.pdf_status = "error"; st.session_state.run_msg = str(e)
                st.session_state.confirm_step = "idle"; st.rerun()

            if st.session_state.run_status == "success":
                spacer(10)
                st.success(f"&#10003;  Excel created: {Path(st.session_state.xlsx_path).name}")
                if st.session_state.pdf_status != "success":
                    st.markdown('<div class="btn-ready">', unsafe_allow_html=True)
                    if st.button("Generate PDF Document", key="gen_pdf_btn", use_container_width=True):
                        st.session_state.confirm_step = "pdf_processing"; st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)
                    if st.session_state.pdf_status == "error":
                        st.error(f"PDF Error: {st.session_state.run_msg}")
                else:
                    st.success(f"&#10003;  PDF created: {Path(st.session_state.pdf_path).name}")
            elif st.session_state.run_status == "error":
                spacer(10); st.error(st.session_state.run_msg)

            spacer(24)
            st.markdown('<div style="padding-top:14px;border-top:1px solid var(--border);"><p style="font-size:10px;color:#8892A4;letter-spacing:0.8px;text-transform:uppercase;text-align:center;margin:0;line-height:1.9;">Nationwide Insurance &nbsp;&middot;&nbsp; BA Analytics Division<br>Internal Use Only</p></div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════════
    # MULTIPLE STATES MODE
    # ══════════════════════════════════════════════════════════════════════════
    else:
        L, R = st.columns([13, 7], gap="large")

        with L:
            st.markdown('<div class="sec-label">&#128194; &nbsp;Source Folder</div>', unsafe_allow_html=True)
            st.markdown('<p class="f-label">&#128193; &nbsp;Main Ratebooks Folder</p>', unsafe_allow_html=True)
            st.markdown('<p class="f-hint">Select the folder that contains one subfolder per state. Each subfolder must contain that state\'s ratebook files.</p>', unsafe_allow_html=True)
            spacer(4)
            typed_src = st.text_input("multi_src_path", value=st.session_state.multi_src_dir, placeholder="Paste path or click Browse", label_visibility="collapsed", key="multi_src_input")
            if typed_src != st.session_state.multi_src_dir:
                st.session_state.multi_src_dir = typed_src; st.session_state.multi_step = "idle"; st.session_state.multi_results = []; st.session_state.multi_sched_map = {}
            if st.button("Browse", key="browse_src_btn"):
                folder = browse_folder()
                if folder: st.session_state.multi_src_dir = folder; st.session_state.multi_step = "idle"; st.session_state.multi_results = []; st.session_state.multi_sched_map = {}; st.rerun()

            if st.session_state.multi_src_dir:
                src_path = Path(st.session_state.multi_src_dir)
                if src_path.exists() and src_path.is_dir():
                    p = str(src_path)
                    st.markdown(f'<p class="f-ok">&#10003; &nbsp;{("…"+p[-38:]) if len(p)>40 else p}</p>', unsafe_allow_html=True)
                    spacer(14)
                    states = scan_states(src_path)
                    if not states:
                        st.markdown('<p class="f-hint">No state subfolders found in this directory.</p>', unsafe_allow_html=True)
                    else:
                        n_ready  = sum(1 for s in states if s["ready"])
                        n_issues = len(states) - n_ready
                        issue_txt = f' &nbsp;&middot;&nbsp; <span style="color:#C8102E;font-weight:700;">{n_issues} issue{"s" if n_issues>1 else ""}</span>' if n_issues else ''
                        state_rows = ""
                        for s in states:
                            found_str = ", ".join(sorted(s["books"].keys())) or "—"
                            if s["ready"]:
                                state_rows += f'<div class="arow arow-ok" style="grid-template-columns:110px 1fr 28px;"><span class="aco">{s["name"]}</span><span class="afile afile-assigned">{found_str}</span><span class="astat astat-ok">&#10003;</span></div>'
                            elif s["conflicts"]:
                                state_rows += f'<div class="arow arow-error" style="grid-template-columns:110px 1fr 28px;"><span class="aco">{s["name"]}</span><span class="afile afile-err">&#9888;&nbsp; Conflicts: {", ".join(s["conflicts"])}</span><span class="astat astat-err">&#9888;</span></div>'
                            elif not s["has_ngic"]:
                                state_rows += f'<div class="arow arow-error" style="grid-template-columns:110px 1fr 28px;"><span class="aco">{s["name"]}</span><span class="afile afile-err">Missing NGIC &nbsp;({found_str})</span><span class="astat astat-err">&#10005;</span></div>'
                            else:
                                state_rows += f'<div class="arow arow-empty" style="grid-template-columns:110px 1fr 28px;"><span class="aco">{s["name"]}</span><span class="afile">{found_str}</span><span class="astat astat-empty">—</span></div>'
                        st.markdown(f'<div class="assign-wrap"><div class="assign-hdr"><span>States Detected &nbsp;({len(states)})</span><span>{n_ready} ready{issue_txt}</span></div>{state_rows}</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<p style="font-size:11px;color:#C8102E;margin:4px 0;">Path not found or not a directory.</p>', unsafe_allow_html=True)

        with R:
            st.markdown('<div class="sec-label">&#9881; &nbsp;Configuration</div>', unsafe_allow_html=True)
            st.markdown('<p class="f-label">&#128193; &nbsp;Save Location</p>', unsafe_allow_html=True)
            st.markdown('<p class="f-hint">Each state\'s Excel will be saved in <code>Save Location / State /</code></p>', unsafe_allow_html=True)
            spacer(4)
            typed_save = st.text_input("multi_save_path", value=st.session_state.multi_save_dir, placeholder="Paste path or click Browse", label_visibility="collapsed", key="multi_save_input")
            if typed_save != st.session_state.multi_save_dir: st.session_state.multi_save_dir = typed_save
            if st.button("Browse", key="browse_save_btn"):
                folder = browse_folder()
                if folder: st.session_state.multi_save_dir = folder; st.rerun()
            if st.session_state.multi_save_dir:
                p = st.session_state.multi_save_dir
                st.markdown(f'<p class="f-ok">&#10003; &nbsp;{("…"+p[-38:]) if len(p)>40 else p}</p>', unsafe_allow_html=True)
            else:
                st.markdown('<p class="f-hint">Choose where all output files will be saved.</p>', unsafe_allow_html=True)

            spacer(6)
            st.markdown('<p class="f-label">&#128202; &nbsp;Schedule Rating Mod &mdash; Per State</p>', unsafe_allow_html=True)
            st.markdown('<p class="f-hint">Rule 417 &middot; Each state can have its own threshold</p>', unsafe_allow_html=True)
            spacer(4)

            # ── Mode toggle ─────────────────────────────────────────────────
            _sched_opts = ["Upload Excel", "Enter Manually"]
            _sched_idx  = 0 if st.session_state.multi_sched_mode == "upload" else 1
            _sched_pick = st.radio(
                "sched_mode_pick",
                options=_sched_opts,
                index=_sched_idx,
                horizontal=True,
                label_visibility="collapsed",
                key="multi_sched_mode_radio",
            )
            _new_mode = "upload" if _sched_pick == "Upload Excel" else "manual"
            if _new_mode != st.session_state.multi_sched_mode:
                st.session_state.multi_sched_mode = _new_mode
                st.rerun()

            # ── Detect states early for this section ────────────────────────
            _src_ok_sched = bool(st.session_state.multi_src_dir) and Path(st.session_state.multi_src_dir).is_dir()
            _states_sched = scan_states(st.session_state.multi_src_dir) if _src_ok_sched else []

            # ── Upload Excel branch ─────────────────────────────────────────
            if st.session_state.multi_sched_mode == "upload":
                spacer(4)
                st.markdown(
                    '<p class="f-hint"><b>Col A</b> = State name &nbsp;&middot;&nbsp; '
                    '<b>Col B</b> = Schedule Rating % &nbsp;(no header row required).<br>'
                    'Only rows whose state name matches a detected folder are applied.</p>',
                    unsafe_allow_html=True,
                )
                _sched_file = st.file_uploader(
                    "Upload schedule rating Excel",
                    type=["xlsx", "xls"],
                    key="multi_sched_excel_uploader",
                    label_visibility="collapsed",
                )
                if _sched_file is not None:
                    import pandas as _pd
                    try:
                        _df = _pd.read_excel(io.BytesIO(_sched_file.read()), header=None)
                        _new_excel = {}
                        for _, _row in _df.iterrows():
                            try:
                                _sn = str(_row.iloc[0]).strip()
                                _sv = int(float(str(_row.iloc[1]).strip()))
                                if _sn and _sn.lower() not in ("nan", "none", ""):
                                    _new_excel[_sn] = max(0, min(100, _sv))
                            except Exception:
                                pass
                        if _new_excel != st.session_state.multi_sched_excel:
                            st.session_state.multi_sched_excel = _new_excel
                            # Pre-populate sched_map for matched detected states
                            for _s in _states_sched:
                                _nm = _s["name"]
                                if _nm in _new_excel:
                                    st.session_state.multi_sched_map[_nm] = _new_excel[_nm]
                            st.rerun()
                    except Exception as _exc:
                        st.markdown(f'<p style="font-size:11px;color:#C8102E;margin:4px 0;">&#9888; Could not parse file: {_exc}</p>', unsafe_allow_html=True)

            # ── Per-state input grid ─────────────────────────────────────────
            if _states_sched:
                spacer(6)
                _n_set = sum(
                    1 for _s in _states_sched
                    if st.session_state.multi_sched_map.get(_s["name"], 0) > 0
                )
                # Apply any excel defaults not yet in sched_map
                if st.session_state.multi_sched_excel and st.session_state.multi_sched_mode == "upload":
                    for _s in _states_sched:
                        _nm = _s["name"]
                        if _nm not in st.session_state.multi_sched_map and _nm in st.session_state.multi_sched_excel:
                            st.session_state.multi_sched_map[_nm] = st.session_state.multi_sched_excel[_nm]

                st.markdown(
                    f'<div class="assign-hdr" style="border:1px solid var(--border);border-radius:8px 8px 0 0;'
                    f'padding:9px 14px;margin-bottom:0;">'
                    f'<span>Per-State Schedule Rating</span>'
                    f'<span style="color:{"var(--ok-fg)" if _n_set==len(_states_sched) else "var(--muted)"};">'
                    f'{_n_set}&thinsp;/&thinsp;{len(_states_sched)} configured</span></div>',
                    unsafe_allow_html=True,
                )
                st.markdown('<div style="border:1px solid var(--border);border-top:none;border-radius:0 0 8px 8px;overflow:hidden;">', unsafe_allow_html=True)
                for _i, _s in enumerate(_states_sched):
                    _nm  = _s["name"]
                    _exc = st.session_state.multi_sched_excel.get(_nm)
                    _cur = int(st.session_state.multi_sched_map.get(_nm, _exc if _exc is not None else 0))

                    _bg   = "#F9FBF9" if _i % 2 == 0 else "#FFFFFF"
                    _src_badge = ""
                    if _exc is not None and st.session_state.multi_sched_mode == "upload":
                        _src_badge = ' &nbsp;<span style="font-size:8px;font-weight:700;color:var(--ok-fg);letter-spacing:0.5px;text-transform:uppercase;">excel</span>'

                    _lc, _ic, _pc = st.columns([5, 3, 1])
                    with _lc:
                        st.markdown(
                            f'<div style="display:flex;align-items:center;height:38px;padding:0 8px;'
                            f'font-size:12px;font-weight:600;color:var(--text);background:{_bg};">'
                            f'{_nm}{_src_badge}</div>',
                            unsafe_allow_html=True,
                        )
                    with _ic:
                        _nv = st.number_input(
                            f"sr_{_nm}",
                            min_value=0, max_value=100,
                            value=_cur,
                            step=1,
                            key=f"msm_{_nm}",
                            label_visibility="collapsed",
                        )
                        if int(_nv) != st.session_state.multi_sched_map.get(_nm):
                            st.session_state.multi_sched_map[_nm] = int(_nv)
                    with _pc:
                        st.markdown(
                            f'<div style="display:flex;align-items:center;height:38px;font-size:12px;'
                            f'color:var(--muted);">%</div>',
                            unsafe_allow_html=True,
                        )
                st.markdown('</div>', unsafe_allow_html=True)
            elif not _src_ok_sched:
                st.markdown('<p class="f-hint">Select a source folder above to configure per-state values.</p>', unsafe_allow_html=True)
            else:
                st.markdown('<p class="f-hint">No state folders detected yet.</p>', unsafe_allow_html=True)

            spacer(6)
            gen_pdf = st.checkbox("Generate PDF for each state", value=st.session_state.multi_gen_pdf, key="multi_pdf_chk")
            if gen_pdf != st.session_state.multi_gen_pdf: st.session_state.multi_gen_pdf = gen_pdf
            if gen_pdf:
                st.markdown('<p class="f-hint">PDFs saved to <code>Save Location / PDF /</code></p>', unsafe_allow_html=True)

            spacer(6)
            st.markdown('<div class="sec-label">&#128203; &nbsp;Readiness</div>', unsafe_allow_html=True)
            src_ok      = bool(st.session_state.multi_src_dir) and Path(st.session_state.multi_src_dir).is_dir()
            save_ok_m   = bool(st.session_state.multi_save_dir)
            states_list = scan_states(st.session_state.multi_src_dir) if src_ok else []
            n_ready_m   = sum(1 for s in states_list if s["ready"])
            ready_m     = src_ok and save_ok_m and n_ready_m > 0

            def rdy(ok, title, sub):
                d = "dot-ok" if ok else "dot-wait"; i = "&#10003;" if ok else "&#9675;"
                return f'<div class="rdy-row"><div class="rdy-dot {d}">{i}</div><div><div class="rdy-title">{title}</div><div class="rdy-sub">{sub}</div></div></div>'

            src_sub  = (f"{len(states_list)} folders, {n_ready_m} ready" if states_list else "No states found") if src_ok else "Not selected"
            save_sub = (("…"+st.session_state.multi_save_dir[-36:]) if len(st.session_state.multi_save_dir)>38 else st.session_state.multi_save_dir) if save_ok_m else "Not yet selected"
            st.markdown('<div class="rdy-card">'
                + rdy(src_ok and bool(states_list), "Source Folder", src_sub)
                + rdy(n_ready_m > 0, f'States Ready &nbsp;<span style="font-size:10px;color:#6B7A9E;">{n_ready_m}/{len(states_list)}</span>', f"{n_ready_m} state{'s' if n_ready_m!=1 else ''} with valid NGIC ratebook")
                + rdy(save_ok_m, "Save Location", save_sub)
                + rdy(True, f'Schedule Mod &nbsp;<span style="font-size:10px;color:#6B7A9E;">Per State</span>', f'{sum(1 for s in states_list if st.session_state.multi_sched_map.get(s["name"],0)>0)} of {len(states_list)} states configured' if states_list else "No states detected")
                + '</div>', unsafe_allow_html=True)

            if st.session_state.multi_step == "idle":
                if ready_m:
                    st.markdown('<div class="btn-ready">', unsafe_allow_html=True)
                    if st.button(f"&#129413;  Create Rate Pages for {n_ready_m} State{'s' if n_ready_m!=1 else ''}", key="multi_run_btn", use_container_width=True):
                        st.session_state.multi_step = "confirm"; st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)
                else:
                    missing_m = (["source folder"] if not src_ok else (["no ready states"] if not n_ready_m else [])) + (["save location"] if not save_ok_m else [])
                    st.markdown('<div class="btn-wait">', unsafe_allow_html=True)
                    st.button(f"Waiting \u2014 {', '.join(missing_m)}", key="multi_run_dis", use_container_width=True, disabled=True)
                    st.markdown('</div>', unsafe_allow_html=True)

            elif st.session_state.multi_step == "confirm":
                st.markdown('<div class="btn-ready">', unsafe_allow_html=True)
                st.button(f"&#129413;  Create Rate Pages for {n_ready_m} State{'s' if n_ready_m!=1 else ''}", key="multi_run_cfm", use_container_width=True, disabled=True)
                st.markdown('</div>', unsafe_allow_html=True)
                st.markdown('<div class="warn-box"><div class="wb-head"><span class="wb-icon">⚠️</span><span class="wb-title">Close &amp; save all open Excel files</span></div><p class="wb-body">The builder needs exclusive access to all workbooks. Please save and close any open <code>.xlsx</code> / <code>.xlsm</code> files before proceeding.</p></div>', unsafe_allow_html=True)
                spacer(8)
                mc1, mc2 = st.columns(2)
                with mc1:
                    if st.button("Cancel", key="multi_cancel_btn", use_container_width=True, type="secondary"):
                        st.session_state.multi_step = "idle"; st.rerun()
                with mc2:
                    if st.button("Proceed", key="multi_proceed_btn", use_container_width=True, type="primary"):
                        st.session_state.multi_step = "processing"; st.rerun()

            elif st.session_state.multi_step == "processing":
                st.markdown('<div class="btn-wait">', unsafe_allow_html=True)
                st.button("Processing states...", key="multi_proc_btn", use_container_width=True, disabled=True)
                st.markdown('</div>', unsafe_allow_html=True)

                from BARatePages import run as run_rate_pages, generate_pdf_only
                src_path_m   = Path(st.session_state.multi_src_dir)
                save_path_m  = Path(st.session_state.multi_save_dir)
                states_to_run = [s for s in scan_states(src_path_m) if s["ready"]]
                pdf_dir = save_path_m / "PDF"
                if st.session_state.multi_gen_pdf: pdf_dir.mkdir(parents=True, exist_ok=True)

                prog_bar    = st.progress(0.0)
                state_ph    = st.empty()
                results = []

                for idx, s in enumerate(states_to_run):
                    sname = s["name"]
                    state_ph.markdown(f'<div class="multi-progress-card"><div class="mp-label">Processing</div><div class="mp-state">{sname}</div><div class="mp-count">{idx+1} of {len(states_to_run)}</div></div>', unsafe_allow_html=True)
                    out_dir = save_path_m / sname
                    out_dir.mkdir(parents=True, exist_ok=True)
                    def _p(k): files = s["books"].get(k); return files[0] if files and len(files)==1 else None
                    try:
                        xlsx_out, _ = run_rate_pages(
                            NGICRatebook=_p("NGIC"), MMRatebook=_p("MM"), NACORatebook=_p("NACO"),
                            NAFFRatebook=_p("NAFF"), NICOFRatebook=_p("NICOF"), HICNJRatebook=_p("HICNJ"),
                            CCMICRatebook=_p("CCMIC"), NWAGRatebook=_p("NWAG"), CWRatebook=None,
                            folder_selected=str(out_dir), SchedRatingMod=int(st.session_state.multi_sched_map.get(sname, 0)) or None, skip_pdf=True)
                        pdf_out = None
                        if st.session_state.multi_gen_pdf:
                            pdf_file = pdf_dir / (Path(xlsx_out).stem + ".pdf")
                            generate_pdf_only(xlsx_out, str(pdf_file))
                            pdf_out = str(pdf_file)
                        results.append({"state": sname, "xlsx": xlsx_out, "pdf": pdf_out, "error": None})
                    except Exception as e:
                        import traceback; traceback.print_exc()
                        results.append({"state": sname, "xlsx": None, "pdf": None, "error": str(e)})
                    prog_bar.progress((idx + 1) / len(states_to_run))

                state_ph.empty()
                st.session_state.multi_results = results
                st.session_state.multi_step = "done"
                st.rerun()

            elif st.session_state.multi_step == "done":
                results   = st.session_state.multi_results
                n_success = sum(1 for r in results if not r["error"])
                n_failed  = len(results) - n_success
                if n_failed == 0:
                    st.success(f"&#10003;  All {n_success} state{'s' if n_success!=1 else ''} completed successfully")
                else:
                    st.warning(f"{n_success} completed &nbsp;&middot;&nbsp; {n_failed} failed")
                result_rows = ""
                for r in results:
                    if r["error"]:
                        result_rows += f'<div class="arow arow-error"><span class="aco">{r["state"]}</span><span class="afile afile-err">{str(r["error"])[:70]}</span><span class="astat astat-err">&#10005;</span></div>'
                    else:
                        pdf_tag = ' &nbsp;<span style="color:var(--ok-fg);font-size:10px;">+ PDF</span>' if r["pdf"] else ''
                        result_rows += f'<div class="arow arow-ok"><span class="aco">{r["state"]}</span><span class="afile afile-assigned">{Path(r["xlsx"]).name}{pdf_tag}</span><span class="astat astat-ok">&#10003;</span></div>'
                st.markdown(f'<div class="assign-wrap"><div class="assign-hdr"><span>Results</span><span>{n_success}/{len(results)} completed</span></div>{result_rows}</div>', unsafe_allow_html=True)
                spacer(8)
                if st.button("Start Over", key="multi_reset_btn", type="secondary"):
                    st.session_state.multi_step = "idle"; st.session_state.multi_results = []; st.rerun()

            spacer(24)
            st.markdown('<div style="padding-top:14px;border-top:1px solid var(--border);"><p style="font-size:10px;color:#8892A4;letter-spacing:0.8px;text-transform:uppercase;text-align:center;margin:0;line-height:1.9;">Nationwide Insurance &nbsp;&middot;&nbsp; BA Analytics Division<br>Internal Use Only</p></div>', unsafe_allow_html=True)


# ─── OTHER LOBs ───────────────────────────────────────────────────────────────
else:
    CS = {
        "General Liability": ("&#9878;",  "General Liability", 'Wire up your GL backend in the <code>elif active_lob == "General Liability"</code> block.'),
        "Farm Auto":         ("&#128668;", "Farm Auto",         'Wire up your FA backend in the <code>elif active_lob == "Farm Auto"</code> block.'),
        "Property":          ("&#127968;", "Property",          'Wire up your Property backend in the <code>elif active_lob == "Property"</code> block.'),
    }
    ic, ti, de = CS[active_lob]
    st.markdown(f"""
    <div class="coming-soon">
      <div class="cs-icon">{ic}</div>
      <div class="cs-title">{ti} Rate Pages</div>
      <div class="cs-sub">{de}</div>
      <div class="cs-tag">Coming Soon</div>
    </div>""", unsafe_allow_html=True)