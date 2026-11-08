"""
Student Marks Data Analysis System - Web Version (v2)
Tech: Python + Streamlit + pandas + openpyxl + matplotlib

Run:  streamlit run app.py
Install:  pip install streamlit pandas openpyxl matplotlib
"""

import html
import io

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# Page config
# --------------------------------------------------------------------------
st.set_page_config(page_title="Markly - Student Marks Analysis",
                   layout="wide")

# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------
REQUIRED = ["RollNo", "Name", "Maths", "English", "Science", "Computer"]
SUBJECTS = ["Maths", "English", "Science", "Computer"]
MAX_MARKS = 100      # maximum marks per subject
PASS_MARKS = 40      # minimum marks needed in EACH subject

NAVY = "#1B1B3A"
PURPLE = "#6D5BD0"
LILAC = "#C9C3F2"
MUTED = "#6B6B8A"

# --------------------------------------------------------------------------
# Styling (lavender SaaS dashboard look)
# --------------------------------------------------------------------------
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"], .stApp { font-family: 'Inter', sans-serif; }
.stApp { background: linear-gradient(180deg, #E4E5F6 0%, #ECECF9 45%, #F3F3FB 100%); color: #1B1B3A; }
[data-testid="stHeader"], [data-testid="stToolbar"], footer, #MainMenu { display: none !important; }
.block-container { max-width: 1120px; padding-top: 1.2rem; padding-bottom: 4rem; }
.hero { text-align:center; padding:36px 0 26px; }
.hero h1 { font-size:54px; line-height:1.08; font-weight:700; letter-spacing:-1.5px; margin:0 0 16px; color:#1B1B3A; }
.hero p { font-size:15px; color:#6B6B8A; max-width:520px; margin:0 auto; line-height:1.6; }
.section-title { text-align:center; font-size:34px; font-weight:700; letter-spacing:-1px; line-height:1.15; margin:46px 0 6px; }
.section-sub { text-align:center; color:#6B6B8A; font-size:14px; margin-bottom:22px; }
.card { background:#fff; border-radius:20px; padding:20px 22px; box-shadow:0 10px 34px rgba(60,55,130,.10); margin-bottom:14px; }
.card h4 { margin:0 0 4px; font-size:15px; font-weight:700; }
.card p { margin:0; color:#6B6B8A; font-size:13px; line-height:1.55; }
.kpi { background:#fff; border-radius:18px; padding:16px 18px; box-shadow:0 8px 28px rgba(60,55,130,.09); }
.kpi .lbl { font-size:12px; color:#6B6B8A; font-weight:500; }
.kpi .val { font-size:28px; font-weight:700; letter-spacing:-.5px; margin-top:2px; }
.kpi .sub { font-size:11px; color:#6D5BD0; font-weight:600; }
.dark-card { background:linear-gradient(160deg,#1B1B3A 0%,#2B2A5C 100%); color:#fff; border-radius:22px;
  padding:24px 26px; box-shadow:0 14px 40px rgba(27,27,58,.28); margin-bottom:14px; }
.dark-card .tag { display:inline-block; background:rgba(255,255,255,.14); border-radius:8px; padding:3px 10px; font-size:11px; font-weight:600; }
.dark-card .nm { font-size:30px; font-weight:700; letter-spacing:-.6px; margin:10px 0 2px; }
.dark-card .rl { color:#BDB9E8; font-size:13px; margin-bottom:16px; }
.dark-card .row { display:flex; gap:34px; }
.dark-card .big { font-size:26px; font-weight:700; }
.dark-card .sm { font-size:11px; color:#BDB9E8; }
.bar-wrap { margin:10px 0 4px; }
.bar-top { display:flex; justify-content:space-between; font-size:13px; font-weight:600; margin-bottom:5px; }
.bar-bg { height:10px; border-radius:6px; background:#ECEBFA; overflow:hidden; }
.bar-fill { height:100%; border-radius:6px; }
.chip { display:inline-block; background:#fff; border-radius:10px; padding:6px 12px; margin:3px 4px 3px 0;
  font-size:12px; font-weight:600; box-shadow:0 4px 14px rgba(60,55,130,.10); }
[data-testid="stFileUploader"] section { background:#fff; border:1.5px dashed #B9B4EC; border-radius:20px; padding:22px; }
[data-testid="stFileUploader"] button, .stDownloadButton button, .stButton button {
  background:#1B1B3A; color:#fff; border:0; border-radius:12px; padding:8px 18px; font-weight:600; }
[data-testid="stFileUploader"] button:hover, .stDownloadButton button:hover, .stButton button:hover { background:#6D5BD0; color:#fff; }
.stDownloadButton button p, .stButton button p { color:#fff; }
[role="tablist"] { gap:30px !important; background:transparent !important; box-shadow:none !important; padding:0 !important; border-radius:0 !important; border-bottom:1px solid #D6D5EE !important; }
[data-baseweb="tab-list"] { gap:30px !important; background:transparent !important; box-shadow:none !important; padding:0 !important; border-radius:0 !important; border-bottom:1px solid #D6D5EE !important; }
[data-baseweb="tab-highlight"], [data-baseweb="tab-border"] { display:none !important; }
[role="tab"] { background:transparent !important; border:0 !important; border-radius:0 !important; box-shadow:none !important; height:auto !important; padding:8px 2px 12px !important; display:inline-flex !important; align-items:center !important; color:#6B6B8A !important; opacity:1 !important; outline:none; }
[role="tab"] * { color:inherit !important; opacity:1 !important; font-size:15px !important; font-weight:500 !important; }
[role="tab"]:hover, [role="tab"]:hover * { color:#1B1B3A !important; }
[role="tab"][aria-selected="true"] { color:#1B1B3A !important; box-shadow:inset 0 -2px 0 #1B1B3A !important; }
[role="tab"][aria-selected="true"] * { color:#1B1B3A !important; font-weight:700 !important; }
[role="tab"]:focus-visible { outline:2px solid #6D5BD0 !important; outline-offset:2px; }
[data-baseweb="tab"] { background:transparent !important; border:0 !important; border-radius:0 !important; box-shadow:none !important; height:auto !important; padding:8px 2px 12px !important; display:inline-flex !important; align-items:center !important; color:#6B6B8A !important; opacity:1 !important; outline:none; }
[data-baseweb="tab"] * { color:inherit !important; opacity:1 !important; font-size:15px !important; font-weight:500 !important; }
[data-baseweb="tab"]:hover, [data-baseweb="tab"]:hover * { color:#1B1B3A !important; }
[data-baseweb="tab"][aria-selected="true"] { color:#1B1B3A !important; box-shadow:inset 0 -2px 0 #1B1B3A !important; }
[data-baseweb="tab"][aria-selected="true"] * { color:#1B1B3A !important; font-weight:700 !important; }
[data-baseweb="tab"]:focus-visible { outline:2px solid #6D5BD0 !important; outline-offset:2px; }
[role="tab"]:nth-child(1)::before { content:""; display:inline-block; flex:0 0 16px; width:16px; height:16px; margin-right:8px; background:currentColor; -webkit-mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Crect%20x%3D%223%22%20y%3D%223%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3Crect%20x%3D%2214%22%20y%3D%223%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3Crect%20x%3D%223%22%20y%3D%2214%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3Crect%20x%3D%2214%22%20y%3D%2214%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3C/svg%3E") center / contain no-repeat; mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Crect%20x%3D%223%22%20y%3D%223%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3Crect%20x%3D%2214%22%20y%3D%223%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3Crect%20x%3D%223%22%20y%3D%2214%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3Crect%20x%3D%2214%22%20y%3D%2214%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3C/svg%3E") center / contain no-repeat; }
[role="tab"]:nth-child(2)::before { content:""; display:inline-block; flex:0 0 16px; width:16px; height:16px; margin-right:8px; background:currentColor; -webkit-mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M4%2020v-8M10%2020V6M16%2020v-6M22%2020H2%22/%3E%3C/svg%3E") center / contain no-repeat; mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M4%2020v-8M10%2020V6M16%2020v-6M22%2020H2%22/%3E%3C/svg%3E") center / contain no-repeat; }
[role="tab"]:nth-child(3)::before { content:""; display:inline-block; flex:0 0 16px; width:16px; height:16px; margin-right:8px; background:currentColor; -webkit-mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M8%2021h8M12%2017v4M7%204h10v5a5%205%200%200%201-10%200V4z%22/%3E%3Cpath%20d%3D%22M7%206H4v2a3%203%200%200%200%203%203M17%206h3v2a3%203%200%200%201-3%203%22/%3E%3C/svg%3E") center / contain no-repeat; mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M8%2021h8M12%2017v4M7%204h10v5a5%205%200%200%201-10%200V4z%22/%3E%3Cpath%20d%3D%22M7%206H4v2a3%203%200%200%200%203%203M17%206h3v2a3%203%200%200%201-3%203%22/%3E%3C/svg%3E") center / contain no-repeat; }
[role="tab"]:nth-child(4)::before { content:""; display:inline-block; flex:0 0 16px; width:16px; height:16px; margin-right:8px; background:currentColor; -webkit-mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Ccircle%20cx%3D%2212%22%20cy%3D%2212%22%20r%3D%229%22/%3E%3Cpath%20d%3D%22M8%2012.5l2.5%202.5L16%209.5%22/%3E%3C/svg%3E") center / contain no-repeat; mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Ccircle%20cx%3D%2212%22%20cy%3D%2212%22%20r%3D%229%22/%3E%3Cpath%20d%3D%22M8%2012.5l2.5%202.5L16%209.5%22/%3E%3C/svg%3E") center / contain no-repeat; }
[role="tab"]:nth-child(5)::before { content:""; display:inline-block; flex:0 0 16px; width:16px; height:16px; margin-right:8px; background:currentColor; -webkit-mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M21%2012A9%209%200%201%201%2012%203v9z%22/%3E%3Cpath%20d%3D%22M15%203.5A9%209%200%200%201%2020.5%209H15z%22/%3E%3C/svg%3E") center / contain no-repeat; mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M21%2012A9%209%200%201%201%2012%203v9z%22/%3E%3Cpath%20d%3D%22M15%203.5A9%209%200%200%201%2020.5%209H15z%22/%3E%3C/svg%3E") center / contain no-repeat; }
[role="tab"]:nth-child(6)::before { content:""; display:inline-block; flex:0 0 16px; width:16px; height:16px; margin-right:8px; background:currentColor; -webkit-mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M12%204v11M7%2010l5%205%205-5M4%2020h16%22/%3E%3C/svg%3E") center / contain no-repeat; mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M12%204v11M7%2010l5%205%205-5M4%2020h16%22/%3E%3C/svg%3E") center / contain no-repeat; }
[data-baseweb="tab"]:nth-child(1)::before { content:""; display:inline-block; flex:0 0 16px; width:16px; height:16px; margin-right:8px; background:currentColor; -webkit-mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Crect%20x%3D%223%22%20y%3D%223%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3Crect%20x%3D%2214%22%20y%3D%223%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3Crect%20x%3D%223%22%20y%3D%2214%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3Crect%20x%3D%2214%22%20y%3D%2214%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3C/svg%3E") center / contain no-repeat; mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Crect%20x%3D%223%22%20y%3D%223%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3Crect%20x%3D%2214%22%20y%3D%223%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3Crect%20x%3D%223%22%20y%3D%2214%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3Crect%20x%3D%2214%22%20y%3D%2214%22%20width%3D%227%22%20height%3D%227%22%20rx%3D%221.5%22/%3E%3C/svg%3E") center / contain no-repeat; }
[data-baseweb="tab"]:nth-child(2)::before { content:""; display:inline-block; flex:0 0 16px; width:16px; height:16px; margin-right:8px; background:currentColor; -webkit-mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M4%2020v-8M10%2020V6M16%2020v-6M22%2020H2%22/%3E%3C/svg%3E") center / contain no-repeat; mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M4%2020v-8M10%2020V6M16%2020v-6M22%2020H2%22/%3E%3C/svg%3E") center / contain no-repeat; }
[data-baseweb="tab"]:nth-child(3)::before { content:""; display:inline-block; flex:0 0 16px; width:16px; height:16px; margin-right:8px; background:currentColor; -webkit-mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M8%2021h8M12%2017v4M7%204h10v5a5%205%200%200%201-10%200V4z%22/%3E%3Cpath%20d%3D%22M7%206H4v2a3%203%200%200%200%203%203M17%206h3v2a3%203%200%200%201-3%203%22/%3E%3C/svg%3E") center / contain no-repeat; mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M8%2021h8M12%2017v4M7%204h10v5a5%205%200%200%201-10%200V4z%22/%3E%3Cpath%20d%3D%22M7%206H4v2a3%203%200%200%200%203%203M17%206h3v2a3%203%200%200%201-3%203%22/%3E%3C/svg%3E") center / contain no-repeat; }
[data-baseweb="tab"]:nth-child(4)::before { content:""; display:inline-block; flex:0 0 16px; width:16px; height:16px; margin-right:8px; background:currentColor; -webkit-mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Ccircle%20cx%3D%2212%22%20cy%3D%2212%22%20r%3D%229%22/%3E%3Cpath%20d%3D%22M8%2012.5l2.5%202.5L16%209.5%22/%3E%3C/svg%3E") center / contain no-repeat; mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Ccircle%20cx%3D%2212%22%20cy%3D%2212%22%20r%3D%229%22/%3E%3Cpath%20d%3D%22M8%2012.5l2.5%202.5L16%209.5%22/%3E%3C/svg%3E") center / contain no-repeat; }
[data-baseweb="tab"]:nth-child(5)::before { content:""; display:inline-block; flex:0 0 16px; width:16px; height:16px; margin-right:8px; background:currentColor; -webkit-mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M21%2012A9%209%200%201%201%2012%203v9z%22/%3E%3Cpath%20d%3D%22M15%203.5A9%209%200%200%201%2020.5%209H15z%22/%3E%3C/svg%3E") center / contain no-repeat; mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M21%2012A9%209%200%201%201%2012%203v9z%22/%3E%3Cpath%20d%3D%22M15%203.5A9%209%200%200%201%2020.5%209H15z%22/%3E%3C/svg%3E") center / contain no-repeat; }
[data-baseweb="tab"]:nth-child(6)::before { content:""; display:inline-block; flex:0 0 16px; width:16px; height:16px; margin-right:8px; background:currentColor; -webkit-mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M12%204v11M7%2010l5%205%205-5M4%2020h16%22/%3E%3C/svg%3E") center / contain no-repeat; mask:url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A//www.w3.org/2000/svg%22%20viewBox%3D%220%200%2024%2024%22%20fill%3D%22none%22%20stroke%3D%22black%22%20stroke-width%3D%222%22%20stroke-linecap%3D%22round%22%20stroke-linejoin%3D%22round%22%3E%3Cpath%20d%3D%22M12%204v11M7%2010l5%205%205-5M4%2020h16%22/%3E%3C/svg%3E") center / contain no-repeat; }
[role="tabpanel"] { padding-top:22px !important; }
.hero h1, .section-title, .card h4, .kpi .val { color:#1B1B3A !important; }
.card p, .kpi .lbl, .hero p, .section-sub { color:#6B6B8A !important; }
[data-testid="stWidgetLabel"] p, .stRadio label p { color:#1B1B3A !important; }
[data-testid="stFileUploader"] section span, [data-testid="stFileUploader"] section small { color:#44445F !important; }
[data-testid="stFileUploader"] section button p, [data-testid="stFileUploader"] section button span { color:#fff !important; }
[data-testid="stDataFrame"] { border-radius:16px; overflow:hidden; box-shadow:0 8px 28px rgba(60,55,130,.09); }
[data-testid="stAlert"] { border-radius:14px; }
.foot { text-align:center; color:#8B8BA8; font-size:12px; margin-top:40px; }
@media (max-width: 720px) {
  .hero h1 { font-size:34px; } .section-title { font-size:26px; }
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def esc(value) -> str:
    return html.escape(str(value))


def fmt(x) -> str:
    """287.0 -> '287', 287.5 -> '287.5' (no silent truncation)."""
    return f"{x:g}"


def show_df(data, **kwargs):
    """st.dataframe that works on old and new Streamlit versions."""
    try:
        st.dataframe(data, width="stretch", **kwargs)
    except TypeError:
        st.dataframe(data, use_container_width=True, **kwargs)


def kpi(label, value, sub=""):
    sub_html = f'<div class="sub">{esc(sub)}</div>' if sub else ""
    return (f'<div class="kpi"><div class="lbl">{esc(label)}</div>'
            f'<div class="val">{esc(value)}</div>{sub_html}</div>')


def progress_bar(label, pct, color):
    return (f'<div class="bar-wrap"><div class="bar-top"><span>{esc(label)}</span>'
            f'<span>{pct:.0f}%</span></div><div class="bar-bg">'
            f'<div class="bar-fill" style="width:{pct:.1f}%;background:{color}"></div>'
            f'</div></div>')


def normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    """'Roll No', 'roll_no', ' maths ' -> RollNo, Maths ..."""
    lookup = {c.lower(): c for c in REQUIRED}
    renamed = []
    for col in df.columns:
        key = "".join(ch for ch in str(col).lower() if ch.isalnum())
        renamed.append(lookup.get(key, str(col).strip()))
    df = df.copy()
    df.columns = renamed
    return df


def fmt_roll(x) -> str:
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return str(x).strip()


@st.cache_data(show_spinner=False)
def get_sheet_names(data: bytes):
    return pd.ExcelFile(io.BytesIO(data)).sheet_names


@st.cache_data(show_spinner=False)
def read_sheet(data: bytes, sheet):
    return pd.read_excel(io.BytesIO(data), sheet_name=sheet)


def make_template() -> bytes:
    sample = pd.DataFrame({
        "RollNo": [1, 2, 3, 4],
        "Name": ["Aarav Sharma", "Diya Verma", "Kabir Singh", "Isha Gupta"],
        "Maths": [78, 92, 35, 88],
        "English": [65, 85, 70, 91],
        "Science": [82, 90, 55, 79],
        "Computer": [90, 95, 60, 84],
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        sample.to_excel(writer, index=False, sheet_name="Marks")
    return buf.getvalue()


def clean_data(df: pd.DataFrame):
    """Returns (valid_df, skipped_df, duplicate_df)."""
    df = df.dropna(how="all").copy()
    df.insert(0, "ExcelRow", df.index + 2)          # header is row 1 in Excel

    for s in SUBJECTS:
        df[s] = pd.to_numeric(df[s], errors="coerce")

    name = df["Name"].where(df["Name"].notna()).astype("string").str.strip()
    blank_marks = df[SUBJECTS].isna().any(axis=1)
    out_of_range = ((df[SUBJECTS] < 0) | (df[SUBJECTS] > MAX_MARKS)).any(axis=1)
    blank_id = df["RollNo"].isna() | name.isna() | (name == "")

    reason = pd.Series("", index=df.index)
    reason[blank_marks] += "Blank or non-numeric marks. "
    reason[out_of_range] += f"Marks outside 0-{MAX_MARKS}. "
    reason[blank_id] += "Missing RollNo or Name. "

    invalid = blank_marks | out_of_range | blank_id
    skipped = df[invalid].copy()
    skipped["Reason"] = reason[invalid].str.strip()
    skipped = skipped[["ExcelRow"] + REQUIRED + ["Reason"]].astype(str)

    valid = df[~invalid].copy()
    valid["Name"] = name[~invalid]
    valid["RollNo"] = valid["RollNo"].apply(fmt_roll)
    valid = valid.drop(columns="ExcelRow").reset_index(drop=True)

    dup_mask = valid["RollNo"].duplicated(keep=False)
    duplicates = valid.loc[dup_mask, REQUIRED].sort_values("RollNo")
    return valid, skipped, duplicates


# --------------------------------------------------------------------------
# Top navigation + hero
# --------------------------------------------------------------------------
st.markdown(
    '<div class="hero"><h1>Take control of<br>your results page</h1>'
    '<p>Upload one Excel sheet and see averages, toppers, pass and fail counts '
    'and charts in a single clean dashboard.</p></div>',
    unsafe_allow_html=True,
)

# --------------------------------------------------------------------------
# Upload
# --------------------------------------------------------------------------
uploaded = st.file_uploader("Upload your marks file (.xlsx)", type=["xlsx"])

if uploaded is None:
    chips = "".join(f'<span class="chip">{c}</span>' for c in REQUIRED)
    st.markdown(
        '<div class="card"><h4>Your sheet needs these columns</h4>'
        f'<p style="margin-bottom:8px">Marks are read out of {MAX_MARKS} per subject. '
        f'A student passes only when every subject is {PASS_MARKS} or above.</p>{chips}</div>',
        unsafe_allow_html=True,
    )
    st.download_button("Download sample template", data=make_template(),
                       file_name="sample_marks.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.stop()

# --------------------------------------------------------------------------
# Read file (with sheet choice and error handling)
# --------------------------------------------------------------------------
file_bytes = uploaded.getvalue()
try:
    sheets = get_sheet_names(file_bytes)
    sheet = sheets[0]
    if len(sheets) > 1:
        sheet = st.selectbox("Sheet to analyse", sheets)
    raw = normalise_columns(read_sheet(file_bytes, sheet))
except Exception as e:
    st.error(f"Could not read the file. Please upload a valid .xlsx file. ({e})")
    st.stop()

missing = [c for c in REQUIRED if c not in raw.columns]
if missing:
    st.error(f"Missing columns: {', '.join(missing)}. "
             f"Found: {', '.join(map(str, raw.columns))}")
    st.stop()

# --------------------------------------------------------------------------
# Clean + validate
# --------------------------------------------------------------------------
df, skipped, duplicates = clean_data(raw)

if not skipped.empty:
    st.warning(f"{len(skipped)} row(s) were skipped. Fix them in Excel and upload again.")
    with st.expander("See skipped rows and reasons"):
        show_df(skipped, hide_index=True)

if not duplicates.empty:
    st.warning(f"{duplicates['RollNo'].nunique()} duplicate Roll No(s) found. "
               "Both rows are kept, so please check them.")
    with st.expander("See duplicate Roll Nos"):
        show_df(duplicates, hide_index=True)

if df.empty:
    st.error("No valid records found after cleaning. Check marks, names and roll numbers.")
    st.stop()

# --------------------------------------------------------------------------
# Derived columns
# --------------------------------------------------------------------------
df["Total"] = df[SUBJECTS].sum(axis=1)
df["Percentage"] = (df["Total"] / (len(SUBJECTS) * MAX_MARKS) * 100).round(2)
df["Failed In"] = df[SUBJECTS].apply(
    lambda r: ", ".join(s for s in SUBJECTS if r[s] < PASS_MARKS) or "-", axis=1)
df["Status"] = df["Failed In"].apply(lambda x: "Pass" if x == "-" else "Fail")

total_students = len(df)
pass_count = int((df["Status"] == "Pass").sum())
fail_count = total_students - pass_count
pass_rate = pass_count / total_students * 100
class_avg = df["Percentage"].mean()
toppers = df[df["Total"] == df["Total"].max()]      # handles ties

st.success(f"Loaded {total_students} records from {uploaded.name}")

# --------------------------------------------------------------------------
# KPI row
# --------------------------------------------------------------------------
k = st.columns(5)
k[0].markdown(kpi("Total students", total_students), unsafe_allow_html=True)
k[1].markdown(kpi("Passed", pass_count, f"{pass_rate:.0f}% pass rate"), unsafe_allow_html=True)
k[2].markdown(kpi("Failed", fail_count), unsafe_allow_html=True)
k[3].markdown(kpi("Class average", f"{class_avg:.1f}%"), unsafe_allow_html=True)
k[4].markdown(kpi("Top score", f"{fmt(toppers['Total'].iloc[0])}",
                  f"{toppers['Percentage'].iloc[0]:.2f}%"), unsafe_allow_html=True)

st.markdown('<div class="section-title">Analytics</div>'
            '<div class="section-sub">Every metric for your class in one place.</div>',
            unsafe_allow_html=True)

# --------------------------------------------------------------------------
# Tabs
# --------------------------------------------------------------------------
tabs = st.tabs(["All Records", "Averages", "Topper", "Pass / Fail", "Charts", "Downloads"])

# Tab 1: Records
with tabs[0]:
    show_df(df, hide_index=True)

# Tab 2: Averages + highest / lowest (with student names)
with tabs[1]:
    cols = st.columns(len(SUBJECTS))
    for col, s in zip(cols, SUBJECTS):
        col.markdown(kpi(f"{s} average", f"{df[s].mean():.2f}"), unsafe_allow_html=True)

    def names_at(subject, value):
        return ", ".join(df.loc[df[subject] == value, "Name"])

    stats = pd.DataFrame({
        "Subject": SUBJECTS,
        "Average": [round(df[s].mean(), 2) for s in SUBJECTS],
        "Highest": [fmt(df[s].max()) for s in SUBJECTS],
        "Scored by": [names_at(s, df[s].max()) for s in SUBJECTS],
        "Lowest": [fmt(df[s].min()) for s in SUBJECTS],
        "Scored by (lowest)": [names_at(s, df[s].min()) for s in SUBJECTS],
    })
    st.markdown("&nbsp;", unsafe_allow_html=True)
    show_df(stats, hide_index=True)

# Tab 3: Topper(s)
with tabs[2]:
    if len(toppers) > 1:
        st.info(f"{len(toppers)} students are tied for the top total.")
    for _, t in toppers.iterrows():
        st.markdown(
            f'<div class="dark-card"><span class="tag">Class topper</span>'
            f'<div class="nm">{esc(t["Name"])}</div>'
            f'<div class="rl">Roll No {esc(t["RollNo"])}</div>'
            f'<div class="row"><div><div class="big">{fmt(t["Total"])}</div>'
            f'<div class="sm">Total marks</div></div>'
            f'<div><div class="big">{t["Percentage"]:.2f}%</div>'
            f'<div class="sm">Percentage</div></div></div></div>',
            unsafe_allow_html=True)

# Tab 4: Pass / Fail
with tabs[3]:
    left, right = st.columns([1, 2])
    with left:
        st.markdown(
            '<div class="card"><h4>Pass / fail split</h4>'
            f'{progress_bar(f"Pass ({pass_count})", pass_rate, PURPLE)}'
            f'{progress_bar(f"Fail ({fail_count})", 100 - pass_rate, NAVY)}'
            f'<p style="margin-top:10px">A student passes only with {PASS_MARKS}+ in every subject.</p></div>',
            unsafe_allow_html=True)
    with right:
        choice = st.radio("Show", ["All", "Pass", "Fail"], horizontal=True,
                          label_visibility="collapsed")
        view = df if choice == "All" else df[df["Status"] == choice]
        show_df(view[["RollNo", "Name", "Total", "Percentage", "Status", "Failed In"]],
                hide_index=True)

# Tab 5: Charts
with tabs[4]:
    c1, c2 = st.columns([3, 2])
    with c1:
        fig, ax = plt.subplots(figsize=(7, 4))
        averages = [df[s].mean() for s in SUBJECTS]
        bars = ax.bar(SUBJECTS, averages, color=PURPLE, width=0.55)
        ax.set_ylim(0, MAX_MARKS)
        ax.set_ylabel("Average marks", color=MUTED)
        ax.set_title("Average marks per subject", color=NAVY, fontweight="bold", loc="left")
        for side in ("top", "right", "left"):
            ax.spines[side].set_visible(False)
        ax.tick_params(colors=MUTED, length=0)
        ax.yaxis.grid(True, color="#ECEBFA")
        ax.set_axisbelow(True)
        for b, v in zip(bars, averages):
            ax.text(b.get_x() + b.get_width() / 2, v + 1.5, f"{v:.1f}",
                    ha="center", fontsize=10, color=NAVY, fontweight="bold")
        st.pyplot(fig)
        plt.close(fig)
    with c2:
        fig2, ax2 = plt.subplots(figsize=(4, 4))
        sizes = [pass_count, fail_count]
        if fail_count == 0:
            sizes = [pass_count]
            colors = [PURPLE]
        else:
            colors = [PURPLE, NAVY]
        ax2.pie(sizes, colors=colors, startangle=90, counterclock=False,
                wedgeprops={"width": 0.32, "edgecolor": "white"})
        ax2.text(0, 0, f"{pass_rate:.0f}%", ha="center", va="center",
                 fontsize=26, fontweight="bold", color=NAVY)
        ax2.set_title("Pass rate", color=NAVY, fontweight="bold")
        st.pyplot(fig2)
        plt.close(fig2)

# Tab 6: Downloads
with tabs[5]:
    topper_lines = "".join(
        f"  {t['Name']} (Roll {t['RollNo']}, Total {fmt(t['Total'])}, {t['Percentage']:.2f}%)\n"
        for _, t in toppers.iterrows())
    failed_lines = "".join(
        f"  {r['Name']} (Roll {r['RollNo']}) - failed in {r['Failed In']}\n"
        for _, r in df[df["Status"] == "Fail"].iterrows()) or "  None\n"

    report = (
        "STUDENT MARKS ANALYSIS REPORT\n"
        + "=" * 40 + "\n"
        + f"Total Students : {total_students}\n"
        + f"Pass Rule      : {PASS_MARKS}+ in every subject (out of {MAX_MARKS})\n\n"
        + "Subject Averages:\n"
    )
    for s in SUBJECTS:
        report += f"  {s:<10}: {df[s].mean():.2f}\n"
    report += (
        f"\nTopper(s):\n{topper_lines}"
        f"\nPass Count     : {pass_count}\n"
        f"Fail Count     : {fail_count}\n"
        f"\nFailed Students:\n{failed_lines}"
    )

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Results")

    d1, d2 = st.columns(2)
    with d1:
        st.markdown('<div class="card"><h4>Text report</h4>'
                    '<p>Summary with averages, toppers and failed students.</p></div>',
                    unsafe_allow_html=True)
        st.download_button("Download report.txt", data=report.encode("utf-8"),
                           file_name="report.txt", mime="text/plain")
    with d2:
        st.markdown('<div class="card"><h4>Results spreadsheet</h4>'
                    '<p>All records with total, percentage and status.</p></div>',
                    unsafe_allow_html=True)
        st.download_button(
            "Download results.xlsx", data=buf.getvalue(), file_name="results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown('<div class="foot">BCA Minor Project - Python, Streamlit, pandas</div>',
            unsafe_allow_html=True)
