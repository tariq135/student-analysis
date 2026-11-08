import html
import io
from urllib.parse import quote

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Student Analysis", layout="wide")

ID_COLS = ["RollNo", "Name"]
PURPLE, NAVY = "#6D5BD0", "#1B1B3A"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
LAYOUT = dict(template="simple_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
              font_family="Inter", font_color=NAVY, margin=dict(l=10, r=10, t=50, b=10))
ICONS = [
    '<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/>'
    '<rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/>',
    '<path d="M4 20v-8M10 20V6M16 20v-6M22 20H2"/>',
    '<path d="M8 21h8M12 17v4M7 4h10v5a5 5 0 0 1-10 0V4z"/><path d="M7 6H4v2a3 3 0 0 0 3 3M17 6h3v2a3 3 0 0 1-3 3"/>',
    '<circle cx="12" cy="12" r="9"/><path d="M8 12.5l2.5 2.5L16 9.5"/>',
    '<path d="M21 12A9 9 0 1 1 12 3v9z"/><path d="M15 3.5A9 9 0 0 1 20.5 9H15z"/>',
    '<path d="M12 3l9.5 17h-19z"/><path d="M12 10v4M12 17.5v.5"/>',
    '<path d="M12 4v11M7 10l5 5 5-5M4 20h16"/>',
]
TAB = ':is([role="tab"],[data-baseweb="tab"])'
TABLIST = ':is([role="tablist"],[data-baseweb="tab-list"])'
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"],.stApp{font-family:'Inter',sans-serif}
.stApp{background:linear-gradient(180deg,#E4E5F6,#ECECF9 45%,#F3F3FB);color:#1B1B3A}
[data-testid="stHeader"],[data-testid="stToolbar"],footer,#MainMenu{display:none!important}
.block-container{max-width:1120px;padding-top:2.5rem;padding-bottom:4rem}
.hero{text-align:center;padding:12px 0 28px}
.hero h1{font-size:54px;line-height:1.08;font-weight:700;letter-spacing:-1.5px;margin:0 0 16px;color:#1B1B3A}
.hero p{color:#6B6B8A;font-size:15px;max-width:520px;margin:0 auto;line-height:1.6}
.card,.kpi,.dark{border-radius:20px;box-shadow:0 10px 34px rgba(60,55,130,.10)}
.card{background:#fff;padding:20px 22px;margin-bottom:14px}
.card p{margin:0;color:#6B6B8A;font-size:13px;line-height:1.55}
.kpi{background:#fff;padding:16px 18px;box-shadow:0 8px 28px rgba(60,55,130,.09)}
.kpi .lbl{font-size:12px;color:#6B6B8A;font-weight:500}
.kpi .val{font-size:28px;font-weight:700;letter-spacing:-.5px;color:#1B1B3A}
.kpi .sub{font-size:11px;color:#6D5BD0;font-weight:600}
.dark{background:linear-gradient(160deg,#1B1B3A,#2B2A5C);color:#fff;padding:24px 26px;margin-bottom:14px}
.dark .nm{font-size:30px;font-weight:700;margin:8px 0 2px}
.dark .rl{color:#BDB9E8;font-size:13px;margin-bottom:16px}
.dark .row{display:flex;gap:34px}
.dark .big{font-size:26px;font-weight:700}
.dark .sm{font-size:11px;color:#BDB9E8}
[data-testid="stFileUploader"] section{background:#fff;border:1.5px dashed #B9B4EC;border-radius:20px}
.stButton button,.stDownloadButton button,[data-testid="stFileUploader"] button{background:#1B1B3A;border:0;border-radius:12px;font-weight:600}
.stButton button:hover,.stDownloadButton button:hover{background:#6D5BD0}
.stButton button p,.stDownloadButton button p,[data-testid="stFileUploader"] button p{color:#fff!important}
[data-testid="stDataFrame"]{border-radius:16px;overflow:hidden}
[data-testid="stWidgetLabel"] p,[data-testid="stFileUploader"] section *{color:#44445F!important}
[data-testid="stFileUploader"] section button,[data-testid="stFileUploader"] section button *{color:#fff!important;fill:#fff!important}
@L@{gap:30px!important;background:transparent!important;box-shadow:none!important;padding:0!important;border-radius:0!important;border-bottom:1px solid #D6D5EE!important}
[data-baseweb="tab-highlight"],[data-baseweb="tab-border"]{display:none!important}
@T@{background:transparent!important;border:0!important;border-radius:0!important;box-shadow:none!important;height:auto!important;padding:8px 2px 12px!important;display:inline-flex!important;align-items:center!important;color:#6B6B8A!important;opacity:1!important}
@T@ *{color:inherit!important;opacity:1!important;font-size:15px!important;font-weight:500!important}
@T@:hover,@T@:hover *{color:#1B1B3A!important}
@T@[aria-selected="true"]{color:#1B1B3A!important;box-shadow:inset 0 -2px 0 #1B1B3A!important}
@T@[aria-selected="true"] *{color:#1B1B3A!important;font-weight:700!important}
[role="tabpanel"]{padding-top:22px!important}
"""


def svg(body):
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="black" '
            f'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">{body}</svg>')


ICON_CSS = "".join(
    f'{TAB}:nth-child({i})::before{{content:"";display:inline-block;flex:0 0 16px;width:16px;height:16px;'
    f'margin-right:8px;background:currentColor;'
    f'-webkit-mask:url("data:image/svg+xml,{quote(svg(b))}") center/contain no-repeat;'
    f'mask:url("data:image/svg+xml,{quote(svg(b))}") center/contain no-repeat}}'
    for i, b in enumerate(ICONS, 1))
st.markdown("<style>" + CSS.replace("@T@", TAB).replace("@L@", TABLIST) + ICON_CSS + "</style>",
            unsafe_allow_html=True)


def esc(v):
    return html.escape(str(v))


def fmt(x):
    return f"{x:g}"


def fmt_roll(x):
    return str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()


def kpis(items):
    for col, (label, value, *sub) in zip(st.columns(len(items)), items):
        sub_html = f'<div class="sub">{esc(sub[0])}</div>' if sub else ""
        col.markdown(f'<div class="kpi"><div class="lbl">{esc(label)}</div>'
                     f'<div class="val">{esc(value)}</div>{sub_html}</div>', unsafe_allow_html=True)


def show(data, **kw):
    try:
        st.dataframe(data, width="stretch", hide_index=True, **kw)
    except TypeError:
        st.dataframe(data, use_container_width=True, hide_index=True, **kw)


def chart(fig):
    try:
        st.plotly_chart(fig, width="stretch")
    except TypeError:
        st.plotly_chart(fig, use_container_width=True)


@st.cache_data(show_spinner=False)
def sheet_names(data):
    return pd.ExcelFile(io.BytesIO(data)).sheet_names


@st.cache_data(show_spinner=False)
def read_sheet(data, sheet):
    return pd.read_excel(io.BytesIO(data), sheet_name=sheet)


def normalise(df):
    lookup = {c.lower(): c for c in ID_COLS}
    df = df.copy()
    df.columns = [lookup.get("".join(ch for ch in str(c).lower() if ch.isalnum()), str(c).strip())
                  for c in df.columns]
    return df


def make_template():
    buf = io.BytesIO()
    pd.DataFrame({
        "RollNo": [1, 2, 3, 4, 5],
        "Name": ["Aarav Sharma", "Diya Verma", "Kabir Singh", "Isha Gupta", "Rohan Mehta"],
        "Maths": [78, 92, 35, 88, 60],
        "English": [65, 85, 70, 91, 30],
        "Science": [82, 90, 55, 79, 45],
        "Computer": [90, 95, 60, 84, 70],
    }).to_excel(buf, index=False, engine="openpyxl")
    return buf.getvalue()


def clean(raw, subjects, max_marks):
    df = raw.dropna(how="all").copy()
    df.insert(0, "Excel Row", df.index + 2)
    for s in subjects:
        df[s] = pd.to_numeric(df[s], errors="coerce")
    name = df["Name"].fillna("").astype(str).str.strip()
    m = df[subjects]
    issues = pd.DataFrame({
        "Blank or non-numeric marks": m.isna().any(axis=1),
        f"Marks outside 0-{max_marks}": ((m < 0) | (m > max_marks)).any(axis=1),
        "Missing RollNo or Name": df["RollNo"].isna() | (name == ""),
    })
    bad = issues.any(axis=1)
    skipped = df.loc[bad, ["Excel Row"] + ID_COLS + subjects].astype(str)
    skipped["Reason"] = [", ".join(r[r].index) for _, r in issues[bad].iterrows()]
    valid = df[~bad].assign(Name=name[~bad], RollNo=df.loc[~bad, "RollNo"].map(fmt_roll))
    valid = valid.drop(columns="Excel Row").reset_index(drop=True)
    dups = valid[valid["RollNo"].duplicated(keep=False)].sort_values("RollNo")[ID_COLS + subjects]
    return valid, skipped, dups


st.markdown('<div class="hero"><h1>Student Analysis</h1>'
            '<p>Upload one Excel sheet to see averages, toppers, pass and fail counts, '
            'charts and students at risk.</p></div>', unsafe_allow_html=True)

uploaded = st.file_uploader("Upload your marks file (.xlsx)", type=["xlsx"])
if uploaded is None:
    st.markdown('<div class="card"><p>Your sheet needs RollNo and Name columns. Every other column is '
                'treated as a subject, so you can add any subject by adding a column.</p></div>',
                unsafe_allow_html=True)
    st.download_button("Download sample template", make_template(),
                       file_name="sample_marks.xlsx", mime=XLSX)
    st.stop()

data = uploaded.getvalue()
try:
    sheets = sheet_names(data)
    raw = normalise(read_sheet(data, st.selectbox("Sheet", sheets) if len(sheets) > 1 else sheets[0]))
except Exception as e:
    st.error(f"Could not read the file. Please upload a valid .xlsx file. ({e})")
    st.stop()

if not set(ID_COLS) <= set(raw.columns):
    st.error(f"Your sheet needs RollNo and Name columns. Found: {', '.join(raw.columns)}")
    st.stop()

others = [c for c in raw.columns if c not in ID_COLS]
numeric = [c for c in others if pd.to_numeric(raw[c], errors="coerce").notna().any()]
c1, c2, c3, c4 = st.columns([3, 1, 1, 1])
subjects = c1.multiselect("Subjects to analyse", others, default=numeric)
max_marks = c2.number_input("Max marks", min_value=1, value=100)
pass_mark = c3.number_input("Pass mark", min_value=0, value=40)
near_mark = c4.number_input("Near-pass margin", min_value=0, value=5)

if not subjects:
    st.error("Select at least one subject.")
    st.stop()
if pass_mark > max_marks:
    st.error("The pass mark cannot be higher than the max marks.")
    st.stop()

df, skipped, dups = clean(raw, subjects, max_marks)
if not skipped.empty:
    st.warning(f"{len(skipped)} row(s) were skipped. Fix them in Excel and upload again.")
    with st.expander("See skipped rows"):
        show(skipped)
if not dups.empty:
    st.warning(f"{dups['RollNo'].nunique()} duplicate Roll No(s) found. Both rows are kept.")
    with st.expander("See duplicate Roll Nos"):
        show(dups)
if df.empty:
    st.error("No valid records found after cleaning.")
    st.stop()

df["Total"] = df[subjects].sum(axis=1)
df["Percentage"] = (df["Total"] / (len(subjects) * max_marks) * 100).round(2)
df["Failed In"] = df[subjects].lt(pass_mark).apply(lambda r: ", ".join(r[r].index) or "-", axis=1)
df["Status"] = df["Failed In"].eq("-").map({True: "Pass", False: "Fail"})
df["Short By"] = (pass_mark - df[subjects]).clip(lower=0).max(axis=1)
df["Risk"] = df["Short By"].map(lambda x: "" if x == 0 else "Near pass" if x <= near_mark else "Far below")

VIEW = ID_COLS + subjects + ["Total", "Percentage", "Failed In", "Status"]
n = len(df)
passed = int((df["Status"] == "Pass").sum())
failed = n - passed
toppers = df[df["Total"] == df["Total"].max()]
top = toppers.iloc[0]
at_risk = df[df["Risk"] != ""].sort_values("Short By", ascending=False)
near = int((at_risk["Risk"] == "Near pass").sum())

st.success(f"Loaded {n} records from {uploaded.name}")
kpis([("Total students", n), ("Passed", passed, f"{passed / n:.0%} pass rate"), ("Failed", failed),
      ("Class average", f"{df['Percentage'].mean():.1f}%"),
      ("Top score", fmt(top["Total"]), f"{top['Percentage']:.2f}%")])

tabs = st.tabs(["All Records", "Averages", "Topper", "Pass / Fail", "Charts", "At Risk", "Downloads"])

with tabs[0]:
    show(df[VIEW])

with tabs[1]:
    kpis([(f"{s} average", f"{df[s].mean():.2f}") for s in subjects])
    show(pd.DataFrame({
        "Subject": subjects,
        "Average": [round(df[s].mean(), 2) for s in subjects],
        "Highest": [fmt(df[s].max()) for s in subjects],
        "Scored by": [", ".join(df.loc[df[s] == df[s].max(), "Name"]) for s in subjects],
        "Lowest": [fmt(df[s].min()) for s in subjects],
        "Lowest by": [", ".join(df.loc[df[s] == df[s].min(), "Name"]) for s in subjects],
    }))

with tabs[2]:
    if len(toppers) > 1:
        st.info(f"{len(toppers)} students are tied for the top total.")
    for _, t in toppers.iterrows():
        st.markdown(
            f'<div class="dark"><div class="sm">Class topper</div><div class="nm">{esc(t["Name"])}</div>'
            f'<div class="rl">Roll No {esc(t["RollNo"])}</div><div class="row">'
            f'<div><div class="big">{fmt(t["Total"])}</div><div class="sm">Total marks</div></div>'
            f'<div><div class="big">{t["Percentage"]:.2f}%</div><div class="sm">Percentage</div></div>'
            f'</div></div>', unsafe_allow_html=True)

with tabs[3]:
    st.caption(f"A student passes only with {pass_mark} or more in every subject.")
    view = st.radio("Show", ["All", "Pass", "Fail"], horizontal=True, label_visibility="collapsed")
    show(df[VIEW] if view == "All" else df.loc[df["Status"] == view, VIEW])

with tabs[4]:
    avgs = df[subjects].mean()
    bar = px.bar(x=subjects, y=avgs.values, title="Average marks per subject",
                 labels={"x": "", "y": "Average marks"}, color_discrete_sequence=[PURPLE])
    bar.update_traces(texttemplate="%{y:.1f}", textposition="outside", cliponaxis=False)
    bar.update_yaxes(range=[0, max_marks]).update_layout(**LAYOUT)
    pie = px.pie(names=["Pass", "Fail"], values=[passed, failed], hole=0.6, title="Pass rate",
                 color_discrete_sequence=[PURPLE, NAVY])
    pie.update_traces(textinfo="label+percent")
    pie.add_annotation(text=f"{passed / n:.0%}", x=0.5, y=0.5, showarrow=False,
                       font=dict(size=24, color=NAVY))
    pie.update_layout(**LAYOUT)
    c1, c2 = st.columns([3, 2])
    with c1:
        chart(bar)
    with c2:
        chart(pie)

with tabs[5]:
    st.caption(f"Students with any subject below {pass_mark}. "
               f"Near pass means the shortfall is {near_mark} marks or less.")
    kpis([("At risk", len(at_risk)), ("Near pass", near), ("Far below", len(at_risk) - near)])
    if at_risk.empty:
        st.success("No students are at risk.")
    else:
        show(at_risk[["RollNo", "Name", "Failed In", "Short By", "Risk"]])

with tabs[6]:
    risk = [f"  {r['Name']} (Roll {r['RollNo']}): {r['Failed In']}, short by {fmt(r['Short By'])} ({r['Risk']})"
            for _, r in at_risk.iterrows()] or ["  None"]
    tops = [f"  {t['Name']} (Roll {t['RollNo']}, total {fmt(t['Total'])}, {t['Percentage']:.2f}%)"
            for _, t in toppers.iterrows()]
    report = "\n".join([
        "STUDENT MARKS ANALYSIS REPORT", "=" * 40,
        f"Total students : {n}", f"Passed         : {passed}", f"Failed         : {failed}", "",
        "Subject averages:", *[f"  {s}: {df[s].mean():.2f}" for s in subjects], "",
        "Topper(s):", *tops, "", "Students at risk:", *risk,
    ])
    xl = io.BytesIO()
    df.to_excel(xl, index=False, engine="openpyxl")
    c1, c2 = st.columns(2)
    c1.download_button("Download report.txt", report, file_name="report.txt", mime="text/plain")
    c2.download_button("Download results.xlsx", xl.getvalue(), file_name="results.xlsx", mime=XLSX)
