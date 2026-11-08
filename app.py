"""
Student Marks Data Analysis System — Web Version
Tech: Python + Streamlit + pandas + openpyxl + matplotlib
"""

import io
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# ---- Page config ----
st.set_page_config(page_title="Student Marks Analysis",
                   page_icon="📊", layout="wide")

# ---- Constants ----
REQUIRED   = ["RollNo", "Name", "Maths", "English", "Science", "Computer"]
SUBJECTS   = ["Maths", "English", "Science", "Computer"]
PASS_MARKS = 40

# ---- Title ----
st.title("Student Marks Data Analysis System")
st.caption("BCA Minor Project — Python · Streamlit · pandas")

# ---- File uploader ----
uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx", "xls"])

if uploaded is None:
    st.info("Please upload an Excel file to begin.")
    st.stop()

# ---- Read file with error handling ----
try:
    df = pd.read_excel(uploaded)
except Exception as e:
    st.error(f"❌ Could not read the file: {e}")
    st.stop()

missing = [c for c in REQUIRED if c not in df.columns]
if missing:
    st.error(f"❌ Missing columns: {', '.join(missing)}")
    st.stop()

# ---- Clean data ----
for s in SUBJECTS:
    df[s] = pd.to_numeric(df[s], errors="coerce")

bad = df[SUBJECTS].isna().any(axis=1).sum()
if bad:
    st.warning(f"⚠️ {bad} row(s) had non-numeric/blank marks and were skipped.")

df = df.dropna(subset=SUBJECTS).reset_index(drop=True)

# ---- Derived columns ----
df["Total"]      = df[SUBJECTS].sum(axis=1)
df["Percentage"] = df["Total"] / len(SUBJECTS)
df["Status"]     = df["Percentage"].apply(
    lambda p: "Pass" if p >= PASS_MARKS else "Fail")

st.success(f"✅ Loaded **{len(df)}** records from **{uploaded.name}**")

# ---- Tabs ----
tabs = st.tabs([
    "All Records",
    "Averages",
    "Topper",
    "Pass / Fail",
    "Chart",
    "Downloads",
])

# Tab 1: Records
with tabs[0]:
    st.subheader("All Records")
    st.dataframe(df, use_container_width=True)

# Tab 2: Averages + High/Low
with tabs[1]:
    st.subheader("Subject-Wise Averages")
    cols = st.columns(4)
    for i, s in enumerate(SUBJECTS):
        cols[i].metric(f"{s} Avg", f"{df[s].mean():.2f}")

    st.divider()
    st.subheader("Highest / Lowest Marks")
    stats = pd.DataFrame({
        "Subject": SUBJECTS,
        "Average": [round(df[s].mean(), 2) for s in SUBJECTS],
        "Highest": [df[s].max() for s in SUBJECTS],
        "Lowest":  [df[s].min() for s in SUBJECTS],
    })
    st.dataframe(stats, use_container_width=True, hide_index=True)

# Tab 3: Topper
with tabs[2]:
    st.subheader("🏆 Class Topper")
    top = df.loc[df["Total"].idxmax()]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Name",       top["Name"])
    c2.metric("Roll No",    top["RollNo"])
    c3.metric("Total Marks", int(top["Total"]))
    c4.metric("Percentage", f"{top['Percentage']:.2f}%")

# Tab 4: Pass/Fail
with tabs[3]:
    st.subheader("Pass / Fail Summary")
    pc = int((df["Status"] == "Pass").sum())
    fc = int((df["Status"] == "Fail").sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Students", len(df))
    c2.metric("✅ Pass", pc)
    c3.metric("❌ Fail", fc)

    st.divider()
    st.dataframe(df[["RollNo", "Name", "Total", "Percentage", "Status"]],
                 use_container_width=True, hide_index=True)

# Tab 5: Chart
with tabs[4]:
    st.subheader("Subject-Wise Average Marks (Bar Chart)")
    fig, ax = plt.subplots(figsize=(9, 4))
    averages = [df[s].mean() for s in SUBJECTS]
    bars = ax.bar(SUBJECTS, averages, color="steelblue")
    ax.set_ylabel("Average Marks")
    ax.set_ylim(0, 100)
    ax.set_title("Average Marks per Subject")
    for b, v in zip(bars, averages):
        ax.text(b.get_x() + b.get_width()/2, v + 1,
                f"{v:.1f}", ha="center", fontsize=10)
    st.pyplot(fig)

# Tab 6: Downloads
with tabs[5]:
    st.subheader("Download Outputs")
    top = df.loc[df["Total"].idxmax()]

    # --- report.txt ---
    report = (
        "STUDENT MARKS ANALYSIS REPORT\n"
        + "=" * 40 + "\n"
        + f"Total Students : {len(df)}\n\n"
        + "Subject Averages:\n"
    )
    for s in SUBJECTS:
        report += f"  {s:<10}: {df[s].mean():.2f}\n"
    report += (
        f"\nTopper         : {top['Name']} "
        f"(Total {int(top['Total'])}, {top['Percentage']:.2f}%)\n"
        f"Pass Count     : {int((df['Status']=='Pass').sum())}\n"
        f"Fail Count     : {int((df['Status']=='Fail').sum())}\n"
    )

    st.download_button(
        "⬇️ Download report.txt",
        data=report.encode("utf-8"),
        file_name="report.txt",
        mime="text/plain",
    )

    # --- results.xlsx ---
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Results")

    st.download_button(
        "⬇️ Download results.xlsx",
        data=buf.getvalue(),
        file_name="results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
