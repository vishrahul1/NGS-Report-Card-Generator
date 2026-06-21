"""
NEO GEETANJALI SCHOOLS – Report Card Generator UI
Run with: streamlit run app.py
"""

import os
import sys
import csv
import time
import tempfile
import warnings
import threading
import io

import streamlit as st
import pandas as pd

# ── Project path setup ─────────────────────────────────────────────────────────
APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

from config import CSV_PATH, OUTPUT_DIR, TEMPLATE_DIR, SUBJECTS

TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "Sample_IIT_Report_Card.docx")
DOCX_OUT_DIR  = os.path.join(OUTPUT_DIR, "docx")

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="NGS Report Card Generator",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Sidebar brand */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #001F60 0%, #003399 100%);
    }
    [data-testid="stSidebar"] * { color: #ffffff !important; }
    [data-testid="stSidebar"] .stSelectbox label { color: #ccd6f6 !important; }

    /* Metric cards */
    [data-testid="metric-container"] {
        background: #f0f4ff;
        border: 1px solid #d0daf8;
        border-radius: 10px;
        padding: 12px 16px;
    }
    [data-testid="stMetricValue"] { color: #001F60 !important; font-weight: 700; }

    /* Section headers */
    .section-header {
        background: linear-gradient(90deg, #001F60, #003399);
        color: white;
        padding: 8px 16px;
        border-radius: 6px;
        font-size: 1rem;
        font-weight: 600;
        margin-bottom: 12px;
    }

    /* Upload zone */
    [data-testid="stFileUploader"] {
        border: 2px dashed #003399;
        border-radius: 10px;
        background: #f8f9ff;
    }

    /* Dataframe */
    [data-testid="stDataFrame"] { border-radius: 8px; }

    /* Status badge */
    .badge-ok   { background:#d4edda; color:#155724; padding:2px 10px; border-radius:12px; font-size:0.8rem; }
    .badge-warn { background:#fff3cd; color:#856404; padding:2px 10px; border-radius:12px; font-size:0.8rem; }
    .badge-err  { background:#f8d7da; color:#721c24; padding:2px 10px; border-radius:12px; font-size:0.8rem; }

    /* Progress label */
    .prog-label { font-size:0.85rem; color:#555; margin-bottom:4px; }
</style>
""", unsafe_allow_html=True)

# ── Session state defaults ─────────────────────────────────────────────────────
for key, default in {
    "ingest_log": [],
    "ingest_done": False,
    "ingest_error": None,
    "gen_log": [],
    "gen_done": False,
    "gen_error": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


# ── Helpers ────────────────────────────────────────────────────────────────────

def load_csv() -> pd.DataFrame | None:
    if not os.path.exists(CSV_PATH):
        return None
    df = pd.read_csv(CSV_PATH, dtype=str)
    for col in ["Physics Score", "Chemistry Score", "Mathematics Score"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def csv_summary(df: pd.DataFrame) -> dict:
    return {
        "total_records":   len(df),
        "unique_students": df["Username"].nunique(),
        "batches":         sorted(df["Batch"].dropna().unique().tolist()),
        "exams":           sorted(df["Exam"].dropna().unique().tolist()),
    }


def section(title: str):
    st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.image(
        "https://img.icons8.com/color/96/000000/graduation-cap.png",
        width=60,
    )
    st.markdown("## NEO GEETANJALI\n### Report Card Generator")
    st.divider()

    page = st.radio(
        "Navigate",
        ["📤 Upload & Ingest", "📊 Data Viewer", "📈 Status Dashboard", "🏆 Batch Toppers", "⚙️ Generate Reports"],
        label_visibility="collapsed",
    )
    st.divider()

    df_check = load_csv()
    if df_check is not None:
        s = csv_summary(df_check)
        st.markdown(f"**{s['total_records']}** records  \n**{s['unique_students']}** students  \n**{len(s['exams'])}** exam(s)")
    else:
        st.markdown("_No data ingested yet_")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1 – UPLOAD & INGEST
# ══════════════════════════════════════════════════════════════════════════════

if page == "📤 Upload & Ingest":
    st.title("📤 Upload Exam Excel Files")
    st.markdown("Upload one or more `.xlsx` exam files. Each file is ingested sequentially and scores are appended to the data store.")

    col_up, col_info = st.columns([3, 2], gap="large")

    with col_up:
        section("Select Excel Files")
        uploaded_files = st.file_uploader(
            "Drag and drop or click to browse — select multiple files at once",
            type=["xlsx"],
            accept_multiple_files=True,
            label_visibility="collapsed",
        )

        if uploaded_files:
            st.success(f"**{len(uploaded_files)} file(s)** selected")

            # Preview each file
            import openpyxl as _openpyxl, warnings as _w
            previews = []
            for uf in uploaded_files:
                try:
                    with _w.catch_warnings():
                        _w.simplefilter("ignore")
                        wb_prev = _openpyxl.load_workbook(
                            io.BytesIO(uf.read()), data_only=True
                        )
                    uf.seek(0)
                    exam_cell = (
                        wb_prev["Class_Wise"]["A4"].value
                        if "Class_Wise" in wb_prev.sheetnames else "?"
                    )
                    previews.append({
                        "file": uf.name,
                        "exam": str(exam_cell) if exam_cell else "?",
                        "size": f"{uf.size / 1024:.1f} KB",
                        "ok": True,
                        "error": "",
                    })
                except Exception as e:
                    previews.append({
                        "file": uf.name,
                        "exam": "—",
                        "size": f"{uf.size / 1024:.1f} KB",
                        "ok": False,
                        "error": str(e),
                    })

            prev_df = pd.DataFrame([
                {"File": p["file"], "Detected Exam": p["exam"], "Size": p["size"],
                 "Status": "✅ Ready" if p["ok"] else f"❌ {p['error']}"}
                for p in previews
            ])
            st.dataframe(prev_df, use_container_width=True, hide_index=True)

            valid_count = sum(1 for p in previews if p["ok"])
            if valid_count == 0:
                st.error("No valid files to ingest.")
            else:
                if st.button(
                    f"▶ Run Ingestion ({valid_count} file{'s' if valid_count > 1 else ''})",
                    type="primary", use_container_width=True
                ):
                    st.session_state.ingest_log   = []
                    st.session_state.ingest_done  = False
                    st.session_state.ingest_error = None

                    from ingest.excel_to_csv import ingest_excel
                    import io as _io
                    from contextlib import redirect_stdout

                    overall_bar = st.progress(0, text="Starting…")
                    file_status = st.empty()
                    log_area    = st.empty()

                    all_log_lines: list[str] = []
                    errors: list[str] = []

                    for fi, uf in enumerate(uploaded_files):
                        if not previews[fi]["ok"]:
                            all_log_lines.append(f"⚠ Skipped {uf.name}: {previews[fi]['error']}")
                            continue

                        pct = int(fi / len(uploaded_files) * 100)
                        overall_bar.progress(pct, text=f"File {fi+1}/{len(uploaded_files)}: {uf.name}")
                        file_status.info(f"Ingesting **{uf.name}** …")

                        tmp_path = None
                        try:
                            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                                tmp.write(uf.read())
                                tmp_path = tmp.name

                            buf = _io.StringIO()
                            with redirect_stdout(buf):
                                ingest_excel(tmp_path, CSV_PATH)

                            lines = buf.getvalue().strip().splitlines()
                            all_log_lines.append(f"── {uf.name} ──")
                            all_log_lines.extend(lines)
                            all_log_lines.append("")

                        except Exception as exc:
                            err_msg = f"❌ {uf.name}: {exc}"
                            all_log_lines.append(err_msg)
                            errors.append(err_msg)
                        finally:
                            if tmp_path:
                                try:
                                    os.unlink(tmp_path)
                                except OSError:
                                    pass

                        log_area.code("\n".join(all_log_lines[-20:]))

                    overall_bar.progress(100, text="All files processed!")
                    file_status.empty()
                    st.session_state.ingest_log  = all_log_lines
                    st.session_state.ingest_done = True
                    if errors:
                        st.session_state.ingest_error = "\n".join(errors)

                    st.rerun()

    with col_info:
        section("Expected Format")
        st.markdown("""
**Sheet: `Score_Sheet`**
| Column | Description |
|--------|-------------|
| `Username` | Student ID |
| `First Name` | Full name |
| `Batch` | Class / section |
| `Physics Score` | Out of 80 |
| `Chemistry Score` | Out of 80 |
| `Mathematics Score` | Out of 80 |

**Sheet: `Class_Wise`**
- Cell **A4** = Exam name (e.g. `IIT FOUNDATION - 12`)

**Multiple files:**
Each file should be a different exam. Duplicate entries are automatically merged.
""")

    # ── Ingestion result ───────────────────────────────────────────────────────
    if st.session_state.ingest_done:
        st.divider()
        section("Ingestion Result")
        log_lines = st.session_state.ingest_log

        # Aggregate metrics across all files
        total_added   = sum(int(l.split(":")[-1].strip()) for l in log_lines if "Rows added" in l)
        total_updated = sum(int(l.split(":")[-1].strip()) for l in log_lines if "Rows updated" in l)
        files_done    = sum(1 for l in log_lines if l.startswith("──"))

        c1, c2, c3 = st.columns(3)
        c1.metric("Files Ingested", files_done)
        c2.metric("Rows Added",     total_added)
        c3.metric("Rows Updated",   total_updated)

        with st.expander("Full log output"):
            st.code("\n".join(log_lines))

        if st.session_state.ingest_error:
            st.warning(f"Some files had errors:\n{st.session_state.ingest_error}")
        else:
            st.success("All files ingested! Switch to **Data Viewer** to inspect the records.")

    elif st.session_state.ingest_error and not st.session_state.ingest_done:
        st.error(f"Ingestion failed: {st.session_state.ingest_error}")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2 – DATA VIEWER
# ══════════════════════════════════════════════════════════════════════════════

elif page == "📊 Data Viewer":
    st.title("📊 Student Exam Data Viewer")

    df = load_csv()
    if df is None or df.empty:
        st.warning("No data found. Upload an Excel file first.")
        st.stop()

    s = csv_summary(df)

    # ── Top metrics ────────────────────────────────────────────────────────────
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Records",    s["total_records"])
    m2.metric("Unique Students",  s["unique_students"])
    m3.metric("Batches",          len(s["batches"]))
    m4.metric("Exams Loaded",     len(s["exams"]))

    st.divider()

    # ── Filters ────────────────────────────────────────────────────────────────
    section("Filters")
    fc1, fc2, fc3 = st.columns([2, 2, 3])

    with fc1:
        batch_opts = ["All"] + s["batches"]
        sel_batch = st.selectbox("Batch", batch_opts)

    with fc2:
        exam_opts = ["All"] + s["exams"]
        sel_exam = st.selectbox("Exam", exam_opts)

    with fc3:
        search = st.text_input("Search by name or username", placeholder="e.g. TESHMITA or 22NGS0286")

    # Apply filters
    fdf = df.copy()
    if sel_batch != "All":
        fdf = fdf[fdf["Batch"] == sel_batch]
    if sel_exam != "All":
        fdf = fdf[fdf["Exam"] == sel_exam]
    if search.strip():
        mask = (
            fdf["Username"].str.contains(search, case=False, na=False) |
            fdf["First Name"].str.contains(search, case=False, na=False)
        )
        fdf = fdf[mask]

    st.markdown(f"**Showing {len(fdf)} of {len(df)} records**")

    # ── Score columns highlight ────────────────────────────────────────────────
    score_cols = ["Physics Score", "Chemistry Score", "Mathematics Score"]

    # Styled dataframe
    def highlight_scores(val):
        try:
            v = float(val)
            if v >= 70: return "background-color: #d4edda; color: #155724"
            if v >= 50: return "background-color: #fff3cd; color: #856404"
            return "background-color: #f8d7da; color: #721c24"
        except (ValueError, TypeError):
            return ""

    styled = (
        fdf.style
        .applymap(highlight_scores, subset=[c for c in score_cols if c in fdf.columns])
        .format({c: "{:.0f}" for c in score_cols if c in fdf.columns}, na_rep="-")
    )

    st.dataframe(styled, use_container_width=True, height=420)

    # ── Download ───────────────────────────────────────────────────────────────
    col_dl1, col_dl2, _ = st.columns([1, 1, 4])
    with col_dl1:
        csv_bytes = fdf.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇ Download Filtered CSV",
            data=csv_bytes,
            file_name="filtered_exam_data.csv",
            mime="text/csv",
        )

    # ── Per-subject average chart ──────────────────────────────────────────────
    if not fdf.empty and all(c in fdf.columns for c in score_cols):
        st.divider()
        section("Subject-wise Score Distribution")

        tab_bar, tab_box = st.tabs(["Average by Batch", "Score Distribution"])

        with tab_bar:
            if sel_batch == "All":
                avg_df = (
                    fdf.groupby("Batch")[score_cols]
                    .mean()
                    .round(1)
                    .reset_index()
                    .rename(columns={"Batch": "index"})
                    .set_index("index")
                )
            else:
                avg_df = (
                    fdf[score_cols].mean().round(1)
                    .rename("Average")
                    .to_frame().T
                    .rename(index={0: sel_batch})
                )

            avg_df.columns = ["Physics", "Chemistry", "Mathematics"]
            st.bar_chart(avg_df, use_container_width=True, height=300)

        with tab_box:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
            import numpy as np

            fig, axes = plt.subplots(1, 3, figsize=(12, 3.5), facecolor="white")
            colors = ["#4472C4", "#ED7D31", "#A5A5A5"]
            labels = ["Physics", "Chemistry", "Mathematics"]
            for ax, col, color, label in zip(axes, score_cols, colors, labels):
                vals = fdf[col].dropna().astype(float)
                ax.hist(vals, bins=12, color=color, edgecolor="white", alpha=0.85)
                ax.axvline(vals.mean(), color="#333", linestyle="--", linewidth=1.2,
                           label=f"Avg: {vals.mean():.1f}")
                ax.set_title(label, fontweight="bold", color="#001F60")
                ax.set_xlabel("Score")
                ax.set_ylabel("Students")
                ax.legend(fontsize=8)
                ax.spines["top"].set_visible(False)
                ax.spines["right"].set_visible(False)
            plt.tight_layout()
            st.pyplot(fig, use_container_width=True)
            plt.close(fig)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 3 – STATUS DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════

elif page == "📈 Status Dashboard":
    st.title("📈 Status Dashboard")

    df = load_csv()

    # ── System checks ──────────────────────────────────────────────────────────
    section("System Status")

    checks = {
        "CSV Data Store":   os.path.exists(CSV_PATH),
        "Word Template":    os.path.exists(TEMPLATE_PATH),
        "Output DOCX Dir":  os.path.isdir(DOCX_OUT_DIR),
    }

    ch_cols = st.columns(len(checks))
    for col, (name, ok) in zip(ch_cols, checks.items()):
        badge_cls = "badge-ok" if ok else "badge-err"
        badge_txt = "✅ Ready" if ok else "❌ Missing"
        col.markdown(
            f"**{name}**  \n<span class='{badge_cls}'>{badge_txt}</span>",
            unsafe_allow_html=True,
        )

    st.divider()

    if df is None or df.empty:
        st.info("No data ingested yet. Upload an Excel file on the **Upload & Ingest** page.")
        st.stop()

    s = csv_summary(df)

    # ── Progress overview ──────────────────────────────────────────────────────
    section("Ingestion Progress")

    total_expected_exams = 14  # template has Test-1 to Test-14
    exams_loaded    = len(s["exams"])
    pct_exams       = min(exams_loaded / total_expected_exams, 1.0)

    pr1, pr2 = st.columns(2)
    with pr1:
        st.markdown('<p class="prog-label">Exams Loaded (of 14 template slots)</p>', unsafe_allow_html=True)
        st.progress(pct_exams, text=f"{exams_loaded} / {total_expected_exams}")

    # Count generated docx
    docx_files = [
        f for f in os.listdir(DOCX_OUT_DIR)
        if f.endswith("_report_card.docx")
    ] if os.path.isdir(DOCX_OUT_DIR) else []
    pct_gen = min(len(docx_files) / max(s["unique_students"], 1), 1.0)

    with pr2:
        st.markdown(f'<p class="prog-label">Report Cards Generated (of {s["unique_students"]} students)</p>', unsafe_allow_html=True)
        st.progress(pct_gen, text=f"{len(docx_files)} / {s['unique_students']}")

    st.divider()

    # ── Per-batch breakdown ────────────────────────────────────────────────────
    section("Batch Breakdown")

    batch_stats = []
    for batch in s["batches"]:
        bdf = df[df["Batch"] == batch]
        students = bdf["Username"].nunique()
        exams    = bdf["Exam"].nunique()

        # Avg scores
        phy_avg  = bdf["Physics Score"].mean() if "Physics Score" in bdf else float("nan")
        chem_avg = bdf["Chemistry Score"].mean() if "Chemistry Score" in bdf else float("nan")
        math_avg = bdf["Mathematics Score"].mean() if "Mathematics Score" in bdf else float("nan")

        batch_stats.append({
            "Batch":         batch,
            "Students":      students,
            "Exams":         exams,
            "Records":       len(bdf),
            "Avg Physics":   round(float(phy_avg),  1) if not pd.isna(phy_avg)  else "-",
            "Avg Chemistry": round(float(chem_avg), 1) if not pd.isna(chem_avg) else "-",
            "Avg Maths":     round(float(math_avg), 1) if not pd.isna(math_avg) else "-",
        })

    st.dataframe(
        pd.DataFrame(batch_stats).set_index("Batch"),
        use_container_width=True,
    )

    st.divider()

    # ── Exam timeline ──────────────────────────────────────────────────────────
    section("Exams Loaded")

    exam_stats = []
    for exam in s["exams"]:
        edf = df[df["Exam"] == exam]
        exam_stats.append({
            "Exam":     exam,
            "Students": edf["Username"].nunique(),
            "Batches":  ", ".join(sorted(edf["Batch"].dropna().unique())),
        })

    for es in exam_stats:
        with st.container(border=True):
            ec1, ec2, ec3 = st.columns([4, 1, 3])
            ec1.markdown(f"**{es['Exam']}**")
            ec2.metric("Students", es["Students"], label_visibility="collapsed")
            ec3.markdown(f"<small>{es['Batches']}</small>", unsafe_allow_html=True)

    st.divider()

    # ── Top performers ─────────────────────────────────────────────────────────
    section("Top 10 Performers (by Total JEE Score)")

    score_cols = ["Physics Score", "Chemistry Score", "Mathematics Score"]
    if all(c in df.columns for c in score_cols):
        df_top = df.copy()
        df_top["Total"] = df_top[score_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1)
        top10 = (
            df_top.sort_values("Total", ascending=False)
            .head(10)[["First Name", "Username", "Batch", "Exam", "Physics Score", "Chemistry Score", "Mathematics Score", "Total"]]
            .reset_index(drop=True)
        )
        top10.index += 1
        st.dataframe(top10, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 4 – GENERATE REPORTS
# ══════════════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 4 – BATCH TOPPERS
# ══════════════════════════════════════════════════════════════════════════════

elif page == "🏆 Batch Toppers":
    st.title("🏆 Batch Toppers")
    st.markdown("Top 3 performers per batch based on average percentage across all exams.")

    df = load_csv()
    if df is None or df.empty:
        st.warning("No data found. Upload an Excel file first.")
        st.stop()

    score_cols = ["Physics Score", "Chemistry Score", "Mathematics Score"]

    # Compute per-student overall average percentage
    # CDF exams: avg / 30 * 100, JEE exams: avg / 80 * 100, then combined
    from config import CDF_MAX_PER_SUBJECT, JEE_MAX_PER_SUBJECT

    student_rows = []
    for uid, grp in df.groupby("Username"):
        first_name = grp["First Name"].iloc[0]

        cdf_s = {s: [] for s in score_cols}
        jee_s = {s: [] for s in score_cols}

        for _, row in grp.iterrows():
            exam = str(row.get("Exam", "")).strip().upper()
            is_cdf = exam.startswith("CDF")
            for col in score_cols:
                try:
                    v = float(row.get(col, ""))
                    if pd.isna(v):
                        continue
                    (cdf_s if is_cdf else jee_s)[col].append(v)
                except (ValueError, TypeError):
                    pass

        # Per-subject average % (combined CDF + JEE per subject)
        subj_pcts = {}
        for col in score_cols:
            parts = []
            if cdf_s[col]:
                parts.append(sum(cdf_s[col]) / len(cdf_s[col]) / CDF_MAX_PER_SUBJECT * 100)
            if jee_s[col]:
                parts.append(sum(jee_s[col]) / len(jee_s[col]) / JEE_MAX_PER_SUBJECT * 100)
            subj_pcts[col] = round(sum(parts) / len(parts), 1) if parts else 0.0

        # Per-subject average raw scores
        subj_cdf_avg = {
            col: round(sum(cdf_s[col]) / len(cdf_s[col]), 1) if cdf_s[col] else "-"
            for col in score_cols
        }
        subj_jee_avg = {
            col: round(sum(jee_s[col]) / len(jee_s[col]), 1) if jee_s[col] else "-"
            for col in score_cols
        }

        # Overall combined %
        cdf_subj = [sum(v)/len(v)/CDF_MAX_PER_SUBJECT*100 for v in cdf_s.values() if v]
        jee_subj = [sum(v)/len(v)/JEE_MAX_PER_SUBJECT*100 for v in jee_s.values() if v]
        overalls = []
        if cdf_subj: overalls.append(sum(cdf_subj) / len(cdf_subj))
        if jee_subj: overalls.append(sum(jee_subj) / len(jee_subj))
        combined = round(sum(overalls) / len(overalls), 1) if overalls else 0.0

        # Primary batch = CDF batch if available
        cdf_rows = grp[grp["Exam"].str.upper().str.startswith("CDF", na=False)]
        primary_batch = cdf_rows["Batch"].iloc[0] if not cdf_rows.empty else grp["Batch"].iloc[0]

        student_rows.append({
            "Username":    uid,
            "Name":        first_name,
            "Batch":       primary_batch,
            "Overall %":   combined,
            "Math %":      subj_pcts["Mathematics Score"],
            "Phy %":       subj_pcts["Physics Score"],
            "Chem %":      subj_pcts["Chemistry Score"],
            "CDF Math":    subj_cdf_avg["Mathematics Score"],
            "CDF Phy":     subj_cdf_avg["Physics Score"],
            "CDF Chem":    subj_cdf_avg["Chemistry Score"],
            "JEE Math":    subj_jee_avg["Mathematics Score"],
            "JEE Phy":     subj_jee_avg["Physics Score"],
            "JEE Chem":    subj_jee_avg["Chemistry Score"],
        })

    if not student_rows:
        st.info("No student data available.")
        st.stop()

    students_df = pd.DataFrame(student_rows)
    batches = sorted(students_df["Batch"].dropna().unique().tolist())

    medals       = ["🥇", "🥈", "🥉"]
    medal_colors = ["#FFD700", "#C0C0C0", "#CD7F32"]

    for batch in batches:
        batch_df = (
            students_df[students_df["Batch"] == batch]
            .sort_values("Overall %", ascending=False)
            .reset_index(drop=True)
            .head(3)
        )

        if batch_df.empty:
            continue

        section(f"Batch: {batch}")

        cols = st.columns(len(batch_df))
        for i, (_, row) in enumerate(batch_df.iterrows()):
            with cols[i]:
                medal = medals[i]
                color = medal_colors[i]

                # Subject rows HTML
                subj_html = ""
                for label, pct_key, cdf_key, jee_key in [
                    ("Maths",   "Math %",  "CDF Math", "JEE Math"),
                    ("Physics", "Phy %",   "CDF Phy",  "JEE Phy"),
                    ("Chem",    "Chem %",  "CDF Chem", "JEE Chem"),
                ]:
                    subj_html += f"""
                    <tr>
                      <td style="text-align:left; color:#555; font-size:0.72rem;
                                 padding:2px 4px;">{label}</td>
                      <td style="text-align:center; font-size:0.72rem;
                                 padding:2px 4px;">{row[cdf_key]}</td>
                      <td style="text-align:center; font-size:0.72rem;
                                 padding:2px 4px;">{row[jee_key]}</td>
                      <td style="text-align:center; font-weight:700;
                                 font-size:0.72rem; color:#001F60;
                                 padding:2px 4px;">{row[pct_key]}%</td>
                    </tr>"""

                st.markdown(
                    f"""
                    <div style="
                        background: linear-gradient(135deg, {color}22, {color}44);
                        border: 2px solid {color};
                        border-radius: 12px;
                        padding: 16px 12px 12px;
                        text-align: center;
                    ">
                        <div style="font-size:2rem;">{medal}</div>
                        <div style="font-weight:700; font-size:0.95rem; color:#001F60;
                                    margin:5px 0 1px;">{row['Name']}</div>
                        <div style="font-size:0.72rem; color:#555;">{row['Username']}</div>
                        <div style="font-size:1.5rem; font-weight:800; color:#001F60;
                                    margin:8px 0 2px;">{row['Overall %']}%</div>
                        <div style="font-size:0.68rem; color:#777; margin-bottom:10px;">
                            Overall Avg %</div>
                        <table style="width:100%; border-collapse:collapse;
                                      border-top:1px solid {color}88; padding-top:6px;">
                          <tr>
                            <th style="font-size:0.65rem; color:#888; text-align:left;
                                       padding:2px 4px;">Subject</th>
                            <th style="font-size:0.65rem; color:#4472C4; text-align:center;
                                       padding:2px 4px;">CDF</th>
                            <th style="font-size:0.65rem; color:#ED7D31; text-align:center;
                                       padding:2px 4px;">JEE</th>
                            <th style="font-size:0.65rem; color:#001F60; text-align:center;
                                       padding:2px 4px;">Avg%</th>
                          </tr>
                          {subj_html}
                        </table>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        st.markdown("")


elif page == "⚙️ Generate Reports":
    st.title("⚙️ Generate Report Cards")

    df = load_csv()
    if df is None or df.empty:
        st.warning("No data found. Upload an Excel file first.")
        st.stop()

    s = csv_summary(df)

    col_gen, col_info = st.columns([2, 3], gap="large")

    with col_gen:
        section("Generation Options")

        gen_mode = st.radio(
            "Generate for",
            ["Specific Batch", "All Batches"],
            horizontal=True,
        )

        sel_batch = None
        if gen_mode == "Specific Batch":
            sel_batch = st.selectbox("Select Batch", s["batches"])
            n_students = df[df["Batch"] == sel_batch]["Username"].nunique()
            st.info(f"**{n_students}** student(s) in this batch")
        else:
            st.info(f"**{s['unique_students']}** students across all batches")

        template_ok = os.path.exists(TEMPLATE_PATH)
        if not template_ok:
            st.error(f"Template not found: {TEMPLATE_PATH}")

        if st.button("▶ Generate Report Cards", type="primary",
                     disabled=not template_ok, use_container_width=True):
            st.session_state.gen_log   = []
            st.session_state.gen_done  = False
            st.session_state.gen_error = None

            log_area  = st.empty()
            prog_bar  = st.progress(0, text="Initialising…")

            batch_filter = None if gen_mode == "All Batches" else sel_batch

            try:
                import io as _io
                from contextlib import redirect_stdout

                # Get student list to estimate total
                with open(CSV_PATH, newline="", encoding="utf-8") as f:
                    all_rows = list(csv.DictReader(f))

                students_to_gen = set()
                for row in all_rows:
                    if batch_filter is None or row.get("Batch", "").strip() == batch_filter:
                        students_to_gen.add(row.get("Username", ""))
                total_gen = len(students_to_gen)

                prog_bar.progress(5, text=f"Generating for {total_gen} student(s)…")

                buf    = _io.StringIO()
                log_lines: list[str] = []

                from generate.template_filler import generate_report_cards

                # We wrap generate_report_cards and capture output line by line
                # by running it synchronously while updating UI per student via
                # a patched print
                def ui_print(*args, **kwargs):
                    line = " ".join(str(a) for a in args)
                    log_lines.append(line)
                    st.session_state.gen_log = log_lines[:]

                    # Update progress if it looks like a student line
                    if line.strip().startswith("[") and "/" in line:
                        try:
                            frac_str = line.split("[")[1].split("]")[0]
                            done, total = map(int, frac_str.split("/"))
                            prog_bar.progress(
                                int(done / total * 95) + 5,
                                text=f"Student {done}/{total}…"
                            )
                        except Exception:
                            pass

                    # Show last 12 lines in log area
                    log_area.code("\n".join(log_lines[-12:]))

                import builtins as _builtins_mod
                original_builtin_print = _builtins_mod.print
                _builtins_mod.print = ui_print

                try:
                    generate_report_cards(
                        csv_path=CSV_PATH,
                        template_path=TEMPLATE_PATH,
                        batch_filter=batch_filter,
                    )
                finally:
                    _builtins_mod.print = original_builtin_print

                prog_bar.progress(100, text="Done!")
                st.session_state.gen_done = True

            except Exception as exc:
                st.session_state.gen_error = str(exc)
                prog_bar.empty()

            st.rerun()

    with col_info:
        section("Output Location")
        st.markdown(f"""
```
{DOCX_OUT_DIR}
```
Each file: `<Username>_report_card.docx`
""")

        # Show existing generated files
        if os.path.isdir(DOCX_OUT_DIR):
            existing = sorted([f for f in os.listdir(DOCX_OUT_DIR) if f.endswith(".docx")])
            if existing:
                section(f"Already Generated ({len(existing)} files)")
                for fname in existing[:20]:
                    st.markdown(f"- {fname}")
                if len(existing) > 20:
                    st.markdown(f"_…and {len(existing)-20} more_")

    # ── Generation result ──────────────────────────────────────────────────────
    if st.session_state.gen_done:
        st.divider()
        st.success(f"Report cards generated successfully!")
        with st.expander("Full generation log"):
            st.code("\n".join(st.session_state.gen_log))

    elif st.session_state.gen_error:
        st.error(f"Generation failed: {st.session_state.gen_error}")
