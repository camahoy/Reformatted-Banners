"""
app.py — Banner Formatter · Streamlit App
Tool 1: Excel banner → Word / Excel output
"""

import streamlit as st
import pandas as pd
import io
from engine import scan_file, get_all_columns, generate_outputs

# ── Page config ───────────────────────────────────────────────
st.set_page_config(
    page_title="Banner Formatter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Styling ───────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Main background */
.stApp {
    background-color: #F7F8FA;
}

/* Hide Streamlit branding */
#MainMenu, footer, header { visibility: hidden; }

/* Custom header */
.app-header {
    background: #0F1923;
    color: white;
    padding: 2rem 2.5rem 1.5rem;
    margin: -1rem -1rem 2rem -1rem;
    border-bottom: 3px solid #2563EB;
}
.app-header h1 {
    font-size: 1.6rem;
    font-weight: 600;
    margin: 0;
    letter-spacing: -0.02em;
}
.app-header p {
    font-size: 0.85rem;
    color: #94A3B8;
    margin: 0.3rem 0 0;
}

/* Step cards */
.step-card {
    background: white;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1.2rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}
.step-label {
    font-size: 0.7rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: #2563EB;
    margin-bottom: 0.4rem;
}
.step-title {
    font-size: 1.05rem;
    font-weight: 600;
    color: #0F1923;
    margin-bottom: 0.8rem;
}

/* Format badge */
.badge {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 4px;
    font-size: 0.7rem;
    font-weight: 600;
    font-family: 'DM Mono', monospace;
    margin-left: 6px;
}
.badge-fmt2 { background: #DBEAFE; color: #1D4ED8; }
.badge-fmt3 { background: #D1FAE5; color: #065F46; }
.badge-fmt4 { background: #FEF3C7; color: #92400E; }
.badge-fmt1 { background: #F1F5F9; color: #64748B; }
.badge-error { background: #FEE2E2; color: #991B1B; }

/* Sheet list */
.sheet-row {
    display: flex;
    align-items: center;
    padding: 0.5rem 0;
    border-bottom: 1px solid #F1F5F9;
    font-size: 0.85rem;
    color: #334155;
}

/* Generate button */
div.stButton > button[kind="primary"] {
    background: #2563EB;
    color: white;
    border: none;
    border-radius: 8px;
    padding: 0.75rem 2.5rem;
    font-size: 1rem;
    font-weight: 600;
    width: 100%;
    transition: background 0.15s;
}
div.stButton > button[kind="primary"]:hover {
    background: #1D4ED8;
}

/* Download buttons */
div.stDownloadButton > button {
    background: #F0FDF4;
    color: #166534;
    border: 1.5px solid #BBF7D0;
    border-radius: 8px;
    font-weight: 600;
    width: 100%;
    padding: 0.65rem;
}

/* Password screen */
.login-wrap {
    max-width: 380px;
    margin: 5rem auto;
    background: white;
    border: 1px solid #E2E8F0;
    border-radius: 16px;
    padding: 2.5rem;
    box-shadow: 0 4px 24px rgba(0,0,0,0.06);
}

/* Metric pills */
.metric-row {
    display: flex;
    gap: 1rem;
    margin: 0.8rem 0;
}
.metric-pill {
    background: #F1F5F9;
    border-radius: 8px;
    padding: 0.5rem 1rem;
    font-size: 0.82rem;
    color: #475569;
}
.metric-pill strong {
    color: #0F1923;
    font-size: 1rem;
}
</style>
""", unsafe_allow_html=True)

# ── Password gate ─────────────────────────────────────────────
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "banners2024")

def check_password():
    if st.session_state.get("authenticated"):
        return True
    st.markdown("""
    <div class="login-wrap">
        <div style="font-size:2rem;margin-bottom:0.5rem">📊</div>
        <div style="font-size:1.3rem;font-weight:600;margin-bottom:0.3rem">Banner Formatter</div>
        <div style="font-size:0.85rem;color:#64748B;margin-bottom:1.5rem">
            Enter your team password to continue.
        </div>
    </div>
    """, unsafe_allow_html=True)
    pwd = st.text_input("Password", type="password", label_visibility="collapsed",
                        placeholder="Enter password…")
    if st.button("Sign in", use_container_width=True):
        if pwd == APP_PASSWORD:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False

if not check_password():
    st.stop()

# ── Header ────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <h1>📊 Banner Formatter</h1>
    <p>Upload an Excel banner · Configure columns & rows · Download Word or Excel output</p>
</div>
""", unsafe_allow_html=True)

# ── Session state init ────────────────────────────────────────
for key, default in [
    ("sheet_metas", None),
    ("file_bytes", None),
    ("file_name", None),
    ("all_cols", None),
    ("results", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ── STEP 1: Upload ────────────────────────────────────────────
st.markdown("""
<div class="step-card">
    <div class="step-label">Step 1</div>
    <div class="step-title">Upload your Excel banner file</div>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Drop your .xlsx file here",
    type=["xlsx"],
    label_visibility="collapsed",
)

if uploaded and uploaded.name != st.session_state.get("file_name"):
    with st.spinner("Reading file and detecting formats…"):
        file_bytes = uploaded.read()
        st.session_state["file_bytes"]  = file_bytes
        st.session_state["file_name"]   = uploaded.name
        st.session_state["sheet_metas"] = scan_file(file_bytes)
        st.session_state["all_cols"]    = get_all_columns(st.session_state["sheet_metas"])
        st.session_state["results"]     = None

# ── STEP 2: Sheet summary + selection ────────────────────────
if st.session_state["sheet_metas"]:
    metas = st.session_state["sheet_metas"]

    fmt_counts = {}
    for m in metas:
        fmt_counts[m["fmt"]] = fmt_counts.get(m["fmt"], 0) + 1

    fmt_labels = {"fmt2": "Standard", "fmt3": "Top Box Summary",
                  "fmt4": "Grid Table", "fmt1": "Legacy", "error": "Error"}

    pills = " ".join([
        f'<span class="metric-pill"><strong>{v}</strong> {fmt_labels.get(k,k)}</span>'
        for k, v in fmt_counts.items()
    ])
    st.markdown(f"""
    <div class="step-card">
        <div class="step-label">Step 2</div>
        <div class="step-title">Review detected sheets</div>
        <div class="metric-row">{pills}</div>
    </div>
    """, unsafe_allow_html=True)

    with st.expander(f"📋 All sheets ({len(metas)}) — click to expand / collapse", expanded=True):
        col_check, col_q, col_fmt = st.columns([0.5, 6, 1.5])
        col_check.markdown("**Include**")
        col_q.markdown("**Question**")
        col_fmt.markdown("**Format**")
        st.divider()

        selected_indices = []
        for m in metas:
            badge_cls = f"badge-{m['fmt']}"
            c1, c2, c3 = st.columns([0.5, 6, 1.5])
            include = c1.checkbox(
                "", value=(m["fmt"] != "error"),
                key=f"inc_{m['index']}",
                label_visibility="collapsed",
            )
            c2.markdown(
                f"<span style='font-size:0.82rem;color:#334155'>"
                f"<b>{m['sheet_name']}</b> — "
                f"{m['question_wording'][:90]}{'…' if len(m['question_wording'])>90 else ''}"
                f"</span>",
                unsafe_allow_html=True,
            )
            c3.markdown(
                f"<span class='badge {badge_cls}'>{m['fmt'].upper()}</span>",
                unsafe_allow_html=True,
            )
            if include:
                selected_indices.append(m["index"])

    st.session_state["selected_indices"] = selected_indices

# ── STEP 3: Configure ─────────────────────────────────────────
if st.session_state.get("all_cols"):
    all_cols = st.session_state["all_cols"]

    st.markdown("""
    <div class="step-card">
        <div class="step-label">Step 3</div>
        <div class="step-title">Configure your output</div>
    </div>
    """, unsafe_allow_html=True)

    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("#### 🗂 Columns to include")
        st.caption("Select which categories appear as columns in your tables.")

        # Quick select buttons
        qc1, qc2 = st.columns(2)
        if qc1.button("Select all", key="sel_all", use_container_width=True):
            for g, _ in all_cols:
                st.session_state[f"col_{g}"] = True
        if qc2.button("Clear all", key="clr_all", use_container_width=True):
            for g, _ in all_cols:
                st.session_state[f"col_{g}"] = False

        selected_cols = []
        for g, s in all_cols:
            default = s.lower() == 'total' or s == ''
            checked = st.checkbox(
                f"{g}" + (f"  ·  *{s}*" if s and s.lower() != 'total' else ''),
                value=st.session_state.get(f"col_{g}", default),
                key=f"col_{g}",
            )
            if checked:
                selected_cols.append(g)

    with col_b:
        st.markdown("#### ⚙️ Output settings")

        output_format = st.radio(
            "Output format",
            options=["Word (.docx)", "Excel (.xlsx)", "Both"],
            index=2,
            horizontal=True,
        )
        output_map = {
            "Word (.docx)": "word",
            "Excel (.xlsx)": "excel",
            "Both": "both",
        }

        excel_mode = "per_question"
        if output_format in ("Excel (.xlsx)", "Both"):
            excel_mode_label = st.radio(
                "Excel sheet layout",
                options=["One sheet per question", "One sheet per table"],
                index=0,
                horizontal=True,
            )
            excel_mode = "per_question" if "question" in excel_mode_label else "per_table"

        st.markdown("---")
        st.markdown("#### 📄 Document settings")

        orientation = st.radio(
            "Page orientation",
            options=["Landscape", "Portrait"],
            index=0,
            horizontal=True,
        )
        portrait_landscape = orientation == "Portrait"

        weighted_data = st.toggle("Weighted data", value=False,
                                  help="Turn on if your banner has weighted counts on a separate row below unweighted counts.")

        st.markdown("---")
        st.markdown("#### 🔢 Row filter")
        st.caption("Choose which answer rows to include in tables.")

        row_filter_label = st.radio(
            "Rows to show",
            options=["All rows", "Nets only (Top/Bottom Box)", "Custom selection"],
            index=0,
        )

        row_filter = "all"
        if row_filter_label == "Nets only (Top/Bottom Box)":
            row_filter = "nets_only"
        elif row_filter_label == "Custom selection":
            st.caption(
                "Enter the exact row labels you want, one per line. "
                "Copy from your banner (e.g. 'Top 2 Box (Net)')."
            )
            custom_text = st.text_area(
                "Row labels",
                placeholder="Top 2 Box (Net)\nTop 3 Box (Net)\nBottom 2 Box (Net)",
                height=120,
                label_visibility="collapsed",
            )
            row_filter = [r.strip() for r in custom_text.splitlines() if r.strip()] or "all"

# ── STEP 4: Generate ──────────────────────────────────────────
if (
    st.session_state.get("file_bytes") is not None
    and st.session_state.get("selected_indices")
    and selected_cols
):
    st.markdown("""
    <div class="step-card">
        <div class="step-label">Step 4</div>
        <div class="step-title">Generate your output</div>
    </div>
    """, unsafe_allow_html=True)

    n_selected = len(st.session_state["selected_indices"])
    st.info(f"Ready to process **{n_selected} sheets** with **{len(selected_cols)} columns**.")

    if st.button("🚀 Generate Document", type="primary"):
        progress_bar = st.progress(0, text="Starting…")

        def update_progress(pct, msg):
            progress_bar.progress(pct, text=msg)

        with st.spinner(""):
            results = generate_outputs(
                file_bytes             = st.session_state["file_bytes"],
                selected_sheet_indices = st.session_state["selected_indices"],
                desired_groups         = selected_cols,
                output_format          = output_map[output_format],
                excel_mode             = excel_mode,
                portrait_landscape     = portrait_landscape,
                weighted_data          = weighted_data,
                row_filter             = row_filter,
                progress_callback      = update_progress,
            )

        progress_bar.progress(1.0, text="Done!")
        st.session_state["results"] = results

    # Show results if available
    if st.session_state["results"]:
        results = st.session_state["results"]

        if results["errors"]:
            with st.expander(f"⚠️ {len(results['errors'])} sheet(s) had errors"):
                for sheet, err in results["errors"]:
                    st.error(f"**{sheet}**: {err}")

        if results["skipped"]:
            with st.expander(f"ℹ️ {len(results['skipped'])} sheet(s) skipped"):
                for s in results["skipped"]:
                    st.write(f"• {s}")

        st.success("✅ Output ready — download below")

        dl1, dl2 = st.columns(2)
        base_name = st.session_state["file_name"].replace(".xlsx", "")

        if results["word_bytes"]:
            dl1.download_button(
                label="⬇️ Download Word (.docx)",
                data=results["word_bytes"],
                file_name=f"{base_name}_formatted.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        if results["excel_bytes"]:
            dl2.download_button(
                label="⬇️ Download Excel (.xlsx)",
                data=results["excel_bytes"],
                file_name=f"{base_name}_formatted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

elif st.session_state.get("file_bytes") and not st.session_state.get("selected_indices"):
    st.warning("Select at least one sheet to continue.")
