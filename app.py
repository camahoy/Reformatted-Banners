"""
app.py — Banner Formatter · Streamlit App
Tool 1: Excel banner → Word / Excel output
"""

import re
import streamlit as st
from engine import (
    scan_file, get_all_columns, get_question_groups,
    scan_rows_for_sheets, generate_outputs,
)

# ── Page config ───────────────────────────────────────────────
st.set_page_config(
    page_title="Banner Formatter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background-color: #F7F8FA; }
#MainMenu, footer, header { visibility: hidden; }
.app-header {
    background: #0F1923; color: white;
    padding: 1.8rem 2.5rem 1.4rem;
    margin: -1rem -1rem 2rem -1rem;
    border-bottom: 3px solid #2563EB;
}
.app-header h1 { font-size: 1.5rem; font-weight: 600; margin: 0; letter-spacing: -0.02em; }
.app-header p  { font-size: 0.83rem; color: #94A3B8; margin: 0.3rem 0 0; }
.step-card {
    background: white; border: 1px solid #E2E8F0;
    border-radius: 12px; padding: 1.4rem 1.6rem;
    margin-bottom: 1rem; box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}
.step-label {
    font-size: 0.68rem; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.08em; color: #2563EB; margin-bottom: 0.3rem;
}
.step-title { font-size: 1rem; font-weight: 600; color: #0F1923; }
.badge {
    display: inline-block; padding: 2px 7px; border-radius: 4px;
    font-size: 0.68rem; font-weight: 700;
    font-family: 'DM Mono', monospace; margin-left: 5px;
}
.badge-fmt2 { background:#DBEAFE; color:#1D4ED8; }
.badge-fmt3 { background:#D1FAE5; color:#065F46; }
.badge-fmt4 { background:#FEF3C7; color:#92400E; }
.badge-fmt1 { background:#F1F5F9; color:#64748B; }
.badge-error{ background:#FEE2E2; color:#991B1B; }
.table-config-box {
    background: #F8FAFC; border: 1px solid #E2E8F0;
    border-radius: 8px; padding: 0.8rem 1rem; margin-bottom: 0.5rem;
}
div.stButton > button[kind="primary"] {
    background: #2563EB; color: white; border: none;
    padding: 0.7rem 2rem; font-size: 0.95rem; font-weight: 600; width: 100%;
}
div.stButton > button[kind="primary"]:hover { background: #1D4ED8; }
div.stDownloadButton > button {
    background: #F0FDF4; color: #166534;
    border: 1.5px solid #BBF7D0; border-radius: 8px;
    font-weight: 600; width: 100%; padding: 0.6rem;
}
.login-wrap {
    max-width: 360px; margin: 5rem auto;
    background: white; border: 1px solid #E2E8F0;
    border-radius: 16px; padding: 2.5rem;
    box-shadow: 0 4px 24px rgba(0,0,0,0.06);
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
        <div style="font-size:0.83rem;color:#64748B;margin-bottom:1.5rem">
            Enter your team password to continue.
        </div>
    </div>
    """, unsafe_allow_html=True)
    pwd = st.text_input("Password", type="password",
                        label_visibility="collapsed", placeholder="Password…")
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
    <p>Upload · Configure · Download Word or Excel output</p>
</div>
""", unsafe_allow_html=True)

# ── Session state init ────────────────────────────────────────
for k, v in [
    ("sheet_metas", None), ("file_bytes", None), ("file_name", None),
    ("all_cols", None), ("q_groups", None), ("results", None),
    ("group_rows_cache", {}),
    ("table_configs", {}),
    ("selected_indices", []),
]:
    if k not in st.session_state:
        st.session_state[k] = v

# ── Default config values ─────────────────────────────────────
selected_cols      = []
output_format      = "both"
excel_mode         = "per_question"
portrait_landscape = False
weighted_data      = False

# ── STEP 1: Upload ────────────────────────────────────────────
st.markdown("""
<div class="step-card">
    <div class="step-label">Step 1</div>
    <div class="step-title">Upload your Excel banner file</div>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("File", type=["xlsx"], label_visibility="collapsed")

if uploaded and uploaded.name != st.session_state.get("file_name"):
    with st.spinner("Reading file and detecting formats…"):
        fb = uploaded.read()
        metas = scan_file(fb)
        st.session_state.update({
            "file_bytes":         fb,
            "file_name":          uploaded.name,
            "sheet_metas":        metas,
            "all_cols":           get_all_columns(metas),
            "q_groups":           get_question_groups(metas),
            "selected_indices":   [m["index"] for m in metas if m["fmt"] != "error"],
            "results":            None,
            "group_rows_cache":   {},
            "table_configs":      {},
        })

# ── STEP 2: Sheet review ──────────────────────────────────────
if st.session_state["sheet_metas"]:
    metas = st.session_state["sheet_metas"]
    fmt_counts = {}
    for m in metas:
        fmt_counts[m["fmt"]] = fmt_counts.get(m["fmt"], 0) + 1
    fmt_labels = {"fmt2":"Standard","fmt3":"Top Box","fmt4":"Grid Table",
                  "fmt1":"Legacy","error":"Error"}

    summary = "  ·  ".join([f"**{v}** {fmt_labels.get(k,k)}" for k,v in fmt_counts.items()])
    st.markdown(f"""
    <div class="step-card">
        <div class="step-label">Step 2</div>
        <div class="step-title">Detected sheets</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown(summary)

    with st.expander(f"📋 Review all {len(metas)} sheets", expanded=False):
        new_selected = []
        for m in metas:
            c1, c2, c3 = st.columns([0.5, 6.5, 1.2])
            inc = c1.checkbox("", value=(m["index"] in st.session_state["selected_indices"]),
                              key=f"inc_{m['index']}", label_visibility="collapsed")
            c2.markdown(
                f"<span style='font-size:0.81rem;color:#334155'>"
                f"<b>{m['sheet_name']}</b> — "
                f"{m['question_wording'][:85]}{'…' if len(m['question_wording'])>85 else ''}"
                f"</span>", unsafe_allow_html=True)
            c3.markdown(
                f"<span class='badge badge-{m['fmt']}'>{m['fmt'].upper()}</span>",
                unsafe_allow_html=True)
            if inc:
                new_selected.append(m["index"])
        st.session_state["selected_indices"] = new_selected

# ── STEP 3: Configure ─────────────────────────────────────────
if st.session_state.get("q_groups"):
    q_groups   = st.session_state["q_groups"]
    all_cols   = st.session_state["all_cols"]
    file_bytes = st.session_state["file_bytes"]

    st.markdown("""
    <div class="step-card">
        <div class="step-label">Step 3</div>
        <div class="step-title">Configure your output</div>
    </div>
    """, unsafe_allow_html=True)

    # Output settings row
    oc1, oc2, oc3, oc4 = st.columns(4)
    with oc1:
        ofl = st.radio("Output format",
                       ["Word (.docx)", "Excel (.xlsx)", "Both"], index=2)
        output_format = {"Word (.docx)":"word","Excel (.xlsx)":"excel","Both":"both"}[ofl]
    with oc2:
        if output_format in ("excel","both"):
            eml = st.radio("Excel layout",
                           ["One sheet per question","One sheet per table"], index=0)
            excel_mode = "per_question" if "question" in eml else "per_table"
        else:
            st.radio("Excel layout",
                     ["One sheet per question","One sheet per table"],
                     index=0, disabled=True)
    with oc3:
        ori = st.radio("Orientation", ["Landscape","Portrait"], index=0)
        portrait_landscape = ori == "Portrait"
    with oc4:
        weighted_data = st.toggle("Weighted data", value=False,
            help="Enable if your banner has weighted counts on a separate row.")

    st.divider()

    # Columns
    st.markdown("#### 🗂 Columns to include")
    st.caption("Applies to Standard and Top Box sheets. Grid tables use their brand columns automatically.")

    qca, qcb = st.columns(2)
    if qca.button("Select all", key="sel_all_cols", use_container_width=True):
        for g, _ in all_cols:
            st.session_state[f"col_{g}"] = True
    if qcb.button("Clear all", key="clr_all_cols", use_container_width=True):
        for g, _ in all_cols:
            st.session_state[f"col_{g}"] = False

    col_grid = st.columns(4)
    selected_cols = []
    for ci, (g, s) in enumerate(all_cols):
        label   = g if (not s or s.lower() == 'total') else f"{g}  ·  {s}"
        default = not s or s.lower() in ('total', '')
        checked = col_grid[ci % 4].checkbox(
            label,
            value=st.session_state.get(f"col_{g}", default),
            key=f"col_{g}",
        )
        if checked:
            selected_cols.append(g)

    st.divider()

    # Per-group table config
    st.markdown("#### 📋 Tables per question group")
    st.caption("Default is to print all rows for each question. "
               "Expand a group to add custom filtered tables.")

    tc = st.session_state["table_configs"]
    # Init defaults for new groups
    for grp in q_groups:
        p = grp["prefix"]
        if p not in tc:
            tc[p] = [{"label": "All rows", "rows": "all"}]

    for grp in q_groups:
        prefix  = grp["prefix"]
        fmt     = grp["fmt"]
        sheets  = grp["sheets"]
        n       = len(sheets)
        sel_in_group = [s["index"] for s in sheets
                        if s["index"] in st.session_state["selected_indices"]]
        if not sel_in_group:
            continue

        with st.expander(
            f"**{prefix}** — {n} sheet{'s' if n>1 else ''} "
            f"[{fmt.upper()}]  ·  {len(tc[prefix])} table(s) configured",
            expanded=False,
        ):
            # Sheet list
            for s in sheets:
                st.markdown(
                    f"<span style='font-size:0.77rem;color:#64748B'>• "
                    f"{s['question_wording'][:90]}</span>",
                    unsafe_allow_html=True)
            st.markdown("")

            # Lazy-load rows
            cache = st.session_state["group_rows_cache"]
            if prefix not in cache:
                with st.spinner(f"Loading rows for {prefix}…"):
                    cache[prefix] = scan_rows_for_sheets(file_bytes, sel_in_group)

            available_rows = cache.get(prefix, [])
            configs = tc[prefix]
            to_delete = []

            for ti, cfg in enumerate(configs):
                st.markdown("<div class='table-config-box'>", unsafe_allow_html=True)

                h1, h2 = st.columns([5, 1])
                cfg["label"] = h1.text_input(
                    "Label", value=cfg["label"],
                    key=f"tlabel_{prefix}_{ti}",
                    label_visibility="collapsed",
                    placeholder="Table name…",
                )
                if ti > 0 and h2.button("✕", key=f"del_{prefix}_{ti}",
                                         use_container_width=True):
                    to_delete.append(ti)

                row_mode = st.radio(
                    "Rows",
                    ["All rows", "Custom selection"],
                    index=0 if cfg["rows"] == "all" else 1,
                    key=f"rmode_{prefix}_{ti}",
                    horizontal=True,
                )

                if row_mode == "Custom selection" and available_rows:
                    st.caption("Check the rows you want in this table:")
                    current = cfg["rows"] if isinstance(cfg["rows"], list) else []
                    new_rows = []
                    rc1, rc2 = st.columns(2)
                    for ri, rl in enumerate(available_rows):
                        col = rc1 if ri % 2 == 0 else rc2
                        if col.checkbox(rl, value=(rl in current),
                                        key=f"row_{prefix}_{ti}_{ri}"):
                            new_rows.append(rl)
                    cfg["rows"] = new_rows if new_rows else "all"
                else:
                    cfg["rows"] = "all"

                st.markdown("</div>", unsafe_allow_html=True)

            for i in sorted(to_delete, reverse=True):
                configs.pop(i)

            if st.button(f"＋ Add table for {prefix}",
                         key=f"add_{prefix}", use_container_width=True):
                configs.append({"label": f"Table {len(configs)+1}", "rows": "all"})
                st.rerun()

    st.session_state["table_configs"] = tc

# ── STEP 4: Generate ──────────────────────────────────────────
if (
    st.session_state.get("file_bytes") is not None
    and st.session_state.get("q_groups")
    and selected_cols
):
    st.markdown("""
    <div class="step-card">
        <div class="step-label">Step 4</div>
        <div class="step-title">Generate your output</div>
    </div>
    """, unsafe_allow_html=True)

    # Build (sheet_index, row_filter) list respecting per-group table configs
    tc = st.session_state["table_configs"]
    sheet_table_configs = []
    for grp in st.session_state["q_groups"]:
        prefix  = grp["prefix"]
        configs = tc.get(prefix, [{"label":"All rows","rows":"all"}])
        for m in grp["sheets"]:
            if m["index"] not in st.session_state.get("selected_indices", []):
                continue
            for cfg in configs:
                sheet_table_configs.append((m["index"], cfg["rows"]))

    n_tables = len(sheet_table_configs)
    st.info(f"Ready to generate **{n_tables} table(s)** with **{len(selected_cols)} column(s)**.")

    if st.button("🚀 Generate Document", type="primary"):
        bar = st.progress(0, text="Starting…")

        def upd(pct, msg):
            bar.progress(min(float(pct), 1.0), text=msg)

        with st.spinner(""):
            results = generate_outputs(
                file_bytes          = st.session_state["file_bytes"],
                sheet_table_configs = sheet_table_configs,
                desired_groups      = selected_cols,
                output_format       = output_format,
                excel_mode          = excel_mode,
                portrait_landscape  = portrait_landscape,
                weighted_data       = weighted_data,
                progress_callback   = upd,
            )
        bar.progress(1.0, text="Done!")
        st.session_state["results"] = results

    if st.session_state["results"]:
        res  = st.session_state["results"]
        base = st.session_state["file_name"].replace(".xlsx","")

        if res.get("errors"):
            with st.expander(f"⚠️ {len(res['errors'])} error(s)"):
                for sh, err in res["errors"]:
                    st.error(f"**{sh}**: {err}")
        if res.get("skipped"):
            with st.expander(f"ℹ️ {len(res['skipped'])} skipped"):
                for s in res["skipped"]:
                    st.write(f"• {s}")

        st.success("✅ Output ready")
        d1, d2 = st.columns(2)
        if res.get("word_bytes"):
            d1.download_button(
                "⬇️ Download Word (.docx)", data=res["word_bytes"],
                file_name=f"{base}_formatted.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        if res.get("excel_bytes"):
            d2.download_button(
                "⬇️ Download Excel (.xlsx)", data=res["excel_bytes"],
                file_name=f"{base}_formatted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

elif st.session_state.get("file_bytes") and not selected_cols:
    st.warning("Select at least one column to continue.")
