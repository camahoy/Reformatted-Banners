"""
app.py — Banner Formatter · Streamlit App
Tool 1: Excel banner → Word / Excel output
"""

import re
import streamlit as st
from engine import (
    scan_file, get_all_columns, get_question_groups,
    scan_rows_for_sheets, generate_outputs,
    scan_multi_source, generate_merged_outputs,
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
    <p>Upload · Configure · Download Word or Excel output &nbsp;·&nbsp; <span style="opacity:0.4;font-size:0.75rem">v1.9</span></p>
</div>
""", unsafe_allow_html=True)

# ── Session state init ────────────────────────────────────────
for k, v in [
    ("sheet_metas", None), ("file_bytes", None), ("file_name", None),
    ("file_hash", None), ("all_cols", None), ("q_groups", None),
    ("results", None),
    ("group_rows_cache", {}),
    ("table_configs", {}),
    ("selected_indices", []),
    # Multi-source state
    ("ms_banners", []),
    ("ms_scan", None),
    ("ms_overrides", {}),
    ("ms_results", None),
]:
    if k not in st.session_state:
        st.session_state[k] = v

# ── Default config values ─────────────────────────────────────
selected_cols        = []
output_format        = "both"
excel_mode           = "per_question"
portrait_landscape   = False
weighted_data        = False
show_sig_flags       = False
heatmap_scheme       = "None"
heatmap_custom_start = None
heatmap_custom_end   = None
use_weighted_base    = False

# ── Mode toggle ───────────────────────────────────────────────
st.markdown("""
<div class="step-card" style="padding:1rem 1.6rem">
    <div class="step-label">Mode</div>
    <div class="step-title" style="margin-bottom:0.6rem">What do you want to do?</div>
</div>
""", unsafe_allow_html=True)

app_mode = st.radio(
    "Mode",
    ["📄 Single banner", "🔀 Multi-source merge"],
    horizontal=True,
    label_visibility="collapsed",
)
is_multi = app_mode == "🔀 Multi-source merge"
st.divider()

# ══════════════════════════════════════════════════════════════
# MULTI-SOURCE MODE
# ══════════════════════════════════════════════════════════════
if is_multi:

    st.markdown("""
    <div class="step-card">
        <div class="step-label">Step 1</div>
        <div class="step-title">Upload banners to merge</div>
    </div>
    """, unsafe_allow_html=True)
    st.caption("Upload 2 or more banners. Give each a short name (e.g. 'Generation', 'Region', 'Total').")

    # Banner slots
    ms_banners = st.session_state["ms_banners"]

    # Add banner slot
    if st.button("＋ Add banner", key="ms_add_banner"):
        ms_banners.append({"name": "", "file_bytes": None, "file_name": None})
        st.session_state["ms_banners"] = ms_banners
        st.rerun()

    changed = False
    for bi, banner in enumerate(ms_banners):
        with st.container():
            c1, c2, c3 = st.columns([2, 4, 0.5])
            new_name = c1.text_input(
                "Banner name", value=banner["name"],
                key=f"ms_name_{bi}", placeholder="e.g. Generation",
                label_visibility="collapsed",
            )
            uploaded_ms = c2.file_uploader(
                "File", type=["xlsx"], key=f"ms_file_{bi}",
                label_visibility="collapsed",
            )
            if c3.button("✕", key=f"ms_del_{bi}"):
                ms_banners.pop(bi)
                st.session_state["ms_banners"] = ms_banners
                st.session_state["ms_scan"] = None
                st.rerun()

            if new_name != banner["name"]:
                ms_banners[bi]["name"] = new_name
                changed = True

            if uploaded_ms and uploaded_ms.name != banner.get("file_name"):
                ms_banners[bi]["file_bytes"] = uploaded_ms.read()
                ms_banners[bi]["file_name"]  = uploaded_ms.name
                if not ms_banners[bi]["name"]:
                    ms_banners[bi]["name"] = uploaded_ms.name.replace(".xlsx","")[:20]
                changed = True

    if changed:
        st.session_state["ms_banners"] = ms_banners
        st.session_state["ms_scan"] = None

    # Scan when 2+ banners are ready
    ready_banners = [b for b in ms_banners if b["file_bytes"] and b["name"]]
    if len(ready_banners) >= 2:
        if st.button("🔍 Scan & match questions", use_container_width=True):
            with st.spinner("Scanning banners and matching questions…"):
                file_list = [(b["file_bytes"], b["name"]) for b in ready_banners]
                st.session_state["ms_scan"] = scan_multi_source(file_list)
                st.session_state["ms_overrides"] = {}
            st.rerun()

    ms_scan = st.session_state.get("ms_scan")

    if ms_scan:
        matched   = ms_scan["matched"]
        unmatched = ms_scan["unmatched"]
        banners   = ms_scan["banners"]

        # Summary
        auto_matched = sum(
            1 for e in matched
            if all(v is not None for v in e["banner_sheets"].values())
        )
        partial = len(matched) - auto_matched

        st.markdown(f"""
        <div class="step-card">
            <div class="step-label">Step 2</div>
            <div class="step-title">Question matching</div>
        </div>
        """, unsafe_allow_html=True)

        mc1, mc2, mc3 = st.columns(3)
        mc1.metric("Auto-matched", auto_matched)
        mc2.metric("Partially matched", partial)
        mc3.metric("Unmatched", len(unmatched))

        # Manual mapping for partial/unmatched
        overrides = st.session_state["ms_overrides"]

        if partial > 0 or unmatched:
            with st.expander(f"⚠️ {partial + len(unmatched)} question(s) need manual mapping", expanded=True):
                st.caption("For each question, select which sheet in the other banner corresponds to it, or choose 'Skip'.")

                # Partial matches
                for entry in matched:
                    missing_banners = [bn for bn, sm in entry["banner_sheets"].items() if sm is None]
                    if not missing_banners:
                        continue

                    st.markdown(f"**{entry['q_id']}** — {entry['wording'][:70]}…")
                    for bn in missing_banners:
                        # Find sheets in that banner
                        b_sheets = next((b["sheets"] for b in banners if b["name"] == bn), [])
                        options  = ["— Skip —"] + [
                            f"{m['sheet_name']}: {m['question_wording'][:50]}"
                            for m in b_sheets if m["fmt"] != "error"
                        ]
                        sheet_metas = [None] + [
                            m for m in b_sheets if m["fmt"] != "error"
                        ]
                        sel = st.selectbox(
                            f"Match in **{bn}**",
                            options, index=0,
                            key=f"ms_map_{entry['q_id']}_{bn}",
                        )
                        sel_idx = options.index(sel)
                        if sel_idx > 0:
                            if entry["q_id"] not in overrides:
                                overrides[entry["q_id"]] = {}
                            overrides[entry["q_id"]][bn] = sheet_metas[sel_idx]["index"]

                # Fully unmatched
                for u_entry in unmatched:
                    st.markdown(f"**{u_entry['q_id']}** (from {u_entry['banner_name']}) — {u_entry['sheet_meta']['question_wording'][:60]}…")
                    for b in banners:
                        if b["name"] == u_entry["banner_name"]:
                            continue
                        b_sheets = b["sheets"]
                        options  = ["— Skip —"] + [
                            f"{m['sheet_name']}: {m['question_wording'][:50]}"
                            for m in b_sheets if m["fmt"] != "error"
                        ]
                        sheet_metas = [None] + [m for m in b_sheets if m["fmt"] != "error"]
                        sel = st.selectbox(
                            f"Match in **{b['name']}**",
                            options, index=0,
                            key=f"ms_umap_{u_entry['q_id']}_{b['name']}",
                        )
                        sel_idx = options.index(sel)
                        if sel_idx > 0:
                            if u_entry["q_id"] not in overrides:
                                overrides[u_entry["q_id"]] = {}
                            overrides[u_entry["q_id"]][b["name"]] = sheet_metas[sel_idx]["index"]

                st.session_state["ms_overrides"] = overrides

        # Step 3 — Column selection per banner
        st.markdown("""
        <div class="step-card">
            <div class="step-label">Step 3</div>
            <div class="step-title">Select columns per banner</div>
        </div>
        """, unsafe_allow_html=True)
        st.caption("Choose which columns from each banner to include in the merged output.")

        selected_cols_per_banner = {}
        for b in banners:
            st.markdown(f"**{b['name']}**")
            all_cols_b = b["all_cols"]
            if not all_cols_b:
                st.caption("No columns found.")
                selected_cols_per_banner[b["name"]] = []
                continue

            qca, qcb = st.columns(2)
            if qca.button("Select all", key=f"ms_selall_{b['name']}", use_container_width=True):
                for g, _ in all_cols_b:
                    st.session_state[f"ms_col_{b['name']}_{g}"] = True
            if qcb.button("Clear all", key=f"ms_clrall_{b['name']}", use_container_width=True):
                for g, _ in all_cols_b:
                    st.session_state[f"ms_col_{b['name']}_{g}"] = False

            col_grid = st.columns(4)
            sel_cols = []
            for ci, (g, s) in enumerate(all_cols_b):
                label   = g if not s or s.lower() == 'total' else f"{g}  ·  {s}"
                default = not s or s.lower() in ('total', '')
                if col_grid[ci % 4].checkbox(
                    label,
                    value=st.session_state.get(f"ms_col_{b['name']}_{g}", default),
                    key=f"ms_col_{b['name']}_{g}",
                ):
                    sel_cols.append(g)
            selected_cols_per_banner[b["name"]] = sel_cols
            st.markdown("---")

        # Step 4 — Output settings
        st.markdown("""
        <div class="step-card">
            <div class="step-label">Step 4</div>
            <div class="step-title">Output settings</div>
        </div>
        """, unsafe_allow_html=True)

        ms_c1, ms_c2, ms_c3, ms_c4 = st.columns(4)
        with ms_c1:
            ms_ofl = st.radio("Output", ["Word (.docx)","Excel (.xlsx)","Both"],
                              index=2, key="ms_output_fmt")
            ms_output = {"Word (.docx)":"word","Excel (.xlsx)":"excel","Both":"both"}[ms_ofl]
        with ms_c2:
            if ms_output in ("excel","both"):
                ms_eml = st.radio("Excel layout",
                                  ["One sheet per question","One sheet per table"],
                                  index=0, key="ms_excel_mode")
                ms_excel_mode = "per_question" if "question" in ms_eml else "per_table"
            else:
                ms_excel_mode = "per_question"
        with ms_c3:
            ms_ori = st.radio("Orientation", ["Landscape","Portrait"],
                              index=0, key="ms_orientation")
            ms_portrait = ms_ori == "Portrait"
        with ms_c4:
            ms_weighted = st.toggle("Weighted data", value=False, key="ms_weighted")

        ms_row_filter = "all"
        ms_rf_label = st.radio(
            "Rows to show",
            ["All rows", "Custom selection"],
            index=0, key="ms_row_filter", horizontal=True,
        )
        if ms_rf_label == "Custom selection":
            ms_custom = st.text_area(
                "Row labels (one per line)",
                placeholder="Top 2 Box (Net)\nBottom 2 Box (Net)",
                height=100, key="ms_custom_rows",
                label_visibility="collapsed",
            )
            ms_row_filter = [r.strip() for r in ms_custom.splitlines() if r.strip()] or "all"

        # Generate
        st.markdown("""
        <div class="step-card">
            <div class="step-label">Step 5</div>
            <div class="step-title">Generate merged output</div>
        </div>
        """, unsafe_allow_html=True)

        any_cols = any(len(v) > 0 for v in selected_cols_per_banner.values())
        if not any_cols:
            st.warning("Select at least one column from at least one banner.")
        else:
            total_q = len([e for e in matched
                           if any(v is not None for v in e["banner_sheets"].values())])
            st.info(f"Ready to merge **{total_q} questions** across **{len(banners)} banners**.")

            if st.button("🚀 Generate Merged Document", type="primary", key="ms_generate"):
                bar = st.progress(0, text="Starting…")
                def ms_upd(pct, msg):
                    bar.progress(min(float(pct), 1.0), text=msg)

                with st.spinner(""):
                    ms_results = generate_merged_outputs(
                        multi_source              = ms_scan,
                        matched_overrides         = st.session_state["ms_overrides"],
                        selected_cols_per_banner  = selected_cols_per_banner,
                        output_format             = ms_output,
                        excel_mode                = ms_excel_mode,
                        portrait_landscape        = ms_portrait,
                        weighted_data             = ms_weighted,
                        row_filter                = ms_row_filter,
                        progress_callback         = ms_upd,
                    )
                bar.progress(1.0, text="Done!")
                st.session_state["ms_results"] = ms_results

        if st.session_state.get("ms_results"):
            res  = st.session_state["ms_results"]
            if res.get("errors"):
                with st.expander(f"⚠️ {len(res['errors'])} error(s)"):
                    for sh, err in res["errors"]:
                        st.error(f"**{sh}**: {err}")
            if res.get("skipped"):
                with st.expander(f"ℹ️ {len(res['skipped'])} skipped"):
                    for s in res["skipped"]:
                        st.write(f"• {s}")

            st.success("✅ Merged output ready")
            d1, d2 = st.columns(2)
            if res.get("word_bytes"):
                d1.download_button(
                    "⬇️ Download Word (.docx)", data=res["word_bytes"],
                    file_name="merged_output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            if res.get("excel_bytes"):
                d2.download_button(
                    "⬇️ Download Excel (.xlsx)", data=res["excel_bytes"],
                    file_name="merged_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    st.stop()  # Don't render single-banner UI in multi mode

# ══════════════════════════════════════════════════════════════
# SINGLE BANNER MODE (original flow continues below)
# ══════════════════════════════════════════════════════════════
st.markdown("""
<div class="step-card">
    <div class="step-label">Step 1</div>
    <div class="step-title">Upload your Excel banner file</div>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("File", type=["xlsx"], label_visibility="collapsed")

# Allow force re-scan if user re-uploads same file after an app update
if st.session_state.get("file_name") and st.button(
    "🔄 Re-scan current file", help="Use this if columns aren't showing after an app update"
):
    st.session_state["file_hash"] = None  # force re-scan on next render
    st.rerun()

if uploaded:
    import hashlib
    file_bytes_new = uploaded.read()
    file_hash = hashlib.md5(file_bytes_new).hexdigest()
    if file_hash != st.session_state.get("file_hash"):
        with st.spinner("Reading file and detecting formats…"):
            metas   = scan_file(file_bytes_new)
            q_grps  = get_question_groups(metas)
            sel_idx = [m["index"] for m in metas if m["fmt"] != "error"]

            # Pre-load row labels for every group so custom selection works immediately
            row_cache = {}
            for grp in q_grps:
                grp_indices = [s["index"] for s in grp["sheets"] if s["index"] in sel_idx]
                if grp_indices:
                    row_cache[grp["prefix"]] = scan_rows_for_sheets(file_bytes_new, grp_indices)

            st.session_state.update({
                "file_bytes":         file_bytes_new,
                "file_name":          uploaded.name,
                "file_hash":          file_hash,
                "sheet_metas":        metas,
                "all_cols":           get_all_columns(metas),
                "q_groups":           q_grps,
                "selected_indices":   sel_idx,
                "results":            None,
                "group_rows_cache":   row_cache,
                "table_configs":      {},
            })

# ── STEP 2: Sheet review ──────────────────────────────────────
if st.session_state["sheet_metas"]:
    metas = st.session_state["sheet_metas"]
    fmt_counts = {}
    for m in metas:
        fmt_counts[m["fmt"]] = fmt_counts.get(m["fmt"], 0) + 1
    fmt_labels = {"fmt2":"Standard","fmt3":"Top Box","fmt4":"Grid Table",
                  "fmt5":"Standard (v2)","fmt1":"Legacy","error":"Error"}

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
        if weighted_data:
            use_weighted_base = st.toggle("Weighted N= in header", value=True,
                help="On = show weighted base (e.g. 7,000) in column headers. Off = show unweighted base (e.g. 7,020).")
        else:
            use_weighted_base = False

    # Heatmap + sig flags row
    st.markdown("#### 🎨 Visual formatting")
    vf1, vf2, vf3, vf4 = st.columns(4)
    with vf1:
        heatmap_scheme = st.selectbox(
            "Heatmap",
            ["None", "Blue scale", "Green scale", "Red-Green diverging", "Custom"],
            index=0,
        )
    with vf2:
        show_sig_flags = st.toggle("▲/▼ sig flags", value=False,
            help="Show ▲ if significantly higher than Total, ▼ if significantly lower.")
    heatmap_custom_start = None
    heatmap_custom_end   = None
    if heatmap_scheme == "Custom":
        with vf3:
            heatmap_custom_start = st.color_picker("Start color (low)", "#FFF7ED")
        with vf4:
            heatmap_custom_end   = st.color_picker("End color (high)", "#1D4ED8")

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

            # Rows pre-loaded at scan time — just retrieve from cache
            cache = st.session_state["group_rows_cache"]
            if prefix not in cache:
                # Fallback: load now if somehow missing
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

    # Build (sheet_index, row_filter, group_prefix, table_label) tuples
    # group_prefix is used as the Excel sheet name in per_question mode
    # so all sub-questions in a group land on the same sheet
    tc = st.session_state["table_configs"]
    sheet_table_configs = []
    for grp in st.session_state["q_groups"]:
        prefix  = grp["prefix"]
        configs = tc.get(prefix, [{"label":"All rows","rows":"all"}])
        for cfg in configs:
            for m in grp["sheets"]:
                if m["index"] not in st.session_state.get("selected_indices", []):
                    continue
                sheet_table_configs.append((
                    m["index"],
                    cfg["rows"],
                    prefix,
                    cfg["label"],
                ))

    n_tables = len(sheet_table_configs)
    st.info(f"Ready to generate **{n_tables} table(s)** with **{len(selected_cols)} column(s)**.")

    if st.button("🚀 Generate Document", type="primary"):
        bar = st.progress(0, text="Starting…")

        def upd(pct, msg):
            bar.progress(min(float(pct), 1.0), text=msg)

        with st.spinner(""):
            results = generate_outputs(
                file_bytes             = st.session_state["file_bytes"],
                sheet_table_configs    = sheet_table_configs,
                desired_groups         = selected_cols,
                output_format          = output_format,
                excel_mode             = excel_mode,
                portrait_landscape     = portrait_landscape,
                weighted_data          = weighted_data,
                use_weighted_base      = use_weighted_base,
                progress_callback      = upd,
                show_sig_flags         = show_sig_flags,
                heatmap_scheme         = heatmap_scheme,
                heatmap_custom_start   = heatmap_custom_start,
                heatmap_custom_end     = heatmap_custom_end,
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
