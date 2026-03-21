"""
engine.py — Banner Formatter core logic
All parsing, detection, and output writing lives here.
The Streamlit app (app.py) calls these functions directly.
"""

import io
import math
import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
import openpyxl
from openpyxl.styles import Font as XLFont, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ── Helper functions ─────────────────────────────────────────

def lowerList(inputList):
    return [x.lower() if isinstance(x, str) else x for x in inputList]

def normal_round(n):
    if n - math.floor(n) < 0.5:
        return math.floor(n)
    return math.ceil(n)

def prettyPrint(number):
    if isinstance(number, str):
        return number
    try:
        if np.isnan(number):
            return ""
    except Exception:
        pass
    if isinstance(number, (int, float)):
        newNumber = str(normal_round(number))
        if len(newNumber) > 3:
            newNumber = newNumber[:-3] + ',' + newNumber[-3:]
        return newNumber
    return number

def removeNaN(theList):
    return [x for x in theList if not (isinstance(x, float) and math.isnan(x))]

def transposeList(theList):
    return pd.DataFrame(theList).T.values.tolist()

def sum2(listSum):
    total = 0
    for i in listSum:
        try:
            total += float(i)
        except Exception:
            pass
    return total


# ── Format detection ─────────────────────────────────────────

def detect_format(sheet_df):
    """
    Returns 'fmt2', 'fmt3', 'fmt4', or 'fmt1' for each sheet.
    fmt2 = standard question (countries as columns)
    fmt3 = top box summary (companies as rows, countries as columns, floating base)
    fmt4 = grid table (brands as columns, scale options as rows)
    fmt1 = original transposed format (legacy / per-country)
    """
    raw = sheet_df.values.tolist()

    # fmt4: brands as columns — row 3 col 1 blank, row 4 col 1 is a brand name not 'Total'
    try:
        row3_col1 = raw[3][1]
        row4_col1 = raw[4][1]
        if (
            isinstance(row4_col1, str) and row4_col1.strip()
            and 'total' not in str(row4_col1).lower()
            and (
                row3_col1 is None
                or not isinstance(row3_col1, str)
                or row3_col1.strip() == ''
                or (isinstance(row3_col1, float) and math.isnan(row3_col1))
            )
        ):
            return 'fmt4'
    except (IndexError, TypeError):
        pass

    # fmt2 / fmt3: row 3 col 1 = 'Total', row 4 = sub-labels
    try:
        row2_col0 = raw[2][0]
        row3_col1 = raw[3][1]
        if (
            isinstance(row2_col0, str) and len(row2_col0.strip()) > 10
            and isinstance(row3_col1, str) and 'total' in row3_col1.lower()
        ):
            row7      = raw[7]
            row7_col0 = row7[0] if row7 else None
            row7_col1 = row7[1] if len(row7) > 1 else None
            if isinstance(row7_col0, str) and row7_col0.strip().lower().startswith('base'):
                if row7_col1 is None or (isinstance(row7_col1, float) and math.isnan(row7_col1)):
                    return 'fmt3'
                return 'fmt2'
    except (IndexError, TypeError):
        pass

    return 'fmt1'


# ── Sheet scanning ───────────────────────────────────────────

def scan_file(file_bytes):
    """
    Read an uploaded Excel file and return metadata for every sheet.
    Returns list of dicts: {index, name, fmt, question_wording, columns}
    """
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    results = []
    for i, sheet_name in enumerate(xl.sheet_names):
        if i == 0:
            continue   # skip table of contents
        try:
            df  = xl.parse(i, header=None, na_values=[''])
            fmt = detect_format(df)
            raw = df.values.tolist()
            question_wording = str(raw[2][0]).strip() if (
                len(raw) > 2 and raw[2] and raw[2][0]
            ) else sheet_name

            # Collect available columns
            columns = []
            if fmt in ('fmt2', 'fmt3'):
                group_row    = raw[3] if len(raw) > 3 else []
                sublabel_row = raw[4] if len(raw) > 4 else []
                for j in range(1, len(group_row)):
                    g = group_row[j]
                    s = sublabel_row[j] if j < len(sublabel_row) else ''
                    if isinstance(g, str) and g.strip():
                        columns.append((g.strip(), s.strip() if isinstance(s, str) else ''))
            elif fmt == 'fmt4':
                brand_row = raw[4] if len(raw) > 4 else []
                for j in range(1, len(brand_row)):
                    v = brand_row[j]
                    if isinstance(v, str) and v.strip():
                        columns.append((v.strip(), ''))

            results.append({
                'index':            i,
                'sheet_name':       sheet_name,
                'fmt':              fmt,
                'question_wording': question_wording,
                'columns':          columns,
            })
        except Exception as e:
            results.append({
                'index':            i,
                'sheet_name':       sheet_name,
                'fmt':              'error',
                'question_wording': f'Error reading sheet: {e}',
                'columns':          [],
            })
    return results


def get_all_columns(sheet_metas):
    """
    Return deduplicated list of (group, sublabel) column tuples
    from fmt2/fmt3 sheets only (not fmt4 — those use brand columns).
    """
    seen = set()
    cols = []
    for s in sheet_metas:
        if s['fmt'] in ('fmt2', 'fmt3'):
            for col in s['columns']:
                if col[0] not in seen:
                    seen.add(col[0])
                    cols.append(col)
    return cols


def scan_rows_for_sheets(file_bytes, sheet_indices):
    """
    Return deduplicated ordered list of answer/row labels
    across all given sheet indices. Used to populate row checkboxes.
    """
    xl   = pd.ExcelFile(io.BytesIO(file_bytes))
    seen = set()
    rows = []
    for i in sheet_indices:
        try:
            df  = xl.parse(i, header=None, na_values=[''])
            fmt = detect_format(df)
            raw = df.values.tolist()
            if fmt == 'fmt2':
                # base row index
                base_row_idx = None
                for ri, row in enumerate(raw):
                    if isinstance(row[0], str) and row[0].strip().lower().startswith('base'):
                        base_row_idx = ri
                        break
                start = (base_row_idx + 2) if base_row_idx is not None else 9
                j = start
                while j < len(raw):
                    label = raw[j][0]
                    if isinstance(label, str) and label.strip():
                        clean = label.strip()
                        if clean.lower() == 'sigma':
                            break
                        if clean not in seen:
                            seen.add(clean)
                            rows.append(clean)
                    j += 3
            elif fmt == 'fmt3':
                end_markers = {'overlap formula used', 'sigma'}
                j = 8
                while j < len(raw):
                    label = raw[j][0]
                    if isinstance(label, str) and label.strip():
                        clean = label.strip()
                        if clean.lower() in end_markers:
                            break
                        if clean not in seen:
                            seen.add(clean)
                            rows.append(clean)
                    j += 3
            elif fmt == 'fmt4':
                j = 9
                while j < len(raw):
                    label = raw[j][0]
                    if isinstance(label, str) and label.strip():
                        clean = label.strip()
                        if clean.lower() == 'sigma':
                            break
                        if clean not in seen:
                            seen.add(clean)
                            rows.append(clean)
                    j += 3
        except Exception:
            pass
    return rows


def get_question_groups(sheet_metas):
    """
    Group sheets by question prefix (A0, A1, S1 etc).
    Returns ordered list of dicts:
      { prefix, sheets: [meta, ...], fmt }
    fmt is taken from first sheet in group.
    """
    import re
    groups = {}
    order  = []
    for m in sheet_metas:
        if m['fmt'] == 'error':
            continue
        # Extract prefix: leading letters + digits e.g. "A0", "S1", "A10"
        match = re.match(r'^([A-Za-z]+\d+)', m['question_wording'].strip())
        if match:
            prefix = match.group(1).upper()
        else:
            prefix = m['sheet_name'][:6]
        if prefix not in groups:
            groups[prefix] = {'prefix': prefix, 'sheets': [], 'fmt': m['fmt']}
            order.append(prefix)
        groups[prefix]['sheets'].append(m)
    return [groups[p] for p in order]


# ── Parsers ──────────────────────────────────────────────────

def _select_cols(all_cols, desired_groups, sublabel_filter='total'):
    """Filter column list to desired_groups (list of group name strings)."""
    if desired_groups is not None:
        desired_lower = {d.lower() for d in desired_groups}
        return [(j, g, s) for (j, g, s) in all_cols if g.lower() in desired_lower]
    # Auto: take columns whose sub_label is 'Total' + first column
    selected = [(j, g, s) for (j, g, s) in all_cols if s.lower() == sublabel_filter]
    if all_cols and all_cols[0] not in selected:
        selected.insert(0, all_cols[0])
    return selected


def parse_fmt2_sheet(sheet_df, desired_groups=None, weighted_data=False):
    raw = sheet_df.values.tolist()
    question_wording = str(raw[2][0]).strip() if raw[2][0] else ''

    group_row    = raw[3]
    sublabel_row = raw[4]
    all_cols     = []
    for j in range(1, len(group_row)):
        g = group_row[j]
        s = sublabel_row[j]
        if isinstance(g, str) and g.strip():
            all_cols.append((j, g.strip(), s.strip() if isinstance(s, str) else ''))

    selected_cols = _select_cols(all_cols, desired_groups)
    col_indices   = [j for (j, g, s) in selected_cols]
    col_labels    = [(g, s) for (j, g, s) in selected_cols]

    # Base row
    base_row_idx = None
    for i, row in enumerate(raw):
        if isinstance(row[0], str) and row[0].strip().lower().startswith('base'):
            base_row_idx = i
            break

    base_offset   = 1 if weighted_data else 0
    base_values   = []
    if base_row_idx is not None:
        base_data_row = raw[base_row_idx + base_offset]
        base_values   = [base_data_row[j] if j < len(base_data_row) else None for j in col_indices]

    # Answers + data
    answers    = []
    data       = []
    base_rows_used = 2 if weighted_data else 1
    data_start = (base_row_idx + base_rows_used + 1) if base_row_idx is not None else 9

    i = data_start
    while i < len(raw):
        label = raw[i][0]
        if isinstance(label, str) and label.strip():
            label_clean = label.strip()
            if label_clean.lower() == 'sigma':
                break
            answers.append(label_clean)
            pct_row  = raw[i + 1] if i + 1 < len(raw) else []
            row_vals = [pct_row[j] if j < len(pct_row) else None for j in col_indices]
            data.append(row_vals)
        i += 3

    return {
        'question_wording': question_wording,
        'columns':          col_labels,
        'base_values':      base_values,
        'answers':          answers,
        'data':             data,
    }


def parse_fmt3_sheet(sheet_df, desired_groups=None):
    raw = sheet_df.values.tolist()
    question_wording = str(raw[2][0]).strip() if raw[2][0] else ''

    group_row    = raw[3]
    sublabel_row = raw[4]
    all_cols     = []
    for j in range(1, len(group_row)):
        g = group_row[j]
        s = sublabel_row[j]
        if isinstance(g, str) and g.strip():
            all_cols.append((j, g.strip(), s.strip() if isinstance(s, str) else ''))

    selected_cols = _select_cols(all_cols, desired_groups)
    col_indices   = [j for (j, g, s) in selected_cols]
    col_labels    = [(g, s) for (j, g, s) in selected_cols]

    end_markers = {'overlap formula used', 'sigma'}

    def get_val(row, j):
        v = row[j] if j < len(row) else None
        if v is None:
            return None
        if isinstance(v, str) and v.strip() in ('-', '', '\xa0'):
            return None
        if isinstance(v, float) and math.isnan(v):
            return None
        return v

    companies     = []
    company_bases = []
    data          = []

    i = 8
    while i < len(raw):
        label = raw[i][0]
        if not isinstance(label, str) or not label.strip():
            i += 1
            continue
        label_clean = label.strip()
        if label_clean.lower() in end_markers:
            break
        companies.append(label_clean)
        base_row = raw[i]
        pct_row  = raw[i + 1] if i + 1 < len(raw) else []
        company_bases.append([get_val(base_row, j) for j in col_indices])
        data.append([get_val(pct_row,  j) for j in col_indices])
        i += 3

    return {
        'question_wording': question_wording,
        'columns':          col_labels,
        'base_values':      None,
        'company_bases':    company_bases,
        'answers':          companies,
        'data':             data,
    }


def parse_fmt4_sheet(sheet_df):
    raw = sheet_df.values.tolist()
    question_wording = str(raw[2][0]).strip() if raw[2][0] else ''

    brand_row     = raw[4]
    brands        = []
    brand_indices = []
    for j in range(1, len(brand_row)):
        v = brand_row[j]
        if isinstance(v, str) and v.strip():
            brands.append(v.strip())
            brand_indices.append(j)

    base_row    = raw[7]
    base_values = [base_row[j] if j < len(base_row) else None for j in brand_indices]

    answers = []
    data    = []
    i = 9
    while i < len(raw):
        label = raw[i][0]
        if isinstance(label, str) and label.strip():
            label_clean = label.strip()
            if label_clean.lower() == 'sigma':
                break
            answers.append(label_clean)
            pct_row  = raw[i + 1] if i + 1 < len(raw) else []
            row_vals = []
            for j in brand_indices:
                v = pct_row[j] if j < len(pct_row) else None
                if v is None or (isinstance(v, float) and math.isnan(v)):
                    row_vals.append(None)
                elif isinstance(v, str) and v.strip() in ('-', '', '\xa0'):
                    row_vals.append(None)
                else:
                    row_vals.append(v)
            data.append(row_vals)
        i += 3

    return {
        'question_wording': question_wording,
        'brands':           brands,
        'base_values':      base_values,
        'answers':          answers,
        'data':             data,
    }


# ── Word output ──────────────────────────────────────────────

def _get_word_template(portrait_landscape=False):
    import os
    template_file = 'template_portrait.docx' if portrait_landscape else 'template_landscape.docx'
    if os.path.exists(template_file):
        return Document(template_file)
    # Fallback to blank document if template not found
    return Document()


def write_table_to_doc(doc, question_wording, col_labels, base_values,
                       answers, data, multiple, is_first):
    if not is_first:
        doc.add_paragraph()
    q_para       = doc.add_paragraph()
    q_para.style = doc.styles['Normal']
    q_para.text  = question_wording

    n_cols = len(col_labels)
    table  = doc.add_table(rows=1, cols=n_cols + 1)
    table.style     = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr = table.rows[0].cells
    hdr[0].text = ''
    for i in range(n_cols):
        base_str = prettyPrint(base_values[i]) if i < len(base_values) else ''
        hdr[i + 1].text = ''
        hdr[i + 1].paragraphs[0].add_run(
            f"{col_labels[i]}\n(N={base_str})"
        ).bold = True
        hdr[i + 1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for r_idx, answer in enumerate(answers):
        if str(answer).strip().lower() == 'sigma':
            continue
        row_cells = table.add_row().cells
        row_cells[0].text = str(answer)
        row_vals  = data[r_idx] if r_idx < len(data) else []
        for c_idx in range(n_cols):
            val  = row_vals[c_idx] if c_idx < len(row_vals) else None
            cell = row_cells[c_idx + 1]
            cell.text = ''
            if val is None or (isinstance(val, float) and math.isnan(val)):
                pass
            elif isinstance(val, str):
                cell.paragraphs[0].add_run(val)
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                cell.paragraphs[0].add_run(
                    f"{normal_round(round(val * multiple, 3))}%"
                )
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for j in range(n_cols + 1):
        for cell in table.columns[j].cells:
            cell.width = Inches(3) if j == 0 else Inches(1.75)
            if j > 0:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


# ── Excel output ─────────────────────────────────────────────

HDR_FILL  = PatternFill("solid", fgColor="1F4E79")
HDR_FONT  = XLFont(bold=True, color="FFFFFF", name="Arial", size=10)
BODY_FONT = XLFont(name="Arial", size=10)
CTR       = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT      = Alignment(horizontal="left",   vertical="center", wrap_text=True)
THIN      = Side(style="thin", color="AAAAAA")
BORDER    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _xl_write_table(ws, start_row, question_wording, col_headers,
                    base_values, answers, data, multiple=100,
                    show_base_in_header=True):
    ws.cell(row=start_row, column=1, value=question_wording)
    ws.cell(row=start_row, column=1).font = XLFont(bold=True, name="Arial", size=10)
    ws.merge_cells(start_row=start_row, start_column=1,
                   end_row=start_row, end_column=len(col_headers) + 1)
    start_row += 1

    ws.cell(row=start_row, column=1, value='').fill   = HDR_FILL
    ws.cell(row=start_row, column=1).border = BORDER
    for ci, label in enumerate(col_headers):
        n_str = ''
        if show_base_in_header and ci < len(base_values) and base_values[ci] is not None:
            n_str = f"\n(N={prettyPrint(base_values[ci])})"
        cell            = ws.cell(row=start_row, column=ci + 2, value=label + n_str)
        cell.font       = HDR_FONT
        cell.fill       = HDR_FILL
        cell.alignment  = CTR
        cell.border     = BORDER
    start_row += 1

    for ri, answer in enumerate(answers):
        if str(answer).strip().lower() == 'sigma':
            continue
        row_vals   = data[ri] if ri < len(data) else []
        label_cell = ws.cell(row=start_row, column=1, value=str(answer))
        label_cell.font      = BODY_FONT
        label_cell.alignment = LEFT
        label_cell.border    = BORDER
        for ci in range(len(col_headers)):
            val  = row_vals[ci] if ci < len(row_vals) else None
            cell = ws.cell(row=start_row, column=ci + 2)
            cell.font      = BODY_FONT
            cell.alignment = CTR
            cell.border    = BORDER
            if val is None or (isinstance(val, float) and math.isnan(val)):
                cell.value = ''
            elif isinstance(val, str):
                cell.value = val
            else:
                cell.value = f"{normal_round(round(val * multiple, 3))}%"
        start_row += 1

    ws.column_dimensions['A'].width = max(ws.column_dimensions['A'].width, 35)
    for ci in range(len(col_headers)):
        col_letter = get_column_letter(ci + 2)
        ws.column_dimensions[col_letter].width = max(
            ws.column_dimensions[col_letter].width, 14)

    return start_row + 1


# ── Main generation function ─────────────────────────────────

def generate_outputs(
    file_bytes,
    sheet_table_configs,   # list of (sheet_index, row_filter)
                           # row_filter: 'all' or list of row label strings
    desired_groups,
    output_format,
    excel_mode,
    portrait_landscape,
    weighted_data,
    progress_callback=None,
):
    """
    Core generation function called by the Streamlit app.
    sheet_table_configs: list of (sheet_index, row_filter) tuples.
    The same sheet_index can appear multiple times (one per table config).
    """
    xl      = pd.ExcelFile(io.BytesIO(file_bytes))
    total   = max(len(sheet_table_configs), 1)
    skipped = []
    errors  = []

    # Cache parsed sheets so we don't re-parse the same sheet multiple times
    parsed_cache = {}

    # Word setup
    word_doc = None
    if output_format in ('word', 'both'):
        try:
            word_doc = _get_word_template(portrait_landscape)
        except Exception:
            word_doc = Document()
        style = word_doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(10)

    # Excel setup
    xl_wb         = None
    xl_sheet_ctr  = [0]
    xl_sheet_rows = {}

    if output_format in ('excel', 'both'):
        xl_wb = openpyxl.Workbook()
        xl_wb.remove(xl_wb.active)

    def _xl_get_sheet(label):
        name = label[:31]
        if excel_mode == 'per_question':
            if name not in xl_wb.sheetnames:
                xl_wb.create_sheet(title=name)
                xl_sheet_rows[name] = 1
            return xl_wb[name], xl_sheet_rows[name]
        else:
            xl_sheet_ctr[0] += 1
            sname = f"{xl_sheet_ctr[0]:03d}_{name}"[:31]
            return xl_wb.create_sheet(title=sname), 1

    def _xl_done(ws, label, next_row):
        if excel_mode == 'per_question':
            xl_sheet_rows[label[:31]] = next_row

    def write_xl(question_wording, col_headers, base_values, answers,
                 data, multiple, sheet_label, show_base=True):
        ws, row = _xl_get_sheet(sheet_label)
        nr = _xl_write_table(ws, row, question_wording, col_headers,
                             base_values, answers, data, multiple, show_base)
        _xl_done(ws, sheet_label, nr)

    def apply_row_filter(answers, data, row_filter):
        if row_filter == 'all' or not row_filter:
            return answers, data
        custom = {r.strip().lower() for r in row_filter}
        fa, fd = [], []
        for a, d in zip(answers, data):
            if a.strip().lower() in custom:
                fa.append(a)
                fd.append(d)
        return fa, fd

    first_word = True

    for idx, (sheet_idx, row_filter) in enumerate(sheet_table_configs):
        if progress_callback:
            progress_callback(idx / total, f"Processing sheet {sheet_idx}…")

        try:
            sheet_name = xl.sheet_names[sheet_idx]
            sheet_label = sheet_name[:31]

            # Parse once per sheet, cache result
            if sheet_idx not in parsed_cache:
                df  = xl.parse(sheet_idx, header=None, na_values=[''])
                fmt = detect_format(df)
                parsed_cache[sheet_idx] = (df, fmt)
            sheet_df, fmt = parsed_cache[sheet_idx]

            # ── fmt2 ─────────────────────────────────────────
            if fmt == 'fmt2':
                parsed  = parse_fmt2_sheet(sheet_df, desired_groups, weighted_data)
                if not parsed['answers']:
                    skipped.append(sheet_label)
                    continue

                col_labels = [g for (g, s) in parsed['columns']]
                answers, data = apply_row_filter(parsed["answers"], parsed["data"], row_filter)

                if not answers:
                    skipped.append(sheet_label)
                    continue

                if word_doc is not None:
                    write_table_to_doc(
                        word_doc, parsed['question_wording'],
                        col_labels, parsed['base_values'],
                        answers, data, 100, first_word,
                    )
                    first_word = False

                if xl_wb is not None:
                    write_xl(parsed['question_wording'], col_labels,
                             parsed['base_values'], answers, data, 100, sheet_label)

            # ── fmt3 ─────────────────────────────────────────
            elif fmt == 'fmt3':
                parsed     = parse_fmt3_sheet(sheet_df, desired_groups)
                if not parsed['answers']:
                    skipped.append(sheet_label)
                    continue

                col_labels    = [g for (g, s) in parsed['columns']]
                companies     = parsed['answers']
                company_bases = parsed['company_bases']
                pct_data      = parsed['data']
                n_cols        = len(col_labels)

                # Apply row filter to companies
                if row_filter not in ('all', None):
                    companies, pct_data, company_bases = zip(
                        *[(c, d, b) for c, d, b in zip(companies, pct_data, company_bases)
                          if row_filter == 'all' or c.strip().lower() in {r.lower() for r in (row_filter if isinstance(row_filter, list) else [])}]
                    ) if companies else ([], [], [])
                    companies     = list(companies)
                    pct_data      = list(pct_data)
                    company_bases = list(company_bases)

                if not companies:
                    skipped.append(sheet_label)
                    continue

                def _write_fmt3_word(show_n, is_first_arg, suffix=''):
                    if word_doc is None:
                        return
                    if not is_first_arg:
                        word_doc.add_paragraph()
                    qp       = word_doc.add_paragraph()
                    qp.style = word_doc.styles['Normal']
                    qp.text  = parsed['question_wording'] + suffix
                    tbl = word_doc.add_table(rows=1, cols=n_cols + 1)
                    tbl.style     = 'Table Grid'
                    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                    hdr = tbl.rows[0].cells
                    hdr[0].text = ''
                    for ci in range(n_cols):
                        hdr[ci+1].text = ''
                        hdr[ci+1].paragraphs[0].add_run(col_labels[ci]).bold = True
                        hdr[ci+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for ri, company in enumerate(companies):
                        rc   = tbl.add_row().cells
                        rc[0].text = company
                        bases = company_bases[ri] if ri < len(company_bases) else []
                        pcts  = pct_data[ri]      if ri < len(pct_data)      else []
                        for ci in range(n_cols):
                            pct  = pcts[ci]  if ci < len(pcts)  else None
                            n    = bases[ci] if ci < len(bases) else None
                            cell = rc[ci + 1]
                            cell.text = ''
                            if pct is None or isinstance(pct, str):
                                cell.paragraphs[0].add_run('-')
                            else:
                                pct_str = f"{normal_round(round(pct * 100, 3))}%"
                                n_str   = f"\n(N={prettyPrint(n)})" if (show_n and n) else ''
                                cell.paragraphs[0].add_run(pct_str + n_str)
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    for ci in range(n_cols + 1):
                        for cell in tbl.columns[ci].cells:
                            cell.width = Inches(3) if ci == 0 else Inches(1.75)

                _write_fmt3_word(show_n=False, is_first_arg=first_word)
                _write_fmt3_word(show_n=True,  is_first_arg=False, suffix=' (with base)')
                first_word = False

                if xl_wb is not None:
                    write_xl(parsed['question_wording'], col_labels,
                             [None]*n_cols, companies, pct_data,
                             100, sheet_label, show_base=False)
                    # second table with N
                    data_with_n = []
                    for ri, company in enumerate(companies):
                        bases = company_bases[ri] if ri < len(company_bases) else []
                        pcts  = pct_data[ri]
                        row   = []
                        for ci in range(n_cols):
                            pct = pcts[ci]  if ci < len(pcts)  else None
                            n   = bases[ci] if ci < len(bases) else None
                            if pct is None or isinstance(pct, str):
                                row.append(None)
                            else:
                                pct_str = f"{normal_round(round(pct * 100, 3))}%"
                                n_str   = f" (N={prettyPrint(n)})" if n else ''
                                row.append(pct_str + n_str)
                        data_with_n.append(row)
                    write_xl(parsed['question_wording'] + ' (with base)',
                             col_labels, [None]*n_cols, companies,
                             data_with_n, 1, sheet_label, show_base=False)

            # ── fmt4 ─────────────────────────────────────────
            elif fmt == 'fmt4':
                parsed  = parse_fmt4_sheet(sheet_df)
                if not parsed['answers']:
                    skipped.append(sheet_label)
                    continue

                answers, data = apply_row_filter(parsed["answers"], parsed["data"], row_filter)
                if not answers:
                    skipped.append(sheet_label)
                    continue

                brands      = parsed['brands']
                base_values = parsed['base_values']
                n_brands    = len(brands)

                if word_doc is not None:
                    if not first_word:
                        word_doc.add_paragraph()
                    qp       = word_doc.add_paragraph()
                    qp.style = word_doc.styles['Normal']
                    qp.text  = parsed['question_wording']
                    tbl = word_doc.add_table(rows=1, cols=n_brands + 1)
                    tbl.style     = 'Table Grid'
                    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                    hdr = tbl.rows[0].cells
                    hdr[0].text = ''
                    for ci, brand in enumerate(brands):
                        n = base_values[ci] if ci < len(base_values) else None
                        hdr[ci+1].text = ''
                        hdr[ci+1].paragraphs[0].add_run(
                            f"{brand}\n(N={prettyPrint(n)})"
                        ).bold = True
                        hdr[ci+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for ri, answer in enumerate(answers):
                        rc = tbl.add_row().cells
                        rc[0].text = answer
                        row_vals = data[ri] if ri < len(data) else []
                        for ci in range(n_brands):
                            val  = row_vals[ci] if ci < len(row_vals) else None
                            cell = rc[ci + 1]
                            cell.text = ''
                            if val is None:
                                cell.paragraphs[0].add_run('-')
                            else:
                                cell.paragraphs[0].add_run(
                                    f"{normal_round(round(val * 100, 3))}%"
                                )
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    for ci in range(n_brands + 1):
                        for cell in tbl.columns[ci].cells:
                            cell.width = Inches(2) if ci == 0 else Inches(1.25)
                    first_word = False

                if xl_wb is not None:
                    write_xl(parsed['question_wording'], brands,
                             base_values, answers, data, 100, sheet_label)

            else:
                skipped.append(f"{sheet_label} (fmt1/grid — skipped)")

        except Exception as e:
            errors.append((xl.sheet_names[sheet_idx], str(e)))

    if progress_callback:
        progress_callback(1.0, "Finalizing…")

    # Serialize outputs
    word_bytes  = None
    excel_bytes = None

    if word_doc is not None and output_format in ('word', 'both'):
        buf = io.BytesIO()
        word_doc.save(buf)
        word_bytes = buf.getvalue()

    if xl_wb is not None and output_format in ('excel', 'both'):
        buf = io.BytesIO()
        xl_wb.save(buf)
        excel_bytes = buf.getvalue()

    return {
        'word_bytes':  word_bytes,
        'excel_bytes': excel_bytes,
        'skipped':     skipped,
        'errors':      errors,
    }
