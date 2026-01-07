from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from typing import Dict, List, Any, Optional
from concurrent.futures import ThreadPoolExecutor

MAX_DYNAMIC_COLS = 16
STAFF_FONT_SIZE = Pt(12)
DEFAULT_FONT_SIZE = Pt(10)
HEADER_FONT_SIZE = Pt(9.5)

SERIAL_COL_WIDTH = Cm(0.9)
STAFF_COL_WIDTH = Cm(5.0)
MONTH_COL_WIDTH = Cm(1.3)
TOTAL_COL_WIDTH = Cm(1.6)    


def set_cell_shading(cell, color_hex: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), color_hex.replace("#", ""))
    tcPr.append(shd)

def set_cell_text_bold(cell):
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.bold = True

def set_cell_font(cell, size: Pt, bold=False, name="Calibri"):
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.size = size
            r.font.bold = bold
            r.font.name = name

def align_cell(cell, align='center'):
    alignment_map = {
        'left': WD_ALIGN_PARAGRAPH.LEFT,
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'right': WD_ALIGN_PARAGRAPH.RIGHT
    }
    for p in cell.paragraphs:
        p.alignment = alignment_map.get(align, WD_ALIGN_PARAGRAPH.CENTER)

def set_cell_vertical_alignment(cell, valign='center'):
    """Set vertical alignment for cell"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = OxmlElement('w:vAlign')
    vAlign.set(qn('w:val'), valign)
    tcPr.append(vAlign)

def set_cell_width(cell, width_cm):
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(width_cm.cm * 567)))  # Convert cm to twentieths of a point
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

def set_table_width_and_alignment(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    
    # Set table alignment to center
    tblJc = OxmlElement('w:jc')
    tblJc.set(qn('w:val'), 'center')
    tblPr.append(tblJc)
    
    # Set fixed table layout to respect column widths
    tblLayout = OxmlElement('w:tblLayout')
    tblLayout.set(qn('w:type'), 'fixed')
    tblPr.append(tblLayout)

def enforce_column_widths(table, num_static_cols=2, num_month_cols=0, has_total=False):
    try:
        if len(table.columns) > 1:
            table.columns[0].width = SERIAL_COL_WIDTH
            table.columns[1].width = STAFF_COL_WIDTH
            
            for i in range(2, len(table.columns)):
                if has_total and i == len(table.columns) - 1:
                    table.columns[i].width = TOTAL_COL_WIDTH
                else:
                    table.columns[i].width = MONTH_COL_WIDTH
    except:
        pass
    
    for row_idx, row in enumerate(table.rows):
        for col_idx, cell in enumerate(row.cells):
            if col_idx == 0:
                set_cell_width(cell, SERIAL_COL_WIDTH)
            elif col_idx == 1:
                set_cell_width(cell, STAFF_COL_WIDTH)
            elif has_total and col_idx == len(row.cells) - 1:
                set_cell_width(cell, TOTAL_COL_WIDTH)
            else:
                set_cell_width(cell, MONTH_COL_WIDTH)

def add_row_with_color(table, values, color=None):
    row = table.add_row().cells
    for i, v in enumerate(values):
        row[i].text = str(v) if v is not None else ""
        if color:
            set_cell_shading(row[i], color)

def add_table_spacing(doc, space_before_pt=12, space_after_pt=12):
    p = doc.add_paragraph()
    p.space_before = Pt(space_before_pt)
    p.space_after = Pt(0)
    
    return p  

def _generate_legend_table(doc, legend, color_map):
    if not legend:
        return None
    
    legend_table = doc.add_table(rows=0, cols=1)
    legend_table.style = "Table Grid"
    set_table_width_and_alignment(legend_table)
    
    # Set column width for legend
    try:
        legend_table.columns[0].width = Cm(4.0)  # Phase name column
    except:
        pass
    
    # Add legend rows
    for legend_item in legend:
        phase_name = legend_item.get('phase', '')
        phase_color = legend_item.get('color', color_map.get(phase_name, '#FFFFFF'))
        
        row = legend_table.add_row()
        cell = row.cells[0]
        
        cell.text = phase_name
        if phase_color:
            set_cell_shading(cell, phase_color)
        set_cell_font(cell, DEFAULT_FONT_SIZE, True)
        align_cell(cell, 'center')
    
    return legend_table

# ================= SDT =================

def find_sdt_by_title(doc: Document, title: str):
    for el in doc.element.body.iter():
        if el.tag.endswith("sdt"):
            for child in el.iter():
                if child.tag.endswith("alias") and child.get(qn("w:val")) == title:
                    return el
                if child.tag.endswith("tag") and child.get(qn("w:val")) == title:
                    return el
    return None


def parse_color_mapping(colors):
    cmap = {}
    for c in colors or []:
        for k, v in c.items():
            cmap[k] = v.get("color") if isinstance(v, dict) else v
    return cmap
def detect_table_format(rows, headers, colors):
    if not rows:
        return 'simple'
    if 'phaseType' in rows[0] or 'phase' in rows[0]:
        return 'phase_based'
    if headers:
        return 'custom'
    return 'simple'


def generate_dynamic_table(doc: Document, tag: str, data: Dict[str, Any]) -> Document:
    
    headers = data.get("headers", [])
    rows = data.get("rows", [])
    colors = data.get("colors", [])
    legend = data.get("legend", [])
    header_color = data.get("headerColor", "#333399")

    if not rows:
        return doc

    color_map = parse_color_mapping(colors)
    table_format = detect_table_format(rows, headers, colors)

    sdt = find_sdt_by_title(doc, tag)
    if sdt is None:
        return doc

    sdt_content = next(el for el in sdt.iter() if el.tag.endswith("sdtContent"))
    
    for child in list(sdt_content):
        sdt_content.remove(child)

    if table_format == 'phase_based':
        table = _generate_phase_based_table(doc, rows, color_map, header_color)
        sdt_content.append(table._tbl)

    elif table_format == 'custom':
        tables = _generate_custom_table(doc, headers, rows, color_map, header_color, legend)
        
        for i, table in enumerate(tables):
            if i > 0: 
                spacing_p = doc.add_paragraph()
                spacing_p.space_before = Pt(12)
                spacing_p.space_after = Pt(6)
                sdt_content.append(spacing_p._element)
            
            sdt_content.append(table._tbl)

    else:
        table = _generate_simple_table(doc, rows, color_map, header_color)
        sdt_content.append(table._tbl)

    if legend:
        legend_spacing_p = doc.add_paragraph()
        legend_spacing_p.space_before = Pt(18)
        legend_spacing_p.space_after = Pt(6)
        sdt_content.append(legend_spacing_p._element)
        
        legend_table = _generate_legend_table(doc, legend, color_map)
        if legend_table:
            sdt_content.append(legend_table._tbl)

    return doc


def _generate_phase_based_table(doc, rows, color_map, header_color):
    headers = [k for k in rows[0] if k not in ('phaseType', 'phase')]
    table = doc.add_table(rows=0, cols=len(headers))
    table.style = "Table Grid"
    
    set_table_width_and_alignment(table)
    
    header_row = table.add_row()
    header_cells = header_row.cells
    for idx, header in enumerate(headers):
        cell = header_cells[idx]
        cell.text = str(header)
        set_cell_shading(cell, header_color)
        set_cell_text_bold(cell)
        set_cell_font(cell, HEADER_FONT_SIZE, True)
        
        if idx == 0:
            align_cell(cell, 'center')  
        elif idx == 1:
            align_cell(cell, 'left')    
        else:
            align_cell(cell, 'center')  
    
    # Add data rows
    for r in rows:
        data_row = table.add_row()
        data_cells = data_row.cells
        phase = r.get('phaseType') or r.get('phase')
        
        for idx, h in enumerate(headers):
            cell = data_cells[idx]
            cell.text = str(r.get(h, ''))
            
            if idx == 0:  # S.No.
                set_cell_font(cell, DEFAULT_FONT_SIZE, False)
                align_cell(cell, 'center')
            elif idx == 1:  # Staff
                set_cell_font(cell, STAFF_FONT_SIZE, False)
                align_cell(cell, 'left')
            else:  # Other columns
                set_cell_font(cell, DEFAULT_FONT_SIZE, False)
                align_cell(cell, 'center')
            
            if phase and phase in color_map:
                set_cell_shading(cell, color_map[phase])
    
    # Enforce column widths
    enforce_column_widths(table, 2, len(headers) - 2, False)
    
    return table


def _generate_simple_table(doc, rows, color_map, header_color):
    headers = list(rows[0].keys())
    table = doc.add_table(rows=0, cols=len(headers))
    table.style = "Table Grid"
    
    set_table_width_and_alignment(table)
    
    header_row = table.add_row()
    header_cells = header_row.cells
    for idx, header in enumerate(headers):
        cell = header_cells[idx]
        cell.text = str(header)
        set_cell_shading(cell, header_color)
        set_cell_text_bold(cell)
        set_cell_font(cell, HEADER_FONT_SIZE, True)
        
        if idx == 0:
            align_cell(cell, 'center')  
        elif idx == 1:
            align_cell(cell, 'left')    
        else:
            align_cell(cell, 'center')  
    
    for r in rows:
        data_row = table.add_row()
        data_cells = data_row.cells
        
        for idx, h in enumerate(headers):
            cell = data_cells[idx]
            cell.text = str(r.get(h, ''))
            
            if idx == 0: 
                set_cell_font(cell, DEFAULT_FONT_SIZE, False)
                align_cell(cell, 'center')
            elif idx == 1:  
                set_cell_font(cell, STAFF_FONT_SIZE, False)
                align_cell(cell, 'left')
            else:  
                set_cell_font(cell, DEFAULT_FONT_SIZE, False)
                align_cell(cell, 'center')
    
    enforce_column_widths(table, 2, len(headers) - 2, False)
    
    return table


def _generate_custom_table(doc, headers, rows, color_map, header_color, legend):
    month_headers = [list(m.keys())[0] for m in rows[0].get('months', [])]
    static_headers = [h for h in headers if h not in month_headers]
    has_total = 'Total' in static_headers
    tables = []
    
    if len(month_headers) > MAX_DYNAMIC_COLS:
        for chunk_idx in range(0, len(month_headers), MAX_DYNAMIC_COLS):
            chunk = month_headers[chunk_idx:chunk_idx + MAX_DYNAMIC_COLS]
            remaining_months_after_this_chunk = len(month_headers) - (chunk_idx + len(chunk))
            is_last_chunk = remaining_months_after_this_chunk == 0
            
            if has_total and is_last_chunk and len(chunk) < MAX_DYNAMIC_COLS:
                table_headers = static_headers[:2] + chunk + ['Total']
                table = _create_single_table(doc, table_headers, rows, color_map, header_color, 
                                           chunk, True, chunk_idx > 0)
            else:
                table_headers = static_headers[:2] + chunk
                table = _create_single_table(doc, table_headers, rows, color_map, header_color, 
                                           chunk, False, chunk_idx > 0)
            
            tables.append(table)
        
        if has_total and len(month_headers) % MAX_DYNAMIC_COLS == 0:
            total_table_headers = static_headers[:2] + ['Total']
            total_table = _create_single_table(doc, total_table_headers, rows, color_map, header_color, 
                                             [], True, True)
            tables.append(total_table)
    else:
        table = _create_single_table(doc, headers, rows, color_map, header_color, 
                                   month_headers, has_total, False)
        tables.append(table)
    
    return tables

def _create_single_table(doc, table_headers, rows, color_map, header_color, 
                        month_chunk, has_total_in_this_table, is_split_table, row_offset=0):
    table = doc.add_table(rows=0, cols=len(table_headers))
    table.style = "Table Grid"
    set_table_width_and_alignment(table)

    header_row = table.add_row()
    header_cells = header_row.cells

    for idx, header_text in enumerate(table_headers):
        cell = header_cells[idx]
        cell.text = str(header_text)
        set_cell_shading(cell, header_color)
        set_cell_text_bold(cell)
        set_cell_font(cell, HEADER_FONT_SIZE, True)
        set_cell_vertical_alignment(cell, 'center')

        if idx == 0:  
            align_cell(cell, 'center')
        elif idx == 1:  
            align_cell(cell, 'left')
        else:  
            align_cell(cell, 'center')

    for row_idx, row_data in enumerate(rows):
        data_row = table.add_row()
        data_cells = data_row.cells

        month_map = {}
        for month_obj in row_data.get('months', []):
            for month_name, month_info in month_obj.items():
                month_map[month_name] = month_info

        for col_idx, header in enumerate(table_headers):
            cell = data_cells[col_idx]

            if col_idx == 0:
                cell_value = str(row_idx + 1)
            elif header == 'Staff':
                cell_value = row_data.get(header, '')
            elif header == 'Total':
                cell_value = row_data.get(header, '')
            else:
                month_info = month_map.get(header, {})
                cell_value = month_info.get('value', '')

            cell.text = str(cell_value)

            if col_idx == 0:
                set_cell_font(cell, DEFAULT_FONT_SIZE, False)
                align_cell(cell, 'center')
                set_cell_vertical_alignment(cell, 'center')
            elif header == 'Staff':
                set_cell_font(cell, STAFF_FONT_SIZE, False)
                align_cell(cell, 'left')
                set_cell_vertical_alignment(cell, 'top')
            elif header == 'Total':
                set_cell_font(cell, DEFAULT_FONT_SIZE, True)
                align_cell(cell, 'center')
                set_cell_vertical_alignment(cell, 'top')
            else:
                phase = month_info.get('phase')
                if phase and phase in color_map:
                    set_cell_shading(cell, color_map[phase])
                set_cell_font(cell, DEFAULT_FONT_SIZE, False)
                align_cell(cell, 'center')
                set_cell_vertical_alignment(cell, 'top')

    month_maps = []
    for row_data in rows:
        month_map = {}
        for month_obj in row_data.get('months', []):
            for month_name, month_info in month_obj.items():
                month_map[month_name] = month_info
        month_maps.append(month_map)

    def calculate_column_total(header):
        column_total = 0.0
        for month_map in month_maps:
            month_info = month_map.get(header, {})
            try:
                value = month_info.get('value', 0)
                if isinstance(value, str):
                    value = float(value) if value.replace('.', '').replace('-', '').isdigit() else 0.0
                else:
                    value = float(value)
                column_total += value
            except (ValueError, TypeError):
                pass
        return column_total

    def calculate_grand_total():
        grand_total = 0.0
        for row_data in rows:
            try:
                total_val = row_data.get('Total', 0)
                if isinstance(total_val, str):
                    total_val = float(total_val) if total_val.replace('.', '').replace('-', '').isdigit() else 0.0
                else:
                    total_val = float(total_val)
                grand_total += total_val
            except (ValueError, TypeError):
                pass
        return grand_total

    with ThreadPoolExecutor() as executor:
        column_totals = list(executor.map(calculate_column_total, table_headers[2:]))
        grand_total = calculate_grand_total()

    totals_row = table.add_row()
    totals_cells = totals_row.cells

    for col_idx, header in enumerate(table_headers):
        cell = totals_cells[col_idx]

        if col_idx == 0: 
            cell.text = str(len(rows) + 1)
            set_cell_font(cell, DEFAULT_FONT_SIZE, False)
            align_cell(cell, 'center')
            set_cell_vertical_alignment(cell, 'center')
        elif col_idx == 1:  
            cell.text = "Total"
            set_cell_font(cell, STAFF_FONT_SIZE, True)
            align_cell(cell, 'left')
            set_cell_vertical_alignment(cell, 'center')
        elif header == 'Total':
            cell.text = str(int(grand_total)) if grand_total == int(grand_total) else str(round(grand_total, 2))
            set_cell_font(cell, DEFAULT_FONT_SIZE, True)
            align_cell(cell, 'center')
            set_cell_vertical_alignment(cell, 'center')
        else:
            column_total = column_totals[col_idx - 2]
            cell.text = str(int(column_total)) if column_total == int(column_total) else str(round(column_total, 2))
            set_cell_font(cell, DEFAULT_FONT_SIZE, True)
            align_cell(cell, 'center')
            set_cell_vertical_alignment(cell, 'center')

    num_static_cols = 2  
    num_month_cols = len(month_chunk)
    enforce_column_widths(table, num_static_cols, num_month_cols, has_total_in_this_table)

    return table
