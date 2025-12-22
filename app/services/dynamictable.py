from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from typing import Dict, List, Any, Optional

def set_cell_shading(cell, color_hex: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), color_hex.replace("#", ""))
    tcPr.append(shd)

def set_cell_text_bold(cell):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.bold = True

def add_row_with_color(table, values: List[Any], color: Optional[str] = None):
    row = table.add_row().cells
    for i, val in enumerate(values):
        row[i].text = str(val) if val is not None else ""
        if color:
            set_cell_shading(row[i], color)

def find_sdt_by_title(doc: Document, title: str):
    for el in doc.element.body.iter():
        if el.tag.endswith("sdt"):
            for child in el.iter():
                if child.tag.endswith("alias") and child.get(qn("w:val")) == title:
                    return el
                if child.tag.endswith("tag") and child.get(qn("w:val")) == title:
                    return el
    return None

def parse_color_mapping(colors: List[Dict[str, str]]) -> Dict[str, str]:
    color_map = {}
    if not colors:
        return color_map
    
    for color_obj in colors:
        for key, value in color_obj.items():
            if isinstance(value, dict) and "color" in value:
                color_map[key] = value["color"]
            else:
                color_map[key] = value
    
    return color_map

def detect_table_format(rows: List[Dict[str, Any]], headers: Optional[List[str]], colors: Optional[List[Dict[str, str]]]) -> str:
    if not rows or len(rows) == 0:
        return 'simple'
    
    first_row = rows[0]
    
    if 'phaseType' in first_row or 'phase' in first_row:
        return 'phase_based'
    
    if headers:
        return 'custom'
    
    return 'simple'

def generate_dynamic_table(doc: Document, tag: str, data: Dict[str, Any]) -> Document:
    import logging
    logger = logging.getLogger(__name__)
    
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
        logger.error(f"Content control with tag '{tag}' not found in document")
        return doc
    
    logger.info(f"Found SDT with tag: {tag}")
    
    sdt_content = None
    for el in sdt.iter():
        if el.tag.endswith("sdtContent"):
            sdt_content = el
            break
    
    if sdt_content is None:
        logger.error(f"sdtContent not found within SDT with tag '{tag}'")
        return doc
    
    for child in list(sdt_content):
        sdt_content.remove(child)
    
    if table_format == 'phase_based':
        table = _generate_phase_based_table(doc, rows, color_map, header_color)
    elif table_format == 'custom' and headers:
        table = _generate_custom_table(doc, headers, rows, color_map, header_color, legend)
    else:
        table = _generate_simple_table(doc, rows, color_map, header_color)
    
    logger.info(f"Table generated with {len(table.rows)} rows, appending to SDT content")
    sdt_content.append(table._tbl)
    logger.info("Table successfully inserted into document")
    
    return doc

def _generate_phase_based_table(doc: Document, rows: List[Dict[str, Any]], color_map: Dict[str, str], header_color: str):
    if not rows:
        return doc.add_table(rows=0, cols=0)
    
    first_row = rows[0]
    headers = [key for key in first_row.keys() if key != 'phaseType' and key != 'phase']
    
    table = doc.add_table(rows=0, cols=len(headers))
    table.style = "Table Grid"
    
    add_row_with_color(table, headers, header_color)
    
    for row_data in rows:
        phase_type = row_data.get('phaseType') or row_data.get('phase', '')
        row_color = color_map.get(phase_type, None)
        values = [row_data.get(h, '') for h in headers]
        add_row_with_color(table, values, row_color)
    
    return table

def _generate_custom_table(doc: Document, headers: List[str], rows: List[Dict[str, Any]], color_map: Dict[str, str], header_color: str, legend: List[Dict[str, str]]):
    table = doc.add_table(rows=0, cols=len(headers))
    table.style = "Table Grid"
    
    # Add header row
    header_row = table.add_row().cells
    for i, header in enumerate(headers):
        header_row[i].text = str(header)
        set_cell_shading(header_row[i], header_color)
        set_cell_text_bold(header_row[i])
    
    for row_data in rows:
        data_row = table.add_row().cells
        for i, header in enumerate(headers):
            cell_data = row_data.get(header, '')
            
            if isinstance(cell_data, dict):
                cell_value = cell_data.get('value', '')
                cell_phase = cell_data.get('phase', '')
                cell_color = color_map.get(cell_phase, None)
                
                data_row[i].text = str(cell_value)
                if cell_color:
                    set_cell_shading(data_row[i], cell_color)
            else:
                data_row[i].text = str(cell_data)
    
    # Add legend if provided
    if legend:
        # Add empty row for spacing
        spacing_row = table.add_row().cells
        for cell in spacing_row:
            cell.text = ""
        
        # Add legend header
        legend_header_row = table.add_row().cells
        legend_header_row[0].text = "Legend:"
        set_cell_text_bold(legend_header_row[0])
        
        # Merge cells for legend header
        for i in range(1, len(headers)):
            legend_header_row[i].text = ""
        
        # Add legend items
        for legend_item in legend:
            legend_row = table.add_row().cells
            phase_name = legend_item.get('phase', '')
            phase_color = legend_item.get('color', '')
            
            # Color indicator cell
            legend_row[0].text = ""
            if phase_color:
                set_cell_shading(legend_row[0], phase_color)
            
            # Phase name
            legend_row[1].text = phase_name
            
            # Clear remaining cells
            for i in range(2, len(headers)):
                legend_row[i].text = ""
    
    return table

def _generate_simple_table(doc: Document, rows: List[Dict[str, Any]], color_map: Dict[str, str], header_color: str):
    if not rows:
        return doc.add_table(rows=0, cols=0)
    
    headers = list(rows[0].keys())
    
    table = doc.add_table(rows=0, cols=len(headers))
    table.style = "Table Grid"
    
    add_row_with_color(table, headers, header_color)
    
    for row_data in rows:
        values = [row_data.get(h, '') for h in headers]
        add_row_with_color(table, values, None)
    
    return table
