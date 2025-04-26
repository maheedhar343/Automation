from docx.shared import Inches, Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
import re

def lighten_color(hex_color, factor=0.4):
    """Lightens the given color by the factor"""
    hex_color = hex_color.lstrip('#')
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    r = min(255, int(r + (255 - r) * factor))
    g = min(255, int(g + (255 - g) * factor))
    b = min(255, int(b + (255 - b) * factor))
    return '{:02x}{:02x}{:02x}'.format(r, g, b)

def set_cell_shading(cell, color_hex):
    """Set the background color of a table cell"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def set_cell_margins(cell, margin_value=100):
    """
    Set margins (padding) for a table cell.
    margin_value: margin width in twips (100 twips = ~0.07 inches)
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    existing_mar = tcPr.find(qn('w:tcMar'))
    if existing_mar is not None:
        tcPr.remove(existing_mar)
    
    tcMar = OxmlElement('w:tcMar')
    for prop in ['top', 'bottom', 'left', 'right']:
        node = OxmlElement(f'w:{prop}')
        node.set(qn('w:w'), str(margin_value))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    
    tcPr.append(tcMar)

def set_cell_border(cell, border_position, border_type='single', border_size=4, border_color='000000'):
    """Set the border for a cell"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    
    border_xml = OxmlElement(f'w:{border_position}')
    border_xml.set(qn('w:val'), border_type)
    border_xml.set(qn('w:sz'), str(border_size))
    border_xml.set(qn('w:color'), border_color)
    
    existing_border = tcBorders.find(qn(f'w:{border_position}'))
    if existing_border is not None:
        tcBorders.remove(existing_border)
    
    tcBorders.append(border_xml)

def set_table_borders(table):
    """Apply 'Table Grid' style and ensure all cell borders are visible and black, except between first and second rows"""
    table.style = 'Table Grid'
    for row_idx, row in enumerate(table.rows):
        for cell in row.cells:
            # Apply borders to all sides by default
            set_cell_border(cell, 'top')
            set_cell_border(cell, 'bottom')
            set_cell_border(cell, 'left')
            set_cell_border(cell, 'right')
            
            # Remove bottom border of first row
            if row_idx == 0:
                set_cell_border(cell, 'bottom', border_type='nil')
            # Remove top border of second row
            if row_idx == 1:
                set_cell_border(cell, 'top', border_type='nil')

def format_text_with_bullets(text, apply_bullets=False):
    """Format text with bullets for lines after the first if apply_bullets is True."""
    if apply_bullets:
        # Split text by line breaks
        lines = text.split('\n')
        formatted_lines = []
        for idx, line in enumerate(lines):
            line = line.strip()  # Remove leading/trailing whitespace
            if line:  # Only add non-empty lines
                if idx == 0:
                    formatted_lines.append(line)  # First line without bullet
                else:
                    formatted_lines.append(f"    • {line}")  # Subsequent lines with bullet and indent
        return '\n'.join(formatted_lines)
    else:
        return text