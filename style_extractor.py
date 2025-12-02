import sys
import json
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn

def emu_to_cm_str(emu):
    if emu is None: return None
    return f"{round(emu / 360000, 2)}cm"

def get_font_name_from_rPr(rPr_element):
    if rPr_element is None: return None
    rFonts = rPr_element.find(qn('w:rFonts'))
    if rFonts is not None:
        eastAsia = rFonts.get(qn('w:eastAsia'))
        if eastAsia: return eastAsia
        ascii_font = rFonts.get(qn('w:ascii'))
        if ascii_font: return ascii_font
    return None

def get_font_name_from_style(style):
    font_name = get_font_name_from_rPr(style.element.rPr)
    if font_name: return font_name
    if style.font.name: return style.font.name
    if style.base_style: return get_font_name_from_style(style.base_style)
    return None

def get_font_size_from_style(style):
    if style.font.size and style.font.size.pt is not None:
        return int(style.font.size.pt * 2)
    if style.base_style: return get_font_size_from_style(style.base_style)
    return None

def get_font_color_from_style(style):
    if style.font.color and style.font.color.rgb:
        return str(style.font.color.rgb)
    if style.base_style: return get_font_color_from_style(style.base_style)
    return None

def get_line_spacing_from_style(style):
    if style.paragraph_format.line_spacing is not None:
        spacing = style.paragraph_format.line_spacing
        if isinstance(spacing, float): return int(spacing * 240)
        elif hasattr(spacing, 'pt') and spacing.pt is not None: return int(spacing.pt * 20)
    if style.base_style: return get_line_spacing_from_style(style.base_style)
    return None

def get_first_line_indent_from_style(style):
    if style.paragraph_format.first_line_indent is not None:
        return int(style.paragraph_format.first_line_indent.twips)
    if style.base_style: return get_first_line_indent_from_style(style.base_style)
    return None

# --- New: Table Analysis Logic ---
def analyze_table_style(doc):
    """
    Analyze the first table in the document to guess style preferences.
    Returns a dict matching the 'table' config structure in docxjs-cli templates.
    """
    if not doc.tables:
        return None # No table found, let JS use default

    table = doc.tables[0]
    
    # 1. Detect Borders (Simple heuristic based on bottom border of first cell)
    # This is tricky because borders can be on table style OR direct formatting.
    # We will check table properties (tblPr)
    border_style = "single"
    border_color = "000000"
    border_size = 4

    try:
        tblBorders = table._element.tblPr.find(qn('w:tblBorders'))
        if tblBorders is not None:
            # Check specific borders, e.g., bottom
            bottom = tblBorders.find(qn('w:bottom'))
            if bottom is not None:
                val = bottom.get(qn('w:val'))
                if val == 'nil' or val == 'none': border_style = "none"
                elif val == 'dotted': border_style = "dotted"
                elif val == 'dashed': border_style = "dashed"
                elif val == 'double': border_style = "double"
                else: border_style = "single" # default to single for 'single', 'thick', etc.
                
                color = bottom.get(qn('w:color'))
                if color and color != 'auto': border_color = color
                
                sz = bottom.get(qn('w:sz'))
                if sz: border_size = int(sz)
    except:
        pass

    # 2. Detect Header Boldness
    header_bold = False
    header_color = "000000"
    
    try:
        first_row = table.rows[0]
        # Check the first paragraph of the first cell
        first_cell = first_row.cells[0]
        if first_cell.paragraphs:
            p = first_cell.paragraphs[0]
            # Check runs for bold
            for run in p.runs:
                if run.bold:
                    header_bold = True
                    break
            # Also check style if not directly bold
            if not header_bold and p.style and p.style.font.bold:
                header_bold = True
                
            # Check color
            for run in p.runs:
                if run.font.color and run.font.color.rgb:
                    header_color = str(run.font.color.rgb)
                    break
    except:
        pass

    return {
        "borderStyle": border_style,
        "borderColor": border_color,
        "borderSize": border_size,
        "headerBold": header_bold,
        "headerColor": header_color,
        "cellAlign": "left" # Hard to detect generic preference, default to left
    }

def extract_styles(docx_path):
    if not os.path.exists(docx_path): return {"error": "File not found"}

    try:
        doc = Document(docx_path)
        styles_data = {}
        
        # 1. Margins
        if doc.sections:
            section = doc.sections[0]
            styles_data['margin'] = {
                "top": emu_to_cm_str(section.top_margin),
                "bottom": emu_to_cm_str(section.bottom_margin),
                "left": emu_to_cm_str(section.left_margin),
                "right": emu_to_cm_str(section.right_margin)
            }
        
        # Helper
        def get_style_robust(style_name):
            if style_name in doc.styles: return doc.styles[style_name]
            if style_name == "Heading 1" and "Heading1" in doc.styles: return doc.styles["Heading1"]
            if style_name == "Heading 2" and "Heading2" in doc.styles: return doc.styles["Heading2"]
            if style_name == "Heading 3" and "Heading3" in doc.styles: return doc.styles["Heading3"]
            return None

        # 2. Fonts
        target_styles_map = {"Normal": "Main", "Heading 1": "H1", "Heading 2": "H2", "Heading 3": "H3"}
        for style_name, prefix in target_styles_map.items():
            style = get_style_robust(style_name)
            if style:
                font_name = get_font_name_from_style(style)
                if font_name: styles_data[f"font{prefix}"] = font_name
                size = get_font_size_from_style(style)
                if size: styles_data[f"fontSize{prefix}"] = size
                if style_name == "Heading 1":
                    color = get_font_color_from_style(style)
                    if color: styles_data["colorHeader1"] = color # Store exact color
                    if color and (color.upper().startswith("FF0000") or color.upper().startswith("E0") or color.upper().startswith("C0")):
                        styles_data["redHeader"] = True
                if style_name == "Normal": 
                    indent = get_first_line_indent_from_style(style)
                    if indent is not None: styles_data["firstLineIndent"] = indent
                    line_spacing = get_line_spacing_from_style(style)
                    if line_spacing is not None: styles_data["lineSpacing"] = line_spacing
                    color = get_font_color_from_style(style)
                    if color: styles_data["colorMain"] = color

        # 3. Table Styles
        table_config = analyze_table_style(doc)
        if table_config:
            styles_data["table"] = table_config

        return styles_data

    except Exception as e:
        print(f"Error processing {docx_path}: {str(e)}", file=sys.stderr)
        return {"error": str(e)}

if __name__ == "__main__":
    if len(sys.argv) < 2: sys.exit(1)
    print(json.dumps(extract_styles(sys.argv[1]), indent=2, ensure_ascii=False))
