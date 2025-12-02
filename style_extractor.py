import sys
import json
import os
from docx import Document
from docx.shared import Pt # Used for converting size

# Docx.js uses half-points for font size. 1pt = 2 half-points
# Docx.js uses twips for margins. 1 cm = 567 twips. 1 inch = 1440 twips
# python-docx margin is in EMUs. 1 EMU = 1/914400 inch. 1 inch = 914400 EMUs.
# 1 cm = 360000 EMUs
def emu_to_cm_str(emu):
    """Convert EMU to cm string (e.g., '3.7cm') for docx.js."""
    if emu is None: return None
    return f"{round(emu / 360000, 2)}cm"


def get_font_name_from_rPr(rPr_element):
    """Extract font name from rPr element, prioritizing eastAsia."""
    if rPr_element is None:
        return None
    rFonts = rPr_element.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    if rFonts is not None:
        eastAsia = rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
        if eastAsia:
            return eastAsia
        ascii_font = rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii')
        if ascii_font:
            return ascii_font
    return None

def get_font_name_from_style(style):
    """Try to extract font name from a style, prioritizing eastAsia, then direct, then base style."""
    # 1. Check style's rPr (run properties) for w:rFonts
    font_name = get_font_name_from_rPr(style.element.rPr)
    if font_name:
        return font_name

    # 2. Fallback to direct style.font.name
    if style.font.name:
        return style.font.name

    # 3. Check base style if available (for inherited fonts)
    if style.base_style:
        return get_font_name_from_style(style.base_style)
    
    return None # Fallback


def get_font_size_from_style(style):
    """Extract font size in half-points (docx.js unit)."""
    # Prefer direct font size
    if style.font.size and style.font.size.pt is not None:
        return int(style.font.size.pt * 2) # Convert Pt to half-points
    
    # Check base style for inherited size
    if style.base_style:
        return get_font_size_from_style(style.base_style)
    
    return None

def get_font_color_from_style(style):
    """Extract font color as hex string."""
    if style.font.color and style.font.color.rgb:
        return str(style.font.color.rgb)
    
    if style.base_style:
        return get_font_color_from_style(style.base_style)
        
    return None

def get_line_spacing_from_style(style):
    """Extract line spacing in half-points (docx.js expects this for Paragraph.spacing.line)."""
    if style.paragraph_format.line_spacing is not None:
        spacing = style.paragraph_format.line_spacing
        if isinstance(spacing, float): # Usually 1.0, 1.5, 2.0 lines
            return int(spacing * 240) # 240 twips = 1 line
        elif hasattr(spacing, 'pt') and spacing.pt is not None: # Explicit point value
            return int(spacing.pt * 20) # 1pt = 20 twips
    
    if style.base_style:
        return get_line_spacing_from_style(style.base_style)
        
    return None

def get_first_line_indent_from_style(style):
    """Extract first line indent in twips (docx.js unit)."""
    if style.paragraph_format.first_line_indent is not None:
        return int(style.paragraph_format.first_line_indent.twips)
    
    if style.base_style:
        return get_first_line_indent_from_style(style.base_style)
        
    return None

def extract_styles(docx_path):
    if not os.path.exists(docx_path):
        return {"error": "File not found"}

    try:
        doc = Document(docx_path)
        styles_data = {}
        
        # 1. Extract Margins (from first section)
        if doc.sections:
            section = doc.sections[0]
            styles_data['margin'] = {
                "top": emu_to_cm_str(section.top_margin),
                "bottom": emu_to_cm_str(section.bottom_margin),
                "left": emu_to_cm_str(section.left_margin),
                "right": emu_to_cm_str(section.right_margin)
            }
        
        # Helper to get style by name, handling common aliases
        def get_style_robust(style_name):
            if style_name in doc.styles:
                return doc.styles[style_name]
            # Try some common aliases if direct name not found (e.g., "Heading1" vs "Heading 1")
            if style_name == "Heading 1" and "Heading1" in doc.styles: return doc.styles["Heading1"]
            if style_name == "Heading 2" and "Heading2" in doc.styles: return doc.styles["Heading2"]
            if style_name == "Heading 3" and "Heading3" in doc.styles: return doc.styles["Heading3"]
            if style_name == "Normal" and "Normal (Web)" in doc.styles: return doc.styles["Normal (Web)"] # Edge case
            return None

        # Detailed style analysis for debugging
        styles_data['detailed_styles_info'] = []
        for style_obj in doc.styles: # Correct way to iterate styles
            if style_obj.type == 1: # Paragraph style (WdStyleType.wdStyleTypeParagraph)
                styles_data['detailed_styles_info'].append({
                    "id": style_obj.style_id,
                    "name": style_obj.name,
                    "font_name": get_font_name_from_style(style_obj),
                    "font_size": get_font_size_from_style(style_obj),
                    "line_spacing": get_line_spacing_from_style(style_obj),
                    "first_line_indent": get_first_line_indent_from_style(style_obj),
                    "base_style": style_obj.base_style.name if style_obj.base_style else None
                })


        # 2. Extract Fonts & Sizes for key styles for actual use
        target_styles_map = {
            "Normal": "Main", 
            "Heading 1": "H1", 
            "Heading 2": "H2", 
            "Heading 3": "H3"
        }

        for style_name, prefix in target_styles_map.items():
            style = get_style_robust(style_name)
            if style:
                font_name = get_font_name_from_style(style)
                if font_name:
                    styles_data[f"font{prefix}"] = font_name
                
                size = get_font_size_from_style(style)
                if size:
                    styles_data[f"fontSize{prefix}"] = size
                
                # Check for red color for H1 (common in official documents)
                if style_name == "Heading 1":
                    color = get_font_color_from_style(style)
                    # Check for a range of red values, not just exact FF0000
                    if color and (color.upper().startswith("FF0000") or color.upper().startswith("E00000") or color.upper().startswith("D00000")):
                        styles_data["redHeader"] = True
                
                # Extract first line indent for paragraph styles
                if style_name == "Normal": 
                    indent = get_first_line_indent_from_style(style)
                    if indent is not None:
                        styles_data["firstLineIndent"] = indent
        
        # 3. Extract Line Spacing (from Normal)
        normal_style = get_style_robust("Normal")
        if normal_style:
            line_spacing = get_line_spacing_from_style(normal_style)
            if line_spacing is not None:
                styles_data["lineSpacing"] = line_spacing

        return styles_data

    except Exception as e:
        # Log to stderr for debugging, return JSON error
        print(f"Error processing {docx_path}: {str(e)}", file=sys.stderr)
        return {"error": str(e)}

if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(1)
    
    path = sys.argv[1]
    result = extract_styles(path)
    print(json.dumps(result, indent=2, ensure_ascii=False))