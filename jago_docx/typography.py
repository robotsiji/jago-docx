"""
Typography and MS Docs styling engine for Jago Docx.
Enforces Arial font family for all headings following the Microsoft Docs hierarchy,
strict typographic rhythm, high-contrast palette, and proper outline levels.
"""

from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn


# Microsoft Docs Palette (High Contrast Slate)
COLOR_TEXT_PRIMARY = RGBColor(0x1F, 0x29, 0x37)    # Slate 800 - Body Text
COLOR_HEADING_1    = RGBColor(0x0F, 0x17, 0x2A)    # Slate 900 - H1 / Title
COLOR_HEADING_2    = RGBColor(0x1E, 0x29, 0x3B)    # Slate 800 - H2
COLOR_HEADING_3    = RGBColor(0x33, 0x41, 0x55)    # Slate 700 - H3
COLOR_HEADING_4    = RGBColor(0x47, 0x55, 0x69)    # Slate 600 - H4
COLOR_MUTED        = RGBColor(0x64, 0x74, 0x8B)    # Slate 500 - Captions / Metadata
COLOR_ACCENT_BLUE  = RGBColor(0x02, 0x84, 0xC7)    # Light blue accent
COLOR_BORDER_LIGHT = "E2E8F0"                      # Light gray hex for borders


def setup_document_typography(doc):
    """
    Configure base styles across the entire document:
    - Normal (Body) -> Arial 10.5pt, 1.15 line spacing, 6pt space-after
    - Headings 1..6 -> Arial Bold with MS Docs proportional scales & outline levels
    - Title & Subtitle -> Arial
    """
    styles = doc.styles

    # 1. Normal (Body) Style
    normal_style = styles["Normal"]
    normal_style.font.name = "Arial"
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = COLOR_TEXT_PRIMARY
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)
    normal_style.paragraph_format.space_before = Pt(0)

    # 2. Configure Headings
    heading_specs = [
        # (Level, Size pt, Bold, Color, Space Before pt, Space After pt, Outline Level)
        (1, 18.0, True,  COLOR_HEADING_1, 16.0, 6.0, 0),
        (2, 14.0, True,  COLOR_HEADING_2, 12.0, 4.0, 1),
        (3, 12.0, True,  COLOR_HEADING_3, 8.0,  3.0, 2),
        (4, 11.0, True,  COLOR_HEADING_4, 6.0,  2.0, 3),
        (5, 10.5, False, COLOR_MUTED,     4.0,  2.0, 4),
        (6, 10.0, False, COLOR_MUTED,     4.0,  2.0, 5),
    ]

    for lvl, size, is_bold, color, sp_before, sp_after, outline_lvl in heading_specs:
        style_name = f"Heading {lvl}"
        try:
            h_style = styles[style_name]
        except KeyError:
            h_style = styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)

        h_style.font.name = "Arial"
        h_style.font.size = Pt(size)
        h_style.font.bold = is_bold
        h_style.font.color.rgb = color
        h_style.paragraph_format.space_before = Pt(sp_before)
        h_style.paragraph_format.space_after = Pt(sp_after)
        h_style.paragraph_format.keep_with_next = True
        h_style.paragraph_format.line_spacing = 1.15


def apply_heading(doc, text: str, level: int = 1):
    """
    Append a heading that strictly follows MS Docs hierarchy and Arial font.
    Ensures keep_with_next is True and outline level is properly set in OpenXML.
    """
    if level < 1:
        level = 1
    elif level > 6:
        level = 6

    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.keep_with_next = True
    
    # Run formatting
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.bold = True if level <= 4 else False
    
    sizes = {1: 18.0, 2: 14.0, 3: 12.0, 4: 11.0, 5: 10.5, 6: 10.0}
    colors = {1: COLOR_HEADING_1, 2: COLOR_HEADING_2, 3: COLOR_HEADING_3, 4: COLOR_HEADING_4, 5: COLOR_MUTED, 6: COLOR_MUTED}
    
    run.font.size = Pt(sizes.get(level, 11.0))
    run.font.color.rgb = colors.get(level, COLOR_TEXT_PRIMARY)

    # Inject OpenXML outlineLvl into paragraph properties
    pPr = p._p.get_or_add_pPr()
    outline_elem = parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="{level - 1}"/>')
    pPr.append(outline_elem)

    return p


def apply_title(doc, text: str, subtitle: str = None):
    """
    Add document title and optional subtitle following MS Docs header guidelines.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4 if subtitle else 16)
    p_title.paragraph_format.keep_with_next = True
    
    run_title = p_title.add_run(text)
    run_title.font.name = "Arial"
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = COLOR_HEADING_1

    p_sub = None
    if subtitle:
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_before = Pt(0)
        p_sub.paragraph_format.space_after = Pt(18)
        p_sub.paragraph_format.keep_with_next = True
        
        run_sub = p_sub.add_run(subtitle)
        run_sub.font.name = "Arial"
        run_sub.font.size = Pt(12)
        run_sub.font.color.rgb = COLOR_MUTED

    return p_title, p_sub
