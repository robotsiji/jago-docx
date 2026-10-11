"""
Core JagoDocument Class for Jago Docx.
Unifies antislop enforcement, Arial MS Docs typography, automated TOC / figure lists,
dynamic OS author metadata, APA footnotes, editable OMML math formulas,
under-element captions, and editable page numbering into an intuitive interface.
"""

import os
from typing import List, Optional, Union

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from .metadata import apply_author_metadata, get_default_username
from .typography import (
    setup_document_typography,
    apply_heading,
    apply_title,
    COLOR_TEXT_PRIMARY,
    COLOR_MUTED,
    COLOR_HEADING_1
)
from .toc import (
    add_table_of_contents_page,
    add_list_of_figures_page,
    add_list_of_tables_page,
    enable_update_fields_on_open
)
from .footnotes import (
    FootnoteManager,
    add_quotation_with_reference,
    format_apa_reference
)
from .math_formula import add_editable_formula
from .captions import add_table_with_caption, add_figure_with_caption
from .page_number import setup_editable_page_number
from .antislop import (
    enforce_antislop_page_setup,
    audit_document_antislop,
    sanitize_prose
)

W_NS = nsdecls("w")


class JagoDocument:
    """
    High-level document builder enforcing Jago Docx quality standards:
    1. Antislop layout, prose, and spacing discipline
    2. Arial font for all headings + auto-generated Table of Contents
    3. Current username as author instead of 'python-docx'
    4. APA 7th edition footnotes and proper quotation blocks
    5. Microsoft Docs heading hierarchy (levels 1-6)
    6. Native editable formula placeholders (OMML)
    7. Captions strictly placed directly underneath tables & figures + List of Tables/Figures
    8. Editable footer page numbers with basic placeholders
    """

    def __init__(
        self,
        title: Optional[str] = None,
        author: Optional[str] = None,
        subject: Optional[str] = None,
        footer_format: str = "page_of_total"
    ):
        self.doc = Document()
        
        # 1. Antislop Page Setup (1-inch margins)
        enforce_antislop_page_setup(self.doc)

        # 2. Typography Setup (Arial headings, MS Docs hierarchy)
        setup_document_typography(self.doc)

        # 3. Dynamic Author Metadata (Username instead of 'python-docx')
        self.author = apply_author_metadata(self.doc, author=author, title=title, subject=subject)

        # 4. Footnote Manager
        self.footnote_mgr = FootnoteManager(self.doc)

        # 5. Editable Footer Page Numbers
        setup_editable_page_number(self.doc, format_type=footer_format)

        # 6. Automatic fields update on open
        enable_update_fields_on_open(self.doc)

    # -------------------------------------------------------------------------
    # Headings & Typography (Requirements 2 & 5)
    # -------------------------------------------------------------------------
    def add_title(self, text: str, subtitle: Optional[str] = None):
        """Add document Title and optional Subtitle."""
        return apply_title(self.doc, text, subtitle)

    def add_heading(self, text: str, level: int = 1):
        """
        Add a heading adhering strictly to MS Docs format and Arial font.
        Levels 1 through 6 are supported.
        """
        return apply_heading(self.doc, text, level)

    def add_paragraph(self, text: str, lead_bold: Optional[str] = None, sanitize: bool = True):
        """
        Add a standard body paragraph with Arial 10.5pt, 1.15 line spacing, 6pt after.
        Optionally sanitizes prose against AI buzzwords.
        """
        clean_text = sanitize_prose(text) if sanitize else text
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15

        if lead_bold:
            r_lead = p.add_run(lead_bold.strip() + " ")
            r_lead.font.name = "Arial"
            r_lead.font.size = Pt(10.5)
            r_lead.font.bold = True
            r_lead.font.color.rgb = COLOR_HEADING_1

        r = p.add_run(clean_text)
        r.font.name = "Arial"
        r.font.size = Pt(10.5)
        r.font.color.rgb = COLOR_TEXT_PRIMARY
        return p

    def add_page_break(self):
        """Insert a clean page break."""
        return self.doc.add_page_break()

    # -------------------------------------------------------------------------
    # Table of Contents & Lists (Requirements 2 & 7)
    # -------------------------------------------------------------------------
    def add_table_of_contents(self, title: str = "Table of Contents"):
        """Auto-generates a dedicated page for the Table of Contents."""
        add_table_of_contents_page(self.doc, title=title)

    def add_list_of_figures(self, title: str = "List of Figures"):
        """Auto-generates a dedicated page for the List of Figures."""
        add_list_of_figures_page(self.doc, title=title)

    def add_list_of_tables(self, title: str = "List of Tables"):
        """Auto-generates a dedicated page for the List of Tables."""
        add_list_of_tables_page(self.doc, title=title)

    # -------------------------------------------------------------------------
    # Quotations & Footnotes with APA Style (Requirement 4)
    # -------------------------------------------------------------------------
    def add_quotation_with_reference(
        self,
        quote_text: str,
        apa_reference: str,
        lead_in: Optional[str] = None
    ):
        """
        Formats quotation (inline for <= 40 words, blockquote for > 40 words)
        and inserts the APA reference as a native Word footnote.
        """
        return add_quotation_with_reference(
            self.doc,
            quote_text=quote_text,
            apa_reference=apa_reference,
            footnote_mgr=self.footnote_mgr,
            lead_in=lead_in
        )

    def add_footnote(self, target_paragraph, apa_reference: str) -> int:
        """
        Attaches an APA footnote reference to an existing paragraph.
        """
        return self.footnote_mgr.add_footnote(target_paragraph, apa_reference)

    # -------------------------------------------------------------------------
    # Editable Formulas (Requirement 6)
    # -------------------------------------------------------------------------
    def add_formula(
        self,
        formula_text: Optional[str] = None,
        equation_number: Optional[int] = None,
        preset: Optional[str] = None,
        description: Optional[str] = None
    ):
        """
        Inserts a native, editable Word equation block (OMML) with optional
        equation numbering and explanatory description.
        """
        return add_editable_formula(
            self.doc,
            formula_text=formula_text,
            equation_number=equation_number,
            preset=preset,
            description=description
        )

    # -------------------------------------------------------------------------
    # Tables and Figures with Captions Right Underneath (Requirement 7)
    # -------------------------------------------------------------------------
    def add_table(
        self,
        headers: List[str],
        rows: List[List[Union[str, int, float]]],
        caption: str,
        col_widths: Optional[List[float]] = None
    ):
        """
        Adds a beautifully styled antislop data table and strictly places
        the caption directly underneath it.
        """
        return add_table_with_caption(
            self.doc,
            headers=headers,
            rows=rows,
            caption=caption,
            col_widths=col_widths
        )

    def add_figure(
        self,
        image_path_or_placeholder: str,
        caption: str,
        width_inches: float = 5.0
    ):
        """
        Adds an image or visual block and strictly places the caption directly underneath it.
        """
        return add_figure_with_caption(
            self.doc,
            image_path_or_text=image_path_or_placeholder,
            caption=caption,
            width_inches=width_inches
        )

    # -------------------------------------------------------------------------
    # Callout / Note Boxes (Antislop Design)
    # -------------------------------------------------------------------------
    def add_callout(self, text: str, title: Optional[str] = None, border_color: str = "0284C7"):
        """
        Adds a modern MS Docs style note or callout box with light background and left border.
        """
        table = self.doc.add_table(rows=1, cols=1)
        table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell = table.cell(0, 0)
        
        tcPr = cell._tc.get_or_add_tcPr()
        # Left border only
        borders = parse_xml(f"""
        <w:tcBorders {W_NS}>
          <w:top w:val="none"/>
          <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>
          <w:bottom w:val="none"/>
          <w:right w:val="none"/>
        </w:tcBorders>
        """)
        tcPr.append(borders)
        
        # Background
        shd = parse_xml(f'<w:shd {W_NS} w:val="clear" w:color="auto" w:fill="F8FAFC"/>')
        tcPr.append(shd)

        # Padding
        tcMar = parse_xml(f"""
        <w:tcMar {W_NS}>
          <w:top w:w="160" w:type="dxa"/>
          <w:bottom w:w="160" w:type="dxa"/>
          <w:left w:w="200" w:type="dxa"/>
          <w:right w:w="200" w:type="dxa"/>
        </w:tcMar>
        """)
        tcPr.append(tcMar)

        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15

        if title:
            r_title = p.add_run(title.strip() + "\n")
            r_title.font.name = "Arial"
            r_title.font.bold = True
            r_title.font.size = Pt(10)
            r_title.font.color.rgb = COLOR_HEADING_1

        r_text = p.add_run(text.strip())
        r_text.font.name = "Arial"
        r_text.font.size = Pt(10)
        r_text.font.color.rgb = COLOR_TEXT_PRIMARY

        # Spacing after callout table
        spacer = self.doc.add_paragraph()
        spacer.paragraph_format.space_before = Pt(0)
        spacer.paragraph_format.space_after = Pt(4)

        return table

    # -------------------------------------------------------------------------
    # Audit & Save
    # -------------------------------------------------------------------------
    def audit(self):
        """Audit the document against antislop rules."""
        return audit_document_antislop(self.doc)

    def save(self, filepath: str):
        """Save the document to disk."""
        directory = os.path.dirname(filepath)
        if directory and not os.path.exists(directory):
            os.makedirs(directory, exist_ok=True)
            
        self.doc.save(filepath)
        return filepath
