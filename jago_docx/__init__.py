"""
Jago Docx - Expert Antislop Word (.docx) Document Generator.
Features:
- Antislop layout & prose hygiene
- Arial font for all headings + MS Docs hierarchy
- Auto-detected OS username as document author (never 'python-docx')
- Native Word footnotes with APA 7th edition formatting
- Native editable formula placeholders (OMML)
- Captions strictly directly underneath tables & figures
- Auto-generated Table of Contents, List of Figures, List of Tables
- Native editable footer page numbering (PAGE & NUMPAGES)
"""

__version__ = "1.0.0"

from .core import JagoDocument
from .metadata import get_default_username, apply_author_metadata
from .typography import (
    apply_heading,
    apply_title,
    setup_document_typography,
    COLOR_TEXT_PRIMARY,
    COLOR_HEADING_1,
    COLOR_HEADING_2,
    COLOR_HEADING_3,
    COLOR_HEADING_4,
    COLOR_MUTED
)
from .footnotes import format_apa_reference, FootnoteManager, add_quotation_with_reference
from .math_formula import add_editable_formula, FORMULA_PRESETS
from .captions import add_table_with_caption, add_figure_with_caption
from .toc import add_table_of_contents_page, add_list_of_figures_page, add_list_of_tables_page
from .page_number import setup_editable_page_number
from .antislop import sanitize_prose, audit_document_antislop

__all__ = [
    "JagoDocument",
    "get_default_username",
    "apply_author_metadata",
    "apply_heading",
    "apply_title",
    "setup_document_typography",
    "format_apa_reference",
    "FootnoteManager",
    "add_quotation_with_reference",
    "add_editable_formula",
    "FORMULA_PRESETS",
    "add_table_with_caption",
    "add_figure_with_caption",
    "add_table_of_contents_page",
    "add_list_of_figures_page",
    "add_list_of_tables_page",
    "setup_editable_page_number",
    "sanitize_prose",
    "audit_document_antislop",
]
