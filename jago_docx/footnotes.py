"""
Footnotes and APA 7th Reference Engine for Jago Docx.
Handles proper quotation formatting (short inline quotes vs blockquotes),
generates APA 7th edition citation strings, and injects true native Word footnotes (footnotes.xml).
"""

from typing import List, Optional, Dict
from docx.shared import Inches, Pt, RGBColor
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from .typography import COLOR_TEXT_PRIMARY, COLOR_MUTED


FOOTNOTE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
FOOTNOTE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"


def format_apa_reference(
    authors: str,
    year: str,
    title: str,
    source: str = "",
    volume_issue: str = "",
    pages: str = "",
    doi_or_url: str = ""
) -> str:
    """
    Format a reference according to APA 7th Edition guidelines.
    Example output:
    Smith, J. A., & Doe, R. (2023). Scalable distributed system design. Journal of Software Engineering, 15(2), 102–115. https://doi.org/10.1000/182
    """
    authors = authors.strip()
    if not authors.endswith("."):
        authors_formatted = f"{authors}."
    else:
        authors_formatted = authors

    year = year.strip("() ")
    title = title.strip().rstrip(".")
    source = source.strip().rstrip(".")

    citation = f"{authors_formatted} ({year}). {title}."
    
    if source:
        citation += f" {source}"
        if volume_issue:
            citation += f", {volume_issue.strip()}"
        if pages:
            citation += f", {pages.strip()}"
        citation += "."
        
    if doi_or_url:
        doi_clean = doi_or_url.strip()
        if not doi_clean.startswith("http"):
            doi_clean = f"https://doi.org/{doi_clean}"
        citation += f" {doi_clean}"

    return citation.strip()


class FootnoteManager:
    """
    Manages native OpenXML footnotes across a Document.
    """
    def __init__(self, doc):
        self.doc = doc
        self.footnotes: List[Dict[str, str]] = []
        self._part_initialized = False

    def _ensure_footnotes_part(self):
        """
        Ensures the footnotes.xml part is registered in the docx package.
        """
        if self._part_initialized:
            return

        w_ns = nsdecls("w")
        base_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes {w_ns}>
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p>
  </w:footnote>
</w:footnotes>"""

        uri = PackURI("/word/footnotes.xml")
        self.footnote_part = Part(
            uri,
            FOOTNOTE_CONTENT_TYPE,
            base_xml.encode("utf-8"),
            self.doc.part.package
        )
        self.doc.part.relate_to(self.footnote_part, FOOTNOTE_REL_TYPE)
        self._part_initialized = True

    def add_footnote(self, target_paragraph, reference_text: str) -> int:
        """
        Appends a native footnote reference to target_paragraph and registers
        the APA footnote content. Returns the 1-based footnote ID.
        """
        self._ensure_footnotes_part()
        fn_id = len(self.footnotes) + 1
        self.footnotes.append({"id": str(fn_id), "text": reference_text.strip()})

        # 1. Add footnote reference run to the paragraph
        w_ns = nsdecls("w")
        ref_xml = f"""
        <w:r {w_ns}>
          <w:rPr>
            <w:rStyle w:val="FootnoteReference"/>
          </w:rPr>
          <w:footnoteReference w:id="{fn_id}"/>
        </w:r>
        """
        target_paragraph._p.append(parse_xml(ref_xml))

        # 2. Update the footnotes.xml part with the new footnote entry
        self._sync_footnotes_xml()
        return fn_id

    def _sync_footnotes_xml(self):
        """
        Serializes all registered footnotes into the footnotes.xml part.
        """
        w_ns = nsdecls("w")
        items_xml = []
        for fn in self.footnotes:
            # Escape XML special characters in reference text
            safe_text = (
                fn["text"]
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;")
            )
            item = f"""
            <w:footnote w:id="{fn['id']}">
              <w:p>
                <w:pPr>
                  <w:pStyle w:val="FootnoteText"/>
                  <w:spacing w:after="40" w:line="240" w:lineRule="auto"/>
                </w:pPr>
                <w:r>
                  <w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
                  <w:footnoteRef/>
                </w:r>
                <w:r>
                  <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                    <w:sz w:val="18"/>
                    <w:color w:val="334155"/>
                  </w:rPr>
                  <w:t xml:space="preserve"> {safe_text}</w:t>
                </w:r>
              </w:p>
            </w:footnote>
            """
            items_xml.append(item)

        full_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes {w_ns}>
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p>
  </w:footnote>
  {"".join(items_xml)}
</w:footnotes>"""

        self.footnote_part._blob = full_xml.encode("utf-8")


def add_quotation_with_reference(
    doc,
    quote_text: str,
    apa_reference: str,
    footnote_mgr: FootnoteManager,
    lead_in: Optional[str] = None
):
    """
    Formats quotes according to standard scholarly guidelines:
    - Short quotes (<= 40 words): Inline format enclosed in double quotes.
    - Block quotes (> 40 words): Dedicated paragraph indented 0.5 inches left, no quotation marks.
    Places the reference as a native footnote.
    """
    word_count = len(quote_text.strip().split())

    if word_count > 40:
        # Long quote -> Blockquote
        if lead_in:
            p_lead = doc.add_paragraph(lead_in.strip())
            p_lead.paragraph_format.space_after = Pt(4)

        p_block = doc.add_paragraph()
        p_block.paragraph_format.left_indent = Inches(0.5)
        p_block.paragraph_format.right_indent = Inches(0.2)
        p_block.paragraph_format.space_before = Pt(4)
        p_block.paragraph_format.space_after = Pt(8)
        p_block.paragraph_format.line_spacing = 1.15

        run = p_block.add_run(quote_text.strip())
        run.font.name = "Arial"
        run.font.size = Pt(10)
        run.font.color.rgb = COLOR_TEXT_PRIMARY

        # Add footnote at the end of blockquote
        footnote_mgr.add_footnote(p_block, apa_reference)
        return p_block
    else:
        # Short quote -> Inline with double quotes
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15

        if lead_in:
            r_lead = p.add_run(lead_in.strip() + " ")
            r_lead.font.name = "Arial"
            r_lead.font.size = Pt(10.5)

        clean_quote = quote_text.strip().strip('"')
        r_quote = p.add_run(f'"{clean_quote}"')
        r_quote.font.name = "Arial"
        r_quote.font.size = Pt(10.5)
        r_quote.font.color.rgb = COLOR_TEXT_PRIMARY

        # Add footnote marker
        footnote_mgr.add_footnote(p, apa_reference)
        return p
