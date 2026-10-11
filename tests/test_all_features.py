"""
Comprehensive unit test suite for Jago Docx verifying all 8 requirements.
"""

import os
import sys
import unittest
import zipfile
import tempfile

# Ensure package is on sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from jago_docx import (
    JagoDocument,
    get_default_username,
    format_apa_reference,
    audit_document_antislop,
    sanitize_prose
)


class TestJagoDocx(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.doc_path = os.path.join(self.temp_dir.name, "test_output.docx")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_01_antislop(self):
        """Rule 1: Antislop prose sanitization and layout auditing."""
        slop_text = "In today's fast-paced digital era, we must delve into the multifaceted tapestry of AI."
        cleaned = sanitize_prose(slop_text)
        self.assertNotIn("In today's fast-paced digital era,", cleaned)
        self.assertNotIn("delve into", cleaned)

        # Audit check
        doc = JagoDocument(title="Antislop Test")
        doc.add_heading("Section 1", level=1)
        # Empty paragraph check
        doc.doc.add_paragraph("")
        violations = doc.audit()
        self.assertTrue(any(v.category == "Layout-Slop" for v in violations))

    def test_02_arial_headings_and_toc(self):
        """Rule 2 & 5: Arial font for all headings, MS Docs hierarchy, and auto TOC."""
        doc = JagoDocument(title="Typography Test")
        doc.add_table_of_contents()
        h1 = doc.add_heading("1. Heading One", level=1)
        h2 = doc.add_heading("1.1 Heading Two", level=2)
        h3 = doc.add_heading("1.1.1 Heading Three", level=3)

        doc.save(self.doc_path)
        with zipfile.ZipFile(self.doc_path, 'r') as z:
            doc_xml = z.read("word/document.xml").decode("utf-8")
            settings_xml = z.read("word/settings.xml").decode("utf-8")
            self.assertIn("Arial", doc_xml)
            self.assertIn('w:val="0"', doc_xml)  # Outline level 0
            self.assertIn("TOC", doc_xml)
            self.assertIn("updateFields", settings_xml)

    def test_03_username_as_author(self):
        """Rule 3: OS username as author, never 'python-docx'."""
        username = get_default_username()
        self.assertTrue(len(username) > 0)

        doc = JagoDocument(title="Metadata Test")
        self.assertEqual(doc.author, username)
        doc.save(self.doc_path)

        with zipfile.ZipFile(self.doc_path, 'r') as z:
            core_xml = z.read("docProps/core.xml").decode("utf-8")
            self.assertIn(username, core_xml)
            self.assertNotIn("python-docx", core_xml.lower())

    def test_04_apa_footnotes_and_quotations(self):
        """Rule 4: Quotation formatting and APA 7th native footnotes."""
        doc = JagoDocument(title="Citations Test")
        apa_ref = format_apa_reference(
            authors="Shannon, C. E.",
            year="1948",
            title="A mathematical theory of communication",
            source="Bell System Technical Journal",
            volume_issue="27(3)",
            pages="379–423",
            doi_or_url="10.1002/j.1538-7305.1948.tb01338.x"
        )
        self.assertIn("Shannon, C. E. (1948).", apa_ref)
        self.assertIn("Bell System Technical Journal", apa_ref)

        # Short quote
        doc.add_quotation_with_reference("Information is the resolution of uncertainty.", apa_ref)

        # Blockquote (> 40 words)
        long_quote = (
            "The fundamental problem of communication is that of reproducing at one point "
            "either exactly or approximately a message selected at another point. "
            "Frequently the messages have meaning; that is they refer to or are correlated according "
            "to some system with certain physical or conceptual entities."
        )
        doc.add_quotation_with_reference(long_quote, apa_ref)

        doc.save(self.doc_path)
        with zipfile.ZipFile(self.doc_path, 'r') as z:
            names = z.namelist()
            self.assertIn("word/footnotes.xml", names)
            footnotes_xml = z.read("word/footnotes.xml").decode("utf-8")
            self.assertIn("Shannon", footnotes_xml)

    def test_05_editable_formula(self):
        """Rule 6: Native editable formula placeholder with OMML."""
        doc = JagoDocument(title="Formula Test")
        doc.add_formula(formula_text="E = mc^2", equation_number=1, description="Mass-energy equivalence")
        doc.add_formula(preset="quadratic", equation_number=2)

        doc.save(self.doc_path)
        with zipfile.ZipFile(self.doc_path, 'r') as z:
            doc_xml = z.read("word/document.xml").decode("utf-8")
            self.assertIn("m:oMath", doc_xml)
            self.assertIn("mc^2", doc_xml)

    def test_06_captions_underneath_and_lists(self):
        """Rule 7: Captions strictly underneath table and figure, plus auto lists."""
        doc = JagoDocument(title="Captions Test")
        doc.add_list_of_figures()
        doc.add_list_of_tables()

        # Add table with caption underneath
        headers = ["Col A", "Col B"]
        rows = [["Data 1", "Data 2"]]
        doc.add_table(headers, rows, caption="Sample Benchmark Data")

        # Add figure with caption underneath
        doc.add_figure("Placeholder: Architecture Chart", caption="Sample Diagram")

        doc.save(self.doc_path)
        with zipfile.ZipFile(self.doc_path, 'r') as z:
            doc_xml = z.read("word/document.xml").decode("utf-8")
            # Verify SEQ fields
            self.assertIn("SEQ Table", doc_xml)
            self.assertIn("SEQ Figure", doc_xml)
            # Verify List of Figures & Tables fields
            self.assertIn('TOC \\c "Figure"', doc_xml)
            self.assertIn('TOC \\c "Table"', doc_xml)

    def test_07_editable_page_numbers(self):
        """Rule 8: Editable page numbers in footer with PAGE and NUMPAGES."""
        doc = JagoDocument(title="Page Number Test")
        doc.save(self.doc_path)

        with zipfile.ZipFile(self.doc_path, 'r') as z:
            names = z.namelist()
            footer_files = [f for f in names if "footer" in f]
            self.assertTrue(len(footer_files) > 0)
            footer_xml = z.read(footer_files[0]).decode("utf-8")
            self.assertIn("PAGE", footer_xml)
            self.assertIn("NUMPAGES", footer_xml)


if __name__ == "__main__":
    unittest.main()
