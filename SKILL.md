---
name: jago-docx
description: Expert antislop Microsoft Word (.docx) document generator. Enforces Arial headings in Microsoft Docs format, auto-generated Table of Contents, List of Figures and Tables, auto OS username as document author (never 'python-docx'), APA 7th edition footnotes, editable OMML math formula placeholders, captions strictly underneath visuals, and editable footer page numbers.
---

# Jago Docx: Publication-Grade Document Generator Skill

`jago-docx` is an advanced AI agent skill and Python engine for creating professional, publication-grade Microsoft Word (`.docx`) documents that completely eliminate AI slop. It adheres strictly to Microsoft Docs typographic standards, scholarly citation protocols, and native OpenXML field automation.

---

## The 8 Non-Negotiable Core Rules

Whenever generating or styling a Word document, you MUST enforce the following 8 core rules:

### 1. Antislop Discipline (Tone, Spacing, and Layout)
* **Zero AI Throat-Clearing**: Ban generic filler openings ("In today's fast-paced digital world...", "Delve into the multifaceted tapestry of...", "It is important to remember that...", "In conclusion, only time will tell...").
* **Substance Over Fluff**: Write with technical density, active voice, concrete numbers, and structured hierarchy.
* **No Manual Spacer Paragraphs**: Never insert empty paragraphs (`doc.add_paragraph("")` or `<w:p/>`) to force spacing. All vertical rhythm MUST be controlled strictly via `paragraph_format.space_before` and `space_after`.
* **Standard Margins**: Exactly 1.0 inch (72pt / 1440 dxa / 25.4mm) on all four sides.
* **Restrained Palette**: Slate neutrals (`#0F172A`, `#1E293B`, `#334155`), high-contrast body (`#1F2937`), subtle borders (`#E2E8F0`), and clean header fills (`#F8FAFC`). Never use default garish cyan/purple headings or heavy black borders.

### 2. Arial Headings & Auto-Generated Table of Contents
* **Font Family**: All headings (Title, Subtitle, Heading 1 through 6) MUST use **`Arial`**.
* **Heading 1**: Arial Bold, 18pt, space before 16pt, space after 6pt, `keep_with_next = True`.
* **Heading 2**: Arial Bold, 14pt, space before 12pt, space after 4pt, `keep_with_next = True`.
* **Heading 3**: Arial Bold, 12pt, space before 8pt, space after 3pt, `keep_with_next = True`.
* **Heading 4**: Arial Semi-Bold / Italic, 11pt, space before 6pt, space after 2pt, `keep_with_next = True`.
* **Dedicated TOC Page**: Place the Table of Contents on its own dedicated page right after the title/abstract.
* **Native Field Code**: Insert `<w:fldSimple w:instr="TOC \o &quot;1-3&quot; \h \z \u"/>`.
* **Auto-Update on Open**: Append `<w:updateFields w:val="true"/>` in `settings.xml` so Word automatically refreshes page numbers and entries when opened. Follow with an explicit page break.

### 3. Dynamic Username as Author (Never 'python-docx')
* **Auto-Detect Username**: Automatically retrieve the OS/environment username via `os.environ.get("USERNAME")`, `os.environ.get("USER")`, or `getpass.getuser()`.
* **Core Properties**: Assign `core_properties.author`, `core_properties.creator`, and `core_properties.last_modified_by` to the detected username.
* **Scrub Generic Defaults**: Ensure the string `"python-docx"` never appears anywhere in core properties, metadata, or document attributes.

### 4. Scholarly Quotations & APA 7th Edition Footnotes
* **Short Quotations (<= 40 words)**: Format inline within body text, enclosed in standard double quotation marks (`"..."`), followed immediately by a footnote reference superscript.
* **Block Quotations (> 40 words)**: Format as a dedicated indented paragraph:
  - Left indent: 0.5 inches (36pt / 720 dxa).
  - Font: Arial 10pt (or body matching), 1.15 line spacing.
  - Do NOT use quotation marks around blockquotes.
  - Place the footnote superscript reference at the end of the block.
* **APA 7th Footnote Format**: References must be placed as native Word footnotes (`word/footnotes.xml`):
  - *Journal*: `Author, A. A., & Author, B. B. (Year). Title of article. Title of Periodical, volume(issue), page–range. https://doi.org/xx.xxx`
  - *Book*: `Author, A. A. (Year). Title of book (edition). Publisher. https://doi.org/xx.xxx`
  - *Website*: `Author, A. A. (Year, Month Day). Title of web page. Website Name. https://url`

### 5. Microsoft Docs Heading Hierarchy
* Align headings with the Microsoft Learn / Docs technical documentation style:
  - Concise, descriptive titles in Sentence Case or Title Case.
  - Explicit outline levels (`<w:outlineLvl w:val="0"/>` for H1, `1` for H2, etc.) for flawless Word Navigation Pane support.
  - `keep_with_next = True` on every heading to prevent orphan titles at page bottoms.

### 6. Native Editable Math Formula Placeholders
* **Office Math Markup Language (OMML)**: Formulas must use native `<m:oMath>` and `<m:oMathPara>` tags in namespace `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"`.
* **Editable in Word**: Never insert static image bitmaps for formulas unless explicitly requested. Users must be able to click and edit equations directly inside Word's Equation Editor.
* **Placeholder Presets**: Provide clear, editable formula placeholders (e.g., `f(x) = ax² + bx + c`, `E = mc²`, `L = λ · W`) with centered alignment and optional right-aligned equation numbers `(1)`, `(2)`.

### 7. Captions Strictly Underneath Visuals & Auto Figure/Table Lists
* **Placement Rule**: Captions for ANY table or figure MUST be placed **DIRECTLY UNDERNEATH** the element. Never place table captions above the table.
* **Native SEQ Fields**:
  - Tables: `<w:fldSimple w:instr="SEQ Table \* ARABIC"/>`
  - Figures: `<w:fldSimple w:instr="SEQ Figure \* ARABIC"/>`
* **Caption Styling**: Arial 9.5pt, color `#475569`, bold prefix (e.g., `Table 1: `, `Figure 1: `), centered or aligned with the element.
* **Auto-Generated Lists**:
  - Auto-generate a dedicated page for **List of Figures** using `<w:fldSimple w:instr="TOC \c &quot;Figure&quot; \h \z"/>`.
  - Auto-generate a dedicated page for **List of Tables** using `<w:fldSimple w:instr="TOC \c &quot;Table&quot; \h \z"/>`.
  - Each list followed by an explicit page break.

### 8. Editable Footer Page Numbers with Basic Placeholders
* **Footer Location**: Placed in the section footer (right-aligned or centered).
* **Native Field Placeholders**: Built with editable Word fields:
  - `Page ` + `<w:fldSimple w:instr="PAGE"/>` + ` of ` + `<w:fldSimple w:instr="NUMPAGES"/>` (or localized `Halaman {PAGE} dari {NUMPAGES}`).
* **Editable in Word**: Because it resides in standard footer XML, users can double-click the footer in Word, retype text, change typography, or move positioning freely.

---

## Python API Quick Reference

Use the `jago_docx` package directly in scripts or generate docs programmatically:

```python
from jago_docx import JagoDocument, format_apa_reference

# 1. Initialize Document (auto-configures margins, Arial styles, username author, footer)
doc = JagoDocument(
    title="Cloud Architecture Assessment",
    subject="Technical Whitepaper"
)

# 2. Add Title & Subtitle
doc.add_title(
    "Cloud Architecture Assessment",
    subtitle="Zero-Trust Network Topology and Multi-Region Resiliency"
)

# 3. Add Auto-Generated TOC and Figure/Table Lists
doc.add_table_of_contents("Table of Contents")
doc.add_list_of_figures("List of Figures")
doc.add_list_of_tables("List of Tables")

# 4. Add Headings (Arial, MS Docs hierarchy)
doc.add_heading("1. Executive Summary and Architecture", level=1)
doc.add_paragraph("This assessment evaluates high-concurrency microservice scaling.", lead_bold="Key Finding:")

# 5. Add APA Footnote Citation (Short quote inline, or blockquote)
apa_ref = format_apa_reference(
    authors="Lamport, L.",
    year="1978",
    title="Time, clocks, and the ordering of events in a distributed system",
    source="Communications of the ACM",
    volume_issue="21(7)",
    pages="558–565",
    doi_or_url="10.1145/359545.359563"
)
doc.add_quotation_with_reference(
    quote_text="The concept of time in distributed systems is relativistic.",
    apa_reference=apa_ref,
    lead_in="As established in fundamental consensus research,"
)

# 6. Add Native Editable Math Formula
doc.add_heading("2. Queue Depth Mathematical Model", level=1)
doc.add_formula(
    formula_text="L = λ · W",
    equation_number=1,
    description="Little's Law for steady-state queueing latency."
)

# 7. Add Table with Caption STRICTLY UNDERNEATH
headers = ["Cluster Node", "vCPU", "Memory (GiB)", "Max RPS", "p99 Latency"]
rows = [
    ["ingress-edge-01", 16, 64, "42,000", "4.2 ms"],
    ["auth-core-01",    32, 128, "65,000", "2.8 ms"],
    ["ledger-db-01",    64, 256, "18,500", "8.1 ms"]
]
doc.add_table(
    headers=headers,
    rows=rows,
    caption="Core Production Node Hardware Sizing and Verified Throughput Benchmarks."
)

# 8. Add Figure with Caption STRICTLY UNDERNEATH
doc.add_figure(
    image_path_or_placeholder="Diagram: Zero-Trust Identity Gateway Mesh",
    caption="Zero-Trust Perimeter Gateway and Service Mesh Mutual-TLS Flow.",
    width_inches=5.5
)

# 9. Audit and Save
violations = doc.audit()
if not violations:
    print("Document is 100% Antislop compliant!")

doc.save("architecture_assessment.docx")
```

---

## Command Line Interface (CLI)

```bash
# Generate complete showcase sample document
jago-docx sample -o sample_report.docx

# Create a clean, pre-structured blank document with TOC & current username
jago-docx new "Enterprise AI Blueprint" -o blueprint.docx --subtitle "Governance & Security"

# Audit an existing docx for antislop violations (filler words, empty spacers, orphaned headings)
jago-docx audit my_document.docx
```

---

## Pre-Flight Generation Checklist

Before delivering or saving any `.docx` file for the user:
- [ ] Document author in `core_properties.author` is set to the current OS username, NOT `"python-docx"`.
- [ ] All headings (Levels 1–6) use `Arial` with MS Docs hierarchy sizing.
- [ ] Table of Contents is on a dedicated page with page break and `updateFields` flag enabled.
- [ ] Citations are formatted in APA 7th edition and placed as native footnotes in `word/footnotes.xml`.
- [ ] Quotes longer than 40 words are formatted as blockquotes (indented 0.5 inches, no quotation marks).
- [ ] Formulas are native Office Math (`m:oMath`) placeholders that can be clicked and edited in Word.
- [ ] Table and figure captions are located **strictly underneath** each element, tagged with `SEQ Table` / `SEQ Figure`.
- [ ] List of Figures and List of Tables pages are auto-generated.
- [ ] Footer contains editable page number placeholders (`PAGE` and `NUMPAGES`).
- [ ] No empty spacer paragraphs (`<w:p/>` without text) and no AI filler cliches.
