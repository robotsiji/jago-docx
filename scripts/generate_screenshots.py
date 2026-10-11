"""
Production Screenshot Generator for Jago Docx Repository.
Generates 8 high-resolution vertical comparison images (UP = BEFORE, DOWN = AFTER)
representing each of the 8 core capabilities.
"""

import os
from PIL import Image, ImageDraw, ImageFont

# Dimensions
WIDTH = 1380
HEIGHT = 880

# Dark outer frame
BG_DARK = (15, 23, 42)          # Slate 900
BG_HEADER = (30, 41, 59)        # Slate 800
TEXT_WHITE = (255, 255, 255)
TEXT_MUTED = (148, 163, 184)     # Slate 400
LINE_DIVIDER = (51, 65, 85)     # Slate 700

# Before / Red theme
BADGE_RED_BG = (254, 226, 226)
BADGE_RED_BORDER = (248, 113, 113)
BADGE_RED_TEXT = (185, 28, 28)
CARD_BG = (255, 255, 255)
CALLOUT_RED_BG = (254, 242, 242)
CALLOUT_RED_BORDER = (239, 68, 68)

# After / Green theme
BADGE_GREEN_BG = (220, 252, 231)
BADGE_GREEN_BORDER = (74, 222, 128)
BADGE_GREEN_TEXT = (21, 128, 61)
CALLOUT_GREEN_BG = (240, 253, 244)
CALLOUT_GREEN_BORDER = (34, 197, 94)

# Document colors
COLOR_HEADING = (15, 23, 42)    # Slate 900
COLOR_BODY = (31, 41, 55)       # Gray 800
COLOR_DOC_MUTED = (100, 116, 139) # Slate 500
COLOR_CALIBRI_CYAN = (46, 116, 181) # Ugly Word default blue heading

# Fonts
FONTS_DIR = r"C:\Windows\Fonts"
font_title = ImageFont.truetype(os.path.join(FONTS_DIR, "arialbd.ttf"), 22)
font_subtitle = ImageFont.truetype(os.path.join(FONTS_DIR, "arial.ttf"), 14)
font_badge = ImageFont.truetype(os.path.join(FONTS_DIR, "arialbd.ttf"), 12)
font_h1 = ImageFont.truetype(os.path.join(FONTS_DIR, "arialbd.ttf"), 17)
font_h2 = ImageFont.truetype(os.path.join(FONTS_DIR, "arialbd.ttf"), 14)
font_body = ImageFont.truetype(os.path.join(FONTS_DIR, "arial.ttf"), 13)
font_italic = ImageFont.truetype(os.path.join(FONTS_DIR, "ariali.ttf"), 13)
font_calibri_h = ImageFont.truetype(os.path.join(FONTS_DIR, "calibrib.ttf"), 17)
font_calibri_b = ImageFont.truetype(os.path.join(FONTS_DIR, "calibri.ttf"), 13)
font_mono = ImageFont.truetype(os.path.join(FONTS_DIR, "consola.ttf"), 12)
font_tag = ImageFont.truetype(os.path.join(FONTS_DIR, "arialbd.ttf"), 11)
font_note = ImageFont.truetype(os.path.join(FONTS_DIR, "ariali.ttf"), 12)


def init_canvas(cap_num: str, title: str, subtitle: str):
    img = Image.new("RGB", (WIDTH, HEIGHT), color=BG_DARK)
    draw = ImageDraw.Draw(img)

    # Header Bar
    draw.rectangle([0, 0, WIDTH, 80], fill=BG_HEADER)
    draw.line([(0, 80), (WIDTH, 80)], fill=LINE_DIVIDER, width=1)

    # Title & Subtitle
    draw.text((40, 18), f"CAPABILITY {cap_num}: {title.upper()}", font=font_title, fill=TEXT_WHITE)
    draw.text((40, 48), subtitle, font=font_subtitle, fill=TEXT_MUTED)

    # UP Card (BEFORE)
    draw.rounded_rectangle([35, 100, WIDTH - 35, 460], radius=10, fill=CARD_BG, outline=(226, 232, 240), width=1)
    draw.rounded_rectangle([55, 115, 340, 142], radius=6, fill=BADGE_RED_BG, outline=BADGE_RED_BORDER, width=1)
    draw.text((68, 120), "BEFORE: DEFAULT / GENERIC SLOP", font=font_badge, fill=BADGE_RED_TEXT)

    # DOWN Card (AFTER)
    draw.rounded_rectangle([35, 485, WIDTH - 35, 845], radius=10, fill=CARD_BG, outline=(226, 232, 240), width=1)
    draw.rounded_rectangle([55, 500, 330, 527], radius=6, fill=BADGE_GREEN_BG, outline=BADGE_GREEN_BORDER, width=1)
    draw.text((68, 505), "AFTER: JAGO DOCX STANDARD", font=font_badge, fill=BADGE_GREEN_TEXT)

    return img, draw


# =============================================================================
# 1. Antislop Discipline
# =============================================================================
def make_card_01(out_dir):
    img, draw = init_canvas(
        "01",
        "Antislop Discipline (Tone, Spacing, Margins)",
        "Eliminates AI throat-clearing openings, enforces 1.0\" margins, and bans manual empty spacer paragraphs."
    )

    # UP: BEFORE
    draw.text((55, 160), "1. Introduction to Cloud Architectures", font=font_calibri_h, fill=COLOR_CALIBRI_CYAN)
    draw.text((55, 190), "In today's fast-paced digital era, it is crucial to delve into the multifaceted tapestry of modern technology...", font=font_calibri_b, fill=COLOR_BODY)
    draw.text((55, 215), "Furthermore, it is important to remember that cloud computing plays a pivotal role in navigating complexities...", font=font_calibri_b, fill=COLOR_BODY)

    # Visual empty paragraph indicators
    draw.rounded_rectangle([55, 250, 420, 275], radius=4, fill=(254, 226, 226), outline=(239, 68, 68), width=1)
    draw.text((65, 254), "[X] <w:p/> Manual Blank Paragraph Spacer (Double Enter)", font=font_tag, fill=BADGE_RED_TEXT)

    draw.rounded_rectangle([55, 285, 420, 310], radius=4, fill=(254, 226, 226), outline=(239, 68, 68), width=1)
    draw.text((65, 289), "[X] <w:p/> Another Empty Spacer Paragraph", font=font_tag, fill=BADGE_RED_TEXT)

    draw.text((55, 330), "In conclusion, only time will tell how distributed systems will continue to evolve.", font=font_calibri_b, fill=COLOR_BODY)

    # Tag callout
    draw.rounded_rectangle([780, 150, WIDTH - 60, 430], radius=6, fill=CALLOUT_RED_BG, outline=CALLOUT_RED_BORDER, width=1)
    draw.text((800, 165), "SLOP PATTERNS DETECTED:", font=font_badge, fill=BADGE_RED_TEXT)
    draw.text((800, 195), "• Throat-clearing cliches: 'In today's fast-paced digital era...'", font=font_body, fill=COLOR_BODY)
    draw.text((800, 225), "• Banned AI buzzwords: 'delve into', 'tapestry', 'pivotal role'", font=font_body, fill=COLOR_BODY)
    draw.text((800, 255), "• Empty paragraphs used for vertical spacing instead of styles", font=font_body, fill=COLOR_BODY)
    draw.text((800, 285), "• Inconsistent page margins and amateur line spacing", font=font_body, fill=COLOR_BODY)
    draw.text((800, 340), "Status: 4 Antislop Violations Flagged", font=font_mono, fill=BADGE_RED_TEXT)

    # DOWN: AFTER
    draw.text((55, 545), "1. Executive Overview and Engineering Constraints", font=font_h1, fill=COLOR_HEADING)
    draw.text((55, 575), "Distributed service meshes enforce partition isolation and sub-15ms p99 latency SLAs under high throughput.", font=font_body, fill=COLOR_BODY)
    
    # Callout Box
    draw.rounded_rectangle([55, 610, 720, 675], radius=4, fill=(248, 250, 252), outline=(2, 132, 199), width=2)
    draw.text((70, 620), "CRITICAL SLA DIRECTIVE", font=font_badge, fill=COLOR_HEADING)
    draw.text((70, 642), "All ingress-to-pod communications must maintain p99 latency < 15ms under 50,000 RPS peak load.", font=font_body, fill=COLOR_BODY)

    draw.text((55, 700), "Telemetry pipelines aggregate OpenTelemetry traces into cold storage for automated audit validation.", font=font_body, fill=COLOR_BODY)

    # Green Callout
    draw.rounded_rectangle([780, 535, WIDTH - 60, 815], radius=6, fill=CALLOUT_GREEN_BG, outline=CALLOUT_GREEN_BORDER, width=1)
    draw.text((800, 550), "JAGO DOCX ENFORCEMENT:", font=font_badge, fill=BADGE_GREEN_TEXT)
    draw.text((800, 580), "• Zero fluff: Concrete metrics, active voice, dense engineering facts", font=font_body, fill=COLOR_BODY)
    draw.text((800, 610), "• Strict typographic rhythm: 6pt space-after, 1.15 line spacing", font=font_body, fill=COLOR_BODY)
    draw.text((800, 640), "• Zero empty paragraphs: all spacing handled via style format", font=font_body, fill=COLOR_BODY)
    draw.text((800, 670), "• Standard 1.0\" (72pt) margins on all 4 document borders", font=font_body, fill=COLOR_BODY)
    draw.text((800, 725), "Status: 100% Antislop Audit Compliant", font=font_mono, fill=BADGE_GREEN_TEXT)

    path = os.path.join(out_dir, "01_antislop.png")
    img.save(path, "PNG")
    print(f"Saved: {path}")


# =============================================================================
# 2. Arial Headings & Auto-Generated TOC
# =============================================================================
def make_card_02(out_dir):
    img, draw = init_canvas(
        "02",
        "Arial Headings & Auto-Generated Table of Contents",
        "All headings strictly use Arial with MS Docs hierarchy; dedicated TOC page with native auto-refresh fields."
    )

    # UP: BEFORE
    draw.text((55, 160), "Heading 1 (Default Calibri Light - Blue Cyan Accent)", font=font_calibri_h, fill=COLOR_CALIBRI_CYAN)
    draw.text((55, 195), "Calibri 11pt unstyled text. No outline level configured.", font=font_calibri_b, fill=COLOR_BODY)
    draw.text((55, 235), "[MISSING TABLE OF CONTENTS]", font=font_calibri_h, fill=BADGE_RED_TEXT)
    draw.text((55, 265), "Static plain-text typed outline that does not update when pages change:", font=font_calibri_b, fill=COLOR_DOC_MUTED)
    draw.text((55, 290), "• Chapter 1 ........................ Page 1 (Hardcoded, stale)", font=font_mono, fill=COLOR_BODY)
    draw.text((55, 315), "• Chapter 2 ........................ Page 2 (Hardcoded, stale)", font=font_mono, fill=COLOR_BODY)

    draw.rounded_rectangle([780, 150, WIDTH - 60, 430], radius=6, fill=CALLOUT_RED_BG, outline=CALLOUT_RED_BORDER, width=1)
    draw.text((800, 165), "DEFICIENCIES:", font=font_badge, fill=BADGE_RED_TEXT)
    draw.text((800, 195), "• Default saturated Calibri cyan headings look amateurish", font=font_body, fill=COLOR_BODY)
    draw.text((800, 225), "• Missing native Word Table of Contents field code", font=font_body, fill=COLOR_BODY)
    draw.text((800, 255), "• Manual typed page numbers break as soon as content shifts", font=font_body, fill=COLOR_BODY)
    draw.text((800, 285), "• Word Navigation Pane cannot parse hierarchy", font=font_body, fill=COLOR_BODY)

    # DOWN: AFTER
    draw.text((55, 545), "Table of Contents", font=font_h1, fill=COLOR_HEADING)
    draw.text((55, 580), "1. Executive Overview and Engineering Constraints .................. 5", font=font_body, fill=COLOR_BODY)
    draw.text((55, 605), "2. Foundational Research & Theoretical Grounding ................. 6", font=font_body, fill=COLOR_BODY)
    draw.text((55, 630), "3. Concurrency and Queueing Formulations ......................... 7", font=font_body, fill=COLOR_BODY)
    draw.text((75, 655), "3.1 Service Queue Equilibrium ..................................... 7", font=font_body, fill=COLOR_DOC_MUTED)
    draw.text((75, 680), "3.2 Amdahl's Speedup Model ........................................ 8", font=font_body, fill=COLOR_DOC_MUTED)
    draw.text((55, 705), "4. Performance Benchmarks and Topology ............................ 9", font=font_body, fill=COLOR_BODY)

    draw.rounded_rectangle([55, 745, 540, 775], radius=4, fill=(241, 245, 249), outline=(148, 163, 184), width=1)
    draw.text((65, 752), "OpenXML: <w:fldSimple w:instr=\"TOC \\o 1-3 \\h \\z \\u\"/>", font=font_mono, fill=COLOR_HEADING)

    draw.rounded_rectangle([780, 535, WIDTH - 60, 815], radius=6, fill=CALLOUT_GREEN_BG, outline=CALLOUT_GREEN_BORDER, width=1)
    draw.text((800, 550), "JAGO DOCX AUTOMATION:", font=font_badge, fill=BADGE_GREEN_TEXT)
    draw.text((800, 580), "• Arial Bold font family applied to all heading levels (1-6)", font=font_body, fill=COLOR_BODY)
    draw.text((800, 610), "• Dedicated TOC page with page break after Title block", font=font_body, fill=COLOR_BODY)
    draw.text((800, 640), "• Native Word TOC field code with leader tab dots & hyperlinking", font=font_body, fill=COLOR_BODY)
    draw.text((800, 670), "• w:updateFields enabled: auto-refreshes page numbers on open", font=font_body, fill=COLOR_BODY)
    draw.text((800, 725), "Status: 100% Native Word TOC Automation", font=font_mono, fill=BADGE_GREEN_TEXT)

    path = os.path.join(out_dir, "02_arial_headings_toc.png")
    img.save(path, "PNG")
    print(f"Saved: {path}")


# =============================================================================
# 3. Dynamic OS Username as Author
# =============================================================================
def make_card_03(out_dir):
    img, draw = init_canvas(
        "03",
        "Dynamic OS Username Metadata (Never 'python-docx')",
        "Detects logged-in system user as document creator, strictly purging default 'python-docx' traces."
    )

    # UP: BEFORE
    draw.rounded_rectangle([55, 160, 700, 360], radius=6, fill=(248, 250, 252), outline=(226, 232, 240), width=1)
    draw.text((75, 175), "Document Properties (docProps/core.xml):", font=font_h2, fill=COLOR_HEADING)
    draw.text((75, 210), "Title:            Untitled Document", font=font_mono, fill=COLOR_BODY)
    draw.text((75, 235), "Author:           python-docx   [X]", font=font_mono, fill=BADGE_RED_TEXT)
    draw.text((75, 260), "Last Modified By: python-docx   [X]", font=font_mono, fill=BADGE_RED_TEXT)
    draw.text((75, 285), "Creator:          python-docx   [X]", font=font_mono, fill=BADGE_RED_TEXT)
    draw.text((75, 310), "Company:          Unknown", font=font_mono, fill=COLOR_DOC_MUTED)

    draw.rounded_rectangle([780, 150, WIDTH - 60, 430], radius=6, fill=CALLOUT_RED_BG, outline=CALLOUT_RED_BORDER, width=1)
    draw.text((800, 165), "METADATA DEFICIENCIES:", font=font_badge, fill=BADGE_RED_TEXT)
    draw.text((800, 195), "• Default python-docx branding leaks into document metadata", font=font_body, fill=COLOR_BODY)
    draw.text((800, 225), "• Instantly reveals to clients/evaluators that doc was generated", font=font_body, fill=COLOR_BODY)
    draw.text((800, 255), "• Core properties remain blank or unformatted", font=font_body, fill=COLOR_BODY)

    # DOWN: AFTER
    draw.rounded_rectangle([55, 545, 700, 755], radius=6, fill=(248, 250, 252), outline=(226, 232, 240), width=1)
    draw.text((75, 560), "Document Properties (docProps/core.xml):", font=font_h2, fill=COLOR_HEADING)
    draw.text((75, 595), "Title:            High-Performance Microservice Mesh Architecture", font=font_mono, fill=COLOR_BODY)
    draw.text((75, 620), "Author:           robotsiji (Current OS Username)   •", font=font_mono, fill=BADGE_GREEN_TEXT)
    draw.text((75, 645), "Last Modified By: robotsiji   •", font=font_mono, fill=BADGE_GREEN_TEXT)
    draw.text((75, 670), "Creator:          robotsiji   •", font=font_mono, fill=BADGE_GREEN_TEXT)
    draw.text((75, 695), "Subject:          Technical Architecture Specification", font=font_mono, fill=COLOR_BODY)

    draw.rounded_rectangle([780, 535, WIDTH - 60, 815], radius=6, fill=CALLOUT_GREEN_BG, outline=CALLOUT_GREEN_BORDER, width=1)
    draw.text((800, 550), "JAGO DOCX METADATA HYGIENE:", font=font_badge, fill=BADGE_GREEN_TEXT)
    draw.text((800, 580), "• Dynamic OS username detection (getpass.getuser() / $USERNAME)", font=font_body, fill=COLOR_BODY)
    draw.text((800, 610), "• Strict scrubbing: 'python-docx' string NEVER appears in metadata", font=font_body, fill=COLOR_BODY)
    draw.text((800, 640), "• Professional enterprise metadata formatting across all core properties", font=font_body, fill=COLOR_BODY)
    draw.text((800, 695), "Author Verified: robotsiji", font=font_mono, fill=BADGE_GREEN_TEXT)

    path = os.path.join(out_dir, "03_author_metadata.png")
    img.save(path, "PNG")
    print(f"Saved: {path}")


# =============================================================================
# 4. Scholarly Quotations & APA 7th Footnotes
# =============================================================================
def make_card_04(out_dir):
    img, draw = init_canvas(
        "04",
        "Scholarly Quotations & APA 7th Edition Footnotes",
        "Formats short quotes inline and long quotes as 0.5\" blockquotes; places citations as native Word footnotes."
    )

    # UP: BEFORE
    draw.text((55, 160), "2. Literature Review", font=font_calibri_h, fill=COLOR_CALIBRI_CYAN)
    draw.text((55, 195), "As stated by Lamport: \"The concept of the time at which an event occurred is fundamentally relativistic\"", font=font_calibri_b, fill=COLOR_BODY)
    draw.text((55, 220), "(Lamport 1978 ACM paper page 558-565). And then another researcher also said that state machines are", font=font_calibri_b, fill=COLOR_BODY)
    draw.text((55, 245), "\"Global state recording algorithms ensure state consistency without halting processing\" (Chandy and Lamport 1985).", font=font_calibri_b, fill=COLOR_BODY)
    draw.text((55, 280), "[X] Citations dumped inline as unstructured text; no interactive footnotes.", font=font_tag, fill=BADGE_RED_TEXT)

    draw.rounded_rectangle([780, 150, WIDTH - 60, 430], radius=6, fill=CALLOUT_RED_BG, outline=CALLOUT_RED_BORDER, width=1)
    draw.text((800, 165), "CITATION DEFICIENCIES:", font=font_badge, fill=BADGE_RED_TEXT)
    draw.text((800, 195), "• Citations clumped inline, breaking reading rhythm", font=font_body, fill=COLOR_BODY)
    draw.text((800, 225), "• Long quotes lack 0.5\" blockquote indentation", font=font_body, fill=COLOR_BODY)
    draw.text((800, 255), "• No native Word footnote part (word/footnotes.xml)", font=font_body, fill=COLOR_BODY)
    draw.text((800, 285), "• Non-compliant citation formatting (missing DOI/volume/issue)", font=font_body, fill=COLOR_BODY)

    # DOWN: AFTER
    draw.text((55, 545), "2. Foundational Research & Theoretical Grounding", font=font_h1, fill=COLOR_HEADING)
    draw.text((55, 575), "As demonstrated in foundational consensus research,", font=font_body, fill=COLOR_BODY)
    draw.text((55, 595), "\"The concept of the time at which an event occurred is fundamentally relativistic.\"^1", font=font_body, fill=COLOR_BODY)

    # Blockquote (>40 words)
    draw.rectangle([75, 625, 77, 685], fill=(2, 132, 199)) # Left bar
    draw.text((85, 625), "A distributed algorithm is designed to execute on a network of processors. Global state", font=font_body, fill=COLOR_BODY)
    draw.text((85, 645), "recording algorithms ensure that state consistency can be checkpointed without halting", font=font_body, fill=COLOR_BODY)
    draw.text((85, 665), "ongoing transaction processing, providing high availability during active state migration.^2", font=font_body, fill=COLOR_BODY)

    # Footnote separator & footnotes
    draw.line([(55, 720), (220, 720)], fill=(148, 163, 184), width=1)
    draw.text((55, 730), "^1 Lamport, L. (1978). Time, clocks, and the ordering of events in a distributed system. CACM, 21(7), 558–565.", font=font_note, fill=COLOR_DOC_MUTED)
    draw.text((55, 750), "^2 Chandy, K. M., & Lamport, L. (1985). Distributed snapshots. ACM TOCS, 3(1), 63–75. https://doi.org/10.1145/214451", font=font_note, fill=COLOR_DOC_MUTED)

    draw.rounded_rectangle([780, 535, WIDTH - 60, 815], radius=6, fill=CALLOUT_GREEN_BG, outline=CALLOUT_GREEN_BORDER, width=1)
    draw.text((800, 550), "JAGO DOCX APA 7TH PROTOCOL:", font=font_badge, fill=BADGE_GREEN_TEXT)
    draw.text((800, 580), "• Short quotes inline with double quotes + superscript link", font=font_body, fill=COLOR_BODY)
    draw.text((800, 610), "• Blockquotes (>40 words) indented 0.5\" without quotation marks", font=font_body, fill=COLOR_BODY)
    draw.text((800, 640), "• True native Word footnotes injected into word/footnotes.xml", font=font_body, fill=COLOR_BODY)
    draw.text((800, 670), "• Strict APA 7th format: Author, Year, Title, Journal, DOI/URL", font=font_body, fill=COLOR_BODY)

    path = os.path.join(out_dir, "04_apa_footnotes.png")
    img.save(path, "PNG")
    print(f"Saved: {path}")


# =============================================================================
# 5. Microsoft Docs Heading Hierarchy
# =============================================================================
def make_card_05(out_dir):
    img, draw = init_canvas(
        "05",
        "Microsoft Docs Heading Hierarchy & Outline Levels",
        "Disciplined typographic scale (H1: 18pt, H2: 14pt, H3: 12pt); keep_with_next prevents orphaned titles."
    )

    # UP: BEFORE
    draw.text((55, 160), "Random Heading (Font size 15.5pt, arbitrary color)", font=font_calibri_h, fill=(168, 85, 247))
    draw.text((55, 190), "Body text separated by 24pt random gap...", font=font_calibri_b, fill=COLOR_BODY)
    draw.text((55, 230), "Subheading with no outline level", font=font_calibri_b, fill=(2, 132, 199))
    draw.text((55, 270), "[X] Missing <w:outlineLvl> - Word Navigation Pane is empty", font=font_tag, fill=BADGE_RED_TEXT)
    draw.text((55, 295), "[X] Missing keep_with_next - Heading gets orphaned at page bottoms", font=font_tag, fill=BADGE_RED_TEXT)

    draw.rounded_rectangle([780, 150, WIDTH - 60, 430], radius=6, fill=CALLOUT_RED_BG, outline=CALLOUT_RED_BORDER, width=1)
    draw.text((800, 165), "TYPOGRAPHY DEFICIENCIES:", font=font_badge, fill=BADGE_RED_TEXT)
    draw.text((800, 195), "• Inconsistent font sizing without proportional typographic scale", font=font_body, fill=COLOR_BODY)
    draw.text((800, 225), "• Headings lack keep_with_next, stranding titles at page bottoms", font=font_body, fill=COLOR_BODY)
    draw.text((800, 255), "• Word sidebar Navigation Pane cannot construct document tree", font=font_body, fill=COLOR_BODY)

    # DOWN: AFTER
    draw.text((55, 545), "1. Executive Overview and Engineering Constraints (H1: 18pt Arial)", font=font_h1, fill=COLOR_HEADING)
    draw.text((55, 575), "Top-level section heading with outlineLvl 0 and keep_with_next=True.", font=font_body, fill=COLOR_BODY)
    draw.text((55, 610), "1.1 Service Queue Equilibrium (H2: 14pt Arial Bold)", font=font_h2, fill=COLOR_HEADING)
    draw.text((55, 635), "Second-level subsection with outlineLvl 1 and proportional spacing.", font=font_body, fill=COLOR_BODY)
    draw.text((55, 670), "1.1.1 Concurrency Boundary Model (H3: 12pt Arial Bold)", font=font_badge, fill=COLOR_HEADING)
    draw.text((55, 695), "Third-level heading with outlineLvl 2 for clean nesting in Navigation Pane.", font=font_body, fill=COLOR_BODY)

    draw.rounded_rectangle([780, 535, WIDTH - 60, 815], radius=6, fill=CALLOUT_GREEN_BG, outline=CALLOUT_GREEN_BORDER, width=1)
    draw.text((800, 550), "MS DOCS TYPOGRAPHIC DISCIPLINE:", font=font_badge, fill=BADGE_GREEN_TEXT)
    draw.text((800, 580), "• Proportional MS Docs scale: H1(18pt), H2(14pt), H3(12pt), H4(11pt)", font=font_body, fill=COLOR_BODY)
    draw.text((800, 610), "• High-contrast slate hierarchy: #0F172A, #1E293B, #334155", font=font_body, fill=COLOR_BODY)
    draw.text((800, 640), "• keep_with_next=True prevents orphaned headings at page breaks", font=font_body, fill=COLOR_BODY)
    draw.text((800, 670), "• Explicit <w:outlineLvl> powers Word's Navigation Pane & TOC", font=font_body, fill=COLOR_BODY)

    path = os.path.join(out_dir, "05_msdocs_hierarchy.png")
    img.save(path, "PNG")
    print(f"Saved: {path}")


# =============================================================================
# 6. Native Editable Math Formulas
# =============================================================================
def make_card_06(out_dir):
    img, draw = init_canvas(
        "06",
        "Native Editable Math Formulas (OMML / Equation Editor)",
        "Leverages Office Math Markup Language; formulas remain 100% editable inside Word's Equation Editor."
    )

    # UP: BEFORE
    draw.rounded_rectangle([55, 160, 650, 310], radius=6, fill=(248, 250, 252), outline=(239, 68, 68), width=1)
    draw.text((75, 175), "Equation (Static Bitmap Screenshot / Low-Res PNG):", font=font_h2, fill=BADGE_RED_TEXT)
    # Simulate blurry formula image
    draw.rectangle([180, 215, 520, 265], fill=(226, 232, 240), outline=(148, 163, 184))
    draw.text((220, 230), "L = λ · W   (Raster Image)", font=font_mono, fill=(100, 116, 139))
    draw.text((75, 280), "[X] Uneditable image bitmap; pixelated on print and high-DPI monitors.", font=font_tag, fill=BADGE_RED_TEXT)

    draw.rounded_rectangle([780, 150, WIDTH - 60, 430], radius=6, fill=CALLOUT_RED_BG, outline=CALLOUT_RED_BORDER, width=1)
    draw.text((800, 165), "BITMAP FORMULA FLAWS:", font=font_badge, fill=BADGE_RED_TEXT)
    draw.text((800, 195), "• Formulas pasted as images cannot be edited in Word", font=font_body, fill=COLOR_BODY)
    draw.text((800, 225), "• Blurry and pixelated when printed or zoomed", font=font_body, fill=COLOR_BODY)
    draw.text((800, 255), "• Cannot copy variables, adjust exponents, or modify equations", font=font_body, fill=COLOR_BODY)

    # DOWN: AFTER
    draw.rounded_rectangle([55, 545, 680, 715], radius=6, fill=(248, 250, 252), outline=(34, 197, 94), width=1)
    draw.text((75, 560), "Native Word Equation (Office Math OMML):", font=font_h2, fill=BADGE_GREEN_TEXT)
    
    # Render crisp equation block
    draw.rounded_rectangle([120, 600, 610, 650], radius=4, fill=(255, 255, 255), outline=(186, 230, 253), width=1)
    draw.text((240, 615), "L = λ · W", font=ImageFont.truetype(os.path.join(FONTS_DIR, "cambria.ttc"), 18), fill=COLOR_HEADING)
    draw.text((560, 618), "(1)", font=font_mono, fill=COLOR_DOC_MUTED)

    draw.text((75, 675), "Description: Little's Law relating queue size (L), arrival rate (λ), and wait time (W).", font=font_note, fill=COLOR_DOC_MUTED)

    draw.rounded_rectangle([780, 535, WIDTH - 60, 815], radius=6, fill=CALLOUT_GREEN_BG, outline=CALLOUT_GREEN_BORDER, width=1)
    draw.text((800, 550), "JAGO DOCX OMML ENGINE:", font=font_badge, fill=BADGE_GREEN_TEXT)
    draw.text((800, 580), "• Native Office Math: <m:oMathPara> & <m:oMath> XML tags", font=font_body, fill=COLOR_BODY)
    draw.text((800, 610), "• 100% editable in Microsoft Word's Equation Editor ribbon", font=font_body, fill=COLOR_BODY)
    draw.text((800, 640), "• Center-aligned display formula with right-aligned numbering (1)", font=font_body, fill=COLOR_BODY)
    draw.text((800, 670), "• Infinite vector sharpness on print and high-DPI displays", font=font_body, fill=COLOR_BODY)

    path = os.path.join(out_dir, "06_editable_formulas.png")
    img.save(path, "PNG")
    print(f"Saved: {path}")


# =============================================================================
# 7. Captions Strictly Underneath Visuals
# =============================================================================
def make_card_07(out_dir):
    img, draw = init_canvas(
        "07",
        "Captions Strictly Underneath Visuals & Auto Figure/Table Lists",
        "All table and figure captions placed directly underneath; native SEQ fields generate List of Figures/Tables."
    )

    # UP: BEFORE
    draw.text((55, 160), "Table: Some Performance Numbers (Caption Placed Above Table)", font=font_calibri_b, fill=BADGE_RED_TEXT)
    draw.rectangle([55, 185, 550, 245], outline=(203, 213, 225), fill=(248, 250, 252))
    draw.text((70, 205), "[Data Table]   Metric A: 14.2ms | Metric B: 99.8%", font=font_mono, fill=COLOR_BODY)
    draw.text((55, 260), "[X] Inconsistent caption placement breaks visual flow", font=font_tag, fill=BADGE_RED_TEXT)
    draw.text((55, 285), "[X] Missing List of Figures / List of Tables pages", font=font_tag, fill=BADGE_RED_TEXT)

    draw.rounded_rectangle([780, 150, WIDTH - 60, 430], radius=6, fill=CALLOUT_RED_BG, outline=CALLOUT_RED_BORDER, width=1)
    draw.text((800, 165), "CAPTION DEFICIENCIES:", font=font_badge, fill=BADGE_RED_TEXT)
    draw.text((800, 195), "• Captions placed haphazardly above tables and below figures", font=font_body, fill=COLOR_BODY)
    draw.text((800, 225), "• Missing native SEQ fields for automated numbering", font=font_body, fill=COLOR_BODY)
    draw.text((800, 255), "• No auto-generated List of Figures or List of Tables", font=font_body, fill=COLOR_BODY)

    # DOWN: AFTER
    # Table first
    draw.rectangle([55, 545, 600, 615], outline=(226, 232, 240), fill=(255, 255, 255))
    draw.rectangle([55, 545, 600, 575], fill=(248, 250, 252))
    draw.text((70, 553), "Service Mesh Pod      Peak RPS      p50 Latency      p99 Latency", font=font_mono, fill=COLOR_HEADING)
    draw.text((70, 585), "edge-gateway-alpha    52,400        1.4 ms           5.8 ms", font=font_mono, fill=COLOR_BODY)
    
    # Caption STRICTLY UNDERNEATH
    draw.text((70, 625), "Table 1: Simulated High-Concurrency Throughput across Primary Pods.", font=font_badge, fill=BADGE_GREEN_TEXT)

    # List of Tables Page preview
    draw.rounded_rectangle([55, 665, 600, 755], radius=4, fill=(248, 250, 252), outline=(148, 163, 184), width=1)
    draw.text((70, 675), "Auto-Generated Front Matter Pages:", font=font_h2, fill=COLOR_HEADING)
    draw.text((70, 702), "• List of Figures (TOC \\c \"Figure\" \\h \\z) ................. Page 3", font=font_body, fill=COLOR_BODY)
    draw.text((70, 725), "• List of Tables (TOC \\c \"Table\" \\h \\z) .................. Page 4", font=font_body, fill=COLOR_BODY)

    draw.rounded_rectangle([780, 535, WIDTH - 60, 815], radius=6, fill=CALLOUT_GREEN_BG, outline=CALLOUT_GREEN_BORDER, width=1)
    draw.text((800, 550), "JAGO DOCX CAPTION STANDARD:", font=font_badge, fill=BADGE_GREEN_TEXT)
    draw.text((800, 580), "• Strict placement rule: Captions ALWAYS directly underneath visuals", font=font_body, fill=COLOR_BODY)
    draw.text((800, 610), "• Native Word SEQ fields: SEQ Table & SEQ Figure auto-numbered", font=font_body, fill=COLOR_BODY)
    draw.text((800, 640), "• Dedicated List of Figures page with page break", font=font_body, fill=COLOR_BODY)
    draw.text((800, 670), "• Dedicated List of Tables page with page break", font=font_body, fill=COLOR_BODY)

    path = os.path.join(out_dir, "07_captions_underneath.png")
    img.save(path, "PNG")
    print(f"Saved: {path}")


# =============================================================================
# 8. Editable Footer Page Numbers
# =============================================================================
def make_card_08(out_dir):
    img, draw = init_canvas(
        "08",
        "Editable Footer Page Numbers (PAGE / NUMPAGES)",
        "Injects native Word footer field placeholders; users can double-click and edit layout directly in Word."
    )

    # UP: BEFORE
    draw.rounded_rectangle([55, 160, 650, 310], radius=6, fill=(248, 250, 252), outline=(239, 68, 68), width=1)
    draw.text((75, 175), "Footer (Static Text):", font=font_h2, fill=BADGE_RED_TEXT)
    draw.text((75, 215), "\"Page 1 of 5\" (Static String)", font=font_mono, fill=COLOR_BODY)
    draw.text((75, 245), "[X] On Page 2, footer still says 'Page 1 of 5'!", font=font_tag, fill=BADGE_RED_TEXT)
    draw.text((75, 270), "[X] On Page 3, footer still says 'Page 1 of 5'!", font=font_tag, fill=BADGE_RED_TEXT)

    draw.rounded_rectangle([780, 150, WIDTH - 60, 430], radius=6, fill=CALLOUT_RED_BG, outline=CALLOUT_RED_BORDER, width=1)
    draw.text((800, 165), "FOOTER DEFICIENCIES:", font=font_badge, fill=BADGE_RED_TEXT)
    draw.text((800, 195), "• Hardcoded strings in footer do not increment across pages", font=font_body, fill=COLOR_BODY)
    draw.text((800, 225), "• Missing native Word PAGE and NUMPAGES field instructions", font=font_body, fill=COLOR_BODY)
    draw.text((800, 255), "• Breaks whenever content grows or shrinks", font=font_body, fill=COLOR_BODY)

    # DOWN: AFTER
    draw.rounded_rectangle([55, 545, 680, 745], radius=6, fill=(248, 250, 252), outline=(34, 197, 94), width=1)
    draw.text((75, 560), "Native Word Footer (OpenXML Field Codes):", font=font_h2, fill=BADGE_GREEN_TEXT)
    
    draw.rounded_rectangle([75, 600, 650, 665], radius=4, fill=(255, 255, 255), outline=(148, 163, 184), width=1)
    draw.text((95, 615), "Footer: Page <w:fldSimple w:instr=\"PAGE\"/> of <w:fldSimple w:instr=\"NUMPAGES\"/>", font=font_mono, fill=COLOR_HEADING)
    draw.text((95, 638), "Computed View: Page 1 of 6 | Page 2 of 6 | Page 3 of 6 ...", font=font_mono, fill=BADGE_GREEN_TEXT)

    draw.text((75, 695), "• 100% Editable: Double-click footer in Word to change font, alignment, or text.", font=font_body, fill=COLOR_BODY)

    draw.rounded_rectangle([780, 535, WIDTH - 60, 815], radius=6, fill=CALLOUT_GREEN_BG, outline=CALLOUT_GREEN_BORDER, width=1)
    draw.text((800, 550), "JAGO DOCX DYNAMIC NUMBERING:", font=font_badge, fill=BADGE_GREEN_TEXT)
    draw.text((800, 580), "• Native OpenXML fields: PAGE (current) and NUMPAGES (total count)", font=font_body, fill=COLOR_BODY)
    draw.text((800, 610), "• Dynamically increments across every page automatically", font=font_body, fill=COLOR_BODY)
    draw.text((800, 640), "• Resides in standard footer XML for frictionless user customization", font=font_body, fill=COLOR_BODY)
    draw.text((800, 670), "• Supports different first page header/footer for cover pages", font=font_body, fill=COLOR_BODY)

    path = os.path.join(out_dir, "08_editable_page_numbers.png")
    img.save(path, "PNG")
    print(f"Saved: {path}")


def generate_all_screenshots():
    out_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "docs", "images"))
    os.makedirs(out_dir, exist_ok=True)

    print("Generating 8 before-and-after comparison screenshots for repo...")
    make_card_01(out_dir)
    make_card_02(out_dir)
    make_card_03(out_dir)
    make_card_04(out_dir)
    make_card_05(out_dir)
    make_card_06(out_dir)
    make_card_07(out_dir)
    make_card_08(out_dir)
    print(f"[OK] All 8 comparison screenshots successfully saved in: {out_dir}")
    return out_dir


if __name__ == "__main__":
    generate_all_screenshots()
