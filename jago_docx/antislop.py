"""
Antislop Engine and Document Guard for Jago Docx.
Enforces strict quality standards on prose and layout:
- Eliminates AI buzzwords, throat-clearing openings, and filler prose
- Bans manual empty spacer paragraphs (w:p with no text)
- Enforces consistent 1-inch margins, disciplined typographic scales, and restrained neutral palettes.
"""

import re
from typing import List, Dict, Tuple
from docx.shared import Inches, Pt
from docx.enum.section import WD_SECTION_START

# List of banned generic AI filler phrases and cliches
BANNED_AI_PHRASES = [
    r"\bin today'?s (?:fast-paced|ever-evolving|rapidly changing|digital) (?:world|era|landscape)\b",
    r"\bdelve into\b",
    r"\btapestry of\b",
    r"\btestament to\b",
    r"\bbeacon of\b",
    r"\bit is important to (?:remember|note|highlight|keep in mind)\b",
    r"\bfurthermore, it is crucial to\b",
    r"\bnavigating the complexities of\b",
    r"\bunlocking the potential of\b",
    r"\ba holistic approach to\b",
    r"\bin conclusion, as we look (?:forward|ahead|to the future)\b",
    r"\bplays a pivotal role in\b",
    r"\bharnessing the power of\b",
]


class AntislopViolation:
    def __init__(self, category: str, message: str, context: str = ""):
        self.category = category
        self.message = message
        self.context = context

    def __repr__(self):
        return f"[{self.category}] {self.message} ({self.context})"


def sanitize_prose(text: str) -> str:
    """
    Strips or flags common AI boilerplate and cliches.
    Replaces common filler phrases with concise, direct alternatives.
    """
    cleaned = text
    # Direct replacements
    replacements = {
        "In today's fast-paced digital era, ": "",
        "In today's fast-paced world, ": "",
        "In the rapidly evolving landscape of ": "In ",
        "It is important to note that ": "",
        "Furthermore, it is crucial to highlight that ": "Notably, ",
        "plays a pivotal role in": "is essential for",
        "delve into": "examine",
        "harnessing the power of": "utilizing",
        "a holistic approach to": "a comprehensive strategy for",
    }
    for pattern, repl in replacements.items():
        cleaned = re.sub(re.escape(pattern), repl, cleaned, flags=re.IGNORECASE)
        
    return cleaned


def enforce_antislop_page_setup(doc, margin_inches: float = 1.0):
    """
    Standardizes page setup across all sections to 1.0-inch margins (72pt)
    preventing inconsistent, amateurish page dimensions.
    """
    for section in doc.sections:
        section.top_margin = Inches(margin_inches)
        section.bottom_margin = Inches(margin_inches)
        section.left_margin = Inches(margin_inches)
        section.right_margin = Inches(margin_inches)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)


def audit_document_antislop(doc) -> List[AntislopViolation]:
    """
    Audits a python-docx Document object for antislop violations:
    1. Empty paragraphs used as vertical spacing spacers.
    2. AI filler words in paragraph text.
    3. Missing keep_with_next on heading paragraphs.
    """
    violations: List[AntislopViolation] = []

    consecutive_empty = 0
    for p_idx, p in enumerate(doc.paragraphs):
        text = p.text.strip()

        # Check for empty spacer paragraphs
        if not text and len(p.runs) == 0:
            consecutive_empty += 1
            if consecutive_empty >= 1:
                violations.append(
                    AntislopViolation(
                        category="Layout-Slop",
                        message=f"Empty paragraph at index {p_idx} used as manual vertical spacer. Use paragraph space_after/space_before instead.",
                        context=f"Paragraph #{p_idx}"
                    )
                )
        else:
            consecutive_empty = 0

            # Check for banned AI phrases
            for pattern in BANNED_AI_PHRASES:
                match = re.search(pattern, text, flags=re.IGNORECASE)
                if match:
                    violations.append(
                        AntislopViolation(
                            category="Tone-Slop",
                            message=f"Detected generic AI phrase: '{match.group(0)}'",
                            context=text[:80] + "..." if len(text) > 80 else text
                        )
                    )

            # Check headings keep_with_next
            if p.style and "Heading" in p.style.name:
                if not p.paragraph_format.keep_with_next:
                    violations.append(
                        AntislopViolation(
                            category="Typography-Slop",
                            message=f"Heading '{text[:30]}' lacks keep_with_next=True. May orphan at page bottoms.",
                            context=p.style.name
                        )
                    )

    return violations
