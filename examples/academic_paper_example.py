"""
Academic Paper Example using Jago Docx.
Demonstrates scholarly research layout with abstract, APA 7th footnotes,
blockquotes, native editable math formulas, tables, and figures.
"""

import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from jago_docx import JagoDocument, format_apa_reference


def build_academic_paper():
    out_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(out_dir, "academic_paper_sample.docx")

    doc = JagoDocument(
        title="Attention Mechanism Efficiency in Ultra-Low Parameter Architectures",
        subject="Deep Learning & Neural Systems Research"
    )

    doc.add_title(
        "Attention Mechanism Efficiency in Ultra-Low Parameter Architectures",
        subtitle="A Comparative Analysis of Linear Attention, State Space Models, and Sparse FlashAttention"
    )

    # Front Matter: TOC, Figures, Tables
    doc.add_table_of_contents("Contents")
    doc.add_list_of_figures("List of Figures")
    doc.add_list_of_tables("List of Tables")

    # Abstract
    doc.add_heading("Abstract", level=1)
    doc.add_paragraph(
        "Transformer-based language models exhibit quadratic time complexity relative to sequence length. "
        "This empirical investigation examines linear attention formulations and state space models (SSM) "
        "operating within resource-constrained edge compute budgets (< 250M parameters). "
        "We demonstrate that selective state spaces maintain 94.2% of full attention perplexity while achieving "
        "a 4.8x increase in token generation throughput.",
        lead_bold="Summary:"
    )

    # Section 1: Introduction & Literature Review
    doc.add_heading("1. Introduction and Related Work", level=1)
    doc.add_paragraph(
        "Standard multi-head self-attention relies on explicit softmax dot products between queries and keys."
    )

    ref_vaswani = format_apa_reference(
        authors="Vaswani, A., Shazeer, N., Parmar, N., Uszkoreit, J., Jones, L., Gomez, A. N., Kaiser, Ł., & Polosukhin, I.",
        year="2017",
        title="Attention is all you need",
        source="Advances in Neural Information Processing Systems",
        volume_issue="30",
        pages="5998–6008",
        doi_or_url="https://doi.org/10.48550/arXiv.1706.03762"
    )
    doc.add_quotation_with_reference(
        quote_text="An attention function can be described as mapping a query and a set of key-value pairs to an output.",
        apa_reference=ref_vaswani,
        lead_in="As formulated in foundational sequence transduction research,"
    )

    ref_gu = format_apa_reference(
        authors="Gu, A., & Dao, T.",
        year="2023",
        title="Mamba: Linear-time sequence modeling with selective state spaces",
        source="arXiv preprint",
        doi_or_url="https://doi.org/10.48550/arXiv.2312.00752"
    )
    long_quote_mamba = (
        "Selective state space models overcome the computational bottleneck of standard attention by allowing "
        "the model to selectively propagate or reset state information based on the current token input. "
        "This design provides linear scaling with respect to context window length while preserving recurrent "
        "inference efficiency."
    )
    doc.add_quotation_with_reference(
        quote_text=long_quote_mamba,
        apa_reference=ref_gu,
        lead_in="Regarding hardware-aware state propagation in continuous-time models:"
    )

    # Section 2: Mathematical Formulations
    doc.add_heading("2. Mathematical Formulation", level=1)
    doc.add_paragraph(
        "Standard scaled dot-product attention computes representations via:"
    )

    doc.add_formula(
        formula_text="Attention(Q, K, V) = softmax((Q K^T) / √(d_k)) V",
        equation_number=1,
        description="Equation 1: Scaled dot-product attention matrix formulation."
    )

    doc.add_paragraph(
        "In contrast, continuous state space dynamics model hidden state transitions via:"
    )

    doc.add_formula(
        formula_text="h'(t) = A h(t) + B x(t),  y(t) = C h(t)",
        equation_number=2,
        description="Equation 2: Continuous linear time-invariant state-space system."
    )

    # Section 3: Empirical Benchmarks
    doc.add_heading("3. Empirical Evaluation", level=1)
    doc.add_paragraph(
        "Experiments were conducted across synthetic needle-in-a-haystack tasks and WikiText-103 benchmarks."
    )

    headers = ["Model Variant", "Active Params", "Context Window", "Perplexity", "Tokens / Sec"]
    rows = [
        ["Baseline Transformer-Tiny", "125M", "4,096", "18.42", "142"],
        ["FlashAttention-v2 (Sparse)", "125M", "8,192", "18.39", "285"],
        ["Linear Attention (ELU)",    "120M", "16,384", "20.15", "512"],
        ["Mamba SSM Selective",       "128M", "32,768", "18.84", "680"],
    ]
    doc.add_table(
        headers=headers,
        rows=rows,
        caption="Comparative Latency, Perplexity, and Context Window Capacity Benchmarks."
    )

    doc.add_figure(
        image_path_or_placeholder="Figure: Throughput vs Sequence Length Scaling Curve",
        caption="Token Generation Throughput vs. Context Length Across Tested Model Architectures.",
        width_inches=5.5
    )

    # Section 4: Conclusion
    doc.add_heading("4. Concluding Remarks", level=1)
    doc.add_paragraph(
        "Selective state space models offer compelling throughput advantages on memory-constrained hardware "
        "without sacrificing perplexity. Future work will investigate hybrid architectures combining SSM layers "
        "with sparse sliding-window attention."
    )

    doc.save(out_path)
    print(f"[OK] Academic paper sample saved at: {out_path}")
    print(f"     Author: {doc.author}")
    return out_path


if __name__ == "__main__":
    build_academic_paper()
