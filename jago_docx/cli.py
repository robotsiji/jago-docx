"""
Command-Line Interface for Jago Docx.
"""

import sys
import os
import argparse
from typing import Optional

from .core import JagoDocument
from .metadata import get_default_username
from .footnotes import format_apa_reference


def generate_sample_document(output_path: str = "jago_docx_sample.docx", author: Optional[str] = None):
    """
    Generates a full showcase document demonstrating all 8 core capabilities.
    """
    doc = JagoDocument(
        title="Microservice Architecture & Distributed Systems Specification",
        author=author,
        subject="Technical Architecture Document"
    )

    doc.add_title(
        "Microservice Architecture & Distributed Systems Specification",
        subtitle="High-Performance Service Mesh, Concurrency Models, and Reliability Engineering"
    )

    # 1. Front Matter - Table of Contents, Figures, Tables
    doc.add_table_of_contents("Table of Contents")
    doc.add_list_of_figures("List of Figures")
    doc.add_list_of_tables("List of Tables")

    # 2. Section 1 - Introduction
    doc.add_heading("1. Executive Overview and Design Principles", level=1)
    doc.add_paragraph(
        "This specification establishes engineering patterns for distributed microservice clusters. "
        "Modern cloud environments require strict latency controls, partition tolerance, and clear "
        "concurrency guarantees under high transaction loads.",
        lead_bold="Operational Mandate:"
    )

    doc.add_callout(
        "All inter-service gRPC communication must maintain p99 latency below 15ms. "
        "Circuit breakers activate whenever error rates exceed 1.5% across a 30-second window.",
        title="CRITICAL SLA BENCHMARK"
    )

    # 3. Section 2 - Scholarly Citations & APA Footnotes (Requirement 4)
    doc.add_heading("2. Theoretical Foundations and Empirical References", level=1)
    doc.add_paragraph(
        "Distributed consensus mechanisms rely on established mathematical proofs regarding network partitions and clock drift."
    )

    # Short quote with APA reference
    ref_lamport = format_apa_reference(
        authors="Lamport, L.",
        year="1978",
        title="Time, clocks, and the ordering of events in a distributed system",
        source="Communications of the ACM",
        volume_issue="21(7)",
        pages="558–565",
        doi_or_url="10.1145/359545.359563"
    )
    doc.add_quotation_with_reference(
        quote_text="The concept of the time at which an event occurred is fundamentally relativistic.",
        apa_reference=ref_lamport,
        lead_in="As demonstrated in foundational consensus research,"
    )

    # Long blockquote (> 40 words) with APA reference
    ref_brewer = format_apa_reference(
        authors="Brewer, E. A.",
        year="2012",
        title="CAP twelve years later: How the 'rules' have changed",
        source="Computer",
        volume_issue="45(2)",
        pages="23–29",
        doi_or_url="10.1109/MC.2012.37"
    )
    long_quote_text = (
        "The modern CAP goal is to maximize combinations of consistency and availability that "
        "make sense for the specific application. Such an approach does not abandon consistency "
        "completely; rather, it uses fine-grained consistency guarantees and error handling to "
        "preserve usability during active network partitions while guaranteeing eventual state synchronization."
    )
    doc.add_quotation_with_reference(
        quote_text=long_quote_text,
        apa_reference=ref_brewer,
        lead_in="Regarding network partition resilience in modern data centers, Brewer formulated:"
    )

    # 4. Section 3 - Mathematical Formulas (Requirement 6)
    doc.add_heading("3. Latency and Queueing Models", level=1)
    doc.add_paragraph(
        "Throughput and queue depth follow Little's Law and exponential service time distributions. "
        "The equations below are native Office Math objects and remain directly editable in Word Equation Editor."
    )

    doc.add_heading("3.1 Service Queue Equilibrium", level=2)
    doc.add_formula(
        formula_text="L = λ · W",
        equation_number=1,
        description="Equation 1: Little's Law relating average queue size (L), arrival rate (λ), and wait time (W)."
    )

    doc.add_heading("3.2 Amdahl's Speedup Model", level=2)
    doc.add_formula(
        formula_text="S(p, s) = 1 / ((1 - p) + (p / s))",
        equation_number=2,
        description="Equation 2: Theoretical speedup as a function of parallel fraction (p) and processor count (s)."
    )

    # 5. Section 4 - Data Tables and Visuals with Captions Right Underneath (Requirement 7)
    doc.add_heading("4. Performance Benchmarks and Validation", level=1)
    doc.add_paragraph(
        "Benchmarking was conducted across three availability zones with isolated VPC peerings. "
        "Results are tabulated below."
    )

    # Table with caption STRICTLY underneath
    headers = ["Service Component", "Throughput (RPS)", "P50 Latency", "P99 Latency", "Error Rate"]
    rows = [
        ["Authentication Gateway", "45,200", "1.8 ms", "6.2 ms", "0.001%"],
        ["Order Orchestrator", "28,400", "3.4 ms", "11.8 ms", "0.004%"],
        ["Inventory Ledger", "19,800", "4.1 ms", "14.2 ms", "0.002%"],
        ["Telemetry Streamer", "82,000", "0.9 ms", "3.1 ms", "0.000%"],
    ]
    doc.add_table(
        headers=headers,
        rows=rows,
        caption="Core Microservice Cluster Latency and Throughput Metrics Under Simulated Peak Load."
    )

    doc.add_paragraph(
        "Physical traffic topology routes ingress packets through redundant edge proxies before reaching container pods."
    )

    # Figure with caption STRICTLY underneath
    doc.add_figure(
        image_path_or_placeholder="Diagram: Dual-Region Ingress and Service Mesh Routing Topology",
        caption="High-Availability Multi-Region Network Topology and Ingress Controller Layout.",
        width_inches=5.5
    )

    # 6. Section 5 - Conclusion
    doc.add_heading("5. Implementation Roadmap and Compliance", level=1)
    doc.add_paragraph(
        "Rollout follows a blue-green canary cycle. Automated rollbacks trigger if health probes fail "
        "three consecutive checks. Logging pipelines export OpenTelemetry traces to cold storage for compliance archiving."
    )

    # Save
    doc.save(output_path)
    print(f"[OK] Jago Docx sample successfully generated at: {output_path}")
    print(f"     Author: {doc.author}")
    print("     All 8 features applied: Antislop, Arial MS Docs, TOC, Username, APA Footnotes, Formulas, Captions Underneath, Editable Page Numbers.")
    return output_path


def main():
    parser = argparse.ArgumentParser(
        prog="jago-docx",
        description="Jago Docx: Anti-Slop, High-End Microsoft Word Document Generator"
    )
    subparsers = parser.add_subparsers(dest="command", help="Available subcommands")

    # Command: sample
    sample_parser = subparsers.add_parser("sample", help="Generate a comprehensive showcase document")
    sample_parser.add_argument("-o", "--output", default="jago_sample_report.docx", help="Output path for docx file")
    sample_parser.add_argument("-a", "--author", default=None, help="Author name (defaults to current OS user)")

    # Command: new
    new_parser = subparsers.add_parser("new", help="Create a clean structured document")
    new_parser.add_argument("title", help="Title of the new document")
    new_parser.add_argument("-o", "--output", required=True, help="Output path for docx file")
    new_parser.add_argument("-a", "--author", default=None, help="Author name (defaults to current OS user)")
    new_parser.add_argument("--subtitle", default=None, help="Optional subtitle")

    # Command: audit
    audit_parser = subparsers.add_parser("audit", help="Audit an existing document for antislop compliance")
    audit_parser.add_argument("file", help="Path to docx file to audit")

    args = parser.parse_args()

    if args.command == "sample":
        generate_sample_document(args.output, args.author)
    elif args.command == "new":
        doc = JagoDocument(title=args.title, author=args.author)
        doc.add_title(args.title, subtitle=args.subtitle)
        doc.add_table_of_contents()
        doc.add_heading("1. Introduction", level=1)
        doc.add_paragraph("Start drafting your content here.")
        doc.save(args.output)
        print(f"[OK] New Jago Docx created at: {args.output}")
        print(f"     Author: {doc.author}")
    elif args.command == "audit":
        from docx import Document
        from .antislop import audit_document_antislop
        if not os.path.exists(args.file):
            print(f"[ERROR] File not found: {args.file}")
            sys.exit(1)
        d = Document(args.file)
        violations = audit_document_antislop(d)
        if not violations:
            print("[OK] Document passed all Antislop checks with zero violations!")
        else:
            print(f"[!] Found {len(violations)} Antislop issues:")
            for v in violations:
                print(f"    - {v}")
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
