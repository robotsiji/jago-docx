"""
Example generator for Jago Docx.
Demonstrates all 8 core capabilities in a production-grade document.
"""

import os
import sys

# Add parent directory to path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from jago_docx import JagoDocument, format_apa_reference


def build_sample():
    out_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(out_dir, "sample_report.docx")

    doc = JagoDocument(
        title="High-Performance Microservice Mesh Architecture",
        subject="Technical Architecture Specification"
    )

    # Title & Subtitle
    doc.add_title(
        "High-Performance Microservice Mesh Architecture",
        subtitle="Zero-Trust Network Topology, Consensus Protocols, and Latency Optimization"
    )

    # Auto-Generated Front Matter Pages
    doc.add_table_of_contents("Table of Contents")
    doc.add_list_of_figures("List of Figures")
    doc.add_list_of_tables("List of Tables")

    # Section 1
    doc.add_heading("1. Executive Overview and Engineering Constraints", level=1)
    doc.add_paragraph(
        "Modern cloud microservices face rigorous throughput, partition tolerance, and tail-latency requirements. "
        "This specification establishes the hardware, network, and concurrency constraints governing production deployments.",
        lead_bold="Primary Objective:"
    )

    doc.add_callout(
        "All ingress-to-pod communications must complete within a strict 15ms p99 budget under 50,000 requests per second.",
        title="ARCHITECTURAL SLA DIRECTIVE"
    )

    # Section 2 - Scholarly References & APA Footnotes (Rule 4)
    doc.add_heading("2. Foundational Research & Theoretical Grounding", level=1)
    doc.add_paragraph(
        "Reliability in distributed state machines stems directly from classical research on logical clocks and consensus boundaries."
    )

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
        lead_in="In his seminal paper on logical clocks, Lamport demonstrated:"
    )

    ref_chandy = format_apa_reference(
        authors="Chandy, K. M., & Lamport, L.",
        year="1985",
        title="Distributed snapshots: Determining global states of distributed systems",
        source="ACM Transactions on Computer Systems",
        volume_issue="3(1)",
        pages="63–75",
        doi_or_url="10.1145/214451.214456"
    )
    long_quote = (
        "A distributed algorithm is designed to execute on a network of processors. "
        "Global state recording algorithms ensure that state consistency can be checkpointed "
        "without halting ongoing transaction processing, providing high availability during active state migration."
    )
    doc.add_quotation_with_reference(
        quote_text=long_quote,
        apa_reference=ref_chandy,
        lead_in="Regarding distributed snapshotting and asynchronous checkpointing:"
    )

    # Section 3 - Editable Formulas (Rule 6)
    doc.add_heading("3. Concurrency and Queueing Formulations", level=1)
    doc.add_paragraph(
        "Steady-state queue capacity and parallel speedup are formulated below. "
        "These equations use native Office Math (OMML) and are completely editable in Word."
    )

    doc.add_formula(
        formula_text="L = λ · W",
        equation_number=1,
        description="Little's Law relating queue length (L), request arrival rate (λ), and wait time (W)."
    )

    doc.add_formula(
        formula_text="S(p, s) = 1 / ((1 - p) + (p / s))",
        equation_number=2,
        description="Amdahl's Law calculating maximum speedup for parallel fraction (p) across (s) cores."
    )

    # Section 4 - Tables and Figures (Captions STRICTLY Underneath - Rule 7)
    doc.add_heading("4. Performance Benchmarks and Topology", level=1)
    doc.add_paragraph(
        "Verification benchmarks were executed across Kubernetes worker nodes in three isolated availability zones."
    )

    headers = ["Service Mesh Pod", "Base Memory", "Peak RPS", "p50 Latency", "p99 Latency"]
    rows = [
        ["edge-gateway-alpha", "4 GiB", "52,400", "1.4 ms", "5.8 ms"],
        ["auth-session-beta",  "8 GiB", "38,900", "2.1 ms", "7.4 ms"],
        ["order-router-gamma", "16 GiB", "61,200", "1.1 ms", "4.9 ms"],
        ["ledger-event-delta", "32 GiB", "24,500", "3.6 ms", "11.2 ms"],
    ]
    doc.add_table(
        headers=headers,
        rows=rows,
        caption="Simulated High-Concurrency Throughput and Response Latency across Primary Pod Clusters."
    )

    doc.add_paragraph(
        "The following diagram illustrates network traffic segregation through the Envoy ingress layer."
    )

    doc.add_figure(
        image_path_or_placeholder="Diagram: Envoy Service Mesh Mutual-TLS Zero-Trust Routing Topology",
        caption="High-Availability Mutual-TLS Zero-Trust Cluster Boundary Layout.",
        width_inches=5.5
    )

    # Section 5 - Conclusion
    doc.add_heading("5. Production Deployment Roadmap", level=1)
    doc.add_paragraph(
        "Canary phases will proceed in 5% increments over 48-hour burn-in periods. "
        "Automated telemetry triggers rollback if error thresholds exceed 0.05%."
    )

    doc.save(out_path)
    print(f"[OK] Generated showcase document at: {out_path}")
    print(f"     Author: {doc.author}")
    return out_path


if __name__ == "__main__":
    build_sample()
