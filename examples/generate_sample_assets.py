"""
Generates clean, professional technical diagram and chart images for sample documents.
"""

import os
from PIL import Image, ImageDraw, ImageFont


def create_sample_assets(output_dir: str):
    os.makedirs(output_dir, exist_ok=True)

    # 1. Architecture Topology Diagram (800 x 420)
    fig1_path = os.path.join(output_dir, "architecture_topology.png")
    img1 = Image.new("RGB", (900, 460), color=(255, 255, 255))
    draw1 = ImageDraw.Draw(img1)

    # Background canvas border
    draw1.rectangle([10, 10, 890, 450], outline=(226, 232, 240), width=2, fill=(248, 250, 252))

    # Header Bar
    draw1.rectangle([10, 10, 890, 60], fill=(15, 23, 42))
    draw1.text((30, 24), "SYSTEM ARCHITECTURE: ZERO-TRUST INGRESS MESH TOPOLOGY", fill=(255, 255, 255))

    # Nodes
    # Ingress Box
    draw1.rectangle([60, 110, 260, 210], outline=(2, 132, 199), width=2, fill=(255, 255, 255))
    draw1.text((80, 135), "Edge Envoy Ingress", fill=(15, 23, 42))
    draw1.text((80, 160), "TLS Termination / WAF", fill=(100, 116, 139))
    draw1.text((80, 180), "Port: 443 -> Envoy Mesh", fill=(100, 116, 139))

    # Service Pod 1: Auth
    draw1.rectangle([360, 90, 560, 180], outline=(30, 41, 59), width=2, fill=(255, 255, 255))
    draw1.text((380, 115), "Auth Core Service", fill=(15, 23, 42))
    draw1.text((380, 140), "mTLS Token Verification", fill=(100, 116, 139))

    # Service Pod 2: Orders
    draw1.rectangle([360, 210, 560, 300], outline=(30, 41, 59), width=2, fill=(255, 255, 255))
    draw1.text((380, 235), "Order Orchestrator", fill=(15, 23, 42))
    draw1.text((380, 260), "gRPC / Event Broker", fill=(100, 116, 139))

    # Service Pod 3: Ledger DB
    draw1.rectangle([660, 150, 840, 250], outline=(16, 185, 129), width=2, fill=(255, 255, 255))
    draw1.text((680, 175), "Ledger Database", fill=(15, 23, 42))
    draw1.text((680, 200), "Raft Quorum (3x Nodes)", fill=(100, 116, 139))

    # Connecting Arrows (Lines)
    draw1.line([(260, 160), (360, 135)], fill=(2, 132, 199), width=3)
    draw1.line([(260, 160), (360, 255)], fill=(2, 132, 199), width=3)
    draw1.line([(560, 135), (660, 200)], fill=(30, 41, 59), width=3)
    draw1.line([(560, 255), (660, 200)], fill=(30, 41, 59), width=3)

    # Status Footer
    draw1.rectangle([60, 350, 840, 410], outline=(226, 232, 240), fill=(241, 245, 249))
    draw1.text((80, 370), "Telemetry: OpenTelemetry Collector | Mutual-TLS: SPIFFE/SPIRE | SLA: p99 < 15ms", fill=(71, 85, 105))

    img1.save(fig1_path, "PNG")
    print(f"Created Figure 1: {fig1_path}")

    # 2. Benchmark Scaling Chart (800 x 400)
    fig2_path = os.path.join(output_dir, "benchmark_chart.png")
    img2 = Image.new("RGB", (900, 420), color=(255, 255, 255))
    draw2 = ImageDraw.Draw(img2)

    draw2.rectangle([10, 10, 890, 410], outline=(226, 232, 240), width=2, fill=(248, 250, 252))
    draw2.rectangle([10, 10, 890, 60], fill=(15, 23, 42))
    draw2.text((30, 24), "BENCHMARK: P99 LATENCY VS THROUGHPUT SCALING", fill=(255, 255, 255))

    # Chart Axis
    draw2.line([(100, 340), (840, 340)], fill=(71, 85, 105), width=2)
    draw2.line([(100, 100), (100, 340)], fill=(71, 85, 105), width=2)

    # Bars
    bar_data = [
        ("Auth Gateway", 180, 140, "45k RPS / 6.2ms"),
        ("Order Engine", 360, 200, "28k RPS / 11.8ms"),
        ("Inventory Ledger", 540, 230, "20k RPS / 14.2ms"),
        ("Telemetry Stream", 720, 90, "82k RPS / 3.1ms")
    ]
    for label, x, height, stat in bar_data:
        draw2.rectangle([x - 50, 340 - height, x + 50, 340], fill=(2, 132, 199), outline=(14, 116, 144))
        draw2.text((x - 45, 350), label, fill=(15, 23, 42))
        draw2.text((x - 45, 340 - height - 25), stat, fill=(71, 85, 105))

    img2.save(fig2_path, "PNG")
    print(f"Created Figure 2: {fig2_path}")

    return fig1_path, fig2_path


if __name__ == "__main__":
    create_sample_assets("examples/assets")
