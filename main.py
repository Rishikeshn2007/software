"""
Certificate of Appreciation Generator
--------------------------------------
Reads student data from an Excel (.xlsx) or CSV file and generates
a PDF Certificate of Appreciation for each student.

Expected columns (case-insensitive): name, university, branch
"""

import os
import sys
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import date
import math


# ─────────────────────────────────────────────
#  CONFIG — tweak these as you like
# ─────────────────────────────────────────────
OUTPUT_DIR   = "certificates"          # folder where PDFs are saved
ISSUER_NAME  = "Tech Fest Committee"   # who is issuing the certificate
EVENT_NAME   = "National Tech Summit 2025"  # event / reason for award
ISSUE_DATE   = date.today().strftime("%B %d, %Y")


# ─────────────────────────────────────────────
#  COLOURS
# ─────────────────────────────────────────────
GOLD_DARK   = colors.HexColor("#B8860B")
GOLD_LIGHT  = colors.HexColor("#FFD700")
GOLD_PALE   = colors.HexColor("#FFF8DC")
NAVY        = colors.HexColor("#0A1628")
NAVY_LIGHT  = colors.HexColor("#1A2B4A")
WHITE       = colors.white
CREAM       = colors.HexColor("#FFFDF5")
SILVER      = colors.HexColor("#C0C0C0")


def draw_border(c, w, h):
    """Draw the decorative multi-layer border."""
    # Outer filled rectangle (navy)
    c.setFillColor(NAVY)
    c.rect(0, 0, w, h, fill=1, stroke=0)

    # Gold outer border frame
    margin = 10 * mm
    c.setStrokeColor(GOLD_DARK)
    c.setLineWidth(4)
    c.rect(margin, margin, w - 2*margin, h - 2*margin, fill=0, stroke=1)

    # Thin gold inner frame
    inner = 14 * mm
    c.setStrokeColor(GOLD_LIGHT)
    c.setLineWidth(1.2)
    c.rect(inner, inner, w - 2*inner, h - 2*inner, fill=0, stroke=1)

    # Cream body fill inside inner frame
    c.setFillColor(CREAM)
    body_pad = 15 * mm
    c.rect(body_pad, body_pad, w - 2*body_pad, h - 2*body_pad, fill=1, stroke=0)

    # Subtle golden gradient stripes at top and bottom
    for i in range(8):
        alpha = 0.04 + i * 0.01
        c.setFillColor(colors.Color(0.72, 0.53, 0.04, alpha=alpha))
        stripe_h = 6 * mm
        c.rect(body_pad, body_pad + i*stripe_h*0.5,
               w - 2*body_pad, stripe_h * 0.4, fill=1, stroke=0)

    # Corner ornaments (simple diamond shapes)
    def diamond(cx, cy, size):
        path = c.beginPath()
        path.moveTo(cx, cy + size)
        path.lineTo(cx + size, cy)
        path.lineTo(cx, cy - size)
        path.lineTo(cx - size, cy)
        path.close()
        c.drawPath(path, fill=1, stroke=0)

    c.setFillColor(GOLD_DARK)
    corner_sz = 5 * mm
    corners = [
        (inner + 4*mm, inner + 4*mm),
        (w - inner - 4*mm, inner + 4*mm),
        (inner + 4*mm, h - inner - 4*mm),
        (w - inner - 4*mm, h - inner - 4*mm),
    ]
    for cx, cy in corners:
        diamond(cx, cy, corner_sz)

    # Horizontal divider lines
    divider_y1 = h * 0.72
    divider_y2 = h * 0.295
    for dy in [divider_y1, divider_y2]:
        c.setStrokeColor(GOLD_DARK)
        c.setLineWidth(1)
        c.line(body_pad + 10*mm, dy, w - body_pad - 10*mm, dy)


def draw_seal(c, cx, cy, radius=22*mm):
    """Draw a decorative circular seal."""
    # Outer ring
    c.setFillColor(GOLD_DARK)
    c.circle(cx, cy, radius, fill=1, stroke=0)

    # Inner ring
    c.setFillColor(GOLD_PALE)
    c.circle(cx, cy, radius * 0.82, fill=1, stroke=0)

    # Star in center
    c.setFillColor(GOLD_DARK)
    c.circle(cx, cy, radius * 0.38, fill=1, stroke=0)

    # Star rays (8 points)
    c.setStrokeColor(GOLD_DARK)
    c.setLineWidth(1.5)
    for i in range(8):
        angle = math.radians(i * 45)
        x1 = cx + math.cos(angle) * radius * 0.42
        y1 = cy + math.sin(angle) * radius * 0.42
        x2 = cx + math.cos(angle) * radius * 0.78
        y2 = cy + math.sin(angle) * radius * 0.78
        c.line(x1, y1, x2, y2)

    # "✓" checkmark in center
    c.setFillColor(CREAM)
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(cx, cy - 5, "✓")

    # Circular text around the seal
    # (simplified — just add "VERIFIED" as a small arc label)
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 6)
    label = "  CERTIFICATE OF APPRECIATION  "
    c.drawCentredString(cx, cy - radius - 5, label[:20])


def generate_certificate(name: str, university: str, branch: str,
                          output_path: str):
    """Generate one certificate PDF and save to output_path."""
    W, H = landscape(A4)   # ~841 × 595 pts
    c = canvas.Canvas(output_path, pagesize=(W, H))

    # ── Background & borders ────────────────
    draw_border(c, W, H)

    # ── Header: "CERTIFICATE OF APPRECIATION" ──
    c.setFillColor(NAVY)
    c.setFont("Times-Bold", 11)
    c.drawCentredString(W/2, H * 0.885, "P R E S E N T S")

    c.setFillColor(NAVY_LIGHT)
    c.setFont("Helvetica", 9)
    c.drawCentredString(W/2, H * 0.915,
                        EVENT_NAME.upper())

    c.setFillColor(GOLD_DARK)
    c.setFont("Times-Bold", 38)
    c.drawCentredString(W/2, H * 0.83, "Certificate of Appreciation")

    # Decorative gold underline
    ul_y = H * 0.815
    c.setStrokeColor(GOLD_LIGHT)
    c.setLineWidth(2)
    c.line(W*0.25, ul_y, W*0.75, ul_y)

    # ── "This is to certify that" ────────────
    c.setFillColor(NAVY)
    c.setFont("Times-Italic", 16)
    c.drawCentredString(W/2, H * 0.755, "This certificate is proudly presented to")

    # ── Recipient Name ───────────────────────
    c.setFillColor(NAVY_LIGHT)
    c.setFont("Times-Bold", 44)
    c.drawCentredString(W/2, H * 0.655, name)

    # Name underline
    c.setStrokeColor(GOLD_DARK)
    c.setLineWidth(1.2)
    c.line(W*0.2, H*0.638, W*0.8, H*0.638)

    # ── University & Branch ──────────────────
    c.setFillColor(NAVY)
    c.setFont("Times-Italic", 14)
    c.drawCentredString(W/2, H * 0.59,
        f"{branch}  ·  {university}")

    # ── Body text ───────────────────────────
    c.setFillColor(NAVY)
    c.setFont("Times-Roman", 12)
    line1 = ("in recognition of outstanding dedication, exceptional performance,")
    line2 = (f"and valuable contribution at  {EVENT_NAME}.")
    c.drawCentredString(W/2, H * 0.535, line1)
    c.drawCentredString(W/2, H * 0.505, line2)

    # ── Footer: Seal + Date + Signature ─────
    seal_x = W * 0.5
    seal_y = H * 0.35
    draw_seal(c, seal_x, seal_y)

    # Left: Date
    c.setFillColor(NAVY)
    c.setFont("Times-Bold", 12)
    c.drawCentredString(W * 0.22, H * 0.36, "Date of Issue")
    c.setFont("Times-Roman", 11)
    c.drawCentredString(W * 0.22, H * 0.335, ISSUE_DATE)
    # Date line
    c.setStrokeColor(GOLD_DARK)
    c.setLineWidth(0.8)
    c.line(W*0.1, H*0.325, W*0.34, H*0.325)

    # Right: Authorised Signatory
    c.setFont("Times-Bold", 12)
    c.drawCentredString(W * 0.78, H * 0.36, "Authorised By")
    c.setFont("Times-Roman", 11)
    c.drawCentredString(W * 0.78, H * 0.335, ISSUER_NAME)
    c.setStrokeColor(GOLD_DARK)
    c.setLineWidth(0.8)
    c.line(W*0.66, H*0.325, W*0.9, H*0.325)

    # Bottom note
    c.setFillColor(GOLD_DARK)
    c.setFont("Times-Italic", 9)
    c.drawCentredString(W/2, H * 0.235,
        "This certificate has been digitally generated and is valid without a physical signature.")

    c.save()


def load_data(filepath: str) -> pd.DataFrame:
    """Load Excel or CSV and normalise column names."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(filepath)
    elif ext == ".csv":
        df = pd.read_csv(filepath)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Use .xlsx or .csv")

    # Normalise column names
    df.columns = [c.strip().lower() for c in df.columns]

    required = {"name", "university", "branch"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"Missing columns: {missing}\n"
            f"Found columns: {list(df.columns)}\n"
            f"Please make sure your file has: name, university, branch"
        )
    return df


def main():
    # ── Resolve input file ───────────────────
    if len(sys.argv) >= 2:
        filepath = sys.argv[1]
    else:
        # Interactive fallback
        filepath = input("Enter path to your Excel/CSV file: ").strip().strip("'\"")

    if not os.path.exists(filepath):
        print(f"❌  File not found: {filepath}")
        sys.exit(1)

    # ── Load data ───────────────────────────
    print(f"📂  Loading data from: {filepath}")
    try:
        df = load_data(filepath)
    except ValueError as e:
        print(f"❌  {e}")
        sys.exit(1)

    print(f"✅  Found {len(df)} records.\n")

    # ── Create output folder ─────────────────
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # ── Generate certificates ────────────────
    success, errors = 0, []
    for idx, row in df.iterrows():
        name       = str(row["name"]).strip()
        university = str(row["university"]).strip()
        branch     = str(row["branch"]).strip()

        if not name or name.lower() == "nan":
            print(f"  ⚠️   Row {idx+2}: Skipping (empty name)")
            continue

        # Sanitise filename
        safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in name)
        filename  = f"Certificate_{safe_name}.pdf"
        out_path  = os.path.join(OUTPUT_DIR, filename)

        try:
            generate_certificate(name, university, branch, out_path)
            print(f"  ✅  [{idx+2}] {name}  →  {filename}")
            success += 1
        except Exception as e:
            print(f"  ❌  [{idx+2}] {name}  →  ERROR: {e}")
            errors.append((name, str(e)))

    # ── Summary ──────────────────────────────
    print(f"\n{'─'*50}")
    print(f"✨  Done!  {success} certificate(s) saved to ./{OUTPUT_DIR}/")
    if errors:
        print(f"⚠️   {len(errors)} error(s):")
        for n, err in errors:
            print(f"      {n}: {err}")


if __name__ == "__main__":
    main()