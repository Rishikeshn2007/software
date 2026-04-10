"""
Certificate of Appreciation Generator + Email Sender
------------------------------------------------------
Reads student data from an Excel (.xlsx) or CSV file, generates
a PDF Certificate of Appreciation for each student, then emails
it to them using Gmail SMTP (or any SMTP provider).

Required columns (case-insensitive): name, university, branch, email
"""

import os
import sys
import math
import smtplib
import getpass
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

#  CONFIG — edit these before running
OUTPUT_DIR   = "certificates"
ISSUER_NAME  = "Tech Fest Committee"
EVENT_NAME   = "National Tech Summit 2025"
ISSUE_DATE   = date.today().strftime("%B %d, %Y")

SMTP_HOST    = "smtp.gmail.com"   # change for Outlook / Yahoo etc.
SMTP_PORT    = 587                # TLS port


#  COLOURS
GOLD_DARK   = colors.HexColor("#B8860B")
GOLD_LIGHT  = colors.HexColor("#FFD700")
GOLD_PALE   = colors.HexColor("#FFF8DC")
NAVY        = colors.HexColor("#0A1628")
NAVY_LIGHT  = colors.HexColor("#1A2B4A")
CREAM       = colors.HexColor("#FFFDF5")


#  CERTIFICATE DRAWING

def draw_border(c, w, h):
    margin = 10 * mm
    c.setFillColor(NAVY)
    c.rect(0, 0, w, h, fill=1, stroke=0)

    c.setStrokeColor(GOLD_DARK)
    c.setLineWidth(4)
    c.rect(margin, margin, w - 2*margin, h - 2*margin, fill=0, stroke=1)

    inner = 14 * mm
    c.setStrokeColor(GOLD_LIGHT)
    c.setLineWidth(1.2)
    c.rect(inner, inner, w - 2*inner, h - 2*inner, fill=0, stroke=1)

    body_pad = 15 * mm
    c.setFillColor(CREAM)
    c.rect(body_pad, body_pad, w - 2*body_pad, h - 2*body_pad, fill=1, stroke=0)

    for i in range(8):
        alpha = 0.04 + i * 0.01
        c.setFillColor(colors.Color(0.72, 0.53, 0.04, alpha=alpha))
        stripe_h = 6 * mm
        c.rect(body_pad, body_pad + i * stripe_h * 0.5,
               w - 2*body_pad, stripe_h * 0.4, fill=1, stroke=0)

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
    for cx, cy in [
        (inner + 4*mm, inner + 4*mm),
        (w - inner - 4*mm, inner + 4*mm),
        (inner + 4*mm, h - inner - 4*mm),
        (w - inner - 4*mm, h - inner - 4*mm),
    ]:
        diamond(cx, cy, corner_sz)

    c.setStrokeColor(GOLD_DARK)
    c.setLineWidth(1)
    for dy in [h * 0.72, h * 0.295]:
        c.line(body_pad + 10*mm, dy, w - body_pad - 10*mm, dy)


def draw_seal(c, cx, cy, radius=22*mm):
    c.setFillColor(GOLD_DARK)
    c.circle(cx, cy, radius, fill=1, stroke=0)

    c.setFillColor(GOLD_PALE)
    c.circle(cx, cy, radius * 0.82, fill=1, stroke=0)

    c.setFillColor(GOLD_DARK)
    c.circle(cx, cy, radius * 0.38, fill=1, stroke=0)

    c.setStrokeColor(GOLD_DARK)
    c.setLineWidth(1.5)
    for i in range(8):
        angle = math.radians(i * 45)
        x1 = cx + math.cos(angle) * radius * 0.42
        y1 = cy + math.sin(angle) * radius * 0.42
        x2 = cx + math.cos(angle) * radius * 0.78
        y2 = cy + math.sin(angle) * radius * 0.78
        c.line(x1, y1, x2, y2)

    c.setFillColor(CREAM)
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(cx, cy - 5, "✓")


def generate_certificate(name: str, university: str, branch: str,
                          output_path: str):
    W, H = landscape(A4)
    c = canvas.Canvas(output_path, pagesize=(W, H))

    draw_border(c, W, H)

    c.setFillColor(NAVY_LIGHT)
    c.setFont("Helvetica", 9)
    c.drawCentredString(W/2, H * 0.915, EVENT_NAME.upper())

    c.setFillColor(NAVY)
    c.setFont("Times-Bold", 11)
    c.drawCentredString(W/2, H * 0.885, "P R E S E N T S")

    c.setFillColor(GOLD_DARK)
    c.setFont("Times-Bold", 38)
    c.drawCentredString(W/2, H * 0.83, "Certificate of Appreciation")

    c.setStrokeColor(GOLD_LIGHT)
    c.setLineWidth(2)
    c.line(W*0.25, H*0.815, W*0.75, H*0.815)

    c.setFillColor(NAVY)
    c.setFont("Times-Italic", 16)
    c.drawCentredString(W/2, H * 0.755, "This certificate is proudly presented to")

    c.setFillColor(NAVY_LIGHT)
    c.setFont("Times-Bold", 44)
    c.drawCentredString(W/2, H * 0.655, name)

    c.setStrokeColor(GOLD_DARK)
    c.setLineWidth(1.2)
    c.line(W*0.2, H*0.638, W*0.8, H*0.638)

    c.setFillColor(NAVY)
    c.setFont("Times-Italic", 14)
    c.drawCentredString(W/2, H * 0.59, f"{branch}  ·  {university}")

    c.setFont("Times-Roman", 12)
    c.drawCentredString(W/2, H * 0.535,
        "in recognition of outstanding dedication, exceptional performance,")
    c.drawCentredString(W/2, H * 0.505,
        f"and valuable contribution at  {EVENT_NAME}.")

    draw_seal(c, W * 0.5, H * 0.35)

    c.setFillColor(NAVY)
    c.setFont("Times-Bold", 12)
    c.drawCentredString(W * 0.22, H * 0.36, "Date of Issue")
    c.setFont("Times-Roman", 11)
    c.drawCentredString(W * 0.22, H * 0.335, ISSUE_DATE)
    c.setStrokeColor(GOLD_DARK)
    c.setLineWidth(0.8)
    c.line(W*0.1, H*0.325, W*0.34, H*0.325)

    c.setFont("Times-Bold", 12)
    c.drawCentredString(W * 0.78, H * 0.36, "Authorised By")
    c.setFont("Times-Roman", 11)
    c.drawCentredString(W * 0.78, H * 0.335, ISSUER_NAME)
    c.line(W*0.66, H*0.325, W*0.9, H*0.325)

    c.setFillColor(GOLD_DARK)
    c.setFont("Times-Italic", 9)
    c.drawCentredString(W/2, H * 0.235,
        "This certificate has been digitally generated and is valid without a physical signature.")

    c.save()


# ─────────────────────────────────────────────
#  DATA LOADING
# ─────────────────────────────────────────────

def load_data(filepath: str) -> pd.DataFrame:
    ext = os.path.splitext(filepath)[1].lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(filepath)
    elif ext == ".csv":
        df = pd.read_csv(filepath)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Use .xlsx or .csv")

    df.columns = [c.strip().lower() for c in df.columns]

    required = {"name", "university", "branch", "email"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"Missing columns: {missing}\n"
            f"Found columns: {list(df.columns)}\n"
            f"Required: name, university, branch, email"
        )
    return df


# ─────────────────────────────────────────────
#  EMAIL
# ─────────────────────────────────────────────

def build_email(sender: str, recipient: str, name: str,
                cert_path: str) -> MIMEMultipart:
    msg = MIMEMultipart()
    msg["From"]    = sender
    msg["To"]      = recipient
    msg["Subject"] = f"Your Certificate of Appreciation – {EVENT_NAME}"

    body = (
        f"Dear {name},\n\n"
        f"Congratulations!\n\n"
        f"We are delighted to present you with your Certificate of Appreciation\n"
        f"for your outstanding contribution and participation at {EVENT_NAME}.\n\n"
        f"Please find your certificate attached to this email.\n\n"
        f"We look forward to seeing you at future events!\n\n"
        f"Warm regards,\n"
        f"{ISSUER_NAME}\n"
    )
    msg.attach(MIMEText(body, "plain"))

    with open(cert_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        'attachment; filename="{}"'.format(os.path.basename(cert_path)),
    )
    msg.attach(part)
    return msg


def send_emails(df: pd.DataFrame, cert_map: dict,
                sender_email: str, sender_password: str):
    sep = "-" * 50
    print(f"\n{sep}")
    print(f"Connecting to {SMTP_HOST}:{SMTP_PORT} ...")

    try:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(sender_email, sender_password)
        print("Login successful.\n")
    except smtplib.SMTPAuthenticationError:
        print("Authentication failed.")
        print("Gmail: use an App Password, not your normal password.")
        print("Create one at https://myaccount.google.com/apppasswords")
        return
    except Exception as e:
        print(f"Could not connect: {e}")
        return

    sent, failed = 0, []

    for idx, row in df.iterrows():
        name  = str(row["name"]).strip()
        email = str(row["email"]).strip()

        if not email or email.lower() == "nan":
            print(f"  [Row {idx+2}] {name}  -> No email, skipped")
            continue

        cert_path = cert_map.get(name)
        if not cert_path or not os.path.exists(cert_path):
            print(f"  [Row {idx+2}] {name}  -> Certificate not found, skipped")
            continue

        try:
            msg = build_email(sender_email, email, name, cert_path)
            server.sendmail(sender_email, email, msg.as_string())
            print(f"  [Row {idx+2}] {name}  -> Sent to {email}")
            sent += 1
        except Exception as e:
            print(f"  [Row {idx+2}] {name}  -> FAILED ({e})")
            failed.append((name, email, str(e)))

    server.quit()
    print(f"\n{sep}")
    print(f"Emails sent: {sent} / {len(df)}")
    if failed:
        print(f"Failed ({len(failed)}):")
        for n, em, err in failed:
            print(f"  {n} <{em}>: {err}")


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────

def main():
    sep = "-" * 50

    # Resolve input file
    if len(sys.argv) >= 2:
        filepath = sys.argv[1]
    else:
        filepath = input("Enter path to your Excel/CSV file: ").strip().strip("'\"")

    if not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        sys.exit(1)

    # Load data
    print(f"Loading data from: {filepath}")
    try:
        df = load_data(filepath)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)

    print(f"Found {len(df)} records.")

    # Collect sender credentials
    print(f"\n{sep}")
    print("Email Sender Setup")
    print("Gmail users: use an App Password (myaccount.google.com/apppasswords)")
    sender_email    = input("Sender email address : ").strip()
    sender_password = getpass.getpass("App Password (hidden) : ")

    # Create output folder
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Generate certificates
    print(f"\n{sep}")
    print("Generating certificates ...\n")

    cert_map = {}
    gen_ok, gen_err = 0, []

    for idx, row in df.iterrows():
        name       = str(row["name"]).strip()
        university = str(row["university"]).strip()
        branch     = str(row["branch"]).strip()

        if not name or name.lower() == "nan":
            print(f"  Row {idx+2}: Skipping (empty name)")
            continue

        safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in name)
        filename  = f"Certificate_{safe_name}.pdf"
        out_path  = os.path.join(OUTPUT_DIR, filename)

        try:
            generate_certificate(name, university, branch, out_path)
            cert_map[name] = out_path
            print(f"  [{idx+2}] {name}  ->  {filename}")
            gen_ok += 1
        except Exception as e:
            print(f"  [{idx+2}] {name}  ->  ERROR: {e}")
            gen_err.append((name, str(e)))

    print(f"\nGenerated: {gen_ok}  |  Errors: {len(gen_err)}")

    # Send emails
    send_emails(df, cert_map, sender_email, sender_password)

    print("\nAll done!")


if __name__ == "__main__":
    main()