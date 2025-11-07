# app.py — Pipe Demand Letter Service (Render-ready)

from flask import Flask, request, send_file, abort, jsonify
from flask_cors import CORS
from io import BytesIO
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pypdf import PdfReader
import pdfplumber
import os
import re

# -----------------------------
# App & Config
# -----------------------------
app = Flask(__name__)
# Allow requests from Google Apps Script and anywhere else (tighten if desired)
CORS(app)

API_KEY = os.getenv("PIPE_DEMAND_API_KEY", "YOUR_SECRET_KEY")
LETTERHEAD_IMAGE = os.getenv("PIPE_LETTERHEAD_IMAGE", "pipe_letterhead.png")  # optional logo file in repo root

# -----------------------------
# Helpers
# -----------------------------

def require_api_key(req) -> None:
    """Abort 401 if header key doesn't match env var."""
    if req.headers.get("X-API-KEY") != API_KEY:
        abort(401, description="Unauthorized: Invalid API key.")

def norm_date(s: str, default: str = "") -> str:
    """Normalize many common date inputs to 'MMM DD YYYY'."""
    if not s:
        return default
    s = str(s).strip()
    fmts = ("%b %d %Y", "%b %d, %Y", "%m %d %Y", "%m/%d/%Y", "%Y-%m-%d")
    for fmt in fmts:
        try:
            return datetime.strptime(s, fmt).strftime("%b %d %Y")
        except Exception:
            pass
    return s  # leave as-is if unknown

def money(val) -> str:
    """Format as $X,XXX.XX (accepts strings with $/commas)."""
    if val in (None, ""):
        return ""
    try:
        v = float(re.sub(r"[^0-9.\-]", "", str(val)))
        return f"${v:,.2f}"
    except Exception:
        return str(val)

def safe_str(v, fallback="") -> str:
    return str(v) if v is not None else fallback

def add_logo_if_present(doc: Document) -> None:
    """Insert letterhead logo if file exists (non-fatal if missing)."""
    try:
        if os.path.isfile(LETTERHEAD_IMAGE):
            p = doc.add_paragraph()
            run = p.add_run()
            run.add_picture(LETTERHEAD_IMAGE, width=Inches(1.6))
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    except Exception:
        # ignore image issues silently
        pass

# -----------------------------
# Routes: health & index
# -----------------------------
@app.get("/")
def index():
    return {"status": "ok", "message": "Use POST /demand-letter (JSON) or POST /agreement-extract (file)."}, 200

@app.get("/healthz")
def healthz():
    return {"status": "ok"}, 200

# -----------------------------
# Route: Generate Demand Letter
# -----------------------------
@app.post("/demand-letter")
def demand_letter():
    require_api_key(request)

    try:
        data = request.get_json(force=True) or {}
    except Exception:
        return jsonify({"error": "Invalid JSON body"}), 400

    # Extract inputs (accept both short and long names)
    business_name = data.get("business_name") or data.get("Business Name") or "BUSINESS NAME"
    business_address = data.get("business_address") or data.get("Business Address") or "Business address"
    contact_name = data.get("contact_name") or data.get("Contact Name") or "Client"
    today = norm_date(data.get("today") or data.get("Today") or datetime.utcnow().strftime("%b %d %Y"))
    effective_date = norm_date(data.get("effective_date") or data.get("Effective Date"))
    default_date = norm_date(data.get("default_date") or data.get("Date of Default Event"))
    last_payment_date = norm_date(data.get("last_payment_date") or data.get("Date of Last Payment"))

    total_adv_plus_fee = money(data.get("total_advance_plus_fee") or data.get("Total Advance + Fee"))
    advance_amount = money(data.get("advance_amount") or data.get("Advance Amount"))
    fee = money(data.get("fee") or data.get("Fee"))
    total_revenue = money(data.get("total_revenue") or data.get("Total Revenue Since Agreement to Today"))
    rr_percent = safe_str(data.get("rr_percent") or data.get("Revenue Share Percentage (RR%)")).strip()
    rr_amount = money(data.get("rr_amount") or data.get("Calculated % of Revenue Payable to Pipe ($)"))
    successful_payments = money(data.get("successful_payments") or data.get("Amount of Successful Payments"))
    percent_or_amount_due = money(
        data.get("percent_or_amount_due")
        or data.get("Payment Percentage or Amount Due ($% of Revenue Amount)")
    )
    shortfall = money(data.get("shortfall") or data.get("Shortfall"))

    # Build DOCX
    doc = Document()
    # Base font
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)

    # Optional logo
    add_logo_if_present(doc)

    # Title
    title = doc.add_paragraph("LETTER OF DEMAND")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].bold = True

    # Address block
    doc.add_paragraph(f"{business_name}\n{business_address}\nUnited States of America\n")

    # SENT VIA EMAIL ON ...
    sent_p = doc.add_paragraph(f"SENT VIA EMAIL ON {today}")
    sent_p.runs[0].bold = True
    doc.add_paragraph("")  # spacer

    # Re: line
    re_p = doc.add_paragraph("Re: Demand for Payment - Pipe Merchant Cash Advance")
    re_p.runs[0].bold = True
    doc.add_paragraph("")  # spacer

    # Dear
    dear = doc.add_paragraph(f"Dear {contact_name},")
    dear.runs[0].bold = True
    doc.add_paragraph("")  # spacer

    # Body — approved copy with fields
    body = (
        f"This is our last attempt and FINAL WARNING to seek payment for {business_name}’s merchant cash advance (“MCA”) "
        f"before we seek all legal remedies available to us. {business_name}(“you”) entered into an MCA Agreement "
        f"(“Agreement”) with Pipe Advancel LLC (the “Company”) dated {effective_date} (the “Effective Date”) for an MCA in "
        f"the total amount of {total_adv_plus_fee} (consisting of a MCA advance of {advance_amount} and a fee of {fee}).\n\n"

        f"Since {default_date}, {business_name} has failed to comply with its terms, by generating revenue and failing to "
        f"deliver and/or preventing Pipe from receiving its share of revenue payments. As of {today}, {business_name} has had "
        f"{total_revenue} in revenue payments of which {rr_percent} ({rr_amount}) are payable to Pipe under the terms of the Agreement. "
        f"We have only received {successful_payments} towards your Total Advance Amount. The last payment to Pipe was on {last_payment_date}.\n\n"

        f"Your failure to pay Pipe the agreed upon percentage of revenue {percent_or_amount_due}, is a breach of the Agreement. "
        f"We have attempted to contact you and resolve this issue informally multiple times. Despite Pipe’s continuous efforts to "
        f"resolve this issue, we have not received a payment.\n\n"

        f"If a payment of {shortfall} is not received within 3 business days of receipt of this letter, we will seek all remedies "
        f"available to us under the Agreement, including referring this matter to a third-party collections firm or seeking appropriate "
        f"legal action. You may also be held liable and subject to additional fees incurred by Pipe in an attempt to pursue these payments.\n\n"

        f"We urge you to treat this matter with the utmost urgency and to cooperate fully in resolving this breach amicably.\n\n"
    )
    doc.add_paragraph(body)

    # Footer/contact
    doc.add_paragraph(
        "Please contact our Servicing and Collections Manager, William, at william@pipe.com immediately within the next 3 business days.\n\n"
        "Thank you for your immediate attention to this critical issue.\n\n"
        "Servicing and Collections\n"
        "Pipe Advance  LLC"
    )

    # Return DOCX
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    fname = f"Demand_Letter_{re.sub(r'\\s+', '_', business_name)}.docx"

    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=fname
    )

# -----------------------------
# Route: Agreement Extraction
# -----------------------------
FIELD_PATTERNS = {
    # Tweak these regexes to match your agreement language precisely
    "business_name": r"(?:Business\s*Name|Merchant)\s*[:\-]\s*(.+)",
    "effective_date": r"(?:Effective\s*Date)\s*[:\-]\s*([A-Za-z]{3}\s+\d{1,2}\s+\d{4}|\d{1,2}[/-]\d{1,2}[/-]\d{4}|[0-9]{4}-[0-9]{2}-[0-9]{2})",
    "advance_amount": r"(?:Advance\s*Amount|Purchase\s*Price
