from flask import Flask, request, send_file, abort, jsonify
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
from datetime import datetime
from flask_cors import CORS
import os
import re

app = Flask(__name__)
CORS(app)  # allow Apps Script origin
API_KEY = os.getenv("PIPE_DEMAND_API_KEY", "YOUR_SECRET_KEY")

# Optional: if you add a file named 'pipe_letterhead.png' to the repo root, we'll place it at the top.
LETTERHEAD_IMAGE = os.getenv("PIPE_LETTERHEAD_IMAGE", "pipe_letterhead.png")

# --- helpers ---
def norm_date(s, default=""):
    if not s:
        return default
    s = s.strip()
    # accept "Nov 06 2025" OR "Nov 06, 2025" OR "11 06 2025"
    for fmt in ("%b %d %Y", "%b %d, %Y", "%m %d %Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%b %d %Y")
        except Exception:
            pass
    return s  # leave as-is if user typed something custom

def money(s):
    if s in (None, ""):
        return ""
    try:
        v = float(re.sub(r"[^0-9.\-]", "", str(s)))
        return f"${v:,.2f}"
    except Exception:
        return str(s)

def require_api_key(req):
    if req.headers.get("X-API-KEY") != API_KEY:
        abort(401, description="Unauthorized: Invalid API key.")

@app.get("/")
def index():
    return {"status": "ok", "message": "Use POST /demand-letter to generate a DOCX."}, 200

@app.get("/healthz")
def healthz():
    return {"status": "ok"}, 200

@app.post("/demand-letter")
def demand_letter():
    require_api_key(request)

    try:
        data = request.get_json(force=True) or {}
    except Exception:
        return jsonify({"error": "Invalid JSON body"}), 400

    # --- pull fields (use your checklist names if present, fallback otherwise) ---
    business_name = data.get("business_name") or data.get("Business Name") or "BUSINESS NAME"
    business_address = data.get("business_address") or data.get("Business Address") or "Business address"
    today = norm_date(data.get("today") or data.get("Today") or datetime.utcnow().strftime("%b %d %Y"))
    effective_date = norm_date(data.get("effective_date") or data.get("Effective Date"))
    total_adv_plus_fee = money(data.get("total_advance_plus_fee") or data.get("Total Advance + Fee"))
    advance_amount = money(data.get("advance_amount") or data.get("Advance Amount"))
    fee = money(data.get("fee") or data.get("Fee"))
    default_date = norm_date(data.get("default_date") or data.get("Date of Default Event"))
    total_revenue = money(data.get("total_revenue") or data.get("Total Revenue Since Agreement to Today"))
    rr_percent = (data.get("rr_percent") or data.get("Revenue Share Percentage (RR%)") or "").strip()
    rr_amount = money(data.get("rr_amount") or data.get("Calculated % of Revenue Payable to Pipe ($)"))
    successful_payments = money(data.get("successful_payments") or data.get("Amount of Successful Payments"))
    last_payment_date = norm_date(data.get("last_payment_date") or data.get("Date of Last Payment"))
    percent_or_amount_due = money(data.get("percent_or_amount_due") or data.get("Payment Percentage or Amount Due ($% of Revenue Amount)"))
    shortfall = money(data.get("shortfall") or data.get("Shortfall"))

    # --- build DOCX ---
    doc = Document()

    # Try to add letterhead image if present
    try:
        if os.path.isfile(LETTERHEAD_IMAGE):
            p_img = doc.add_paragraph()
            run_img = p_img.add_run()
            run_img.add_picture(LETTERHEAD_IMAGE, width=Inches(1.6))
            p_img.alignment = WD_ALIGN_PARAGRAPH.LEFT
    except Exception:
        pass  # ignore logo failures

    # Title
    title = doc.add_paragraph("LETTER OF DEMAND")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].bold = True

    # Address block
    doc.add_paragraph(f"{business_name}\n{business_address}\nUnited States of America\n\n")

    # SENT VIA...
    sent_line = doc.add_paragraph(f"SENT VIA EMAIL ON {today}")
    sent_line.runs[0].bold = True
    doc.add_paragraph("\n")

    # Re: line
    re_line = doc.add_paragraph("Re: Demand for Payment - Pipe Merchant Cash Advance")
    re_line.runs[0].bold = True
    doc.add_paragraph("\n")

    # Dear Client
    dear = doc.add_paragraph("Dear Client,")
    dear.runs[0].bold = True
    doc.add_paragraph("\n")

    # BODY — exact approved copy with your fields filled in
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

    # return file
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

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
