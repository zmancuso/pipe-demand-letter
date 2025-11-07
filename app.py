from flask import Flask, request, send_file, abort, jsonify
from io import BytesIO
from docx import Document
from docx.shared import Pt
from datetime import datetime
from flask_cors import CORS
import os

# -----------------------------
# Flask App Setup
# -----------------------------
app = Flask(__name__)

# Enable CORS for all origins (Google Apps Script hosted pages need this)
CORS(app)

# Load API key from environment or use fallback
API_KEY = os.getenv('PIPE_DEMAND_API_KEY', 'YOUR_SECRET_KEY')

# -----------------------------
# Routes
# -----------------------------

@app.get("/")
def index():
    """Simple index message so visiting root doesn't 404."""
    return {"status": "ok", "message": "Use POST /demand-letter to generate a DOCX."}, 200

@app.get("/healthz")
def health():
    """Health check endpoint for Render."""
    return {"status": "ok"}, 200

@app.post("/demand-letter")
def demand_letter():
    """Main endpoint that generates the Demand Letter DOCX."""
    # --- Security check ---
    if request.headers.get("X-API-KEY") != API_KEY:
        return abort(401, description="Unauthorized: Invalid API key.")

    # --- Parse incoming data ---
    try:
        data = request.get_json(force=True)
    except Exception:
        return jsonify({"error": "Invalid JSON body"}), 400

    # --- Extract fields with defaults ---
    name = data.get("business_name", "Business")
    address = data.get("business_address", "Unknown Address")
    contact = data.get("contact_name", "Client")
    today = data.get("today", datetime.utcnow().strftime("%b %d %Y"))
    effective_date = data.get("effective_date", "N/A")
    default_date = data.get("default_date", "N/A")
    last_payment = data.get("last_payment_date", "N/A")

    # --- Create the DOCX letter ---
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)

    doc.add_paragraph(f"{name}\n{address}\nUnited States of America\n\n")
    doc.add_paragraph(f"SENT VIA EMAIL ON {today}\n\n")
    doc.add_paragraph("Re: Demand for Payment - Pipe Merchant Cash Advance\n\n")
    doc.add_paragraph(f"Dear {contact},\n\n")
    doc.add_paragraph(
        f"This is our last attempt to seek payment for {name}'s merchant cash advance. "
        f"You entered into an agreement dated {effective_date}. "
        f"You defaulted on {default_date}, and your last payment was on {last_payment}. "
        "Please remit payment immediately.\n\n"
        "Servicing and Collections\n"
        "Pipe Advance LLC"
    )

    # --- Convert to downloadable file ---
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)

    filename = f"Demand_Letter_{name.replace(' ', '_')}.docx"

    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=filename
    )

# -----------------------------
# Run the app locally (for testing)
# -----------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
