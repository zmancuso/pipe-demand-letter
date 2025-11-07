from flask import Flask, request, send_file, abort, jsonify
from io import BytesIO
from docx import Document
from docx.shared import Pt
from datetime import datetime
from flask_cors import CORS
import os

app = Flask(__name__)
CORS(app)

API_KEY = os.getenv('PIPE_DEMAND_API_KEY', 'YOUR_SECRET_KEY')

@app.get("/")
def index():
    return {"status": "ok", "message": "Use POST /demand-letter to generate a DOCX."}

@app.get("/healthz")
def health():
    return {"status": "ok"}

@app.post("/demand-letter")
def demand_letter():
    if request.headers.get("X-API-KEY") != API_KEY:
        return abort(401)

    data = request.get_json(force=True)
    name = data.get("business_name", "Business")
    address = data.get("business_address", "Unknown Address")
    contact = data.get("contact_name", "Client")
    today = data.get("today", datetime.utcnow().strftime("%b %d %Y"))
    effective_date = data.get("effective_date", "N/A")
    default_date = data.get("default_date", "N/A")
    last_payment = data.get("last_payment_date", "N/A")

    # Create DOCX
    doc = Docu
