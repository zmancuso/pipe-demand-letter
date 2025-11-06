from flask import Flask, request, send_file, abort, jsonify
from io import BytesIO
from docx import Document
from docx.shared import Pt
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from flask_cors import CORS
import os, re

app = Flask(__name__)
CORS(app)

API_KEY = os.getenv('PIPE_DEMAND_API_KEY', 'YOUR_SECRET_KEY')

@app.get()
def index()
    return {status ok, message POST demand-letter to generate a DOCX}

@app.get(healthz)
def health()
    return {status ok}

@app.post(demand-letter)
def demand_letter()
    if request.headers.get(X-API-KEY) != API_KEY
        return abort(401)

    data = request.get_json(force=True)
    name = data.get(business_name, Business)
    address = data.get(business_address, Unknown Address)
    contact = data.get(contact_name, Client)
    today = data.get(today, datetime.utcnow().strftime(%b %d %Y))
    effective_date = data.get(effective_date, NA)
    default_date = data.get(default_date, NA)
    last_payment = data.get(last_payment_date, NA)

    doc = Document()
    doc.styles[Normal].font.name = Times New Roman
    doc.styles[Normal].font.size = Pt(12)

    doc.add_paragraph(f{name}n{address}nUnited States of Americann)
    doc.add_paragraph(fSENT VIA EMAIL ON {today}nn)
    doc.add_paragraph(Re Demand for Payment - Pipe Merchant Cash Advancenn)
    doc.add_paragraph(fDear {contact},nn)
    doc.add_paragraph(
        fThis is our last attempt to seek payment for {name}'s merchant cash advance. 
        fYou entered into an agreement dated {effective_date}. 
        fYou defaulted on {default_date}, and your last payment was on {last_payment}. 
        Please remit payment immediately.
    )
    doc.add_paragraph(nServicing and CollectionsnPipe Advance LLC)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        mimetype=applicationvnd.openxmlformats-officedocument.wordprocessingml.document,
        as_attachment=True,
        download_name=fDemand_Letter_{name.replace(' ', '_')}.docx,
    )

if __name__ == __main__
    app.run(host=0.0.0.0, port=8080)
