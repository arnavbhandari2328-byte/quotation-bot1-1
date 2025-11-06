import os
import json
import re
import base64
from datetime import datetime

import requests
from flask import Flask, request, jsonify

# ---------- Gmail (OAuth via Gmail API) ----------
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ---------- Flask ----------
app = Flask(__name__)

# =========================
#  Environment Variables
# =========================
# WhatsApp Cloud API
WA_VERIFY_TOKEN = os.getenv("WA_VERIFY_TOKEN", "verifyme")  # used by FB webhook verify
WA_ACCESS_TOKEN = os.getenv("WA_ACCESS_TOKEN")              # permanent token
WA_PHONE_NUMBER_ID = os.getenv("WA_PHONE_NUMBER_ID")        # your sender phone number id

# Gmail OAuth (no SMTP)
GMAIL_CLIENT_ID = os.getenv("GMAIL_CLIENT_ID")
GMAIL_CLIENT_SECRET = os.getenv("GMAIL_CLIENT_SECRET")
GMAIL_REFRESH_TOKEN = os.getenv("GMAIL_REFRESH_TOKEN")
GMAIL_SENDER = os.getenv("GMAIL_SENDER")  # the Gmail account you’ll send from

# =========================
#  Helpers
# =========================

def ok_health():
    return {
        "service": "quotation-bot",
        "status": "ok",
        "time": datetime.utcnow().isoformat() + "Z",
    }

def _gmail_creds():
    """Build Google OAuth Credentials using a refresh token."""
    missing = [k for k,v in {
        "GMAIL_CLIENT_ID": GMAIL_CLIENT_ID,
        "GMAIL_CLIENT_SECRET": GMAIL_CLIENT_SECRET,
        "GMAIL_REFRESH_TOKEN": GMAIL_REFRESH_TOKEN,
        "GMAIL_SENDER": GMAIL_SENDER,
    }.items() if not v]
    if missing:
        raise RuntimeError(f"Missing Gmail env vars: {', '.join(missing)}")

    return Credentials(
        token=None,
        refresh_token=GMAIL_REFRESH_TOKEN,
        client_id=GMAIL_CLIENT_ID,
        client_secret=GMAIL_CLIENT_SECRET,
        token_uri="https://oauth2.googleapis.com/token",
        scopes=["https://www.googleapis.com/auth/gmail.send"],
    )

def send_gmail(to_email: str, subject: str, html_body: str, text_body: str = None):
    """Send email via Gmail API."""
    creds = _gmail_creds()
    service = build("gmail", "v1", credentials=creds)

    msg = MIMEMultipart("alternative")
    msg["From"] = GMAIL_SENDER
    msg["To"] = to_email
    msg["Subject"] = subject

    if text_body:
        msg.attach(MIMEText(text_body, "plain"))
    msg.attach(MIMEText(html_body, "html"))

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    message = {"raw": raw}
    sent = service.users().messages().send(userId="me", body=message).execute()
    return sent.get("id")

def send_whatsapp_text(to_phone_e164: str, body: str):
    """Send a text message via WhatsApp Cloud API."""
    if not all([WA_ACCESS_TOKEN, WA_PHONE_NUMBER_ID]):
        print("WhatsApp keys missing; cannot send reply.")
        return None

    url = f"https://graph.facebook.com/v19.0/{WA_PHONE_NUMBER_ID}/messages"
    headers = {
        "Authorization": f"Bearer {WA_ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }
    payload = {
        "messaging_product": "whatsapp",
        "to": to_phone_e164,
        "type": "text",
        "text": {"body": body}
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=30)
    try:
        return resp.json()
    except Exception:
        return {"status_code": resp.status_code, "text": resp.text}

def parse_fields(msg: str):
    """
    Very simple local parser (no Gemini). Accepts free-form or key:value lines.
    Fields: customer_name, product, quantity, rate, units, email.
    """
    text = msg.strip()

    # Key-value style
    kv = {}
    for line in text.splitlines():
        if ":" in line:
            k, v = line.split(":", 1)
            kv[k.strip().lower()] = v.strip()

    # Try KV first
    name = kv.get("name") or kv.get("customer") or kv.get("customer name")
    product = kv.get("product")
    quantity = kv.get("quantity") or kv.get("qty")
    rate = kv.get("rate") or kv.get("price")
    units = kv.get("units") or kv.get("unit")
    email = kv.get("email")

    # Fallback regex for single line messages
    def r(pattern):
        m = re.search(pattern, text, re.I)
        return m.group(1).strip() if m else None

    if not name:
        name = r(r"name\s*[:\-]\s*([^\n,]+)")
    if not product:
        product = r(r"product\s*[:\-]\s*([^\n,]+)")
    if not quantity:
        quantity = r(r"quantity\s*[:\-]\s*([0-9]+)")
    if not rate:
        rate = r(r"rate\s*[:\-]\s*([0-9,]+)")
    if not units:
        units = r(r"units\s*[:\-]\s*([^\n,]+)")
    if not email:
        email = r(r"email\s*[:\-]\s*([\w\.\-\+]+@[\w\.\-]+\.[A-Za-z]{2,})")

    # Normalize
    if rate:
        rate = re.sub(r"[^\d]", "", rate) or rate
    if quantity:
        quantity = re.sub(r"[^\d]", "", quantity) or quantity

    # Build context
    ctx = {
        "customer_name": name,
        "product": product,
        "quantity": quantity,
        "rate": rate,
        "units": units,
        "email": email
    }
    return ctx

def missing_fields(ctx):
    req = ["customer_name", "product", "quantity", "rate", "units", "email"]
    return [k for k in req if not ctx.get(k)]

def html_quote(ctx):
    total = ""
    try:
        if ctx.get("quantity") and ctx.get("rate"):
            total_v = int(ctx["quantity"]) * int(ctx["rate"])
            total = f"{total_v:,}"
    except Exception:
        pass

    return f"""
    <div>
      <p>Dear {ctx['customer_name']},</p>
      <p>Please find your quotation below:</p>
      <table border="1" cellspacing="0" cellpadding="6">
        <tr><th align="left">Product</th><td>{ctx['product']}</td></tr>
        <tr><th align="left">Quantity</th><td>{ctx['quantity']} {ctx['units']}</td></tr>
        <tr><th align="left">Rate</th><td>{ctx['rate']} per {ctx['units']}</td></tr>
        <tr><th align="left">Total</th><td>{total if total else '-'}</td></tr>
      </table>
      <p>Regards,<br/>NIVEE METAL PRODUCTS PVT LTD</p>
    </div>
    """.strip()

PROMPT_EXAMPLE = (
    "Sorry, I couldn't read all details. Please send like:\n\n"
    "Name: Raju\n"
    "Product: 5 inch SS 316L sheets\n"
    "Quantity: 5\n"
    "Rate: 25000\n"
    "Units: Pcs\n"
    "Email: raju@example.com"
)

# =========================
#  Routes
# =========================

@app.route("/", methods=["GET"])
def root():
    return jsonify(ok_health())

@app.route("/webhook", methods=["GET"])
def webhook_verify():
    """
    Facebook / WhatsApp webhook verification:
    GET /webhook?hub.mode=subscribe&hub.verify_token=...&hub.challenge=...
    """
    mode = request.args.get("hub.mode")
    token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")

    if mode == "subscribe" and token == WA_VERIFY_TOKEN:
        return challenge, 200
    return "Forbidden", 403

@app.route("/webhook", methods=["POST"])
def webhook_receive():
    """
    Handle WhatsApp webhook messages.
    """
    data = request.get_json(silent=True) or {}
    # Basic extraction of the incoming message
    try:
        entry = data["entry"][0]
        changes = entry["changes"][0]
        value = changes["value"]
        messages = value.get("messages")
        if not messages:
            return "ok", 200
        msg = messages[0]
        from_phone = msg["from"]          # E.164
        txt = msg.get("text", {}).get("body", "").strip()
    except Exception:
        return "ok", 200

    print(f"Incoming text from {from_phone} : {txt}")

    # Parse locally
    ctx = parse_fields(txt)
    missing = missing_fields(ctx)

    if missing:
        # Ask the user to send in the example format
        send_whatsapp_text(from_phone, PROMPT_EXAMPLE)
        return "ok", 200

    # Build and send email
    try:
        subject = f"Quotation from NIVEE METAL PRODUCTS PVT LTD"
        html = html_quote(ctx)
        plain = (
            f"Dear {ctx['customer_name']},\n\n"
            f"Product: {ctx['product']}\n"
            f"Quantity: {ctx['quantity']} {ctx['units']}\n"
            f"Rate: {ctx['rate']} per {ctx['units']}\n\n"
            f"Regards,\nNIVEE METAL PRODUCTS PVT LTD\n"
        )
        send_gmail(ctx["email"], subject, html, plain)
        send_whatsapp_text(from_phone, f"Quotation sent to {ctx['email']} ✅")
    except Exception as e:
        print("Email send error:", e)
        send_whatsapp_text(from_phone, f"Sorry, I created the quote but couldn't send the email to {ctx['email']}.")

    return "ok", 200

# =========================
#  Run (for local dev)
# =========================
if __name__ == "__main__":
    port = int(os.getenv("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
