# app.py
import os
import json
import re
import datetime
from flask import Flask, request, Response, jsonify
import requests
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
from docxtpl import DocxTemplate

# ---------- ENV ----------
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")  # Gemini
# WhatsApp / Meta
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN") or os.getenv("META_ACCESS_TOKEN")
PHONE_NUMBER_ID = os.getenv("BUSINESS_PHONE_NUMBER_ID") or os.getenv("PHONE_NUMBER_ID")
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN") or os.getenv("META_VERIFY_TOKEN")
# Zoho Mail
ZOHO_EMAIL = os.getenv("ZOHO_EMAIL")
ZOHO_APP_PASSWORD = os.getenv("ZOHO_APP_PASSWORD")
# Template
TEMPLATE_FILE = os.getenv("TEMPLATE_FILE", "Template.docx")

# ---------- GUARDRAILS ----------
missing = []
for k, v in {
    "GOOGLE_API_KEY": GOOGLE_API_KEY,
    "WHATSAPP_TOKEN": WHATSAPP_TOKEN,
    "PHONE_NUMBER_ID": PHONE_NUMBER_ID,
    "VERIFY_TOKEN": VERIFY_TOKEN,
    "ZOHO_EMAIL": ZOHO_EMAIL,
    "ZOHO_APP_PASSWORD": ZOHO_APP_PASSWORD,
}.items():
    if not v:
        missing.append(k)
if missing:
    print("!!! WARNING: Missing environment variables:", ", ".join(missing))

# ---------- Gemini (new SDK) ----------
from google import genai
from google.genai.types import GenerateContentConfig
gclient = genai.Client(api_key=GOOGLE_API_KEY) if GOOGLE_API_KEY else None

# ---------- Flask ----------
app = Flask(__name__)

@app.get("/")
def health():
    return jsonify({
        "service": "quotation-bot",
        "status": "ok",
        "time": datetime.datetime.utcnow().isoformat() + "Z"
    })

# ---------- WhatsApp helpers ----------
def send_whatsapp(to_phone, text):
    """
    Send a simple text reply to the user via Meta Graph API.
    """
    if not (WHATSAPP_TOKEN and PHONE_NUMBER_ID):
        print("!!! ERROR: Missing WHATSAPP_TOKEN or PHONE_NUMBER_ID")
        return

    url = f"https://graph.facebook.com/v20.0/{PHONE_NUMBER_ID}/messages"
    headers = {
        "Authorization": f"Bearer {WHATSAPP_TOKEN}",
        "Content-Type": "application/json"
    }
    payload = {
        "messaging_product": "whatsapp",
        "to": to_phone,
        "type": "text",
        "text": {"body": text}
    }
    try:
        r = requests.post(url, headers=headers, json=payload, timeout=20)
        r.raise_for_status()
        print(f"✅ WhatsApp reply sent to {to_phone}")
    except Exception as e:
        print("!!! ERROR sending WhatsApp:", e)
        if 'r' in locals():
            print("Status:", r.status_code, "Body:", r.text)

# ---------- Email (Zoho SMTP) ----------
def send_email_via_zoho(recipient, subject, body, attachment_path=None):
    if not all([ZOHO_EMAIL, ZOHO_APP_PASSWORD]):
        print("!!! ERROR: Missing ZOHO_EMAIL or ZOHO_APP_PASSWORD")
        return False

    msg = MIMEMultipart()
    msg["From"] = ZOHO_EMAIL
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    # optional attachment
    if attachment_path:
        from email.mime.base import MIMEBase
        from email import encoders
        with open(attachment_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{os.path.basename(attachment_path)}"'
        )
        msg.attach(part)

    try:
        # Zoho India: smtp.zoho.in; global: smtp.zoho.com
        with smtplib.SMTP_SSL("smtp.zoho.in", 465, timeout=30) as server:
            server.login(ZOHO_EMAIL, ZOHO_APP_PASSWORD)
            server.send_message(msg)
        print(f"✅ Email sent to {recipient}")
        return True
    except Exception as e:
        print("❌ Email send error:", e)
        return False

# ---------- DOCX ----------
def create_docx(context):
    """
    Renders the quotation using docxtpl.
    Writes to /tmp on Render. Returns filepath or None.
    """
    try:
        doc = DocxTemplate(TEMPLATE_FILE)
    except Exception as e:
        print(f"!!! ERROR opening template '{TEMPLATE_FILE}':", e)
        return None

    try:
        doc.render(context)
        safe_name = re.sub(r"[^A-Za-z0-9 _-]", "", context.get("customer_name", "Customer")).strip() or "Customer"
        fn = f"Quotation_{safe_name}_{datetime.date.today()}.docx"
        out_path = os.path.join("/tmp", fn)
        doc.save(out_path)
        print("✅ DOCX created:", out_path)
        return out_path
    except Exception as e:
        print("!!! ERROR rendering/saving DOCX:", e)
        return None

# ---------- Parsing (Gemini + fallback) ----------
SYSTEM_INSTRUCTIONS = """
You extract quotation details from a user's message.

Return a SINGLE-MINIFIED JSON object (no backticks, no prose) with keys:
q_no (string or empty if absent),
date (Month DD, YYYY; default today's date),
company_name,
customer_name,
product,
quantity (number only),
rate (number only),
units (default "Nos" if absent),
hsn (string or empty),
email.

Example output:
{"q_no":"101","date":"November 05, 2025","company_name":"Raj Pvt Ltd","customer_name":"Raju","product":"3 inch pipe","quantity":"500","rate":"600","units":"Pcs","hsn":"7304","email":"raju@gmail.com"}
"""

def parse_with_rules(txt: str):
    """
    Very small regex fallback if AI fails.
    Looks for common fields in "Name: ..., Product: ..., Quantity: ..., Rate: ..., Units: ..., Email: ..."
    """
    def pick(pat):
        m = re.search(pat, txt, re.IGNORECASE)
        return (m.group(1) or "").strip() if m else ""

    return {
        "q_no": "",
        "date": datetime.date.today().strftime("%B %d, %Y"),
        "company_name": "",
        "customer_name": pick(r"(?:name|customer|person)\s*:\s*(.+)"),
        "product": pick(r"(?:product|item)\s*:\s*(.+)"),
        "quantity": re.sub(r"\D", "", pick(r"(?:quantity|qty)\s*:\s*([0-9,\.]+)")),
        "rate": re.sub(r"[^\d.]", "", pick(r"(?:rate|price)\s*:\s*([0-9,\.]+)")),
        "units": pick(r"(?:units?)\s*:\s*([A-Za-z]+)") or "Nos",
        "hsn": pick(r"(?:hsn)\s*:\s*([0-9A-Za-z]+)"),
        "email": pick(r"(?:email|mail)\s*:\s*([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})")
    }

def parse_command(text: str):
    """
    Use Gemini (new client) to extract fields. Fall back to simple regex rules.
    """
    # 1) Gemini
    if gclient:
        try:
            prompt = f"{SYSTEM_INSTRUCTIONS}\n\nUser message:\n{text}"
            resp = gclient.models.generate_content(
                model="gemini-1.5-flash",
                contents=prompt,
                config=GenerateContentConfig(max_output_tokens=300)
            )
            raw = (resp.text or "").strip()
            raw = raw.replace("```json", "").replace("```", "").strip()
            print("AI raw:", raw)

            data = json.loads(raw)

            # Validate requireds
            need = ["product", "customer_name", "email", "rate", "quantity"]
            if not all(data.get(k) for k in need):
                raise ValueError("Missing required field(s)")

            # Normalize numeric formatting
            try:
                rate_num = float(str(data["rate"]).replace(",", ""))
                qty_num = int(float(str(data["quantity"]).replace(",", "")))
            except Exception:
                raise ValueError("Invalid numbers in rate/quantity")

            data["rate_formatted"] = f"₹{rate_num:,.2f}"
            data["total"] = f"₹{rate_num * qty_num:,.2f}"
            data["rate"] = data["rate_formatted"]
            data["quantity"] = str(qty_num)

            # Defaults
            data.setdefault("date", datetime.date.today().strftime("%B %d, %Y"))
            data.setdefault("units", "Nos")
            data.setdefault("company_name", "")
            data.setdefault("hsn", "")
            data.setdefault("q_no", "")

            return data
        except Exception as e:
            print("!!! ERROR during AI parse:", e)

    # 2) Fallback
    data = parse_with_rules(text)
    need = ["product", "customer_name", "email", "rate", "quantity"]
    if not all(data.get(k) for k in need):
        return None

    # format amounts
    try:
        rate_num = float(str(data["rate"]).replace(",", ""))
        qty_num = int(float(str(data["quantity"]).replace(",", "")))
        data["rate_formatted"] = f"₹{rate_num:,.2f}"
        data["total"] = f"₹{rate_num * qty_num:,.2f}"
        data["rate"] = data["rate_formatted"]
        data["quantity"] = str(qty_num)
    except Exception:
        return None
    return data

# ---------- Webhook ----------
@app.route("/webhook", methods=["GET", "POST"])
def webhook():
    # GET verify
    if request.method == "GET":
        mode = request.args.get("hub.mode")
        token = request.args.get("hub.verify_token")
        challenge = request.args.get("hub.challenge")
        if mode == "subscribe" and token and token == VERIFY_TOKEN:
            print("✅ Webhook verified.")
            return Response(challenge, status=200)
        print("❌ Webhook verification failed.")
        return Response("Verification token mismatch", status=403)

    # POST (message/status)
    try:
        data = request.get_json() or {}
    except Exception:
        return Response(status=200)

    # WhatsApp inbound parsing (Meta format)
    try:
        change = data.get("entry", [{}])[0].get("changes", [{}])[0]
        value = change.get("value", {})
        if "messages" in value and value["messages"]:
            msg = value["messages"][0]
            msg_type = msg.get("type")
            from_phone = msg.get("from")
            if msg_type == "text":
                user_text = msg["text"]["body"]
            else:
                print("Non-text message received; ignoring.")
                return Response(status=200)

            # Try AI parse
            context = parse_command(user_text)
            if not context:
                # Ask user to use the helper format
                helper = (
                    "Sorry, I couldn't read all details. Please send like:\n\n"
                    "Name: Raju\n"
                    "Product: 5 inch SS 316L sheets\n"
                    "Quantity: 5\n"
                    "Rate: 25000\n"
                    "Units: Pcs\n"
                    "Email: raju@example.com"
                )
                send_whatsapp(from_phone, helper)
                return Response(status=200)

            # Create docx
            doc_path = create_docx(context)

            # Email it
            subject = f"Quotation from NIVEE METAL PRODUCTS PVT LTD (Ref: {context.get('q_no','N/A')})"
            body = (
                f"Dear {context['customer_name']},\n\n"
                "Thank you for your enquiry.\n\n"
                f"Product: {context['product']}\n"
                f"Qty: {context['quantity']} {context.get('units','')}\n"
                f"Rate: {context['rate']}\n"
                f"Total: {context['total']}\n\n"
                "Please find our official quotation attached.\n\n"
                "Regards,\n"
                "Harsh Bhandari\n"
                "Nivee Metal Products Pvt. Ltd."
            )

            ok = send_email_via_zoho(
                recipient=context["email"],
                subject=subject,
                body=body,
                attachment_path=doc_path
            )

            if ok:
                send_whatsapp(
                    from_phone,
                    f"Success! Your quotation for {context['product']} was emailed to {context['email']}."
                )
            else:
                send_whatsapp(
                    from_phone,
                    f"Sorry, I created the quote but couldn't send the email to {context['email']}."
                )

        elif "statuses" in value:
            # Delivery/read receipts etc. — ignore
            pass

    except Exception as e:
        print("!!! ERROR handling webhook POST:", e)

    return Response(status=200)

# ---------- Run ----------
if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    print(f"Starting on 0.0.0.0:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)
