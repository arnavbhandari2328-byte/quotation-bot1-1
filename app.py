# app.py
import os
import re
import gc
import ssl
import json
import time
import smtplib
import datetime
from email.message import EmailMessage
from flask import Flask, request, Response, jsonify
from docxtpl import DocxTemplate
import requests

# -------- Optional (but recommended) --------
# Gemini (Google Generative AI)
# pip install google-generativeai
import google.generativeai as genai


# =========================
# Environment & Constants
# =========================
TEMPLATE_FILE = "Template.docx"

# WhatsApp / Meta
META_ACCESS_TOKEN = os.environ.get("META_ACCESS_TOKEN") or os.environ.get("WHATSApp_TOKEN") or os.environ.get("WHATSAPP_TOKEN")
PHONE_NUMBER_ID   = os.environ.get("PHONE_NUMBER_ID")
META_VERIFY_TOKEN = os.environ.get("META_VERIFY_TOKEN") or os.environ.get("VERIFY_TOKEN")

# Gemini
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

# Zoho SMTP
ZOHO_EMAIL        = os.environ.get("ZOHO_EMAIL")                         # e.g. arnavbhandari2328@zohomail.in
ZOHO_APP_PASSWORD = os.environ.get("ZOHO_APP_PASSWORD")                  # App-specific password
SMTP_SERVER       = os.environ.get("SMTP_SERVER", "smtp.zoho.in")        # or smtp.zoho.com
SMTP_PORT         = int(os.environ.get("SMTP_PORT", "587"))              # we’ll try both 587 and 465 anyway
SENDER_NAME       = os.environ.get("ZOHO_SENDER_NAME", "NIVEE METAL PRODUCTS PVT LTD")

# Conservative worker settings for Render free tier
WEB_CONCURRENCY = int(os.environ.get("WEB_CONCURRENCY", "1"))

# Flask
app = Flask(__name__)


# =========================
# Init Gemini
# =========================
if GEMINI_API_KEY:
    try:
        genai.configure(api_key=GEMINI_API_KEY)
        GEMINI_MODEL = genai.GenerativeModel("gemini-pro")
        print("Gemini configured.")
    except Exception as e:
        GEMINI_MODEL = None
        print(f"!!! CRITICAL: Could not configure Gemini: {e}")
else:
    GEMINI_MODEL = None
    print("!!! WARNING: GEMINI_API_KEY not set. AI parsing will not work.")


# =========================
# Helpers
# =========================
def send_whatsapp_reply(to_phone_number: str, message_text: str):
    """Send a simple WhatsApp text message back to the user via Meta."""
    if not (META_ACCESS_TOKEN and PHONE_NUMBER_ID):
        print("!!! ERROR: Missing META_ACCESS_TOKEN or PHONE_NUMBER_ID. Cannot send WhatsApp reply.")
        return

    url = f"https://graph.facebook.com/v19.0/{PHONE_NUMBER_ID}/messages"
    headers = {
        "Authorization": f"Bearer {META_ACCESS_TOKEN}",
        "Content-Type": "application/json",
    }
    payload = {
        "messaging_product": "whatsapp",
        "to": to_phone_number,
        "type": "text",
        "text": {"body": message_text},
    }
    try:
        r = requests.post(url, headers=headers, json=payload, timeout=15)
        r.raise_for_status()
        print(f"WhatsApp reply sent to {to_phone_number}")
    except Exception as e:
        print(f"!!! ERROR sending WhatsApp reply: {e}")
        try:
            print(f"Response: {r.status_code} {r.text}")
        except Exception:
            pass


def _clean_json_text(text: str) -> str:
    """Remove code fences if model decides to return ```json ... ```."""
    if not text:
        return ""
    t = text.strip()
    t = t.replace("```json", "").replace("```", "").strip()
    return t


def parse_command_with_ai(user_text: str):
    """Ask Gemini to extract a structured quotation. Return dict or None."""
    print("Sending command to Google AI (Gemini) for parsing...")

    if not GEMINI_MODEL:
        print("!!! ERROR: Gemini model unavailable.")
        return None

    today_str = datetime.date.today().strftime("%B %d, %Y")

    prompt = f"""
You are a quotation data extractor. Read the user's WhatsApp message and return ONLY a compact JSON object.

Message:
{user_text}

Rules:
- Return JSON with ALL keys: "q_no","date","company_name","customer_name","product","quantity","rate","units","hsn","email".
- If a field is missing, use an empty string "" (do not omit fields).
- "date": If missing, use "{today_str}".
- "quantity": only the integer amount (no unit). If unclear, use "".
- "rate": only the numeric amount per unit (no currency). If unclear, use "".
- "units": typical values are "Pcs","Nos","Kgs". Default "Nos" if not specified.
- DO NOT include extra text, code fences, or comments.

Return strictly the JSON.
"""

    try:
        resp = GEMINI_MODEL.generate_content(prompt)
        ai_text = _clean_json_text(getattr(resp, "text", "") or "")
        print(f"AI response received: {ai_text}")
        data = json.loads(ai_text or "{}")

        # basic validation + normalization
        for k in ["q_no","date","company_name","customer_name","product","quantity","rate","units","hsn","email"]:
            data.setdefault(k, "")

        # Required fields to continue
        required = ["customer_name", "product", "quantity", "rate", "email"]
        for k in required:
            if not str(data.get(k, "")).strip():
                print(f"!!! ERROR: Missing field {k}")
                return None

        # Clean numbers
        try:
            qty_num = int(re.sub(r"[^\d]", "", str(data["quantity"])))
        except Exception:
            print("!!! ERROR: quantity not integer.")
            return None

        try:
            rate_num = float(re.sub(r"[^\d.]", "", str(data["rate"])))
        except Exception:
            print("!!! ERROR: rate not number.")
            return None

        total_num = qty_num * rate_num
        data["quantity"] = str(qty_num)
        data["rate"] = f"{rate_num:.2f}"
        data["rate_formatted"] = f"₹{rate_num:,.2f}"
        data["total"] = f"₹{total_num:,.2f}"
        if not data["date"]:
            data["date"] = today_str
        if not data["units"]:
            data["units"] = "Nos"

        print(f"Parsed context: {data}")
        return data

    except Exception as e:
        print(f"!!! ERROR during AI parse: {e}")
        return None


def create_quotation_from_template(context: dict) -> str | None:
    """Render Template.docx with context and save to /tmp. Return path or None."""
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(script_dir, TEMPLATE_FILE)
        doc = DocxTemplate(template_path)
    except Exception as e:
        print(f"!!! ERROR: Could not load template '{TEMPLATE_FILE}'. Error: {e}")
        return None

    try:
        doc.render(context)
        safe_name = "".join(c for c in context["customer_name"] if c.isalnum() or c in " _-").rstrip() or "Customer"
        filename = f"Quotation_{safe_name}_{datetime.date.today()}.docx"
        out_path = os.path.join("/tmp", filename)  # use ephemeral disk on Render
        doc.save(out_path)
        print(f"✅ DOCX created: '{out_path}'")
        return out_path
    except Exception as e:
        print(f"!!! ERROR rendering or saving the document: {e}")
        return None
    finally:
        gc.collect()


def send_email_with_attachment(recipient_email: str, subject: str, body: str, attachment_path: str) -> bool:
    """Send email via Zoho. Try STARTTLS:587, then SSL:465 fallback."""
    if not attachment_path:
        print("Cannot send email, no attachment was created.")
        return False

    if not (ZOHO_EMAIL and ZOHO_APP_PASSWORD):
        print("!!! ERROR: ZOHO_EMAIL or ZOHO_APP_PASSWORD missing.")
        return False

    # Build message
    msg = EmailMessage()
    display_from = f"{SENDER_NAME} <{ZOHO_EMAIL}>" if SENDER_NAME else ZOHO_EMAIL
    msg["From"] = display_from
    msg["To"] = recipient_email
    msg["Subject"] = subject
    msg.set_content(body)

    try:
        with open(attachment_path, "rb") as f:
            data = f.read()
        msg.add_attachment(
            data,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=os.path.basename(attachment_path),
        )
    except Exception as e:
        print("Attach error:", e)
        return False

    # 1) STARTTLS on 587
    try:
        host = SMTP_SERVER
        print(f"SMTP: trying STARTTLS {host}:587 …")
        with smtplib.SMTP(host, 587, timeout=20) as smtp:
            smtp.ehlo()
            smtp.starttls(context=ssl.create_default_context())
            smtp.ehlo()
            smtp.login(ZOHO_EMAIL, ZOHO_APP_PASSWORD)
            smtp.send_message(msg)
        print(f"Email sent to {recipient_email} via TLS:587")
        return True
    except Exception as e1:
        print("SMTP STARTTLS failed:", e1)

    # 2) SSL on 465 fallback
    try:
        print(f"SMTP: trying SSL {host}:465 …")
        with smtplib.SMTP_SSL(host, 465, timeout=20, context=ssl.create_default_context()) as smtp:
            smtp.login(ZOHO_EMAIL, ZOHO_APP_PASSWORD)
            smtp.send_message(msg)
        print(f"Email sent to {recipient_email} via SSL:465")
        return True
    except Exception as e2:
        print("SMTP SSL failed:", e2)
        return False
    finally:
        # Always try to cleanup the temp file
        try:
            os.remove(attachment_path)
            print(f"Cleaned up '{attachment_path}'")
        except Exception:
            pass


# =========================
# Flask Routes
# =========================
@app.route("/", methods=["GET"])
def health():
    return jsonify({
        "service": "quotation-bot",
        "status": "ok",
        "time": datetime.datetime.utcnow().isoformat() + "Z"
    })


@app.route("/webhook", methods=["GET", "POST"])
def webhook():
    # --- Verification (GET) ---
    if request.method == "GET":
        print("Webhook GET (verify) received")
        mode = request.args.get("hub.mode")
        token = request.args.get("hub.verify_token")
        challenge = request.args.get("hub.challenge")
        if mode == "subscribe" and token and challenge:
            if token == META_VERIFY_TOKEN:
                print("Verification successful.")
                return Response(challenge, status=200)
            print("Verification token mismatch.")
            return Response("Verification token mismatch", status=403)
        return Response("Bad verify request", status=400)

    # --- Messages (POST) ---
    print("Webhook POST received")
    data = request.json or {}

    # Navigate Meta's structure
    try:
        change = data["entry"][0]["changes"][0]
    except Exception:
        print("No entry/changes in payload.")
        return Response(status=200)

    # New incoming message
    if "messages" in change.get("value", {}) and change["value"]["messages"]:
        msg = change["value"]["messages"][0]
        if msg.get("type") != "text":
            print(f"Ignoring non-text message type: {msg.get('type')}")
            return Response(status=200)

        from_number = msg["from"]
        user_text = msg["text"]["body"]
        print(f"Incoming text from {from_number}: {user_text!r}")

        context = parse_command_with_ai(user_text)
        gc.collect()  # free memory after AI call

        if not context:
            send_whatsapp_reply(from_number,
                                "Sorry, I couldn't read all details. Please send like:\n"
                                "Name: Raju\nProduct: 5 inch SS 316L sheets\nQuantity: 5\nRate: 25000\nUnits: Pcs\nEmail: raju@example.com")
            return Response(status=200)

        # Create the DOCX
        print(f"Generating quote for {context['customer_name']}...")
        doc_file = create_quotation_from_template(context)
        if not doc_file:
            send_whatsapp_reply(from_number, "Sorry, I couldn't create the quotation DOCX.")
            return Response(status=200)

        # Build email
        subject = f"Quotation from {SENDER_NAME} (Ref: {context.get('q_no', 'N/A')})"
        body = (
            f"Dear {context['customer_name']},\n\n"
            f"Thank you for your enquiry.\n\n"
            f"Please find our official quotation attached for:\n"
            f"• Product: {context['product']}\n"
            f"• Quantity: {context['quantity']} {context['units']}\n"
            f"• Rate: {context['rate_formatted']} per {context['units']}\n"
            f"• Total: {context['total']}\n\n"
            f"Regards,\n{SENDER_NAME}\n"
        )

        ok = send_email_with_attachment(context["email"], subject, body, doc_file)

        if ok:
            send_whatsapp_reply(
                from_number,
                f"✅ Success! Your quotation for {context['product']} was created and emailed to {context['email']}."
            )
        else:
            send_whatsapp_reply(
                from_number,
                f"⚠️ Created the quotation but couldn't send email to {context['email']}. "
                f"Please check the email or try again."
            )

        return Response(status=200)

    # Status updates (sent/delivered/read)
    if "statuses" in change.get("value", {}):
        st = change["value"]["statuses"][0]
        print(f"Status update: {st.get('status')} for message {st.get('id')}.")
        return Response(status=200)

    print("Change without messages or statuses. Ignoring.")
    return Response(status=200)


# =========================
# Run (for local dev)
# =========================
if __name__ == "__main__":
    if not all([META_ACCESS_TOKEN, PHONE_NUMBER_ID, META_VERIFY_TOKEN]):
        print("!!! WARNING: Some WhatsApp env vars missing.")
    if not (ZOHO_EMAIL and ZOHO_APP_PASSWORD):
        print("!!! WARNING: Zoho SMTP env vars missing.")

    port = int(os.environ.get("PORT", 5000))
    print(f"Starting Flask on 0.0.0.0:{port}")
    app.run(host="0.0.0.0", port=port)
