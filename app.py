import os
import re
import json
import base64
import datetime
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

from flask import Flask, request, Response, jsonify
import requests

# ---------- ENV ----------
GEMINI_API_KEY    = os.environ.get("GEMINI_API_KEY")

META_ACCESS_TOKEN = os.environ.get("META_ACCESS_TOKEN")
PHONE_NUMBER_ID   = os.environ.get("PHONE_NUMBER_ID")
META_VERIFY_TOKEN = os.environ.get("META_VERIFY_TOKEN")

# Zoho Mail SMTP
ZOHO_EMAIL        = os.environ.get("ZOHO_EMAIL")
ZOHO_APP_PASSWORD = os.environ.get("ZOHO_APP_PASSWORD")
SMTP_SERVER       = os.environ.get("SMTP_SERVER", "smtp.zoho.in")
SMTP_PORT         = int(os.environ.get("SMTP_PORT", "465"))

TEMPLATE_FILE = "Template.docx"

# ---------- APP ----------
app = Flask(__name__)

@app.get("/")
def root():
    return jsonify(service="quotation-bot", status="ok",
                   time=str(datetime.datetime.utcnow()) + "Z")

@app.get("/health")
def health():
    missing = [k for k, v in {
        "GEMINI_API_KEY": GEMINI_API_KEY,
        "META_ACCESS_TOKEN": META_ACCESS_TOKEN,
        "PHONE_NUMBER_ID": PHONE_NUMBER_ID,
        "META_VERIFY_TOKEN": META_VERIFY_TOKEN,
        "ZOHO_EMAIL": ZOHO_EMAIL,
        "ZOHO_APP_PASSWORD": ZOHO_APP_PASSWORD,
    }.items() if not v]
    return jsonify(ok=len(missing) == 0, missing=missing)


# ---------- WhatsApp reply ----------
def send_whatsapp_reply(to_phone_number: str, message_text: str) -> None:
    if not META_ACCESS_TOKEN or not PHONE_NUMBER_ID:
        print("!!! ERROR: Meta API keys missing; cannot send reply.")
        return
    url = f"https://graph.facebook.com/v19.0/{PHONE_NUMBER_ID}/messages"
    headers = {
        "Authorization": f"Bearer {META_ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }
    payload = {
        "messaging_product": "whatsapp",
        "to": to_phone_number,
        "type": "text",
        "text": {"body": message_text}
    }
    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=20)
        resp.raise_for_status()
        print(f"✅ WhatsApp reply sent to {to_phone_number}")
    except requests.exceptions.RequestException as e:
        print(f"!!! WhatsApp send error: {e} :: {getattr(e, 'response', None) and e.response.text}")


# ---------- AI parse (Gemini; lazy import) ----------
def parse_command_with_ai(command_text: str):
    print("Sending command to Google AI (Gemini) for parsing...")
    try:
        import google.generativeai as genai  # lazy import to reduce startup cost
        if not GEMINI_API_KEY:
            print("!!! ERROR: GEMINI_API_KEY not set.")
            return None
        genai.configure(api_key=GEMINI_API_KEY)

        model = genai.GenerativeModel('models/gemini-pro-latest')
        today = datetime.date.today().strftime('%B %d, %Y')

        system_prompt = f"""
        You are an assistant for a stainless steel trader. Extract a quotation.

        Date today: {today}

        Extract:
        - q_no
        - date (default: today's date)
        - company_name
        - customer_name
        - product
        - quantity (ONLY the number)
        - rate (number)
        - units (default "Nos")
        - hsn
        - email

        Return ONLY a single minified JSON string. No extra words or code fences.

        Example:
        User: "quote 101 for Raju at Raj pvt ltd, 500 pcs 3in pipe at 600, hsn 7304, email raju@gmail.com"
        AI: {{"q_no":"101","date":"{today}","company_name":"Raj pvt ltd","customer_name":"Raju","product":"3in pipe","quantity":"500","rate":"600","units":"Pcs","hsn":"7304","email":"raju@gmail.com"}}
        """
        response = model.generate_content(system_prompt + "\n\nUser: " + command_text)
        ai_text = (response.text or "").strip().replace("```json", "").replace("```", "").strip()
        print(f"AI response received: {ai_text}")

        context = json.loads(ai_text)

        # Required fields
        for f in ['product', 'customer_name', 'email', 'rate', 'quantity']:
            if not str(context.get(f, "")).strip():
                print(f"!!! ERROR: Missing field {f}")
                return None

        # Numbers & totals
        try:
            qty  = int(re.sub(r"[^\d]", "", str(context['quantity'])))
            rate = float(str(context['rate']).replace(",", "").strip())
            total = qty * rate
            context['quantity'] = str(qty)
            context['rate_formatted'] = f"₹{rate:,.2f}"
            context['total'] = f"₹{total:,.2f}"
            context['rate'] = context['rate_formatted']
        except ValueError:
            print("!!! ERROR: Invalid rate/quantity.")
            return None

        # Defaults
        context.setdefault('date', today)
        context.setdefault('company_name', "")
        context.setdefault('hsn', "")
        context.setdefault('q_no', "")
        context['units'] = context.get('units') or "Nos"

        print(f"Parsed context: {context}")
        return context

    except Exception as e:
        print(f"!!! AI error: {e}")
        return None


# ---------- DOCX generation (docxtpl) ----------
def create_quotation_from_template(context) -> str | None:
    try:
        from docxtpl import DocxTemplate
        script_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(script_dir, TEMPLATE_FILE)
        doc = DocxTemplate(template_path)
        doc.render(context)

        safe = "".join(c for c in context['customer_name'] if c.isalnum() or c in " _-").rstrip()
        filename = f"Quotation_{safe}_{datetime.date.today()}.docx"
        out_path = os.path.join("/tmp", filename)  # temp dir on Render
        doc.save(out_path)
        print(f"✅ DOCX created: '{out_path}'")
        return out_path
    except Exception as e:
        print(f"!!! DOCX error: {e}")
        return None


# ---------- Email via Zoho SMTP ----------
def send_email_with_attachment(recipient_email: str, subject: str, body_html: str, attachment_path: str) -> bool:
    if not all([ZOHO_EMAIL, ZOHO_APP_PASSWORD, attachment_path]):
        print("Email prerequisites missing.")
        return False
    try:
        msg = MIMEMultipart()
        msg["From"] = ZOHO_EMAIL
        msg["To"] = recipient_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body_html, "html"))

        with open(attachment_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
            msg.attach(part)

        ctx = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=ctx) as server:
            server.login(ZOHO_EMAIL, ZOHO_APP_PASSWORD)
            server.send_message(msg)

        print(f"✅ Email sent via Zoho Mail to {recipient_email}")
        # cleanup temp file
        try:
            os.remove(attachment_path)
        except Exception as e:
            print(f"Warn: could not delete temp file: {e}")
        return True
    except Exception as e:
        print(f"❌ Email send error: {e}")
        return False


# ---------- Webhook (Meta verification + events) ----------
@app.route("/webhook", methods=["GET", "POST"])
def webhook():
    if request.method == "GET":
        # Webhook verification
        mode = request.args.get('hub.mode')
        token = request.args.get('hub.verify_token')
        challenge = request.args.get('hub.challenge')
        if mode == 'subscribe' and token == META_VERIFY_TOKEN:
            print("✅ WEBHOOK VERIFIED SUCCESSFULLY!")
            return Response(challenge, status=200)
        print("❌ Webhook verification failed")
        return Response("Verification token mismatch", status=403)

    # POST — message or status update
    print("Webhook POST received")
    try:
        data = request.get_json(silent=True) or {}
        change = (data.get('entry', [{}])[0].get('changes') or [{}])[0]
        val = change.get('value', {})

        # Received user message
        if val.get('messages'):
            msg = val['messages'][0]
            if msg.get('type') != 'text':
                print(f"Ignoring non-text message: {msg.get('type')}")
                return Response(status=200)

            customer_phone = msg['from']
            user_text = msg['text']['body']

            # Parse
            context = parse_command_with_ai(user_text)
            if not context:
                send_whatsapp_reply(customer_phone, "Sorry, I couldn't understand your request. Please re-check and try again.")
                return Response(status=200)

            # Create DOCX
            doc_file = create_quotation_from_template(context)
            if not doc_file:
                send_whatsapp_reply(customer_phone, "Sorry, an internal error occurred while creating your document.")
                return Response(status=200)

            # Email
            subject = f"Quotation from NIVEE METAL PRODUCTS PVT LTD (Ref: {context.get('q_no','N/A')})"
            body_html = f"""
            <p>Dear {context['customer_name']},</p>
            <p>Thank you for your enquiry.</p>
            <p>Please find our official quotation attached.</p>
            <p>Regards,<br>Harsh Bhandari<br>Nivee Metal Products Pvt. Ltd.</p>
            """
            ok = send_email_with_attachment(context['email'], subject, body_html, doc_file)

            if ok:
                send_whatsapp_reply(customer_phone, f"Success! Your quotation for {context['product']} has been emailed to {context['email']}.")
            else:
                send_whatsapp_reply(customer_phone, f"Sorry, I created the quote but couldn't send the email to {context['email']}.")

            return Response(status=200)

        # Ignore delivery/read/status events
        if val.get('statuses'):
            st = val['statuses'][0]
            print(f"Status update: {st.get('status')} for message {st.get('id')}")
            return Response(status=200)

        print("No messages or statuses in webhook; ignoring.")
        return Response(status=200)

    except Exception as e:
        print(f"Webhook handling error: {e}")
        print(f"Raw body: {request.data}")
        return Response(status=200)


# ---------- Local run (Render uses Gunicorn) ----------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting Flask on 0.0.0.0:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)
