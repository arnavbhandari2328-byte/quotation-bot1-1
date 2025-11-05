# app.py
import os
import json
import smtplib
import datetime
import traceback
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders

from flask import Flask, request, Response, jsonify
import requests

# === Gemini (correct import) ===
import google.generativeai as genai

# === DOCX template ===
from docxtpl import DocxTemplate

# --------------------
# Environment
# --------------------
GOOGLE_API_KEY     = os.environ.get("GOOGLE_API_KEY", "")
WHATSAPP_TOKEN     = os.environ.get("WHATSAPP_TOKEN", "")
PHONE_NUMBER_ID    = os.environ.get("PHONE_NUMBER_ID", "")
VERIFY_TOKEN       = os.environ.get("VERIFY_TOKEN", "")
ZOHO_EMAIL         = os.environ.get("ZOHO_EMAIL", "")
ZOHO_APP_PASSWORD  = os.environ.get("ZOHO_APP_PASSWORD", "")
TEMPLATE_FILE      = os.environ.get("TEMPLATE_FILE", "Template.docx")

app = Flask(__name__)

# --------------------
# Configure Gemini
# --------------------
if GOOGLE_API_KEY:
    try:
        genai.configure(api_key=GOOGLE_API_KEY)
        GEM_MODEL = genai.GenerativeModel("gemini-1.5-flash")
    except Exception as e:
        print("Gemini init error:", e)
        GEM_MODEL = None
else:
    print("!!! GOOGLE_API_KEY is missing")
    GEM_MODEL = None


# --------------------
# Helpers
# --------------------
def send_whatsapp(to, text):
    """Send a WhatsApp text reply via Meta API."""
    if not (WHATSAPP_TOKEN and PHONE_NUMBER_ID):
        print("WhatsApp keys missing; cannot send reply.")
        return

    url = f"https://graph.facebook.com/v20.0/{PHONE_NUMBER_ID}/messages"
    headers = {
        "Authorization": f"Bearer {WHATSAPP_TOKEN}",
        "Content-Type": "application/json",
    }
    payload = {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "text",
        "text": {"body": text},
    }
    try:
        r = requests.post(url, headers=headers, json=payload, timeout=20)
        if r.status_code >= 300:
            print("WhatsApp send error:", r.status_code, r.text)
        else:
            print("WhatsApp reply sent to", to)
    except Exception as e:
        print("WhatsApp send exception:", e)


def parse_with_gemini(user_text):
    """Call Gemini to extract required fields as a minified JSON string."""
    if not GEM_MODEL:
        print("Gemini model not available.")
        return None

    today = datetime.date.today().strftime("%B %d, %Y")

    system_prompt = f"""
You are an assistant for a stainless steel trading company. Extract fields from the user's text.

Required keys:
- q_no (string, default "")
- date (string, default today's date: {today})
- company_name (string, default "")
- customer_name (string, REQUIRED)
- product (string, REQUIRED)
- quantity (integer as string, REQUIRED) -> extract only digits
- rate (float/integer as string, REQUIRED)
- units (string, default "Nos")
- hsn (string, default "")
- email (string, REQUIRED)

Return ONLY a single minified JSON object with those keys. No code fences, no extra text.
"""

    prompt = f"{system_prompt}\n\nUser: {user_text}"
    try:
        resp = GEM_MODEL.generate_content(prompt)
        text = (resp.text or "").strip()
        # Remove accidental code fences if any
        text = text.replace("```json", "").replace("```", "").strip()
        print("AI raw:", text)

        data = json.loads(text)

        # Basic validation
        required = ["customer_name", "product", "quantity", "rate", "email"]
        for k in required:
            if k not in data or not str(data[k]).strip():
                print(f"Missing field: {k}")
                return None

        # Normalize numbers
        try:
            qty = int("".join(ch for ch in str(data["quantity"]) if ch.isdigit()))
        except Exception:
            return None

        try:
            price = float(str(data["rate"]).replace(",", ""))
        except Exception:
            return None

        total = qty * price
        data["quantity"] = str(qty)
        data["rate"] = f"₹{price:,.2f}"
        data["total"] = f"₹{total:,.2f}"
        data["date"] = data.get("date") or today
        data["q_no"] = data.get("q_no", "")
        data["company_name"] = data.get("company_name", "")
        data["hsn"] = data.get("hsn", "")
        data["units"] = data.get("units") or "Nos"

        print("Parsed context:", data)
        return data
    except Exception as e:
        print("Gemini parse error:", e)
        traceback.print_exc()
        return None


def create_docx(context):
    """Fill the DOCX template with the context and save to /tmp."""
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(here, TEMPLATE_FILE)
        doc = DocxTemplate(template_path)
    except Exception as e:
        print("Template load error:", e)
        return None

    try:
        doc.render(context)
        safe_name = "".join(c for c in context["customer_name"] if c.isalnum() or c in " _-").strip() or "Customer"
        filename = f"Quotation_{safe_name}_{datetime.date.today()}.docx"
        out_path = os.path.join("/tmp", filename)
        doc.save(out_path)
        print("DOCX created:", out_path)
        return out_path
    except Exception as e:
        print("DOCX render/save error:", e)
        traceback.print_exc()
        return None


def send_email_zoho(to_email, subject, body, attachment_path):
    """Send email with DOCX attachment via Zoho SMTP (app password)."""
    if not (ZOHO_EMAIL and ZOHO_APP_PASSWORD):
        print("Zoho creds missing; cannot send email.")
        return False

    msg = MIMEMultipart()
    msg["From"] = ZOHO_EMAIL
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    if attachment_path:
        try:
            with open(attachment_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(attachment_path)}"')
            msg.attach(part)
        except Exception as e:
            print("Attach error:", e)
            return False

    try:
        server = smtplib.SMTP("smtp.zoho.in", 587, timeout=30)
        server.starttls()
        server.login(ZOHO_EMAIL, ZOHO_APP_PASSWORD)
        server.sendmail(ZOHO_EMAIL, [to_email], msg.as_string())
        server.quit()
        print("Email sent to", to_email)
        return True
    except Exception as e:
        print("SMTP error:", e)
        traceback.print_exc()
        return False


# --------------------
# Routes
# --------------------
@app.route("/", methods=["GET"])
def health():
    return jsonify({
        "service": "quotation-bot",
        "status": "ok",
        "time": datetime.datetime.utcnow().isoformat() + "Z"
    })


@app.route("/webhook", methods=["GET", "POST"])
def webhook():
    # --- Meta verification (GET)
    if request.method == "GET":
        mode = request.args.get("hub.mode")
        token = request.args.get("hub.verify_token")
        challenge = request.args.get("hub.challenge")
        if mode == "subscribe" and token == VERIFY_TOKEN:
            return Response(challenge, status=200)
        return Response("Verification failed", status=403)

    # --- Incoming messages (POST)
    try:
        data = request.get_json(force=True, silent=True) or {}
        print("Webhook POST received")
        # Meta structure
        change = (data.get("entry", [{}])[0]).get("changes", [{}])[0]
        value = change.get("value", {})

        # Message?
        if "messages" in value and value["messages"]:
            msg = value["messages"][0]
            if msg.get("type") == "text":
                from_number = msg.get("from")
                text = msg["text"]["body"].strip()
                print("Incoming text from", from_number, ":", text)

                # 1) Parse with Gemini
                context = parse_with_gemini(text)
                if not context:
                    send_whatsapp(
                        from_number,
                        "Sorry, I couldn't read all details. Please send like:\n\n"
                        "Name: Raju\nProduct: 5 inch SS 316L sheets\nQuantity: 5\nRate: 25000\nUnits: Pcs\nEmail: raju@example.com"
                    )
                    return Response(status=200)

                # 2) Create DOCX
                docx_path = create_docx(context)
                if not docx_path:
                    send_whatsapp(from_number, "Sorry, I created the quote but failed to build the document.")
                    return Response(status=200)

                # 3) Send email
                subject = f"Quotation from NIVEE METAL PRODUCTS PVT LTD (Ref: {context.get('q_no', 'N/A')})"
                body = f"""Dear {context['customer_name']},

Thank you for your enquiry.

Please find our official quotation attached.

Regards,
NIVEE METAL PRODUCTS PVT LTD
"""
                ok = send_email_zoho(context["email"], subject, body, docx_path)
                if ok:
                    send_whatsapp(
                        from_number,
                        f"Success! Your quotation for {context['product']} has been emailed to {context['email']}."
                    )
                else:
                    send_whatsapp(
                        from_number,
                        f"Sorry, I created the quote but couldn't send the email to {context['email']}."
                    )
                return Response(status=200)

            else:
                print("Non-text message; ignoring.")
                return Response(status=200)

        # Status updates, etc.
        if "statuses" in value:
            st = value["statuses"][0]
            print("Status update:", st.get("status"))
            return Response(status=200)

        print("Change without messages/status; ignoring.")
        return Response(status=200)

    except Exception as e:
        print("Webhook error:", e)
        traceback.print_exc()
        return Response(status=200)


# --------------------
# Gunicorn entry
# --------------------
if __name__ == "__main__":
    # Local run (Render uses gunicorn app:app)
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
