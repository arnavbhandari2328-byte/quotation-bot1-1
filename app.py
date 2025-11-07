import os
import re
import json
import datetime
import logging
from flask import Flask, request, jsonify, Response
from dotenv import load_dotenv
import google.generativeai as genai
from docxtpl import DocxTemplate
import yagmail

# Load environment variables from .env
load_dotenv()

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

app = Flask(__name__)
TEMPLATE_FILE = "Template.docx"

# Configure Gemini using the GEMINI_API_KEY environment variable
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if not GEMINI_API_KEY:
    logger.warning("GEMINI_API_KEY not found in environment; requests to Gemini will fail until you set it in .env")

genai.configure(api_key=GEMINI_API_KEY)

# Use the specified model
MODEL_NAME = "gemini-1.5-flash"


def _safe_filename(name: str) -> str:
    """Create a filesystem-safe short filename fragment from customer name."""
    keep = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-"
    return ''.join(c if c in keep else '_' for c in name).strip('_')[:64]


def _normalize_context(ctx):
    # coerce numbers and add derived fields
    try:
        qty = int(str(ctx.get("quantity", "")).strip())
        rate = float(str(ctx.get("rate", "")).strip())
        total = qty * rate
        ctx["quantity"] = str(qty)
        ctx["rate_formatted"] = f"₹{rate:,.2f}"
        ctx["total"] = f"₹{total:,.2f}"
        ctx["rate"] = ctx["rate_formatted"]
    except Exception:
        return None

    # fill defaults
    ctx.setdefault("date", datetime.date.today().strftime("%B %d, %Y"))
    ctx.setdefault("company_name", "")
    ctx.setdefault("hsn", "")
    ctx.setdefault("q_no", "")
    if not ctx.get("units"):
        ctx["units"] = "Nos"
    return ctx


def _regex_fallback(text: str):
    """
    Very simple fallback:
    "quote 110 for Rudra at Nivee Metal, 5 pcs 5 inch SS 316L sheets at 25000, hsn 7219, email vip@example.com"
    """
    email_m = re.search(r'[\w\.-]+@[\w\.-]+', text)
    qno_m   = re.search(r'\bquote\s+(\w+)', text, re.I)
    hsn_m   = re.search(r'\bhsn\s+(\w+)', text, re.I)

    # qty + units + product + rate
    # ex: "5 pcs 5 inch SS 316L sheets at 25000"
    qty_prod_rate = re.search(
        r'(\d+)\s+(\w+)\s+(.+?)\s+at\s+(\d+(?:\.\d+)?)',
        text, re.I
    )

    # "for NAME at COMPANY"
    name_co = re.search(r'\bfor\s+(.+?)\s+at\s+(.+?)(?:,|$)', text, re.I)

    if not (email_m and qty_prod_rate and name_co):
        return None

    ctx = {
        "q_no": qno_m.group(1) if qno_m else "",
        "customer_name": name_co.group(1).strip(),
        "company_name": name_co.group(2).strip(),
        "quantity": qty_prod_rate.group(1),
        "units": qty_prod_rate.group(2),
        "product": qty_prod_rate.group(3).strip(),
        "rate": qty_prod_rate.group(4),
        "hsn": hsn_m.group(1) if hsn_m else "",
        "email": email_m.group(0)
    }
    return _normalize_context(ctx)


# Helper to create a full-path quotation using the existing create_quotation_doc
def create_quotation(context: dict):
    """Compatibility wrapper: returns full path to the generated docx or None."""
    filename = create_quotation_doc(context)
    if not filename:
        return None
    return os.path.join(os.getcwd(), filename)


def send_email_with_attachment(recipient: str, subject: str, body: str, attachment_path: str) -> bool:
    """Send an email with attachment using yagmail if credentials are available in env.

    Returns True on success, False otherwise.
    """
    gmail_user = os.getenv("GMAIL_USER")
    gmail_pass = os.getenv("GMAIL_PASS")
    if not gmail_user or not gmail_pass:
        logger.warning("GMAIL_USER or GMAIL_PASS not set; skipping email send")
        return False

    try:
        yag = yagmail.SMTP(gmail_user, gmail_pass)
        yag.send(to=recipient, subject=subject, contents=body, attachments=attachment_path)
        logger.info("Email sent to %s", recipient)
        return True
    except Exception:
        logger.exception("Failed to send email to %s", recipient)
        return False


def parse_command_with_ai(command_text: str):
    # 1) Try Gemini
    try:
        model = genai.GenerativeModel("gemini-1.5-flash")
        prompt = f"""
You are an assistant that extracts quotation data as compact JSON only (no code fences).
Fields: q_no, date, company_name, customer_name, product, quantity, rate, units, hsn, email.
If a field is missing, use an empty string. Date default is today's date ({datetime.date.today().strftime('%B %d, %Y')}).

Text: {command_text}
"""
        resp = model.generate_content(prompt)
        raw = resp.text.strip()
        # strip any accidental code fencing
        raw = raw.replace("```json", "").replace("```", "").strip()
        data = json.loads(raw)
        ctx = _normalize_context(data)
        if ctx:
            return ctx
        logger.warning("Gemini returned JSON but failed normalization; will fallback.")
    except Exception as e:
        logger.exception("Gemini parse failed; will fallback. Error: %s", e)

    # 2) Fallback to regex so the flow continues
    ctx = _regex_fallback(command_text)
    return ctx


# (request/Response already imported at top; os and datetime are imported earlier)

# WhatsApp webhook: supports both Meta format and a simple test format
@app.route("/webhook", methods=["GET", "POST"])
def webhook():
    # --- Meta verification (GET) ---
    if request.method == "GET":
        mode = request.args.get("hub.mode")
        token = request.args.get("hub.verify_token")
        challenge = request.args.get("hub.challenge")
        if mode == "subscribe" and token == os.getenv("META_VERIFY_TOKEN"):
            return Response(challenge or "", status=200)
        return Response("Verification token mismatch", status=403)

    # --- Message delivery (POST) ---
    data = request.get_json(silent=True) or {}
    # Try Meta’s structure first
    text = None
    try:
        change = data["entry"][0]["changes"][0]
        if "messages" in change["value"] and change["value"]["messages"]:
            msg = change["value"]["messages"][0]
            if msg.get("type") == "text":
                text = msg["text"]["body"]
    except Exception:
        pass

    # Fallback: simple emulator format { "message": "..." }
    if not text:
        text = data.get("message")

    if not text:
        return jsonify({"status": "ignored", "reason": "no text"}), 200

    # Use your existing pipeline
    context = parse_command_with_ai(text)
    if not context:
        return jsonify({"status": "error", "message": "Failed to parse message with Gemini"}), 200

    doc_path = create_quotation_from_template(context) if "create_quotation_from_template" in globals() else create_quotation(context)
    if not doc_path:
        return jsonify({"status": "error", "message": "Failed to generate document"}), 200

    subject = f"Quotation from NIVEE METAL PRODUCTS PVT LTD (Ref: {context.get('q_no', 'N/A')})"
    body = f"""Dear {context['customer_name']},

Thank you for your enquiry. Please find the quotation attached.

Regards,
Nivee Metal Products Pvt. Ltd.
"""
    email_ok = send_email_with_attachment(context.get("email", ""), subject, body, doc_path)

    return jsonify({
        "status": "ok" if email_ok else "mail_failed",
        "file": os.path.basename(doc_path)
    }), 200


def create_quotation_doc(context: dict) -> str:
    """Render `Template.docx` with the provided context and save the file.

    Returns the filename (relative) on success or None on failure.
    """
    try:
        if not os.path.exists(TEMPLATE_FILE):
            logger.error("Template file %s not found", TEMPLATE_FILE)
            return None

        doc = DocxTemplate(TEMPLATE_FILE)

        # Ensure some default values
        ctx = {k: (v if v is not None else '') for k, v in context.items()}
        if not ctx.get('date'):
            ctx['date'] = datetime.date.today().isoformat()

        customer = ctx.get('customer_name') or 'Customer'
        safe_customer = _safe_filename(customer)
        date_str = datetime.date.today().isoformat()
        filename = f"Quotation_{safe_customer}_{date_str}.docx"

        doc.render(ctx)
        output_path = os.path.join(os.getcwd(), filename)
        doc.save(output_path)
        logger.info("Saved quotation to %s", output_path)
        return filename
    except Exception:
        logger.exception("Failed to create quotation document")
        return None


@app.route("/", methods=["GET"])
def index():
    return "Quotation Bot is running ✅", 200


@app.route("/quote", methods=["POST"])
def quote():
    """Accepts JSON: { "message": "<user text>" }

    Uses Gemini to parse the text, fills the Word template, and returns a JSON response
    with the generated filename on success.
    """
    try:
        data = request.get_json(force=True)
    except Exception:
        logger.exception("Invalid JSON in request")
        return jsonify({"status": "error", "message": "Invalid JSON"}), 400

    message = (data or {}).get('message')
    if not message:
        return jsonify({"status": "error", "message": "Missing 'message' in JSON body"}), 400

    parsed = parse_command_with_ai(message)
    if not parsed:
        return jsonify({"status": "error", "message": "Failed to parse message with Gemini"}), 500

    # Create the docx
    filename = create_quotation_doc(parsed)
    if not filename:
        return jsonify({"status": "error", "message": "Failed to generate document"}), 500

    return jsonify({"status": "success", "file": filename}), 200


# Run instructions
if __name__ == "__main__":
    # To run locally:
    # python app.py
    app.run(host="0.0.0.0", port=5000, debug=True)
