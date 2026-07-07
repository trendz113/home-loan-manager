"""
Home Loan Manager — Flask backend
SalaryBit (salarybit.in)

Mirrors the Tax Notice Shield / Insurance Mitra Railway deployment pattern:
- Deterministic rules engine (EMI math) in Python, no AI needed for this tool
- Flask + gunicorn, single service on Railway
- Razorpay order creation done server-side only (key_secret never touches the frontend)
- Leads appended as rows to an Excel file (leads.xlsx) — see /admin/leads to download

Env vars required (set these in Railway, and locally in .env):
  RAZORPAY_KEY_ID
  RAZORPAY_KEY_SECRET
  LEADS_FILE_PATH            -> optional, defaults to ./leads.xlsx
                                 point this at a Railway Volume mount (e.g. /data/leads.xlsx)
                                 so leads survive redeploys — see README.
  RATES_FILE_PATH            -> optional, defaults to ./rates.json
                                 point this at the same Volume (e.g. /data/rates.json) so
                                 lender rate edits survive redeploys too — see README.
  ADMIN_TOKEN                -> secret used to protect the /admin/leads and /admin/rates routes
  REPORT_PRICE_PAISE         -> optional, defaults to 14900 (₹149)
  PORT                       -> set by Railway automatically
"""

import os
import json
import logging
import threading

from flask import Flask, render_template, request, jsonify, send_file, abort, Response
from flask_cors import CORS
import razorpay
from openpyxl import Workbook, load_workbook
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("home-loan-manager")

app = Flask(__name__)
CORS(app)

# ── CONFIG ──────────────────────────────────────────────────────────────
RAZORPAY_KEY_ID = os.environ.get("RAZORPAY_KEY_ID", "")
RAZORPAY_KEY_SECRET = os.environ.get("RAZORPAY_KEY_SECRET", "")
REPORT_PRICE_PAISE = int(os.environ.get("REPORT_PRICE_PAISE", "14900"))  # ₹149

# Point this at a Railway Volume mount (e.g. /data/leads.xlsx) so leads survive
# redeploys. Defaults to a local file for quick testing, which will NOT persist
# on Railway without a volume attached.
LEADS_FILE_PATH = os.environ.get("LEADS_FILE_PATH", os.path.join(os.path.dirname(__file__), "leads.xlsx"))
RATES_FILE_PATH = os.environ.get("RATES_FILE_PATH", os.path.join(os.path.dirname(__file__), "rates.json"))
ADMIN_TOKEN = os.environ.get("ADMIN_TOKEN", "")

LEADS_HEADERS = [
    "timestamp", "name", "phone", "email", "city",
    "loan_amount", "current_rate", "tenure_years",
    "best_lender", "total_savings",
    "payment_id", "order_id", "amount_paid_inr",
    "product", "status",
]

_leads_lock = threading.Lock()  # gunicorn can run multiple threads/workers; guard file writes

razorpay_client = None
if RAZORPAY_KEY_ID and RAZORPAY_KEY_SECRET:
    razorpay_client = razorpay.Client(auth=(RAZORPAY_KEY_ID, RAZORPAY_KEY_SECRET))
else:
    logger.warning("Razorpay keys not set — /api/create-order and /api/verify-payment will fail until configured.")

# ── LENDER RATE DATA ─────────────────────────────────────────────────────
# Rates live in a JSON file (RATES_FILE_PATH), not hardcoded here, so you can
# update them from the /admin/rates web form without touching code or
# redeploying. On first run (no file yet), these defaults seed the file.
# NOTE: there is no reliable free API for live Indian bank lending rates —
# banks revise floating rates mainly after RBI's ~bimonthly MPC meetings, so
# checking and updating this every couple of months (or when your DSA
# partners flag a change) is the realistic, sustainable approach.
DEFAULT_LENDER_RATES = [
    {"name": "SBI (Public Sector)", "rate": 7.35},
    {"name": "PNB / Bank of Baroda", "rate": 7.55},
    {"name": "HDFC Bank", "rate": 7.90},
    {"name": "ICICI Bank", "rate": 8.10},
    {"name": "Kotak Mahindra Bank", "rate": 8.30},
    {"name": "Axis Bank", "rate": 8.60},
]

_rates_lock = threading.Lock()


def load_lender_rates() -> list:
    if os.path.exists(RATES_FILE_PATH):
        try:
            with open(RATES_FILE_PATH, "r") as f:
                data = json.load(f)
            if isinstance(data, list) and data:
                return data
        except Exception:
            logger.exception("Failed to read %s — falling back to defaults", RATES_FILE_PATH)
    return DEFAULT_LENDER_RATES


def save_lender_rates(rates: list) -> None:
    with _rates_lock:
        os.makedirs(os.path.dirname(RATES_FILE_PATH) or ".", exist_ok=True)
        with open(RATES_FILE_PATH, "w") as f:
            json.dump(rates, f, indent=2)


# ── EMI RULES ENGINE ────────────────────────────────────────────────────
def calculate_emi(principal: float, annual_rate: float, years: float) -> float:
    r = (annual_rate / 12) / 100
    n = years * 12
    if r == 0:
        return principal / n
    return (principal * r * (1 + r) ** n) / ((1 + r) ** n - 1)


def build_calculation(principal: float, current_rate: float, years: float) -> dict:
    lender_rates = load_lender_rates()

    current_emi = calculate_emi(principal, current_rate, years)
    current_total_payment = current_emi * years * 12
    current_total_interest = current_total_payment - principal

    best_lender = min(lender_rates, key=lambda l: l["rate"])
    new_emi = calculate_emi(principal, best_lender["rate"], years)
    new_total_payment = new_emi * years * 12
    total_savings = current_total_payment - new_total_payment
    monthly_savings = current_emi - new_emi

    rate_table = sorted(lender_rates, key=lambda l: l["rate"])
    rate_table_out = [
        {
            "name": l["name"],
            "rate": l["rate"],
            "emi": round(calculate_emi(principal, l["rate"], years)),
            "is_lowest": i == 0,
        }
        for i, l in enumerate(rate_table)
    ]

    return {
        "current_emi": round(current_emi),
        "current_total_interest": round(current_total_interest),
        "current_rate": current_rate,
        "best_rate": best_lender["rate"],
        "best_lender_name": best_lender["name"],
        "total_savings": round(total_savings),
        "monthly_savings": round(monthly_savings),
        "worth_switching": total_savings > 0,
        "rate_table": rate_table_out,
    }


# ── ROUTES ───────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html", razorpay_key_id=RAZORPAY_KEY_ID)


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/api/calculate-emi", methods=["POST"])
def api_calculate_emi():
    data = request.get_json(silent=True) or {}
    try:
        principal = float(data.get("loan_amount", 0))
        current_rate = float(data.get("current_rate", 0))
        years = float(data.get("tenure_years", 0))
    except (TypeError, ValueError):
        return jsonify({"error": "Invalid input. Please send numeric values."}), 400

    if principal <= 0 or current_rate <= 0 or years <= 0:
        return jsonify({"error": "loan_amount, current_rate, and tenure_years must all be positive numbers."}), 400

    result = build_calculation(principal, current_rate, years)
    return jsonify(result)


@app.route("/api/create-order", methods=["POST"])
def api_create_order():
    if not razorpay_client:
        return jsonify({"error": "Payments are not configured on this server yet."}), 503

    data = request.get_json(silent=True) or {}
    # Optional: allow the calc inputs to be stashed in notes so you can see them in the Razorpay dashboard
    notes = {
        "product": "home-loan-manager-report",
        "loan_amount": str(data.get("loan_amount", "")),
        "current_rate": str(data.get("current_rate", "")),
        "tenure_years": str(data.get("tenure_years", "")),
    }

    try:
        order = razorpay_client.order.create(
            {
                "amount": REPORT_PRICE_PAISE,
                "currency": "INR",
                "receipt": f"hlm_{os.urandom(6).hex()}",
                "notes": notes,
                "payment_capture": 1,
            }
        )
    except Exception as exc:
        logger.exception("Razorpay order creation failed")
        return jsonify({"error": "Could not create payment order.", "detail": str(exc)}), 502

    return jsonify(
        {
            "order_id": order["id"],
            "amount": order["amount"],
            "currency": order["currency"],
            "key_id": RAZORPAY_KEY_ID,
        }
    )


@app.route("/api/verify-payment", methods=["POST"])
def api_verify_payment():
    """
    Call this from the Razorpay checkout `handler` callback with:
      { razorpay_order_id, razorpay_payment_id, razorpay_signature, lead: {...}, calc: {...} }
    Verifies the signature server-side, then forwards the lead (with payment_id)
    to leads.xlsx so it becomes an actual DSA-routable lead.
    """
    if not razorpay_client:
        return jsonify({"error": "Payments are not configured on this server yet."}), 503

    data = request.get_json(silent=True) or {}
    params = {
        "razorpay_order_id": data.get("razorpay_order_id", ""),
        "razorpay_payment_id": data.get("razorpay_payment_id", ""),
        "razorpay_signature": data.get("razorpay_signature", ""),
    }

    if not all(params.values()):
        return jsonify({"error": "Missing Razorpay verification fields."}), 400

    try:
        razorpay_client.utility.verify_payment_signature(params)
    except razorpay.errors.SignatureVerificationError:
        return jsonify({"error": "Payment signature verification failed.", "verified": False}), 400

    lead = data.get("lead", {})
    calc = data.get("calc", {})
    lead_record = {
        "name": lead.get("name", ""),
        "phone": lead.get("phone", ""),
        "email": lead.get("email", ""),
        "city": lead.get("city", ""),
        "loan_amount": calc.get("loan_amount", ""),
        "current_rate": calc.get("current_rate", ""),
        "tenure_years": calc.get("tenure_years", ""),
        "best_lender": calc.get("best_lender_name", ""),
        "total_savings": calc.get("total_savings", ""),
        "payment_id": params["razorpay_payment_id"],
        "order_id": params["razorpay_order_id"],
        "amount_paid_inr": REPORT_PRICE_PAISE / 100,
        "product": "home-loan-manager-report",
        "status": "paid",
    }

    sheet_ok = push_lead_to_sheet(lead_record)

    return jsonify({"verified": True, "lead_saved": sheet_ok})


@app.route("/api/lead-capture", methods=["POST"])
def api_lead_capture():
    """
    Free-standing lead capture (e.g. if you ever add a pre-payment 'notify me'
    form, or want to log a lead even before checkout completes).
    """
    data = request.get_json(silent=True) or {}
    name = (data.get("name") or "").strip()
    phone = (data.get("phone") or "").strip()

    if not name or not phone:
        return jsonify({"error": "name and phone are required."}), 400

    lead_record = {
        "name": name,
        "phone": phone,
        "email": (data.get("email") or "").strip(),
        "city": (data.get("city") or "").strip(),
        "loan_amount": data.get("loan_amount", ""),
        "current_rate": data.get("current_rate", ""),
        "tenure_years": data.get("tenure_years", ""),
        "payment_id": "",
        "order_id": "",
        "amount_paid_inr": 0,
        "product": "home-loan-manager-lead",
        "status": "unpaid_lead",
    }

    sheet_ok = push_lead_to_sheet(lead_record)
    return jsonify({"saved": sheet_ok})


def _get_or_create_workbook():
    if os.path.exists(LEADS_FILE_PATH):
        return load_workbook(LEADS_FILE_PATH)
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"
    ws.append(LEADS_HEADERS)
    return wb


def push_lead_to_sheet(record: dict) -> bool:
    """
    Appends a lead as a new row in leads.xlsx (kept at LEADS_FILE_PATH).
    NOTE: on Railway, LEADS_FILE_PATH must point at a mounted Volume or this
    data is lost on every redeploy/restart. See README for the volume setup.
    """
    import datetime

    row = [datetime.datetime.utcnow().isoformat()] + [record.get(h, "") for h in LEADS_HEADERS[1:]]

    try:
        with _leads_lock:
            os.makedirs(os.path.dirname(LEADS_FILE_PATH) or ".", exist_ok=True)
            wb = _get_or_create_workbook()
            ws = wb["Leads"] if "Leads" in wb.sheetnames else wb.active
            ws.append(row)
            wb.save(LEADS_FILE_PATH)
        return True
    except Exception:
        logger.exception("Failed to append lead to Excel file at %s", LEADS_FILE_PATH)
        return False


@app.route("/admin/rates", methods=["GET", "POST"])
def admin_rates():
    """
    Simple password-protected page to view/edit lender rates without touching
    code or redeploying. Visit /admin/rates?token=YOUR_ADMIN_TOKEN
    """
    if not ADMIN_TOKEN:
        abort(503, description="ADMIN_TOKEN is not configured on this server.")
    token = request.args.get("token") or request.form.get("token")
    if token != ADMIN_TOKEN:
        abort(403)

    message = ""
    if request.method == "POST":
        names = request.form.getlist("name")
        rates = request.form.getlist("rate")
        new_rates = []
        try:
            for name, rate in zip(names, rates):
                name = name.strip()
                if not name:
                    continue
                new_rates.append({"name": name, "rate": float(rate)})
            if not new_rates:
                raise ValueError("At least one lender row is required.")
            save_lender_rates(new_rates)
            message = "Rates updated successfully."
        except ValueError as exc:
            message = f"Error: {exc}"

    current_rates = load_lender_rates()

    rows_html = "".join(
        f"""
        <tr>
          <td><input type="text" name="name" value="{l['name']}" style="width:220px;padding:6px"></td>
          <td><input type="number" step="0.01" name="rate" value="{l['rate']}" style="width:90px;padding:6px"></td>
        </tr>
        """
        for l in current_rates
    )

    # A few blank rows so you can add new lenders without editing HTML
    blank_rows = "".join(
        """
        <tr>
          <td><input type="text" name="name" value="" style="width:220px;padding:6px" placeholder="Lender name"></td>
          <td><input type="number" step="0.01" name="rate" value="" style="width:90px;padding:6px" placeholder="Rate %"></td>
        </tr>
        """
        for _ in range(3)
    )

    html = f"""
    <html>
    <head><title>Home Loan Manager — Update Rates</title>
    <style>
      body {{ font-family: sans-serif; max-width: 600px; margin: 40px auto; color: #1a1a2e; }}
      h1 {{ font-size: 20px; }}
      table {{ border-collapse: collapse; margin: 20px 0; }}
      th, td {{ padding: 6px 10px; text-align: left; }}
      button {{ background: #5b6af0; color: #fff; border: none; padding: 10px 20px; border-radius: 6px; font-size: 14px; cursor: pointer; }}
      .msg {{ padding: 10px; border-radius: 6px; background: #eeeef6; margin-bottom: 10px; }}
    </style>
    </head>
    <body>
      <h1>Lender Rates</h1>
      <p>Update starting rates here whenever RBI revises the repo rate or a DSA partner flags a change. No redeploy needed — saved instantly.</p>
      {'<div class="msg">' + message + '</div>' if message else ''}
      <form method="POST">
        <input type="hidden" name="token" value="{token}">
        <table>
          <tr><th>Lender Name</th><th>Rate (%)</th></tr>
          {rows_html}
          {blank_rows}
        </table>
        <button type="submit">Save Rates</button>
      </form>
    </body>
    </html>
    """
    return Response(html, mimetype="text/html")


@app.route("/admin/leads")
def admin_download_leads():
    """
    Download the current leads.xlsx file.
    Protect with ?token=YOUR_ADMIN_TOKEN (set ADMIN_TOKEN env var).
    """
    if not ADMIN_TOKEN:
        abort(503, description="ADMIN_TOKEN is not configured on this server.")
    if request.args.get("token") != ADMIN_TOKEN:
        abort(403)
    if not os.path.exists(LEADS_FILE_PATH):
        abort(404, description="No leads recorded yet.")
    return send_file(LEADS_FILE_PATH, as_attachment=True, download_name="leads.xlsx")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
