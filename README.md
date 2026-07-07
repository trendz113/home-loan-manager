# Home Loan Manager — SalaryBit

Flask backend + frontend for the free EMI/balance-transfer calculator and the
₹149 paid "Personalized Switching Report" upsell. Matches the same
architecture pattern as Tax Notice Shield and Insurance Mitra: deterministic
Python rules engine, single Flask service, deployed on Railway.

## Routes

| Route | Method | Purpose |
|---|---|---|
| `/` | GET | Renders the page (injects the public Razorpay `key_id` into the template) |
| `/health` | GET | Health check for Railway |
| `/api/calculate-emi` | POST | Free EMI + balance-transfer savings calculator |
| `/api/create-order` | POST | Creates a Razorpay order server-side (₹149 by default) |
| `/api/verify-payment` | POST | Verifies the Razorpay signature after checkout, then appends the lead to `leads.xlsx` |
| `/api/lead-capture` | POST | Stand-alone lead capture, independent of payment (e.g. for a future "notify me" form) |
| `/admin/leads?token=...` | GET | Downloads the current `leads.xlsx` file (token-protected) |

## Local setup

```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
cp .env.example .env   # then fill in real values
python app.py
```

Visit `http://localhost:5000`.

## Environment variables

Set these in Railway's Variables tab (same as your other services):

- `RAZORPAY_KEY_ID` — public key id (also injected into the frontend)
- `RAZORPAY_KEY_SECRET` — **never** exposed to the frontend, used only server-side to create orders and verify signatures
- `LEADS_FILE_PATH` — where `leads.xlsx` is written, e.g. `/data/leads.xlsx` (see Railway Volume setup below)
- `ADMIN_TOKEN` — long random secret that protects the `/admin/leads` download route
- `REPORT_PRICE_PAISE` — optional, defaults to `14900` (₹149)

## Lead storage: Excel file + Railway Volume

Leads (both paid and free lead-capture) are appended as rows to an `.xlsx`
file using `openpyxl` — no third-party webhook needed.

**Important — Railway's filesystem is ephemeral.** Anything written to disk
is wiped on every redeploy or restart *unless* you attach a Volume:

1. In your Railway project → the service → **Settings → Volumes → New Volume**.
2. Mount path: `/data` (any empty path works, just match it below).
3. Set `LEADS_FILE_PATH=/data/leads.xlsx` in the service's environment variables.
4. Redeploy. From then on, `leads.xlsx` persists across deploys and restarts.

To download the current leads file at any time, visit:

```
https://<your-railway-url>/admin/leads?token=<your ADMIN_TOKEN value>
```

Open it in Excel, Google Sheets (File → Import), or LibreOffice — it's a
standard `.xlsx` with one row per lead and headers:
`timestamp, name, phone, email, city, loan_amount, current_rate, tenure_years, best_lender, total_savings, payment_id, order_id, amount_paid_inr, product, status`.

**Backup tip:** since this is your only copy of lead data, periodically hit
the `/admin/leads` URL and save a dated copy locally, or script a scheduled
download (e.g. a small cron on your machine, or a Railway cron job hitting
the endpoint and emailing you the file) so a lost Volume never means lost leads.

## Automated daily backup (`backup/backup_leads.py`)

A small standalone script (run on your own laptop/PC, not on Railway) that
downloads `/admin/leads` every day and saves it as `leads_YYYY-MM-DD.xlsx`,
pruning anything older than `KEEP_LAST_N` days (default 90).

### One-time setup

```bash
cd backup
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env
# edit .env: set APP_URL to your Railway URL, ADMIN_TOKEN to match the server
python backup_leads.py          # test it once manually
```

Check `backup/backups/leads_<today>.xlsx` was created.

### Schedule it — macOS/Linux (cron)

```bash
crontab -e
```

Add a line to run daily at 9 AM (adjust paths to match your machine):

```
0 9 * * * cd /full/path/to/home-loan-manager/backup && /full/path/to/venv/bin/python backup_leads.py >> backup.log 2>&1
```

### Schedule it — Windows (Task Scheduler)

1. Open **Task Scheduler** → **Create Basic Task**.
2. Trigger: **Daily**, pick a time (e.g. 9:00 AM).
3. Action: **Start a program**.
   - Program/script: full path to `venv\Scripts\python.exe`
   - Add arguments: `backup_leads.py`
   - Start in: full path to the `backup` folder
4. Finish. Test it once with "Run" in Task Scheduler to confirm a file appears in `backup\backups`.

### Notes

- The script exits cleanly (code 0, just a warning) if no leads exist yet —
  safe to run before you have any traffic.
- If `ADMIN_TOKEN` doesn't match the server, it fails loudly (exit code 1)
  instead of silently skipping — check `backup.log` if a scheduled run seems
  to have done nothing.
- Since backups land on your own machine, they're a genuine second copy —
  independent of the Railway Volume, so even a volume failure or accidental
  file deletion on the server doesn't cost you more than a day's leads.

## Deploying to Railway

1. Push this repo to a new GitHub repo (e.g. `trendz113/home-loan-manager`).
2. In Railway: New Project → Deploy from GitHub repo.
3. Railway auto-detects the `Procfile` and Python buildpack; no extra config needed.
4. Add the environment variables above in the Railway dashboard.
5. Once deployed, note the generated `*.up.railway.app` URL — same pattern as
   `web-production-796d.up.railway.app` (Tax Notice Shield) and
   `web-production-b0a7.up.railway.app` (Insurance Mitra).
6. Point a CNAME / link from salarybit.in (or GitHub Pages nav) to this service.

## Notes

- Lender rate table lives in `app.py` (`LENDER_RATES`) — the single source of
  truth. Update it whenever RBI revises the repo rate or you get fresh DSA
  rate sheets; the frontend no longer hardcodes rates.
- Razorpay `key_secret` never reaches the browser — only `key_id` is templated
  into `index.html`, and order creation + signature verification happen in
  `app.py`.
- `/api/verify-payment` appends to `leads.xlsx` only after the Razorpay
  signature verifies, so every row in the file corresponds to a real,
  confirmed payment (except rows from `/api/lead-capture`, marked
  `status = unpaid_lead`).
- A `threading.Lock` guards writes so concurrent requests (e.g. gunicorn
  running multiple workers/threads) don't corrupt the file — fine at low
  volume; if lead volume grows a lot, consider moving to a real database.
