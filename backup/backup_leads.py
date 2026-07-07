"""
backup_leads.py — pulls /admin/leads from your deployed Home Loan Manager
service and saves a dated .xlsx copy locally. Run this daily (see cron /
Task Scheduler setup in README below) so a lost Railway Volume or accidental
overwrite never means lost leads.

Usage:
    python backup_leads.py

Config (env vars, or a .env file in this folder):
    APP_URL       e.g. https://web-production-xxxx.up.railway.app
    ADMIN_TOKEN   same value as set on the Railway service
    BACKUP_DIR    optional, defaults to ./backups
    KEEP_LAST_N   optional, defaults to 90 (old backups beyond this are deleted)
"""

import os
import sys
import glob
import datetime
import logging

import requests
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("backup_leads")

APP_URL = os.environ.get("APP_URL", "").rstrip("/")
ADMIN_TOKEN = os.environ.get("ADMIN_TOKEN", "")
BACKUP_DIR = os.environ.get("BACKUP_DIR", os.path.join(os.path.dirname(__file__), "backups"))
KEEP_LAST_N = int(os.environ.get("KEEP_LAST_N", "90"))


def main():
    if not APP_URL or not ADMIN_TOKEN:
        logger.error("APP_URL and ADMIN_TOKEN must be set (env vars or .env file). Aborting.")
        sys.exit(1)

    os.makedirs(BACKUP_DIR, exist_ok=True)

    today = datetime.date.today().isoformat()  # e.g. 2026-07-07
    dest_path = os.path.join(BACKUP_DIR, f"leads_{today}.xlsx")

    url = f"{APP_URL}/admin/leads"
    try:
        resp = requests.get(url, params={"token": ADMIN_TOKEN}, timeout=30)
    except requests.RequestException as exc:
        logger.error("Could not reach %s: %s", url, exc)
        sys.exit(1)

    if resp.status_code == 404:
        logger.warning("No leads recorded yet on the server (404). Nothing to back up today.")
        sys.exit(0)
    if resp.status_code == 403:
        logger.error("Admin token rejected (403). Check ADMIN_TOKEN matches the Railway env var.")
        sys.exit(1)
    if resp.status_code != 200:
        logger.error("Unexpected response %s from %s", resp.status_code, url)
        sys.exit(1)

    with open(dest_path, "wb") as f:
        f.write(resp.content)

    logger.info("Saved backup: %s (%d KB)", dest_path, len(resp.content) // 1024)

    _prune_old_backups()


def _prune_old_backups():
    files = sorted(glob.glob(os.path.join(BACKUP_DIR, "leads_*.xlsx")))
    excess = len(files) - KEEP_LAST_N
    if excess > 0:
        for f in files[:excess]:
            os.remove(f)
            logger.info("Pruned old backup: %s", f)


if __name__ == "__main__":
    main()
