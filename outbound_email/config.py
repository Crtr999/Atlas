"""
Configuration for the outbound email system.
Loads from environment variables with sensible defaults.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ── Paths ─────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = os.environ.get("EMAIL_DB_PATH", str(BASE_DIR / "outbound_email.db"))

# ── Email Sending (SMTP) ─────────────────────────────────────────
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USERNAME = os.environ.get("SMTP_USERNAME", "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")
SMTP_USE_TLS = os.environ.get("SMTP_USE_TLS", "true").lower() == "true"

# ── Email Sending (SendGrid - alternative) ────────────────────────
SENDGRID_API_KEY = os.environ.get("SENDGRID_API_KEY", "")
EMAIL_PROVIDER = os.environ.get("EMAIL_PROVIDER", "smtp")  # "smtp" or "sendgrid"

# ── Email Receiving (IMAP) ────────────────────────────────────────
IMAP_HOST = os.environ.get("IMAP_HOST", "imap.gmail.com")
IMAP_PORT = int(os.environ.get("IMAP_PORT", "993"))
IMAP_USERNAME = os.environ.get("IMAP_USERNAME", "")
IMAP_PASSWORD = os.environ.get("IMAP_PASSWORD", "")
IMAP_USE_SSL = os.environ.get("IMAP_USE_SSL", "true").lower() == "true"
IMAP_FOLDER = os.environ.get("IMAP_FOLDER", "INBOX")

# ── Sender Identity ──────────────────────────────────────────────
SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "")
SENDER_NAME = os.environ.get("SENDER_NAME", "Deed Street Capital")

# ── Rate Limiting ─────────────────────────────────────────────────
DAILY_SEND_LIMIT = int(os.environ.get("DAILY_SEND_LIMIT", "500"))
# Delay between individual emails (seconds) to avoid spam flags
SEND_DELAY_SECONDS = float(os.environ.get("SEND_DELAY_SECONDS", "5"))
# Warmup: start with fewer emails and ramp up over days
WARMUP_ENABLED = os.environ.get("WARMUP_ENABLED", "true").lower() == "true"
WARMUP_START_COUNT = int(os.environ.get("WARMUP_START_COUNT", "25"))
WARMUP_DAILY_INCREMENT = int(os.environ.get("WARMUP_DAILY_INCREMENT", "25"))

# ── Claude AI ─────────────────────────────────────────────────────
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL = os.environ.get("CLAUDE_MODEL", "claude-sonnet-4-5-20250929")

# ── Calendly ──────────────────────────────────────────────────────
CALENDLY_LINK = os.environ.get("CALENDLY_LINK", "")

# ── Polling ───────────────────────────────────────────────────────
# Check for replies every 10 minutes (600s) for fast response turnaround
IMAP_POLL_INTERVAL = int(os.environ.get("IMAP_POLL_INTERVAL", "600"))  # seconds

# ── Company Info (for AI context) ─────────────────────────────────
COMPANY_NAME = "Deed Street Capital"
COMPANY_DESCRIPTION = (
    "Deed Street Capital is a national leader in the secondary market for "
    "seller-financed, privately-issued mortgage notes. We purchase private "
    "mortgage notes secured by commercial and residential properties with "
    "note balances from $25,000 up to $3,000,000. As a direct buyer with "
    "over 20 years of experience, we eliminate middlemen and offer transparent, "
    "competitive pricing with fast closings — most within 30 days. We operate "
    "in all 50 states."
)
COMPANY_VALUE_PROPS = [
    "Direct buyer — no middlemen, best price guaranteed",
    "Fast closings — most transactions completed within 30 days",
    "Nationwide service — we buy notes in all 50 states",
    "No hidden fees — transparent pricing from start to finish",
    "Flexible options — full or partial note purchases available",
    "20+ years of experience in mortgage note acquisitions",
    "Proprietary underwriting model for accurate, competitive pricing",
]


def validate_config():
    """Check that required configuration is set. Returns list of missing items."""
    missing = []
    if not SMTP_USERNAME and EMAIL_PROVIDER == "smtp":
        missing.append("SMTP_USERNAME")
    if not SMTP_PASSWORD and EMAIL_PROVIDER == "smtp":
        missing.append("SMTP_PASSWORD")
    if not SENDGRID_API_KEY and EMAIL_PROVIDER == "sendgrid":
        missing.append("SENDGRID_API_KEY")
    if not IMAP_USERNAME:
        missing.append("IMAP_USERNAME")
    if not IMAP_PASSWORD:
        missing.append("IMAP_PASSWORD")
    if not SENDER_EMAIL:
        missing.append("SENDER_EMAIL")
    if not ANTHROPIC_API_KEY:
        missing.append("ANTHROPIC_API_KEY")
    if not CALENDLY_LINK:
        missing.append("CALENDLY_LINK")
    return missing
