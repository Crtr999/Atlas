"""
SQLite database layer for prospect tracking, campaign management,
and email conversation history.
"""

import sqlite3
import json
import logging
from datetime import datetime, date
from pathlib import Path
from contextlib import contextmanager

from . import config

logger = logging.getLogger("outbound-email")

SCHEMA = """
CREATE TABLE IF NOT EXISTS prospects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT UNIQUE NOT NULL,
    first_name TEXT DEFAULT '',
    last_name TEXT DEFAULT '',
    company TEXT DEFAULT '',
    phone TEXT DEFAULT '',
    note_type TEXT DEFAULT '',
    note_balance TEXT DEFAULT '',
    property_state TEXT DEFAULT '',
    source TEXT DEFAULT '',
    tags TEXT DEFAULT '',
    status TEXT DEFAULT 'new'
        CHECK(status IN ('new','contacted','engaged','calendly_sent',
                         'meeting_booked','closed','opted_out','bounced')),
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS campaigns (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    subject_template TEXT NOT NULL,
    body_template TEXT NOT NULL,
    is_active INTEGER DEFAULT 1,
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS emails_sent (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    prospect_id INTEGER NOT NULL REFERENCES prospects(id),
    campaign_id INTEGER REFERENCES campaigns(id),
    message_id TEXT UNIQUE,
    subject TEXT NOT NULL,
    body TEXT NOT NULL,
    sent_at TEXT DEFAULT (datetime('now')),
    is_auto_reply INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS emails_received (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    prospect_id INTEGER REFERENCES prospects(id),
    from_email TEXT NOT NULL,
    subject TEXT DEFAULT '',
    body TEXT NOT NULL,
    raw_message_id TEXT,
    in_reply_to TEXT,
    received_at TEXT DEFAULT (datetime('now')),
    processed INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS conversations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    prospect_id INTEGER UNIQUE NOT NULL REFERENCES prospects(id),
    status TEXT DEFAULT 'initial_outreach'
        CHECK(status IN ('initial_outreach','awaiting_reply','engaged',
                         'calendly_sent','meeting_booked','closed','opted_out')),
    ai_summary TEXT DEFAULT '',
    last_activity TEXT DEFAULT (datetime('now')),
    reply_count INTEGER DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS daily_send_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    send_date TEXT NOT NULL,
    count INTEGER DEFAULT 0,
    UNIQUE(send_date)
);

CREATE INDEX IF NOT EXISTS idx_prospects_email ON prospects(email);
CREATE INDEX IF NOT EXISTS idx_prospects_status ON prospects(status);
CREATE INDEX IF NOT EXISTS idx_emails_sent_prospect ON emails_sent(prospect_id);
CREATE INDEX IF NOT EXISTS idx_emails_sent_message_id ON emails_sent(message_id);
CREATE INDEX IF NOT EXISTS idx_emails_received_prospect ON emails_received(prospect_id);
CREATE INDEX IF NOT EXISTS idx_emails_received_in_reply_to ON emails_received(in_reply_to);
CREATE INDEX IF NOT EXISTS idx_conversations_prospect ON conversations(prospect_id);
CREATE INDEX IF NOT EXISTS idx_conversations_status ON conversations(status);
CREATE INDEX IF NOT EXISTS idx_daily_send_log_date ON daily_send_log(send_date);
"""


class Database:
    """SQLite database for outbound email tracking."""

    def __init__(self, db_path: str = None):
        self.db_path = db_path or config.DB_PATH
        self._init_db()

    def _init_db(self):
        with self._connect() as conn:
            conn.executescript(SCHEMA)

    @contextmanager
    def _connect(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    # ── Prospects ──────────────────────────────────────────────────

    def add_prospect(self, email: str, first_name: str = "", last_name: str = "",
                     company: str = "", phone: str = "", note_type: str = "",
                     note_balance: str = "", property_state: str = "",
                     source: str = "", tags: str = "") -> int:
        with self._connect() as conn:
            cursor = conn.execute(
                """INSERT OR IGNORE INTO prospects
                   (email, first_name, last_name, company, phone,
                    note_type, note_balance, property_state, source, tags)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (email.lower().strip(), first_name, last_name, company, phone,
                 note_type, note_balance, property_state, source, tags),
            )
            if cursor.rowcount == 0:
                row = conn.execute(
                    "SELECT id FROM prospects WHERE email = ?",
                    (email.lower().strip(),),
                ).fetchone()
                return row["id"]
            return cursor.lastrowid

    def import_prospects_csv(self, csv_path: str) -> dict:
        """Import prospects from a CSV file. Returns counts."""
        import csv
        added = 0
        skipped = 0
        with open(csv_path, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                email = row.get("email", "").strip()
                if not email:
                    skipped += 1
                    continue
                pid = self.add_prospect(
                    email=email,
                    first_name=row.get("first_name", ""),
                    last_name=row.get("last_name", ""),
                    company=row.get("company", ""),
                    phone=row.get("phone", ""),
                    note_type=row.get("note_type", ""),
                    note_balance=row.get("note_balance", ""),
                    property_state=row.get("property_state", ""),
                    source=row.get("source", ""),
                    tags=row.get("tags", ""),
                )
                if pid:
                    added += 1
                else:
                    skipped += 1
        return {"added": added, "skipped": skipped}

    def get_unsent_prospects(self, campaign_id: int, limit: int = 500) -> list:
        """Get prospects who haven't been emailed for a given campaign."""
        with self._connect() as conn:
            rows = conn.execute(
                """SELECT p.* FROM prospects p
                   WHERE p.status NOT IN ('opted_out', 'bounced')
                   AND p.id NOT IN (
                       SELECT prospect_id FROM emails_sent
                       WHERE campaign_id = ? AND is_auto_reply = 0
                   )
                   ORDER BY p.created_at ASC
                   LIMIT ?""",
                (campaign_id, limit),
            ).fetchall()
            return [dict(r) for r in rows]

    def get_prospect_by_email(self, email: str) -> dict | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM prospects WHERE email = ?",
                (email.lower().strip(),),
            ).fetchone()
            return dict(row) if row else None

    def get_prospect_by_id(self, prospect_id: int) -> dict | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM prospects WHERE id = ?", (prospect_id,)
            ).fetchone()
            return dict(row) if row else None

    def update_prospect_status(self, prospect_id: int, status: str):
        with self._connect() as conn:
            conn.execute(
                "UPDATE prospects SET status = ?, updated_at = datetime('now') WHERE id = ?",
                (status, prospect_id),
            )

    def get_all_prospects(self, status_filter: str = None) -> list:
        with self._connect() as conn:
            if status_filter:
                rows = conn.execute(
                    "SELECT * FROM prospects WHERE status = ? ORDER BY created_at DESC",
                    (status_filter,),
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM prospects ORDER BY created_at DESC"
                ).fetchall()
            return [dict(r) for r in rows]

    # ── Campaigns ─────────────────────────────────────────────────

    def create_campaign(self, name: str, subject_template: str,
                        body_template: str) -> int:
        with self._connect() as conn:
            cursor = conn.execute(
                """INSERT INTO campaigns (name, subject_template, body_template)
                   VALUES (?, ?, ?)""",
                (name, subject_template, body_template),
            )
            return cursor.lastrowid

    def get_campaign(self, campaign_id: int) -> dict | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM campaigns WHERE id = ?", (campaign_id,)
            ).fetchone()
            return dict(row) if row else None

    def get_campaign_by_name(self, name: str) -> dict | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM campaigns WHERE name = ?", (name,)
            ).fetchone()
            return dict(row) if row else None

    def list_campaigns(self) -> list:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM campaigns ORDER BY created_at DESC"
            ).fetchall()
            return [dict(r) for r in rows]

    # ── Emails Sent ───────────────────────────────────────────────

    def log_sent_email(self, prospect_id: int, subject: str, body: str,
                       message_id: str = None, campaign_id: int = None,
                       is_auto_reply: bool = False) -> int:
        with self._connect() as conn:
            cursor = conn.execute(
                """INSERT INTO emails_sent
                   (prospect_id, campaign_id, message_id, subject, body, is_auto_reply)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (prospect_id, campaign_id, message_id, subject, body,
                 1 if is_auto_reply else 0),
            )
            return cursor.lastrowid

    def get_sent_emails_for_prospect(self, prospect_id: int) -> list:
        with self._connect() as conn:
            rows = conn.execute(
                """SELECT * FROM emails_sent WHERE prospect_id = ?
                   ORDER BY sent_at ASC""",
                (prospect_id,),
            ).fetchall()
            return [dict(r) for r in rows]

    def get_sent_email_by_message_id(self, message_id: str) -> dict | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM emails_sent WHERE message_id = ?",
                (message_id,),
            ).fetchone()
            return dict(row) if row else None

    # ── Emails Received ───────────────────────────────────────────

    def log_received_email(self, from_email: str, subject: str, body: str,
                           raw_message_id: str = None, in_reply_to: str = None,
                           prospect_id: int = None) -> int:
        with self._connect() as conn:
            cursor = conn.execute(
                """INSERT INTO emails_received
                   (prospect_id, from_email, subject, body, raw_message_id, in_reply_to)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (prospect_id, from_email.lower().strip(), subject, body,
                 raw_message_id, in_reply_to),
            )
            return cursor.lastrowid

    def get_unprocessed_replies(self) -> list:
        with self._connect() as conn:
            rows = conn.execute(
                """SELECT * FROM emails_received
                   WHERE processed = 0
                   ORDER BY received_at ASC"""
            ).fetchall()
            return [dict(r) for r in rows]

    def mark_reply_processed(self, reply_id: int):
        with self._connect() as conn:
            conn.execute(
                "UPDATE emails_received SET processed = 1 WHERE id = ?",
                (reply_id,),
            )

    def get_received_emails_for_prospect(self, prospect_id: int) -> list:
        with self._connect() as conn:
            rows = conn.execute(
                """SELECT * FROM emails_received WHERE prospect_id = ?
                   ORDER BY received_at ASC""",
                (prospect_id,),
            ).fetchall()
            return [dict(r) for r in rows]

    # ── Conversations ─────────────────────────────────────────────

    def get_or_create_conversation(self, prospect_id: int) -> dict:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM conversations WHERE prospect_id = ?",
                (prospect_id,),
            ).fetchone()
            if row:
                return dict(row)
            conn.execute(
                "INSERT INTO conversations (prospect_id) VALUES (?)",
                (prospect_id,),
            )
            row = conn.execute(
                "SELECT * FROM conversations WHERE prospect_id = ?",
                (prospect_id,),
            ).fetchone()
            return dict(row)

    def update_conversation(self, prospect_id: int, status: str = None,
                            ai_summary: str = None):
        with self._connect() as conn:
            updates = ["last_activity = datetime('now')"]
            params = []
            if status:
                updates.append("status = ?")
                params.append(status)
            if ai_summary is not None:
                updates.append("ai_summary = ?")
                params.append(ai_summary)
            updates.append("reply_count = reply_count + 1")
            params.append(prospect_id)
            conn.execute(
                f"UPDATE conversations SET {', '.join(updates)} WHERE prospect_id = ?",
                params,
            )

    def get_conversation(self, prospect_id: int) -> dict | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM conversations WHERE prospect_id = ?",
                (prospect_id,),
            ).fetchone()
            return dict(row) if row else None

    def get_full_thread(self, prospect_id: int) -> list:
        """Get the full email thread (sent + received) for a prospect, chronologically."""
        sent = self.get_sent_emails_for_prospect(prospect_id)
        received = self.get_received_emails_for_prospect(prospect_id)
        thread = []
        for e in sent:
            thread.append({
                "role": "sent",
                "subject": e["subject"],
                "body": e["body"],
                "timestamp": e["sent_at"],
            })
        for e in received:
            thread.append({
                "role": "received",
                "subject": e["subject"],
                "body": e["body"],
                "timestamp": e["received_at"],
            })
        thread.sort(key=lambda x: x["timestamp"])
        return thread

    # ── Daily Send Tracking ───────────────────────────────────────

    def get_today_send_count(self) -> int:
        today = date.today().isoformat()
        with self._connect() as conn:
            row = conn.execute(
                "SELECT count FROM daily_send_log WHERE send_date = ?",
                (today,),
            ).fetchone()
            return row["count"] if row else 0

    def increment_send_count(self):
        today = date.today().isoformat()
        with self._connect() as conn:
            conn.execute(
                """INSERT INTO daily_send_log (send_date, count) VALUES (?, 1)
                   ON CONFLICT(send_date) DO UPDATE SET count = count + 1""",
                (today,),
            )

    def get_daily_limit(self) -> int:
        """Calculate today's send limit, accounting for warmup."""
        if not config.WARMUP_ENABLED:
            return config.DAILY_SEND_LIMIT

        with self._connect() as conn:
            row = conn.execute(
                "SELECT COUNT(DISTINCT send_date) as days FROM daily_send_log"
            ).fetchone()
            days_active = row["days"] if row else 0

        warmup_limit = config.WARMUP_START_COUNT + (days_active * config.WARMUP_DAILY_INCREMENT)
        return min(warmup_limit, config.DAILY_SEND_LIMIT)

    def get_remaining_sends_today(self) -> int:
        return max(0, self.get_daily_limit() - self.get_today_send_count())

    # ── Stats ─────────────────────────────────────────────────────

    def get_stats(self) -> dict:
        with self._connect() as conn:
            total_prospects = conn.execute("SELECT COUNT(*) as c FROM prospects").fetchone()["c"]
            by_status = {}
            for row in conn.execute(
                "SELECT status, COUNT(*) as c FROM prospects GROUP BY status"
            ).fetchall():
                by_status[row["status"]] = row["c"]

            total_sent = conn.execute("SELECT COUNT(*) as c FROM emails_sent WHERE is_auto_reply = 0").fetchone()["c"]
            total_replies = conn.execute("SELECT COUNT(*) as c FROM emails_received").fetchone()["c"]
            auto_replies_sent = conn.execute("SELECT COUNT(*) as c FROM emails_sent WHERE is_auto_reply = 1").fetchone()["c"]

            conv_stats = {}
            for row in conn.execute(
                "SELECT status, COUNT(*) as c FROM conversations GROUP BY status"
            ).fetchall():
                conv_stats[row["status"]] = row["c"]

            today_sent = self.get_today_send_count()
            today_limit = self.get_daily_limit()

            return {
                "total_prospects": total_prospects,
                "prospects_by_status": by_status,
                "total_outreach_sent": total_sent,
                "total_replies_received": total_replies,
                "auto_replies_sent": auto_replies_sent,
                "conversations_by_status": conv_stats,
                "today_sent": today_sent,
                "today_limit": today_limit,
                "today_remaining": max(0, today_limit - today_sent),
            }
