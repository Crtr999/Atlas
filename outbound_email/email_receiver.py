"""
IMAP email receiver for processing incoming replies.
Polls the inbox, matches replies to prospects, and queues them for AI processing.
"""

import imaplib
import email
import logging
import re
from email.header import decode_header
from email.utils import parseaddr
from datetime import datetime
from typing import Optional

from . import config

logger = logging.getLogger("outbound-email")


class EmailReceiver:
    """Connects to IMAP server and fetches new replies."""

    def __init__(self, db):
        self.db = db
        self._connection = None

    def connect(self):
        """Establish IMAP connection."""
        try:
            if config.IMAP_USE_SSL:
                self._connection = imaplib.IMAP4_SSL(
                    config.IMAP_HOST, config.IMAP_PORT
                )
            else:
                self._connection = imaplib.IMAP4(
                    config.IMAP_HOST, config.IMAP_PORT
                )
            self._connection.login(config.IMAP_USERNAME, config.IMAP_PASSWORD)
            logger.info(f"Connected to IMAP server {config.IMAP_HOST}")
        except imaplib.IMAP4.error as e:
            logger.error(f"IMAP connection failed: {e}")
            raise

    def disconnect(self):
        """Close IMAP connection."""
        if self._connection:
            try:
                self._connection.close()
                self._connection.logout()
            except Exception:
                pass
            self._connection = None

    def fetch_new_replies(self) -> list:
        """
        Fetch unread emails from the inbox that are replies to our outreach.
        Returns list of parsed email dicts.
        """
        if not self._connection:
            self.connect()

        replies = []
        try:
            self._connection.select(config.IMAP_FOLDER)

            # Search for unseen messages
            status, message_ids = self._connection.search(None, "UNSEEN")
            if status != "OK" or not message_ids[0]:
                return replies

            ids = message_ids[0].split()
            logger.info(f"Found {len(ids)} unread emails")

            for msg_id in ids:
                try:
                    parsed = self._fetch_and_parse(msg_id)
                    if parsed:
                        replies.append(parsed)
                except Exception as e:
                    logger.error(f"Error parsing message {msg_id}: {e}")
                    continue

        except imaplib.IMAP4.error as e:
            logger.error(f"IMAP fetch error: {e}")
            # Reconnect on failure
            self.disconnect()
            self.connect()

        return replies

    def _fetch_and_parse(self, msg_id: bytes) -> Optional[dict]:
        """Fetch a single email by ID and parse it."""
        status, data = self._connection.fetch(msg_id, "(RFC822)")
        if status != "OK":
            return None

        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        # Parse headers
        from_name, from_email_addr = parseaddr(msg.get("From", ""))
        subject = self._decode_header(msg.get("Subject", ""))
        message_id = msg.get("Message-ID", "")
        in_reply_to = msg.get("In-Reply-To", "")
        references = msg.get("References", "")
        date_str = msg.get("Date", "")

        # Extract body
        body_text = ""
        body_html = ""

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition", ""))

                # Skip attachments
                if "attachment" in content_disposition:
                    continue

                try:
                    payload = part.get_payload(decode=True)
                    if payload is None:
                        continue
                    charset = part.get_content_charset() or "utf-8"
                    text = payload.decode(charset, errors="replace")

                    if content_type == "text/plain":
                        body_text = text
                    elif content_type == "text/html":
                        body_html = text
                except Exception:
                    continue
        else:
            try:
                payload = msg.get_payload(decode=True)
                if payload:
                    charset = msg.get_content_charset() or "utf-8"
                    text = payload.decode(charset, errors="replace")
                    if msg.get_content_type() == "text/html":
                        body_html = text
                    else:
                        body_text = text
            except Exception:
                pass

        # Prefer plain text, fall back to stripped HTML
        body = body_text or self._strip_html(body_html)

        # Clean up the reply (remove quoted text from previous messages)
        body = self._extract_reply_body(body)

        if not body.strip():
            return None

        return {
            "from_email": from_email_addr.lower().strip(),
            "from_name": from_name,
            "subject": subject,
            "body": body.strip(),
            "message_id": message_id.strip(),
            "in_reply_to": in_reply_to.strip(),
            "references": references.strip(),
            "date": date_str,
        }

    def process_replies(self) -> list:
        """
        Fetch new replies and match them to prospects in the database.
        Returns list of matched replies ready for AI processing.
        """
        raw_replies = self.fetch_new_replies()
        matched = []

        for reply in raw_replies:
            from_email = reply["from_email"]

            # Look up the prospect
            prospect = self.db.get_prospect_by_email(from_email)
            if not prospect:
                logger.info(f"Received email from unknown sender: {from_email} — skipping")
                continue

            # Check if this is a reply to one of our sent emails
            is_reply = False
            if reply["in_reply_to"]:
                sent = self.db.get_sent_email_by_message_id(reply["in_reply_to"])
                if sent:
                    is_reply = True

            # Even without In-Reply-To header, if the sender is a known prospect,
            # treat it as a reply
            if not is_reply and prospect["status"] in ("contacted", "engaged", "calendly_sent"):
                is_reply = True

            if not is_reply:
                logger.info(f"Email from {from_email} doesn't match any outreach — skipping")
                continue

            # Log in database
            reply_id = self.db.log_received_email(
                from_email=from_email,
                subject=reply["subject"],
                body=reply["body"],
                raw_message_id=reply["message_id"],
                in_reply_to=reply["in_reply_to"],
                prospect_id=prospect["id"],
            )

            matched.append({
                "reply_id": reply_id,
                "prospect": prospect,
                "subject": reply["subject"],
                "body": reply["body"],
                "message_id": reply["message_id"],
                "in_reply_to": reply["in_reply_to"],
            })

            logger.info(f"Matched reply from {from_email} (prospect #{prospect['id']})")

        return matched

    def _decode_header(self, header_value: str) -> str:
        """Decode an email header value."""
        if not header_value:
            return ""
        parts = decode_header(header_value)
        decoded = []
        for part, charset in parts:
            if isinstance(part, bytes):
                decoded.append(part.decode(charset or "utf-8", errors="replace"))
            else:
                decoded.append(part)
        return " ".join(decoded)

    def _strip_html(self, html: str) -> str:
        """Convert HTML to plain text."""
        if not html:
            return ""
        # Remove HTML tags
        text = re.sub(r"<br\s*/?>", "\n", html, flags=re.IGNORECASE)
        text = re.sub(r"<p[^>]*>", "\n", text, flags=re.IGNORECASE)
        text = re.sub(r"</p>", "\n", text, flags=re.IGNORECASE)
        text = re.sub(r"<[^>]+>", "", text)
        # Decode HTML entities
        import html as html_module
        text = html_module.unescape(text)
        return text.strip()

    def _extract_reply_body(self, body: str) -> str:
        """
        Extract only the new reply content, stripping quoted previous messages.
        Common patterns: "On ... wrote:", "> " prefixed lines, "-----Original Message-----"
        """
        lines = body.split("\n")
        reply_lines = []

        for line in lines:
            stripped = line.strip()
            # Stop at common reply delimiters
            if re.match(r"^On .+ wrote:$", stripped):
                break
            if stripped.startswith("-----Original Message-----"):
                break
            if stripped.startswith("From:") and "sent:" in body.lower():
                break
            if re.match(r"^>+\s", stripped):
                # Skip quoted lines but continue looking for new content
                continue
            if stripped == "--":
                # Email signature delimiter
                break
            reply_lines.append(line)

        return "\n".join(reply_lines).strip()
