"""
Email sender with SMTP and SendGrid support.
Handles rate limiting, warmup, and delivery tracking.
"""

import smtplib
import logging
import time
import uuid
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr, formatdate, make_msgid

from . import config

logger = logging.getLogger("outbound-email")


class EmailSender:
    """Sends emails via SMTP or SendGrid with rate limiting."""

    def __init__(self, db):
        self.db = db
        self.provider = config.EMAIL_PROVIDER

    def send_email(self, to_email: str, subject: str, body_html: str,
                   body_text: str = None, prospect_id: int = None,
                   campaign_id: int = None, is_auto_reply: bool = False,
                   in_reply_to: str = None, references: str = None) -> dict:
        """
        Send a single email. Returns dict with success status and message_id.
        Respects daily rate limits.
        """
        # Check rate limit (skip for auto-replies to avoid missing conversations)
        if not is_auto_reply:
            remaining = self.db.get_remaining_sends_today()
            if remaining <= 0:
                logger.warning("Daily send limit reached. Skipping.")
                return {"success": False, "error": "Daily send limit reached"}

        if self.provider == "sendgrid":
            result = self._send_via_sendgrid(
                to_email, subject, body_html, body_text,
                in_reply_to=in_reply_to, references=references,
            )
        else:
            result = self._send_via_smtp(
                to_email, subject, body_html, body_text,
                in_reply_to=in_reply_to, references=references,
            )

        if result["success"]:
            # Log in database
            self.db.log_sent_email(
                prospect_id=prospect_id,
                subject=subject,
                body=body_html or body_text,
                message_id=result.get("message_id"),
                campaign_id=campaign_id,
                is_auto_reply=is_auto_reply,
            )
            if not is_auto_reply:
                self.db.increment_send_count()
            logger.info(f"Sent email to {to_email} (msg_id={result.get('message_id', 'N/A')})")
        else:
            logger.error(f"Failed to send email to {to_email}: {result.get('error')}")

        return result

    def _send_via_smtp(self, to_email: str, subject: str, body_html: str,
                       body_text: str = None, in_reply_to: str = None,
                       references: str = None) -> dict:
        """Send email via SMTP."""
        msg = MIMEMultipart("alternative")
        message_id = make_msgid(domain=config.SENDER_EMAIL.split("@")[-1] if config.SENDER_EMAIL else "deedstreetcapital.com")

        msg["From"] = formataddr((config.SENDER_NAME, config.SENDER_EMAIL))
        msg["To"] = to_email
        msg["Subject"] = subject
        msg["Date"] = formatdate(localtime=True)
        msg["Message-ID"] = message_id

        # Thread tracking headers for replies
        if in_reply_to:
            msg["In-Reply-To"] = in_reply_to
        if references:
            msg["References"] = references

        # Add plain text version
        if body_text:
            msg.attach(MIMEText(body_text, "plain", "utf-8"))
        else:
            # Strip HTML tags for a basic plain text version
            import re
            plain = re.sub(r"<[^>]+>", "", body_html)
            plain = re.sub(r"\n\s*\n", "\n\n", plain).strip()
            msg.attach(MIMEText(plain, "plain", "utf-8"))

        # Add HTML version
        msg.attach(MIMEText(body_html, "html", "utf-8"))

        try:
            if config.SMTP_USE_TLS:
                server = smtplib.SMTP(config.SMTP_HOST, config.SMTP_PORT)
                server.ehlo()
                server.starttls()
                server.ehlo()
            else:
                server = smtplib.SMTP(config.SMTP_HOST, config.SMTP_PORT)

            server.login(config.SMTP_USERNAME, config.SMTP_PASSWORD)
            server.sendmail(config.SENDER_EMAIL, to_email, msg.as_string())
            server.quit()

            return {"success": True, "message_id": message_id}

        except smtplib.SMTPRecipientsRefused:
            return {"success": False, "error": "Recipient refused (invalid address)"}
        except smtplib.SMTPAuthenticationError:
            return {"success": False, "error": "SMTP authentication failed"}
        except smtplib.SMTPException as e:
            return {"success": False, "error": f"SMTP error: {e}"}
        except Exception as e:
            return {"success": False, "error": f"Send failed: {e}"}

    def _send_via_sendgrid(self, to_email: str, subject: str, body_html: str,
                           body_text: str = None, in_reply_to: str = None,
                           references: str = None) -> dict:
        """Send email via SendGrid API."""
        try:
            import sendgrid
            from sendgrid.helpers.mail import (
                Mail, Email, To, Content, Header,
            )
        except ImportError:
            return {"success": False, "error": "sendgrid package not installed. Run: pip install sendgrid"}

        sg = sendgrid.SendGridAPIClient(api_key=config.SENDGRID_API_KEY)
        message_id = make_msgid(domain=config.SENDER_EMAIL.split("@")[-1] if config.SENDER_EMAIL else "deedstreetcapital.com")

        mail = Mail(
            from_email=Email(config.SENDER_EMAIL, config.SENDER_NAME),
            to_emails=To(to_email),
            subject=subject,
        )

        if body_text:
            mail.add_content(Content("text/plain", body_text))
        mail.add_content(Content("text/html", body_html))

        # Add threading headers
        mail.add_header(Header("Message-ID", message_id))
        if in_reply_to:
            mail.add_header(Header("In-Reply-To", in_reply_to))
        if references:
            mail.add_header(Header("References", references))

        try:
            response = sg.send(mail)
            if response.status_code in (200, 201, 202):
                return {"success": True, "message_id": message_id}
            else:
                return {"success": False, "error": f"SendGrid status {response.status_code}"}
        except Exception as e:
            return {"success": False, "error": f"SendGrid error: {e}"}

    def send_batch(self, prospects: list, campaign: dict,
                   render_fn=None) -> dict:
        """
        Send emails to a batch of prospects for a campaign.
        render_fn(prospect, campaign) -> (subject, body_html)
        Returns summary counts.
        """
        sent = 0
        failed = 0
        skipped = 0

        for prospect in prospects:
            remaining = self.db.get_remaining_sends_today()
            if remaining <= 0:
                skipped += len(prospects) - sent - failed
                logger.info(f"Daily limit reached. Sent {sent}, skipping rest.")
                break

            try:
                if render_fn:
                    subject, body_html = render_fn(prospect, campaign)
                else:
                    subject = self._render_template(campaign["subject_template"], prospect)
                    body_html = self._render_template(campaign["body_template"], prospect)

                result = self.send_email(
                    to_email=prospect["email"],
                    subject=subject,
                    body_html=body_html,
                    prospect_id=prospect["id"],
                    campaign_id=campaign["id"],
                )

                if result["success"]:
                    self.db.update_prospect_status(prospect["id"], "contacted")
                    # Create conversation record
                    self.db.get_or_create_conversation(prospect["id"])
                    sent += 1
                else:
                    if "refused" in result.get("error", "").lower():
                        self.db.update_prospect_status(prospect["id"], "bounced")
                    failed += 1

            except Exception as e:
                logger.error(f"Error sending to {prospect['email']}: {e}")
                failed += 1

            # Delay between sends to avoid spam triggers
            if config.SEND_DELAY_SECONDS > 0:
                time.sleep(config.SEND_DELAY_SECONDS)

        return {"sent": sent, "failed": failed, "skipped": skipped}

    def _render_template(self, template: str, prospect: dict) -> str:
        """Render a template string with prospect data."""
        replacements = {
            "{{first_name}}": prospect.get("first_name", ""),
            "{{last_name}}": prospect.get("last_name", ""),
            "{{company}}": prospect.get("company", ""),
            "{{email}}": prospect.get("email", ""),
            "{{note_type}}": prospect.get("note_type", ""),
            "{{note_balance}}": prospect.get("note_balance", ""),
            "{{property_state}}": prospect.get("property_state", ""),
            "{{sender_name}}": config.SENDER_NAME,
            "{{sender_email}}": config.SENDER_EMAIL,
            "{{calendly_link}}": config.CALENDLY_LINK,
        }
        result = template
        for key, value in replacements.items():
            result = result.replace(key, str(value) if value else "")
        return result
