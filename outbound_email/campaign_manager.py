"""
Campaign manager — orchestrates outbound sending, reply processing,
and AI-powered auto-responses.
"""

import logging
import time
import signal
import sys
from datetime import datetime

from . import config
from .database import Database
from .email_sender import EmailSender
from .email_receiver import EmailReceiver
from .ai_responder import AIResponder
from .templates import render_template, get_template, ALL_TEMPLATES

logger = logging.getLogger("outbound-email")


class CampaignManager:
    """Orchestrates the full outbound email lifecycle."""

    def __init__(self):
        self.db = Database()
        self.sender = EmailSender(self.db)
        self._receiver = None
        self._ai = None
        self._running = False

    @property
    def receiver(self):
        """Lazy-init EmailReceiver (needs IMAP credentials)."""
        if self._receiver is None:
            self._receiver = EmailReceiver(self.db)
        return self._receiver

    @property
    def ai(self):
        """Lazy-init AIResponder (needs ANTHROPIC_API_KEY)."""
        if self._ai is None:
            self._ai = AIResponder()
        return self._ai

    # ── Campaign Sending ──────────────────────────────────────────

    def send_campaign(self, campaign_id: int, limit: int = None,
                      force: bool = False) -> dict:
        """
        Send a campaign to unsent prospects.
        force=True re-sends to prospects already contacted (e.g. after a bad import).
        Returns summary of results.
        """
        campaign = self.db.get_campaign(campaign_id)
        if not campaign:
            return {"error": f"Campaign {campaign_id} not found"}

        if not campaign["is_active"]:
            return {"error": f"Campaign '{campaign['name']}' is not active"}

        remaining = self.db.get_remaining_sends_today()
        if remaining <= 0:
            return {"error": "Daily send limit reached", "sent": 0}

        batch_size = min(remaining, limit) if limit else remaining
        prospects = self.db.get_unsent_prospects(campaign_id, limit=batch_size,
                                                 force=force)

        if not prospects:
            return {"sent": 0, "error": "no_unsent",
                    "message": (
                        "No prospects found to send to.\n"
                        "  These prospects may have already been contacted for this campaign.\n"
                        "  To re-send, run:  python3 -m outbound_email send <id> --force"
                    )}

        logger.info(
            f"Sending campaign '{campaign['name']}' to {len(prospects)} prospects "
            f"({remaining} sends remaining today)"
        )

        results = self.sender.send_batch(prospects, campaign)
        logger.info(
            f"Campaign batch complete: sent={results['sent']}, "
            f"failed={results['failed']}, skipped={results['skipped']}"
        )
        return results

    def create_campaign_from_template(self, template_name: str,
                                      campaign_name: str = None) -> dict:
        """Create a new campaign from a built-in template."""
        template = get_template(template_name)
        if not template:
            return {"error": f"Template '{template_name}' not found"}

        name = campaign_name or template["name"]

        existing = self.db.get_campaign_by_name(name)
        if existing:
            return {"campaign_id": existing["id"], "message": "Campaign already exists"}

        campaign_id = self.db.create_campaign(
            name=name,
            subject_template=template["subject"],
            body_template=template["body"],
        )
        return {"campaign_id": campaign_id, "message": f"Campaign '{name}' created"}

    # ── Reply Processing ──────────────────────────────────────────

    def process_replies(self) -> dict:
        """
        Check for new replies, analyze them with AI, and send auto-responses.
        Returns processing summary.
        """
        processed = 0
        auto_replied = 0
        calendly_sent = 0
        opt_outs = 0
        errors = 0

        try:
            matched_replies = self.receiver.process_replies()
        except Exception as e:
            logger.error(f"Error fetching replies: {e}")
            return {"error": str(e)}

        for reply in matched_replies:
            try:
                result = self._handle_single_reply(reply)
                processed += 1
                if result.get("auto_replied"):
                    auto_replied += 1
                if result.get("calendly_sent"):
                    calendly_sent += 1
                if result.get("opted_out"):
                    opt_outs += 1
            except Exception as e:
                logger.error(f"Error processing reply from {reply['prospect']['email']}: {e}")
                errors += 1
            finally:
                # Mark as processed regardless
                self.db.mark_reply_processed(reply["reply_id"])

        return {
            "processed": processed,
            "auto_replied": auto_replied,
            "calendly_sent": calendly_sent,
            "opt_outs": opt_outs,
            "errors": errors,
        }

    def _handle_single_reply(self, reply: dict) -> dict:
        """Process a single reply: analyze with AI and send response."""
        prospect = reply["prospect"]
        prospect_id = prospect["id"]

        # Get conversation history
        thread = self.db.get_full_thread(prospect_id)

        # Analyze with AI
        ai_result = self.ai.analyze_and_respond(
            prospect=prospect,
            reply_body=reply["body"],
            reply_subject=reply["subject"],
            conversation_history=thread,
        )

        classification = ai_result["classification"]
        result = {"auto_replied": False, "calendly_sent": False, "opted_out": False}

        # Update conversation status based on classification
        status_map = {
            "INTERESTED": "engaged",
            "SCHEDULING": "calendly_sent",
            "OBJECTION": "engaged",
            "QUESTION": "engaged",
            "NOT_INTERESTED": "closed",
            "ALREADY_SOLD": "closed",
            "OPT_OUT": "opted_out",
        }

        new_status = status_map.get(classification, "engaged")

        # If AI says send Calendly, upgrade status
        if ai_result["should_send_calendly"]:
            new_status = "calendly_sent"
            result["calendly_sent"] = True

        # Update prospect and conversation
        prospect_status_map = {
            "engaged": "engaged",
            "calendly_sent": "calendly_sent",
            "closed": "closed" if prospect["status"] != "meeting_booked" else "meeting_booked",
            "opted_out": "opted_out",
        }
        self.db.update_prospect_status(
            prospect_id,
            prospect_status_map.get(new_status, "engaged")
        )
        self.db.update_conversation(
            prospect_id,
            status=new_status,
            ai_summary=ai_result["summary"],
        )

        if classification == "OPT_OUT":
            result["opted_out"] = True

        # Send auto-reply if AI recommends it
        if ai_result["should_reply"] and ai_result["reply_body"]:
            # Build reply HTML
            reply_html = self._plain_to_html(ai_result["reply_body"])

            # Get the original message ID for threading
            in_reply_to = reply.get("message_id", "")
            references = reply.get("in_reply_to", "")
            if references and in_reply_to:
                references = f"{references} {in_reply_to}"
            elif in_reply_to:
                references = in_reply_to

            send_result = self.sender.send_email(
                to_email=prospect["email"],
                subject=ai_result["reply_subject"],
                body_html=reply_html,
                body_text=ai_result["reply_body"],
                prospect_id=prospect_id,
                is_auto_reply=True,
                in_reply_to=in_reply_to,
                references=references,
            )

            if send_result["success"]:
                result["auto_replied"] = True
                logger.info(
                    f"Auto-replied to {prospect['email']} "
                    f"(classification={classification}, calendly={result['calendly_sent']})"
                )
            else:
                logger.error(
                    f"Failed to auto-reply to {prospect['email']}: "
                    f"{send_result.get('error')}"
                )

        return result

    def _plain_to_html(self, text: str) -> str:
        """Convert plain text email to basic HTML."""
        import html
        escaped = html.escape(text)
        paragraphs = escaped.split("\n\n")
        html_parts = []
        for p in paragraphs:
            lines = p.replace("\n", "<br>\n")
            html_parts.append(f"<p>{lines}</p>")

        return (
            '<div style="font-family: Arial, sans-serif; font-size: 14px; '
            'color: #333; line-height: 1.6;">\n'
            + "\n".join(html_parts)
            + "\n</div>"
        )

    # ── Daemon Mode ───────────────────────────────────────────────

    def run_daemon(self):
        """
        Run continuously: periodically check for replies and process them.
        Runs until interrupted (Ctrl+C).
        """
        self._running = True

        def handle_signal(signum, frame):
            logger.info("Shutdown signal received. Stopping...")
            self._running = False

        signal.signal(signal.SIGINT, handle_signal)
        signal.signal(signal.SIGTERM, handle_signal)

        logger.info("=" * 60)
        logger.info("DEED STREET CAPITAL — EMAIL RESPONSE DAEMON")
        logger.info("=" * 60)
        logger.info(f"Polling interval: {config.IMAP_POLL_INTERVAL}s")
        logger.info(f"IMAP: {config.IMAP_HOST} ({config.IMAP_USERNAME})")
        logger.info("Listening for replies... (Ctrl+C to stop)")

        while self._running:
            try:
                result = self.process_replies()
                if result.get("processed", 0) > 0:
                    logger.info(
                        f"Processed {result['processed']} replies: "
                        f"{result.get('auto_replied', 0)} auto-replied, "
                        f"{result.get('calendly_sent', 0)} Calendly sent, "
                        f"{result.get('opt_outs', 0)} opt-outs"
                    )
            except Exception as e:
                logger.error(f"Daemon loop error: {e}")

            # Sleep in small intervals so we can respond to shutdown quickly
            for _ in range(config.IMAP_POLL_INTERVAL):
                if not self._running:
                    break
                time.sleep(1)

        if self._receiver:
            self._receiver.disconnect()
        logger.info("Daemon stopped.")

    # ── Reporting ─────────────────────────────────────────────────

    def get_report(self) -> str:
        """Generate a human-readable status report."""
        stats = self.db.get_stats()
        lines = [
            "=" * 50,
            "  DEED STREET CAPITAL — OUTBOUND EMAIL REPORT",
            "=" * 50,
            "",
            f"  Total Prospects:         {stats['total_prospects']}",
            f"  Outreach Emails Sent:    {stats['total_outreach_sent']}",
            f"  Replies Received:        {stats['total_replies_received']}",
            f"  Auto-Replies Sent:       {stats['auto_replies_sent']}",
            "",
            f"  Today's Sends:           {stats['today_sent']} / {stats['today_limit']}",
            f"  Remaining Today:         {stats['today_remaining']}",
            "",
            "  Prospects by Status:",
        ]

        for status, count in sorted(stats.get("prospects_by_status", {}).items()):
            lines.append(f"    {status:20s} {count}")

        conv = stats.get("conversations_by_status", {})
        if conv:
            lines.append("")
            lines.append("  Conversations by Status:")
            for status, count in sorted(conv.items()):
                lines.append(f"    {status:20s} {count}")

        reply_rate = 0
        if stats["total_outreach_sent"] > 0:
            reply_rate = (stats["total_replies_received"] / stats["total_outreach_sent"]) * 100

        lines.extend([
            "",
            f"  Reply Rate:              {reply_rate:.1f}%",
            "",
            "=" * 50,
        ])
        return "\n".join(lines)
