"""
Claude AI-powered email response engine.
Analyzes incoming replies, classifies intent, and generates contextual responses.
Sends Calendly link when prospects are engaged.
"""

import json
import logging
import anthropic

from . import config

logger = logging.getLogger("outbound-email")

SYSTEM_PROMPT = f"""You are the AI email assistant for Deed Street Capital, a national leader in purchasing privately-held mortgage notes.

## About Deed Street Capital
{config.COMPANY_DESCRIPTION}

## Key Value Propositions
{chr(10).join(f"- {v}" for v in config.COMPANY_VALUE_PROPS)}

## Your Role
You handle email replies from prospects who received our outreach about selling their mortgage notes. Your job is to:

1. **Analyze** the prospect's reply to understand their intent and sentiment
2. **Classify** their response into one of these categories:
   - INTERESTED: They want to learn more, asked questions, or expressed willingness to talk
   - OBJECTION: They have concerns (price, timing, trust, etc.) that can be addressed
   - NOT_INTERESTED: Polite decline but door might be open later
   - OPT_OUT: They EXPLICITLY asked to stop receiving emails, unsubscribe, or said "remove me." This is ONLY for clear, unambiguous opt-out requests.
   - QUESTION: They asked a specific question about the process, our company, or how we got their information
   - ALREADY_SOLD: They already sold their note or no longer hold one
   - SCHEDULING: They're ready to schedule a call or meeting
3. **Generate** a professional, warm, and concise reply email

## CRITICAL Classification Rules
- "How did you get my information?" or "Where did you find me?" or "How did you get my email?" is a QUESTION, NOT an OPT_OUT. Respond by briefly explaining that their information is part of publicly available mortgage note records, then pivot back to the value proposition. Ask if they are in first position on a mortgage note and would be interested in receiving a quote for a lump sum of cash.
- Only classify as OPT_OUT if the prospect explicitly says "stop emailing me", "unsubscribe", "remove me from your list", "do not contact me again", or similar unambiguous language.
- When in doubt between OPT_OUT and QUESTION, always classify as QUESTION.

## Response Guidelines
- Keep emails SHORT but well-structured. Nobody reads long emails.
- Be conversational and human. Never sound like a template or robot.
- Use the prospect's first name if available.
- Never be pushy or aggressive. Deed Street Capital is a trusted, professional firm.
- If they have questions, answer them clearly and specifically.
- For objections, acknowledge their concern and provide a brief, honest counter.
- Match the prospect's tone -- if they're casual, be casual. If formal, be formal.
- Always end with a clear next step (schedule a call, answer another question, etc.).
- NEVER fabricate specific numbers, rates, or terms -- speak in general terms about our competitive pricing.
- If the prospect seems ready to talk, include the Calendly link naturally.
- For INTERESTED or engaged replies, break your response into 2-3 short paragraphs separated by blank lines. Do NOT write one long block of text. Each paragraph should cover a distinct point. This makes the email look professional and easy to read.
- For OPT_OUT: confirm their removal gracefully. Say something like "Thank you for your understanding. If you ever are interested in receiving a quote for your mortgage note, please don't hesitate to reach out." Do NOT say "sorry for the intrusion" or anything that sounds apologetic or intrusive.

## Important Rules
- Do NOT use emoji in emails.
- Do NOT use em dashes (--) in emails. Use commas, periods, or restructure the sentence instead.
- Do NOT use exclamation marks excessively (one max per email).
- Do NOT use marketing jargon or corporate buzzwords.
- Do NOT sound overly AI-generated or robotic. Avoid stiff, formulaic phrasing like "I'd be happy to assist you with that" or "I completely understand your concern." Write like a real person having a normal conversation.
- Do NOT make promises about specific pricing without speaking to the prospect first.
- Keep subject lines natural -- reply to their subject, don't create a new one.
- Sign off as the sender's name, not "Deed Street Capital Team" or similar.
"""


class AIResponder:
    """Uses Claude to analyze replies and generate email responses."""

    def __init__(self):
        if not config.ANTHROPIC_API_KEY:
            raise ValueError("ANTHROPIC_API_KEY is required for AI responder")
        self.client = anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)

    def analyze_and_respond(self, prospect: dict, reply_body: str,
                            reply_subject: str, conversation_history: list,
                            sender_name: str = None) -> dict:
        """
        Analyze a prospect's reply and generate an appropriate response.

        Returns:
            {
                "classification": str,  # INTERESTED, OBJECTION, OPT_OUT, etc.
                "should_send_calendly": bool,
                "should_reply": bool,
                "reply_subject": str,
                "reply_body": str,
                "summary": str,  # Brief AI summary of the interaction
            }
        """
        sender_name = sender_name or config.SENDER_NAME

        # Build conversation context
        thread_context = self._format_thread(conversation_history)

        prospect_context = self._format_prospect(prospect)

        user_prompt = f"""## Prospect Information
{prospect_context}

## Conversation History
{thread_context}

## Latest Reply From Prospect
Subject: {reply_subject}
Body:
{reply_body}

## Instructions
Analyze this reply and generate an appropriate response email. Your response must be valid JSON with these fields:

{{
    "classification": "INTERESTED | OBJECTION | NOT_INTERESTED | OPT_OUT | QUESTION | ALREADY_SOLD | SCHEDULING",
    "should_send_calendly": true/false,
    "should_reply": true/false,
    "reply_subject": "Re: ...",
    "reply_body": "The email body text (plain text, not HTML)",
    "summary": "One-sentence summary of where this conversation stands"
}}

Rules for your decision:
- Set should_send_calendly=true ONLY when the prospect is clearly INTERESTED or in SCHEDULING mode
- Set should_reply=false ONLY for OPT_OUT (we still send a gracious goodbye) or if the message is spam/irrelevant
- For OPT_OUT: set should_reply=true but should_send_calendly=false, and write a brief gracious removal confirmation
- The reply_body should be the complete email body ready to send, signed off with "{sender_name}"
- If should_send_calendly is true, naturally include this scheduling link in your reply: {config.CALENDLY_LINK}
- The reply_subject should be "Re: " followed by the original subject (without stacking Re: Re: Re:)

Return ONLY the JSON object. No other text."""

        try:
            response = self.client.messages.create(
                model=config.CLAUDE_MODEL,
                max_tokens=1024,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": user_prompt}],
            )

            response_text = response.content[0].text.strip()

            # Parse JSON from response (handle potential markdown code blocks)
            if response_text.startswith("```"):
                response_text = response_text.split("```")[1]
                if response_text.startswith("json"):
                    response_text = response_text[4:]
                response_text = response_text.strip()

            result = json.loads(response_text)

            # Validate required fields
            required = ["classification", "should_send_calendly", "should_reply",
                        "reply_subject", "reply_body", "summary"]
            for field in required:
                if field not in result:
                    raise ValueError(f"Missing field: {field}")

            logger.info(
                f"AI classified reply from {prospect['email']} as "
                f"{result['classification']} (calendly={result['should_send_calendly']})"
            )

            return result

        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse AI response as JSON: {e}")
            logger.error(f"Raw response: {response_text[:500]}")
            return self._fallback_response(prospect, reply_subject, sender_name)
        except Exception as e:
            logger.error(f"AI responder error: {e}")
            return self._fallback_response(prospect, reply_subject, sender_name)

    def _format_thread(self, conversation_history: list) -> str:
        """Format conversation history for the AI prompt."""
        if not conversation_history:
            return "(No prior conversation history)"

        parts = []
        for entry in conversation_history:
            direction = "WE SENT" if entry["role"] == "sent" else "PROSPECT REPLIED"
            parts.append(
                f"[{entry['timestamp']}] {direction}:\n"
                f"Subject: {entry['subject']}\n"
                f"{entry['body']}\n"
            )
        return "\n---\n".join(parts)

    def _format_prospect(self, prospect: dict) -> str:
        """Format prospect data for the AI prompt."""
        parts = [f"Email: {prospect['email']}"]
        if prospect.get("first_name"):
            parts.append(f"Name: {prospect['first_name']} {prospect.get('last_name', '')}")
        if prospect.get("company"):
            parts.append(f"Company: {prospect['company']}")
        if prospect.get("note_type"):
            parts.append(f"Note Type: {prospect['note_type']}")
        if prospect.get("note_balance"):
            parts.append(f"Note Balance: {prospect['note_balance']}")
        if prospect.get("property_state"):
            parts.append(f"Property State: {prospect['property_state']}")
        if prospect.get("status"):
            parts.append(f"Current Status: {prospect['status']}")
        return "\n".join(parts)

    def _fallback_response(self, prospect: dict, reply_subject: str,
                           sender_name: str) -> dict:
        """Fallback response if AI fails — flags for manual review."""
        name = prospect.get("first_name", "")
        clean_subject = reply_subject
        if not clean_subject.lower().startswith("re:"):
            clean_subject = f"Re: {clean_subject}"

        return {
            "classification": "QUESTION",
            "should_send_calendly": False,
            "should_reply": False,
            "reply_subject": clean_subject,
            "reply_body": "",
            "summary": "AI processing failed — flagged for manual review",
        }
