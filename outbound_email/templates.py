"""
Email templates for Deed Street Capital outbound campaigns.
Each template is designed for a specific outreach scenario.
Templates use {{variable}} syntax for personalization.
"""

from typing import Optional

# ── Initial Outreach Templates ────────────────────────────────────

INITIAL_OUTREACH_V1 = {
    "name": "initial_outreach_v1",
    "subject": "Quick question about your note, {{first_name}}",
    "body": """\
<div style="font-family: Arial, sans-serif; font-size: 14px; color: #333; line-height: 1.6;">
<p>Hi {{first_name}},</p>

<p>I came across your information and wanted to reach out. I'm with Deed Street Capital — we're a direct buyer of privately-held mortgage notes nationwide.</p>

<p>If you currently hold a seller-financed note and have ever considered selling it (or even just a portion of it), I'd love to have a quick conversation. We offer competitive pricing with no hidden fees, and most of our transactions close within 30 days.</p>

<p>Would you be open to a brief call this week?</p>

<p>Best,<br>
Deed Street Capital<br>
{{sender_email}}<br>
Deedstreetcapital.com</p>
</div>
""",
}

INITIAL_OUTREACH_V2 = {
    "name": "initial_outreach_v2",
    "subject": "Selling your mortgage note?",
    "body": """\
<div style="font-family: Arial, sans-serif; font-size: 14px; color: #333; line-height: 1.6;">
<p>Hi {{first_name}},</p>

<p>I'll keep this short. Deed Street Capital purchases private mortgage notes — residential and commercial — in all 50 states. We're a direct buyer, which means no middlemen and better pricing for you.</p>

<p>If you hold a seller-financed note and are thinking about converting it to cash, we can usually provide a quote within 24 hours and close within 30 days.</p>

<p>Interested in getting a no-obligation quote? Just reply to this email and I'll walk you through how it works.</p>

<p>Deed Street Capital<br>
{{sender_email}}<br>
Deedstreetcapital.com</p>
</div>
""",
}

INITIAL_OUTREACH_V3 = {
    "name": "initial_outreach_v3",
    "subject": "{{first_name}} — have you considered selling your note?",
    "body": """\
<div style="font-family: Arial, sans-serif; font-size: 14px; color: #333; line-height: 1.6;">
<p>{{first_name}},</p>

<p>Many note holders don't realize they can sell their mortgage note for a lump sum of cash — or even sell just a portion of their payments while keeping the rest.</p>

<p>At Deed Street Capital, we specialize in purchasing privately-held mortgage notes. We've been doing this for over 20 years and buy notes in every state. No hidden fees, no middlemen, and most deals close in under 30 days.</p>

<p>If you've ever thought about cashing out your note, I'd be happy to provide a free, no-obligation quote. Just reply here and we can get started.</p>

<p>Deed Street Capital<br>
{{sender_email}}<br>
Deedstreetcapital.com</p>
</div>
""",
}

# ── Follow-Up Templates ───────────────────────────────────────────

FOLLOW_UP_V1 = {
    "name": "follow_up_v1",
    "subject": "Re: Quick question about your note, {{first_name}}",
    "body": """\
<div style="font-family: Arial, sans-serif; font-size: 14px; color: #333; line-height: 1.6;">
<p>Hi {{first_name}},</p>

<p>I sent a note last week about potentially purchasing your mortgage note. I know inboxes get busy, so I wanted to follow up.</p>

<p>If selling your note is something you'd consider — even partially — I'm happy to provide a quick, no-obligation quote. The process is straightforward and we handle all the paperwork.</p>

<p>If this isn't relevant to you, no worries at all. Just let me know and I won't follow up again.</p>

<p>Deed Street Capital<br>
{{sender_email}}<br>
Deedstreetcapital.com</p>
</div>
""",
}

FOLLOW_UP_V2 = {
    "name": "follow_up_v2",
    "subject": "One last thing, {{first_name}}",
    "body": """\
<div style="font-family: Arial, sans-serif; font-size: 14px; color: #333; line-height: 1.6;">
<p>Hi {{first_name}},</p>

<p>I'll keep this brief — I've reached out a couple times about purchasing your mortgage note. I understand if the timing isn't right.</p>

<p>If you ever decide to explore selling your note in the future, feel free to reach out. We buy notes of all sizes ($25K to $3M) in every state, and we're always here.</p>

<p>All the best,<br>
Deed Street Capital<br>
{{sender_email}}<br>
Deedstreetcapital.com</p>
</div>
""",
}

# ── Template Registry ─────────────────────────────────────────────

ALL_TEMPLATES = {
    "initial_outreach_v1": INITIAL_OUTREACH_V1,
    "initial_outreach_v2": INITIAL_OUTREACH_V2,
    "initial_outreach_v3": INITIAL_OUTREACH_V3,
    "follow_up_v1": FOLLOW_UP_V1,
    "follow_up_v2": FOLLOW_UP_V2,
}


def get_template(name: str) -> Optional[dict]:
    """Get a template by name."""
    return ALL_TEMPLATES.get(name)


def list_templates() -> list:
    """List all available template names."""
    return list(ALL_TEMPLATES.keys())


def render_template(template: dict, prospect: dict,
                    extra_vars: dict = None) -> tuple:
    """
    Render a template with prospect data.
    Returns (subject, body_html).
    """
    from . import config

    variables = {
        "{{first_name}}": prospect.get("first_name") or "there",
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

    if extra_vars:
        variables.update(extra_vars)

    subject = template["subject"]
    body = template["body"]

    for key, value in variables.items():
        val = str(value) if value else ""
        subject = subject.replace(key, val)
        body = body.replace(key, val)

    # Clean up any empty first names that slipped through
    if not prospect.get("first_name"):
        subject = subject.replace(", there", "").replace("there — ", "")
        body = body.replace(">there,<", "><").replace(">there<", "><")

    return subject, body
