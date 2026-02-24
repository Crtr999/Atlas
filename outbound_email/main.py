#!/usr/bin/env python3
"""
Deed Street Capital — Outbound Email System CLI

Usage:
    python -m outbound_email <command> [options]

Commands:
    setup           Validate configuration and initialize database
    import-csv      Import prospects from a CSV file
    import-xlsx     Import prospects from an Excel (.xlsx) file
    export-report   Export analytics report to Excel (.xlsx)
    add-prospect    Add a single prospect
    list-prospects  List all prospects (optionally filtered by status)
    campaigns       List available campaigns/templates
    create-campaign Create a campaign from a template
    send            Send a campaign to unsent prospects
    check-replies   Check for and process new replies (one-time)
    daemon          Run reply-processing daemon (continuous)
    report          Show campaign statistics
    thread          View full email thread for a prospect
"""

import argparse
import logging
import sys
import os

# Ensure the parent directory is on the path so we can run as a module
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from outbound_email import config
from outbound_email.database import Database
from outbound_email.campaign_manager import CampaignManager
from outbound_email.templates import list_templates, ALL_TEMPLATES

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("outbound-email")


def cmd_setup(args):
    """Validate config and initialize the database."""
    print("Checking configuration...")
    missing = config.validate_config()
    if missing:
        print(f"\nMissing required config: {', '.join(missing)}")
        print("Set these in your .env file. See .env.example for reference.")
        return 1

    print("Configuration OK.")
    print(f"  Email provider:  {config.EMAIL_PROVIDER}")
    print(f"  SMTP host:       {config.SMTP_HOST}")
    print(f"  IMAP host:       {config.IMAP_HOST}")
    print(f"  Sender:          {config.SENDER_NAME} <{config.SENDER_EMAIL}>")
    print(f"  Daily limit:     {config.DAILY_SEND_LIMIT}")
    print(f"  Warmup enabled:  {config.WARMUP_ENABLED}")
    print(f"  Calendly:        {config.CALENDLY_LINK}")

    db = Database()
    print(f"\nDatabase initialized at: {config.DB_PATH}")
    stats = db.get_stats()
    print(f"  Prospects: {stats['total_prospects']}")
    print(f"  Emails sent: {stats['total_outreach_sent']}")
    print(f"  Replies: {stats['total_replies_received']}")
    print("\nSetup complete.")
    return 0


def cmd_import_csv(args):
    """Import prospects from CSV."""
    db = Database()
    print(f"Importing prospects from {args.file}...")
    result = db.import_prospects_csv(args.file)
    print(f"  Added: {result['added']}")
    print(f"  Skipped (duplicate/invalid): {result['skipped']}")
    return 0


def cmd_import_xlsx(args):
    """Import prospects from an Excel file."""
    db = Database()
    print(f"Importing prospects from {args.file}...")
    result = db.import_prospects_xlsx(args.file, sheet_name=args.sheet)
    print(f"  Added: {result['added']}")
    print(f"  Skipped (duplicate/invalid): {result['skipped']}")
    return 0


def cmd_export_report(args):
    """Export campaign report to an Excel file."""
    from datetime import datetime
    db = Database()
    output = args.output or f"deed_street_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    print(f"Generating report...")
    db.export_report_xlsx(output)
    print(f"Report exported to: {output}")
    print("Sheets included:")
    print("  1. Analytics    — summary metrics, reply rate, status breakdowns")
    print("  2. Prospects    — full contact list with statuses")
    print("  3. Emails Sent  — every outbound email (campaigns + auto-replies)")
    print("  4. Replies Received — all inbound replies")
    print("  5. Conversations — per-prospect thread summaries with AI notes")
    return 0


def cmd_add_prospect(args):
    """Add a single prospect."""
    db = Database()
    pid = db.add_prospect(
        email=args.email,
        first_name=args.first_name or "",
        last_name=args.last_name or "",
        company=args.company or "",
        note_type=args.note_type or "",
        note_balance=args.note_balance or "",
        property_state=args.state or "",
        source=args.source or "manual",
    )
    print(f"Prospect added (id={pid}): {args.email}")
    return 0


def cmd_list_prospects(args):
    """List prospects."""
    db = Database()
    prospects = db.get_all_prospects(status_filter=args.status)
    if not prospects:
        print("No prospects found.")
        return 0

    print(f"{'ID':>5}  {'Email':30s}  {'Name':20s}  {'Status':15s}  {'Source':10s}")
    print("-" * 90)
    for p in prospects:
        name = f"{p['first_name']} {p['last_name']}".strip()
        print(f"{p['id']:>5}  {p['email']:30s}  {name:20s}  {p['status']:15s}  {p['source']:10s}")

    print(f"\nTotal: {len(prospects)}")
    return 0


def cmd_campaigns(args):
    """List campaigns and templates."""
    print("Available Templates:")
    print("-" * 50)
    for name, tpl in ALL_TEMPLATES.items():
        print(f"  {name}")
        print(f"    Subject: {tpl['subject'][:60]}...")
        print()

    db = Database()
    campaigns = db.list_campaigns()
    if campaigns:
        print("\nCreated Campaigns:")
        print("-" * 50)
        for c in campaigns:
            status = "active" if c["is_active"] else "inactive"
            print(f"  [{c['id']}] {c['name']} ({status}) — created {c['created_at']}")
    else:
        print("\nNo campaigns created yet. Use 'create-campaign' to create one.")
    return 0


def cmd_create_campaign(args):
    """Create a campaign from a template."""
    mgr = CampaignManager()
    result = mgr.create_campaign_from_template(
        template_name=args.template,
        campaign_name=args.name,
    )
    if "error" in result:
        print(f"Error: {result['error']}")
        return 1
    print(f"Campaign created: id={result['campaign_id']} — {result['message']}")
    return 0


def cmd_send(args):
    """Send campaign emails."""
    mgr = CampaignManager()

    campaign = mgr.db.get_campaign(args.campaign_id)
    if not campaign:
        print(f"Campaign {args.campaign_id} not found.")
        return 1

    remaining = mgr.db.get_remaining_sends_today()
    limit = min(remaining, args.limit) if args.limit else remaining
    print(f"Campaign: {campaign['name']}")
    print(f"Sends remaining today: {remaining} (limit for this batch: {limit})")

    if limit <= 0:
        print("Daily send limit reached. Try again tomorrow.")
        return 0

    if not args.yes:
        confirm = input(f"Send up to {limit} emails? [y/N] ").strip().lower()
        if confirm != "y":
            print("Cancelled.")
            return 0

    print(f"Sending...")
    result = mgr.send_campaign(args.campaign_id, limit=limit)

    if "error" in result:
        print(f"Error: {result['error']}")
        return 1

    print(f"  Sent: {result.get('sent', 0)}")
    print(f"  Failed: {result.get('failed', 0)}")
    print(f"  Skipped: {result.get('skipped', 0)}")
    return 0


def cmd_check_replies(args):
    """Check for and process replies once."""
    mgr = CampaignManager()
    print("Checking for new replies...")
    result = mgr.process_replies()

    if "error" in result:
        print(f"Error: {result['error']}")
        return 1

    print(f"  Processed: {result.get('processed', 0)}")
    print(f"  Auto-replied: {result.get('auto_replied', 0)}")
    print(f"  Calendly sent: {result.get('calendly_sent', 0)}")
    print(f"  Opt-outs: {result.get('opt_outs', 0)}")
    print(f"  Errors: {result.get('errors', 0)}")
    return 0


def cmd_daemon(args):
    """Run the reply-processing daemon."""
    mgr = CampaignManager()
    mgr.run_daemon()
    return 0


def cmd_report(args):
    """Show campaign statistics."""
    mgr = CampaignManager()
    print(mgr.get_report())
    return 0


def cmd_thread(args):
    """View full email thread for a prospect."""
    db = Database()

    if args.email:
        prospect = db.get_prospect_by_email(args.email)
    else:
        prospect = db.get_prospect_by_id(args.prospect_id)

    if not prospect:
        print("Prospect not found.")
        return 1

    name = f"{prospect['first_name']} {prospect['last_name']}".strip()
    print(f"Thread for: {name} <{prospect['email']}> (status: {prospect['status']})")

    conv = db.get_conversation(prospect["id"])
    if conv:
        print(f"Conversation status: {conv['status']}")
        if conv.get("ai_summary"):
            print(f"AI Summary: {conv['ai_summary']}")

    print("-" * 60)

    thread = db.get_full_thread(prospect["id"])
    if not thread:
        print("(No emails in thread)")
        return 0

    for entry in thread:
        direction = ">>> SENT" if entry["role"] == "sent" else "<<< RECEIVED"
        print(f"\n{direction} [{entry['timestamp']}]")
        print(f"Subject: {entry['subject']}")
        print(entry["body"][:500])
        if len(entry["body"]) > 500:
            print("... (truncated)")
        print("-" * 40)

    return 0


def main():
    parser = argparse.ArgumentParser(
        description="Deed Street Capital — Outbound Email System",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # setup
    subparsers.add_parser("setup", help="Validate config and init database")

    # import-csv
    p_import = subparsers.add_parser("import-csv", help="Import prospects from CSV")
    p_import.add_argument("file", help="Path to CSV file")

    # import-xlsx
    p_xlsx = subparsers.add_parser("import-xlsx", help="Import prospects from Excel (.xlsx)")
    p_xlsx.add_argument("file", help="Path to .xlsx file")
    p_xlsx.add_argument("--sheet", default=None, help="Sheet name (default: first sheet)")

    # export-report
    p_export = subparsers.add_parser("export-report", help="Export report to Excel (.xlsx)")
    p_export.add_argument("-o", "--output", help="Output file path (default: auto-generated)")

    # add-prospect
    p_add = subparsers.add_parser("add-prospect", help="Add a single prospect")
    p_add.add_argument("email", help="Prospect email address")
    p_add.add_argument("--first-name", help="First name")
    p_add.add_argument("--last-name", help="Last name")
    p_add.add_argument("--company", help="Company name")
    p_add.add_argument("--note-type", help="Type of note (residential, commercial)")
    p_add.add_argument("--note-balance", help="Approximate note balance")
    p_add.add_argument("--state", help="Property state")
    p_add.add_argument("--source", default="manual", help="Lead source")

    # list-prospects
    p_list = subparsers.add_parser("list-prospects", help="List prospects")
    p_list.add_argument("--status", help="Filter by status")

    # campaigns
    subparsers.add_parser("campaigns", help="List campaigns and templates")

    # create-campaign
    p_create = subparsers.add_parser("create-campaign", help="Create campaign from template")
    p_create.add_argument("template", choices=list_templates(), help="Template name")
    p_create.add_argument("--name", help="Custom campaign name")

    # send
    p_send = subparsers.add_parser("send", help="Send campaign emails")
    p_send.add_argument("campaign_id", type=int, help="Campaign ID")
    p_send.add_argument("--limit", type=int, help="Max emails to send")
    p_send.add_argument("-y", "--yes", action="store_true", help="Skip confirmation")

    # check-replies
    subparsers.add_parser("check-replies", help="Check for and process replies")

    # daemon
    subparsers.add_parser("daemon", help="Run reply-processing daemon")

    # report
    subparsers.add_parser("report", help="Show campaign statistics")

    # thread
    p_thread = subparsers.add_parser("thread", help="View email thread for a prospect")
    group = p_thread.add_mutually_exclusive_group(required=True)
    group.add_argument("--email", help="Prospect email")
    group.add_argument("--id", dest="prospect_id", type=int, help="Prospect ID")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 1

    commands = {
        "setup": cmd_setup,
        "import-csv": cmd_import_csv,
        "import-xlsx": cmd_import_xlsx,
        "export-report": cmd_export_report,
        "add-prospect": cmd_add_prospect,
        "list-prospects": cmd_list_prospects,
        "campaigns": cmd_campaigns,
        "create-campaign": cmd_create_campaign,
        "send": cmd_send,
        "check-replies": cmd_check_replies,
        "daemon": cmd_daemon,
        "report": cmd_report,
        "thread": cmd_thread,
    }

    handler = commands.get(args.command)
    if handler:
        return handler(args)
    else:
        parser.print_help()
        return 1


if __name__ == "__main__":
    sys.exit(main())
