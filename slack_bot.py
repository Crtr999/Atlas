#!/usr/bin/env python3
"""
CBC Atlas Slack Bot

A Slack bot powered by Claude that can:
  - Answer questions about CBC Settlement Funding court data
  - Edit the Excel workbook (add cases, update notes, etc.)
  - Save and retrieve institutional knowledge (judge patterns, court tips, etc.)
  - Run analytical queries and risk scores over the case history
  - Regenerate the HTML application

Usage:
    1. Copy .env.example to .env and fill in your tokens
    2. pip install -r requirements.txt
    3. python slack_bot.py

In Slack, mention @Atlas or send a DM:
    @Atlas What are the redaction levels in Ramsey, MN?
    @Atlas Note that Judge Smith in Cook County, IL always denies cases with liens > 33%
    @Atlas What's our denial rate in Texas?
    @Atlas Show me a risk score for Hennepin County, MN
    @Atlas What do we know about Judge Williams in Tarrant County, TX?
"""

import json
import logging
import os
import re
import traceback

from dotenv import load_dotenv
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler
import anthropic

from analytics import AtlasAnalytics
from data_reader import AtlasDataReader
from knowledge_store import KnowledgeStore

# ── Setup ──────────────────────────────────────────────────────────

load_dotenv()
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("atlas-bot")

SLACK_BOT_TOKEN = os.environ.get("SLACK_BOT_TOKEN")
SLACK_APP_TOKEN = os.environ.get("SLACK_APP_TOKEN")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
EXCEL_PATH = os.environ.get("EXCEL_PATH", "./CBC_Settlement_Funding_Master_v4.xlsx")
HTML_OUTPUT_PATH = os.environ.get("HTML_OUTPUT_PATH", "./CBCAssistant_Complete.html")

if not SLACK_BOT_TOKEN:
    raise ValueError("SLACK_BOT_TOKEN is required. See .env.example")
if not SLACK_APP_TOKEN:
    raise ValueError("SLACK_APP_TOKEN is required. See .env.example")
if not ANTHROPIC_API_KEY:
    raise ValueError("ANTHROPIC_API_KEY is required. See .env.example")

app = App(token=SLACK_BOT_TOKEN)
claude_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
data_reader = AtlasDataReader(EXCEL_PATH)
knowledge_store = KnowledgeStore(repo_dir=".")
analytics = AtlasAnalytics(data_reader)

# ── Tool definitions for Claude ────────────────────────────────────

TOOLS = [
    # ── Existing structured-data tools ─────────────────────────────
    {
        "name": "query_county",
        "description": (
            "Look up all structured information about a specific county including "
            "redaction level, court access, cases, and notes from the Excel file. "
            "Also pair this with get_entity_knowledge to surface team insights."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "county_name": {
                    "type": "string",
                    "description": "The county name, e.g. 'Ramsey' or 'Ramsey County'",
                },
                "state_code": {
                    "type": "string",
                    "description": "Two-letter state code, e.g. 'MN'. Optional but recommended.",
                },
            },
            "required": ["county_name"],
        },
    },
    {
        "name": "query_state",
        "description": (
            "Look up all structured information about a state including rate caps, "
            "IPA requirements, legal fees, and case summary from the Excel file."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "state_code": {
                    "type": "string",
                    "description": "Two-letter state code, e.g. 'MN'",
                },
            },
            "required": ["state_code"],
        },
    },
    {
        "name": "get_all_data_summary",
        "description": (
            "Get a full summary of all structured data in the Excel file. Use this when "
            "the user asks a broad question spanning multiple counties/states, "
            "or when you need to compare across the dataset."
        ),
        "input_schema": {
            "type": "object",
            "properties": {},
        },
    },
    {
        "name": "add_case",
        "description": (
            "Add a new case to the Cases sheet in the Excel file. "
            "Use when the user wants to record a new court case."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "state": {"type": "string", "description": "Two-letter state code"},
                "county": {"type": "string", "description": "County name"},
                "judge": {"type": "string", "description": "Judge name"},
                "client_name": {"type": "string", "description": "Client name"},
                "case_number": {"type": "string", "description": "Case number"},
                "court_date": {"type": "string", "description": "Court date"},
                "status": {
                    "type": "string",
                    "description": "Case status: Approved, Denied, or Dismissed",
                },
                "notes": {"type": "string", "description": "Optional case notes"},
            },
            "required": [
                "state", "county", "judge", "client_name",
                "case_number", "court_date", "status",
            ],
        },
    },
    {
        "name": "update_county_notes",
        "description": (
            "Update or add structured notes for a specific county in the "
            "Jurisdictions sheet of the Excel file."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "state": {"type": "string", "description": "Two-letter state code"},
                "county": {"type": "string", "description": "County name"},
                "notes": {"type": "string", "description": "The notes to set"},
            },
            "required": ["state", "county", "notes"],
        },
    },
    {
        "name": "update_redaction_level",
        "description": (
            "Update or add the redaction level for a county in the Jurisdictions sheet."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "state": {"type": "string", "description": "Two-letter state code"},
                "county": {"type": "string", "description": "County name"},
                "redaction_level": {
                    "type": "string",
                    "description": (
                        "Redaction level: 'Full Redaction', 'Partial Redaction', "
                        "or 'No Redaction'"
                    ),
                },
                "redaction_notes": {
                    "type": "string",
                    "description": "Optional notes about the redaction",
                },
            },
            "required": ["state", "county", "redaction_level"],
        },
    },
    {
        "name": "update_case_status",
        "description": "Update the status of an existing case in the Excel file.",
        "input_schema": {
            "type": "object",
            "properties": {
                "state": {"type": "string", "description": "Two-letter state code"},
                "county": {"type": "string", "description": "County name"},
                "client_name": {"type": "string", "description": "Client name to find"},
                "new_status": {
                    "type": "string",
                    "description": "New status: Approved, Denied, or Dismissed",
                },
            },
            "required": ["state", "county", "client_name", "new_status"],
        },
    },
    {
        "name": "regenerate_html",
        "description": (
            "Regenerate the CBC Assistant HTML file from the current Excel data. "
            "Call this after making any edits to the Excel file."
        ),
        "input_schema": {
            "type": "object",
            "properties": {},
        },
    },

    # ── Knowledge base tools ────────────────────────────────────────
    {
        "name": "save_knowledge",
        "description": (
            "Save an institutional knowledge snippet about a judge, county, state, "
            "insurance carrier, or general court insight. "
            "Use this when the user says things like 'note that...', 'remember that...', "
            "'save this...', 'add a note that...', or shares an observation or lesson "
            "learned about a court, judge, or process. "
            "Always use the current user's name as the author."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "content": {
                    "type": "string",
                    "description": "The knowledge text to save",
                },
                "author": {
                    "type": "string",
                    "description": (
                        "Name of the person sharing this knowledge. "
                        "Use the current user's display name."
                    ),
                },
                "entity_type": {
                    "type": "string",
                    "description": (
                        "Category: 'county', 'judge', 'state', 'insurance', or 'general'"
                    ),
                },
                "entity_name": {
                    "type": "string",
                    "description": (
                        "The specific entity this knowledge is about, e.g. "
                        "'Judge Williams', 'Tarrant County', 'TX'. "
                        "Leave empty for general knowledge."
                    ),
                },
                "state": {
                    "type": "string",
                    "description": "Two-letter state code if applicable, e.g. 'TX'",
                },
            },
            "required": ["content", "author"],
        },
    },
    {
        "name": "search_knowledge",
        "description": (
            "Search the team knowledge base for institutional insights, notes, and "
            "lessons learned. Use this alongside structured data lookups to give "
            "comprehensive answers. Useful for questions about judge behavior, "
            "court quirks, process tips, or anything the team has observed over time. "
            "Always search knowledge when answering questions about specific counties, "
            "judges, or states."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "What to search for, e.g. 'Judge Williams denial patterns'",
                },
                "entity_type": {
                    "type": "string",
                    "description": (
                        "Optional filter: 'county', 'judge', 'state', 'insurance', 'general'"
                    ),
                },
                "state": {
                    "type": "string",
                    "description": "Optional two-letter state code to narrow results",
                },
            },
            "required": ["query"],
        },
    },
    {
        "name": "get_entity_knowledge",
        "description": (
            "Get all saved knowledge entries for a specific entity (county, judge, state, "
            "or insurance carrier). Use alongside query_county or query_state to "
            "surface everything the team knows about a specific entity."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "entity_type": {
                    "type": "string",
                    "description": "'county', 'judge', 'state', or 'insurance'",
                },
                "entity_name": {
                    "type": "string",
                    "description": (
                        "Name of the entity, e.g. 'Judge Williams', 'Tarrant County'"
                    ),
                },
                "state": {
                    "type": "string",
                    "description": "Optional two-letter state code",
                },
            },
            "required": ["entity_type", "entity_name"],
        },
    },
    {
        "name": "delete_knowledge",
        "description": (
            "Delete a knowledge entry by its ID. Use when a user asks to remove or "
            "correct a specific knowledge entry. The ID is the short code shown in "
            "brackets, e.g. [a1b2c3d4]."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "knowledge_id": {
                    "type": "string",
                    "description": "The 8-character knowledge entry ID to delete",
                },
            },
            "required": ["knowledge_id"],
        },
    },
    {
        "name": "list_recent_knowledge",
        "description": (
            "List the most recently added knowledge entries. Use when the user asks "
            "what's been recently saved, wants to review the knowledge base, or wants "
            "to see what the team has been learning."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "limit": {
                    "type": "integer",
                    "description": "How many entries to return (default 10, max 25)",
                },
            },
        },
    },

    # ── Analytics tools ─────────────────────────────────────────────
    {
        "name": "get_risk_score",
        "description": (
            "Get the historical risk score for a county, judge, state, or combination "
            "based on actual case outcomes. Use when the user asks 'is it safe to fund "
            "in X', 'what's our history in Y', 'how does Judge Z typically rule', "
            "'should we take this case in...', or similar risk questions."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "county": {
                    "type": "string",
                    "description": "County name (optional)",
                },
                "state": {
                    "type": "string",
                    "description": "Two-letter state code (optional)",
                },
                "judge": {
                    "type": "string",
                    "description": "Judge name (optional)",
                },
            },
        },
    },
    {
        "name": "run_analytics",
        "description": (
            "Run an analytical query over the full case history. Use for questions "
            "like 'what states have the highest denial rate', 'which judges approve "
            "most cases', 'show me all denied cases in Texas', 'give me a portfolio "
            "overview', or any broad data analysis question."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "query_type": {
                    "type": "string",
                    "description": (
                        "One of: 'denial_patterns', 'judge_analytics', "
                        "'state_analytics', 'portfolio_summary', 'cases_by_status'"
                    ),
                },
                "state": {
                    "type": "string",
                    "description": "Optional two-letter state code to filter results",
                },
                "judge": {
                    "type": "string",
                    "description": "Optional judge name to filter results",
                },
                "status": {
                    "type": "string",
                    "description": (
                        "Optional status filter for cases_by_status: "
                        "'Approved', 'Denied', or 'Dismissed'"
                    ),
                },
            },
            "required": ["query_type"],
        },
    },
]


# ── Tool execution ──────────────────────────────────────────────────

def execute_tool(tool_name: str, tool_input: dict) -> str:
    """Execute a tool call and return the result as a string."""
    try:
        # --- Structured data tools ---
        if tool_name == "query_county":
            return data_reader.query_county(
                tool_input["county_name"],
                tool_input.get("state_code"),
            )

        elif tool_name == "query_state":
            return data_reader.query_state(tool_input["state_code"])

        elif tool_name == "get_all_data_summary":
            return data_reader.get_all_data_summary()

        elif tool_name == "add_case":
            return data_reader.add_case(**tool_input)

        elif tool_name == "update_county_notes":
            return data_reader.update_county_notes(
                tool_input["state"],
                tool_input["county"],
                tool_input["notes"],
            )

        elif tool_name == "update_redaction_level":
            return data_reader.update_redaction_level(
                tool_input["state"],
                tool_input["county"],
                tool_input["redaction_level"],
                tool_input.get("redaction_notes", ""),
            )

        elif tool_name == "update_case_status":
            return data_reader.update_case_status(
                tool_input["state"],
                tool_input["county"],
                tool_input["client_name"],
                tool_input["new_status"],
            )

        elif tool_name == "regenerate_html":
            return data_reader.regenerate_html()

        # --- Knowledge base tools ---
        elif tool_name == "save_knowledge":
            kid = knowledge_store.add_knowledge(
                content=tool_input["content"],
                author=tool_input.get("author", "Team Member"),
                entity_type=tool_input.get("entity_type", "general"),
                entity_name=tool_input.get("entity_name", ""),
                state=tool_input.get("state", ""),
            )
            return (
                f"Knowledge saved! Entry ID: [{kid}]\n"
                f"Content: {tool_input['content']}\n"
                f"Category: {tool_input.get('entity_type', 'general')}"
                + (f" — {tool_input.get('entity_name', '')}" if tool_input.get("entity_name") else "")
                + (f", {tool_input.get('state', '')}" if tool_input.get("state") else "")
            )

        elif tool_name == "search_knowledge":
            items = knowledge_store.search_knowledge(
                query=tool_input["query"],
                entity_type=tool_input.get("entity_type"),
                state=tool_input.get("state"),
            )
            if not items:
                return "No knowledge entries found matching that query."
            stats = knowledge_store.get_stats()
            header = (
                f"Found {len(items)} relevant knowledge entries "
                f"(knowledge base has {stats['total']} total):\n"
            )
            return header + KnowledgeStore.format_for_slack(items)

        elif tool_name == "get_entity_knowledge":
            items = knowledge_store.get_by_entity(
                entity_type=tool_input["entity_type"],
                entity_name=tool_input["entity_name"],
                state=tool_input.get("state"),
            )
            if not items:
                name = tool_input["entity_name"]
                etype = tool_input["entity_type"]
                return f"No knowledge entries found for {etype} '{name}'."
            return (
                f"Team knowledge for {tool_input['entity_name']} "
                f"({tool_input['entity_type']}):\n"
                + KnowledgeStore.format_for_slack(items)
            )

        elif tool_name == "delete_knowledge":
            kid = tool_input["knowledge_id"]
            success = knowledge_store.delete_knowledge(kid)
            if success:
                return f"Knowledge entry [{kid}] deleted successfully."
            return f"Knowledge entry [{kid}] not found — no changes made."

        elif tool_name == "list_recent_knowledge":
            limit = min(int(tool_input.get("limit", 10)), 25)
            items = knowledge_store.list_recent(limit=limit)
            stats = knowledge_store.get_stats()
            if not items:
                return "The knowledge base is empty. Start adding entries with 'save this...' or 'note that...'"
            header = (
                f"Most recent {len(items)} knowledge entries "
                f"(of {stats['total']} total, by type: "
                + ", ".join(f"{k}: {v}" for k, v in sorted(stats["by_type"].items()))
                + "):\n"
            )
            return header + KnowledgeStore.format_for_slack(items)

        # --- Analytics tools ---
        elif tool_name == "get_risk_score":
            return analytics.get_risk_score(
                county=tool_input.get("county"),
                state=tool_input.get("state"),
                judge=tool_input.get("judge"),
            )

        elif tool_name == "run_analytics":
            return analytics.run_analytics_query(
                query_type=tool_input["query_type"],
                state=tool_input.get("state"),
                judge=tool_input.get("judge"),
                status=tool_input.get("status"),
            )

        else:
            return f"Unknown tool: {tool_name}"

    except Exception as e:
        logger.error(f"Tool execution error ({tool_name}): {e}", exc_info=True)
        return f"Error executing {tool_name}: {e}"


# ── System prompt ───────────────────────────────────────────────────

BASE_SYSTEM_PROMPT = """You are the CBC Atlas Assistant, a helpful AI integrated into Slack \
for the CBC Settlement Funding Court Research team.

You have access to two layers of knowledge:

**1. Structured Excel Data (via tools)**
- Cases — settlement funding court cases with State, County, Judge, Client, \
Status (Approved/Denied/Dismissed)
- Jurisdictions — county-level info: Redaction Levels, Redaction Notes, County Notes
- State Data — rate caps, IPA requirements, affidavit/declaration requirements, legal fees
- Insurance Carriers — carrier admin fees, contacts, and procedures
- Court Access — info for 3,140+ counties: Westlaw coverage, websites, fees, logins

**2. Team Knowledge Base (via tools)**
- Institutional insights saved by the team: judge patterns, court quirks, process tips, \
lessons learned from past cases
- Searchable by query, entity type, and state
- Grows over time as the team contributes observations

**Analytics & Risk Scoring (via tools)**
- Historical risk scores by county, state, judge, or combination
- Denial pattern analysis — identify high-risk jurisdictions and judges
- Portfolio summary, state analytics, judge performance breakdowns

**What you can do:**
- Answer questions about court data, redaction levels, state rules, and more
- Save and retrieve team knowledge when people share insights or ask what we know
- Run risk assessments and analytical queries over case history
- Edit the Excel file (add cases, update notes, change redaction levels, update statuses)
- Regenerate the HTML application after edits

**Core guidelines:**
- For county/state questions: always call BOTH the structured query tool AND search_knowledge \
or get_entity_knowledge — give a complete answer combining both sources
- For "note that / remember / save this" phrasing: always use save_knowledge
- For risk and "is this a good jurisdiction" questions: use get_risk_score, optionally \
combined with search_knowledge
- For broad analytics: use run_analytics with the appropriate query_type
- After any Excel edit, always call regenerate_html to keep the HTML in sync
- Format responses for Slack: use *bold*, bullet points, concise language
- Never guess — always use a tool to look up data before answering
"""


def _build_system_prompt(user_name: str) -> str:
    """Inject the current user's name so Claude knows who is speaking."""
    return BASE_SYSTEM_PROMPT + f"\n\n**Current user:** {user_name} — use this name as the author when saving knowledge."


# ── Claude conversation with tool use ──────────────────────────────

def ask_claude(user_message: str, user_name: str = "Team Member") -> str:
    """
    Send a message to Claude with tool use and return the final text response.

    Handles multi-turn tool loops (Claude calls tools, we execute, Claude continues)
    up to MAX_ITERATIONS rounds.
    """
    system_prompt = _build_system_prompt(user_name)
    messages = [{"role": "user", "content": user_message}]

    max_iterations = 12
    for _ in range(max_iterations):
        response = claude_client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4096,
            system=system_prompt,
            tools=TOOLS,
            messages=messages,
        )

        if response.stop_reason == "tool_use":
            assistant_content = response.content
            messages.append({"role": "assistant", "content": assistant_content})

            tool_results = []
            for block in assistant_content:
                if block.type == "tool_use":
                    logger.info(f"Tool call: {block.name}({json.dumps(block.input)})")
                    result = execute_tool(block.name, block.input)
                    tool_results.append(
                        {
                            "type": "tool_result",
                            "tool_use_id": block.id,
                            "content": result,
                        }
                    )
            messages.append({"role": "user", "content": tool_results})

        else:
            text_parts = [
                block.text
                for block in response.content
                if hasattr(block, "text")
            ]
            return "\n".join(text_parts) if text_parts else (
                "I processed your request but have nothing further to add."
            )

    return "I reached the maximum number of tool calls. Try asking a more specific question."


# ── Slack helpers ───────────────────────────────────────────────────

def _get_user_display_name(user_id: str) -> str:
    """
    Look up a Slack user's display name from their user ID.

    Falls back to the user_id string if the lookup fails.
    Requires the users:read OAuth scope on the Slack app.
    """
    try:
        result = app.client.users_info(user=user_id)
        profile = result["user"].get("profile", {})
        name = (
            profile.get("display_name")
            or profile.get("real_name")
            or result["user"].get("name")
            or user_id
        )
        return name.strip() or user_id
    except Exception:
        return user_id


# ── Slack event handlers ────────────────────────────────────────────

@app.event("app_mention")
def handle_mention(event, say):
    """Handle @Atlas mentions in channels."""
    try:
        user_message = event.get("text", "")
        user_id = event.get("user", "")
        channel = event.get("channel", "")
        print(f"[BOT] Mention from {user_id} in {channel}")

        # Strip the bot mention tag from the message text
        user_message = re.sub(r"<@[A-Z0-9]+>", "", user_message).strip()
        print(f"[BOT] Message: {user_message}")

        if not user_message:
            say(
                text=(
                    "Hi! Ask me anything about CBC Settlement Funding data, "
                    "save team knowledge with 'note that...', or ask for a risk score."
                ),
                channel=channel,
            )
            return

        user_name = _get_user_display_name(user_id)
        print(f"[BOT] User display name: {user_name}")

        thinking = say(text="Looking into that...", channel=channel)
        print("[BOT] Calling Claude API...")

        response = ask_claude(user_message, user_name=user_name)
        print(f"[BOT] Response ready ({len(response)} chars)")

        app.client.chat_update(
            channel=channel,
            ts=thinking["ts"],
            text=response,
        )
        print("[BOT] Response sent.")

    except Exception as e:
        print(f"[ERROR] {type(e).__name__}: {e}")
        traceback.print_exc()


@app.event("message")
def handle_direct_message(event, say):
    """Handle direct messages to the bot."""
    # Ignore bot messages, edits, and anything in public channels
    if event.get("subtype"):
        return
    if event.get("channel_type") != "im":
        return

    try:
        user_message = event.get("text", "").strip()
        user_id = event.get("user", "")
        channel = event.get("channel", "")

        if not user_message:
            return

        print(f"[BOT] DM from {user_id}: {user_message}")

        user_name = _get_user_display_name(user_id)

        thinking = say(text="Looking into that...", channel=channel)

        response = ask_claude(user_message, user_name=user_name)
        print(f"[BOT] Response ready ({len(response)} chars)")

        app.client.chat_update(
            channel=channel,
            ts=thinking["ts"],
            text=response,
        )
        print("[BOT] Response sent.")

    except Exception as e:
        print(f"[ERROR] {type(e).__name__}: {e}")
        traceback.print_exc()


# ── Main ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    kb_stats = knowledge_store.get_stats()

    print("=" * 60)
    print("CBC ATLAS SLACK BOT")
    print("=" * 60)
    print(f"\nExcel file:       {EXCEL_PATH}")
    print(f"HTML output:      {HTML_OUTPUT_PATH}")
    print(f"Knowledge base:   {kb_stats['total']} entries")
    if kb_stats["by_type"]:
        for etype, count in sorted(kb_stats["by_type"].items()):
            print(f"  {etype}: {count}")
    print(f"\nModel:            claude-sonnet-4-6")
    print(f"Tools available:  {len(TOOLS)}")
    print("\nBot is starting... Connect to Slack via Socket Mode.")
    print("Mention @Atlas in a channel or send a DM to interact.\n")

    handler = SocketModeHandler(app, SLACK_APP_TOKEN)
    handler.start()
