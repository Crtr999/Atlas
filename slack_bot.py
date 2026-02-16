#!/usr/bin/env python3
"""
CBC Atlas Slack Bot

A Slack bot powered by Claude that can answer questions about the CBC
Settlement Funding Excel data, edit the spreadsheet, and regenerate HTML.

Usage:
    1. Copy .env.example to .env and fill in your tokens
    2. pip install -r requirements.txt
    3. python slack_bot.py

In Slack, mention @Claude with your question:
    @Claude What are the redaction levels in Ramsey, MN?
    @Claude Add a new case for John Doe in Hennepin County, MN - Approved
    @Claude Update the county notes for Ramsey, MN
    @Claude Regenerate the HTML
"""

import os
import re
import json
import logging
from pathlib import Path

from dotenv import load_dotenv
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler
import anthropic

from data_reader import AtlasDataReader

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

# ── Tool definitions for Claude ────────────────────────────────────

TOOLS = [
    {
        "name": "query_county",
        "description": (
            "Look up all information about a specific county including "
            "redaction level, court access, cases, and notes. "
            "Use this when the user asks about a specific county."
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
            "Look up all information about a state including rate caps, "
            "IPA requirements, legal fees, and case summary."
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
            "Get a full summary of all data in the Excel file. Use this when "
            "the user asks a broad question that spans multiple counties/states, "
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
            "required": ["state", "county", "judge", "client_name",
                         "case_number", "court_date", "status"],
        },
    },
    {
        "name": "update_county_notes",
        "description": (
            "Update or add notes for a specific county in the Jurisdictions sheet."
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
                    "description": "Redaction level, e.g. 'Full Redaction', 'Partial Redaction', 'No Redaction'",
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
        "description": "Update the status of an existing case.",
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
]

# ── Tool execution ─────────────────────────────────────────────────


def execute_tool(tool_name: str, tool_input: dict) -> str:
    """Execute a tool call and return the result string."""
    try:
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
            result = data_reader.add_case(**tool_input)
            return result
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
        else:
            return f"Unknown tool: {tool_name}"
    except Exception as e:
        logger.error(f"Tool execution error: {e}", exc_info=True)
        return f"Error executing {tool_name}: {e}"


# ── Claude conversation with tool use ──────────────────────────────

SYSTEM_PROMPT = """You are the CBC Atlas Assistant, a helpful bot integrated into Slack.
You have access to the CBC Settlement Funding Master Excel workbook which contains:

1. **Cases** - Settlement funding court cases with State, County, Judge, Client, Status (Approved/Denied/Dismissed)
2. **Jurisdictions** - County-level info with Redaction Levels, Redaction Notes, and County Notes
3. **State Data** - State rules: rate caps, IPA requirements, affidavit/declaration requirements, legal fees
4. **Insurance Carriers** - Carrier admin fees, contacts, and procedures
5. **Court Access** - Court access info for 3,140 counties: Westlaw coverage, websites, fees, login info

You can:
- **Answer questions** about any data in the spreadsheet (redaction levels, cases, state rules, etc.)
- **Edit the Excel file** (add cases, update county notes, change redaction levels, update case statuses)
- **Regenerate the HTML** application after making edits

Guidelines:
- Always use the appropriate tool to look up data before answering. Do not guess.
- When the user asks about a county, use query_county. For states, use query_state.
- For broad questions spanning multiple counties/states, use get_all_data_summary.
- After any edit (add_case, update_county_notes, etc.), always call regenerate_html to keep the HTML in sync.
- Give clear, concise answers. Format nicely for Slack (use *bold*, bullet points, etc.).
- When giving feedback on court processes, base it on the actual data in the spreadsheet.
- If the user asks about "our court process" or similar, look up the relevant data and provide analysis.
"""


def ask_claude(user_message: str) -> str:
    """Send a message to Claude with tool use and return the final response."""
    messages = [{"role": "user", "content": user_message}]

    # Loop to handle tool use (Claude may call multiple tools in sequence)
    max_iterations = 10
    for _ in range(max_iterations):
        response = claude_client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            tools=TOOLS,
            messages=messages,
        )

        # Check if Claude wants to use tools
        if response.stop_reason == "tool_use":
            # Process all tool calls in the response
            assistant_content = response.content
            messages.append({"role": "assistant", "content": assistant_content})

            tool_results = []
            for block in assistant_content:
                if block.type == "tool_use":
                    logger.info(f"Tool call: {block.name}({json.dumps(block.input)})")
                    result = execute_tool(block.name, block.input)
                    tool_results.append({
                        "type": "tool_result",
                        "tool_use_id": block.id,
                        "content": result,
                    })

            messages.append({"role": "user", "content": tool_results})
        else:
            # Extract the final text response
            text_parts = []
            for block in response.content:
                if hasattr(block, "text"):
                    text_parts.append(block.text)
            return "\n".join(text_parts) if text_parts else "I processed your request but have no additional response."

    return "I hit the maximum number of tool calls. Please try a simpler question."


# ── Slack event handlers ───────────────────────────────────────────

@app.event("app_mention")
def handle_mention(event, say):
    """Handle @Claude mentions in channels."""
    user_message = event.get("text", "")
    user_id = event.get("user", "")
    channel = event.get("channel", "")

    # Strip the bot mention from the message
    user_message = re.sub(r"<@[A-Z0-9]+>", "", user_message).strip()

    if not user_message:
        say(
            text="Hi! Ask me anything about the CBC Settlement Funding data. "
                 "For example:\n"
                 "- *What are the redaction levels in Ramsey, MN?*\n"
                 "- *Give me feedback on our court process in Minnesota*\n"
                 "- *Add a new case for John Doe in Hennepin County, MN*\n"
                 "- *Regenerate the HTML*",
            channel=channel,
        )
        return

    logger.info(f"Message from <@{user_id}>: {user_message}")

    # Send a "thinking" message
    thinking = say(text=":hourglass_flowing_sand: Looking into that...", channel=channel)

    try:
        response = ask_claude(user_message)

        # Update the thinking message with the actual response
        app.client.chat_update(
            channel=channel,
            ts=thinking["ts"],
            text=response,
        )
    except Exception as e:
        logger.error(f"Error processing message: {e}", exc_info=True)
        app.client.chat_update(
            channel=channel,
            ts=thinking["ts"],
            text=f":x: Sorry, I ran into an error: {e}",
        )


@app.event("message")
def handle_direct_message(event, say):
    """Handle direct messages to the bot."""
    # Skip bot messages, message_changed events, etc.
    if event.get("subtype"):
        return

    # Only respond to DMs (channel type 'im')
    if event.get("channel_type") != "im":
        return

    user_message = event.get("text", "").strip()
    user_id = event.get("user", "")
    channel = event.get("channel", "")

    if not user_message:
        return

    logger.info(f"DM from <@{user_id}>: {user_message}")

    # Send a "thinking" message
    thinking = say(text=":hourglass_flowing_sand: Looking into that...", channel=channel)

    try:
        response = ask_claude(user_message)

        app.client.chat_update(
            channel=channel,
            ts=thinking["ts"],
            text=response,
        )
    except Exception as e:
        logger.error(f"Error processing DM: {e}", exc_info=True)
        app.client.chat_update(
            channel=channel,
            ts=thinking["ts"],
            text=f":x: Sorry, I ran into an error: {e}",
        )


# ── Main ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 60)
    print("CBC ATLAS SLACK BOT")
    print("=" * 60)
    print(f"\nExcel file: {EXCEL_PATH}")
    print(f"HTML output: {HTML_OUTPUT_PATH}")
    print("\nBot is starting... Connect to Slack via Socket Mode.")
    print("Mention @Claude in a channel or send a DM to interact.\n")

    handler = SocketModeHandler(app, SLACK_APP_TOKEN)
    handler.start()
