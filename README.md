# CBC Atlas — Court Research Intelligence System

A complete system for CBC Settlement Funding's Court Research team.
Query court data, build institutional knowledge, analyze risk, and share insights — all from Slack.

---

## What's Included

| File | Purpose |
|---|---|
| `CBC_Settlement_Funding_Master_v4.xlsx` | Master data file (6 sheets) |
| `CBCAssistant_Complete.html` | Generated web application |
| `generate_cbc_assistant.py` | Regenerates the HTML from Excel |
| `slack_bot.py` | Slack bot (Claude-powered, all capabilities) |
| `data_reader.py` | Excel read/write layer |
| `analytics.py` | Risk scoring and analytical queries |
| `knowledge_store.py` | Persistent institutional knowledge base |
| `knowledge_entries.json` | Knowledge base data (auto-created, git-tracked) |
| `.env.example` | Environment variable template |

---

## Capabilities

### 1. Structured Data Queries
Ask anything about the 6 Excel sheets: counties, redaction levels, state rules, court access, insurance carriers, and case history.

### 2. Team Knowledge Base
Save and retrieve institutional insights — judge patterns, court quirks, process tips — that live alongside the structured data. Knowledge is stored in `knowledge_entries.json`, committed to git automatically, and persists across restarts.

### 3. Risk Scoring
Get historical approval/denial rates for any county, state, judge, or combination. Four risk tiers: LOW / MODERATE / HIGH / VERY HIGH.

### 4. Analytics
Run analytical queries across the full case history: denial patterns, judge performance, state breakdowns, portfolio summary.

### 5. Excel Editing
Add cases, update county notes, change redaction levels, and update case statuses — all from a Slack message.

### 6. HTML Regeneration
The web application regenerates automatically after any Excel edit.

---

## Regenerating the HTML

```bash
python3 generate_cbc_assistant.py
```

---

## Slack Bot Setup

### Required OAuth Scopes
Go to **OAuth & Permissions > Bot Token Scopes** and add:

| Scope | Purpose |
|---|---|
| `app_mentions:read` | Respond to @mentions |
| `chat:write` | Send messages |
| `im:history` | Read DMs |
| `im:read` | Access DM channels |
| `im:write` | Send DMs |
| `users:read` | Look up user display names for knowledge authorship |

### Setup Steps

1. Go to [api.slack.com/apps](https://api.slack.com/apps) → **Create New App** → **From scratch**
2. Name it (e.g. "CBC Atlas") and pick your workspace

3. **Enable Socket Mode:**
   - Go to **Socket Mode**, toggle on
   - Generate an **App-Level Token** with `connections:write` scope → this is your `SLACK_APP_TOKEN` (starts with `xapp-`)

4. **Set Bot Permissions** (see table above)

5. **Enable Events:**
   - Go to **Event Subscriptions**, toggle on
   - Under **Subscribe to bot events**, add: `app_mention`, `message.im`

6. **Install to Workspace:**
   - Go to **Install App** and install
   - Copy the **Bot User OAuth Token** → this is your `SLACK_BOT_TOKEN` (starts with `xoxb-`)

7. **Anthropic API Key:**
   - Go to [console.anthropic.com](https://console.anthropic.com/) and create an API key

### Running the Bot

```bash
pip install -r requirements.txt
cp .env.example .env
# Fill in SLACK_BOT_TOKEN, SLACK_APP_TOKEN, ANTHROPIC_API_KEY
python slack_bot.py
```

---

## Using the Bot in Slack

Mention `@Atlas` in any channel or send a DM.

### Structured Data Queries
```
@Atlas What are the redaction levels in Ramsey, MN?
@Atlas What are the state rules for Texas?
@Atlas Tell me everything about Hennepin County, MN
@Atlas What's the court access info for Cook County, IL?
```

### Knowledge Base — Saving Insights
```
@Atlas Note that Judge Williams in Tarrant County, TX consistently denies cases with attorney liens over 40%
@Atlas Remember that Harris County, TX courthouse closed every third Friday
@Atlas Save this: State Farm requires 30-day notice before any hearing in FL
@Atlas Add a note that Judge Lopez in Miami-Dade approves most cases under $50K
```

### Knowledge Base — Retrieving Insights
```
@Atlas What do we know about Judge Williams?
@Atlas Show me all knowledge about Texas courts
@Atlas Search for notes about attorney lien issues
@Atlas What's been recently saved to the knowledge base?
```

### Risk Scoring
```
@Atlas What's the risk score for Cook County, IL?
@Atlas Give me a risk assessment for Judge Smith in Dallas County, TX
@Atlas Is Minnesota a safe state to fund in?
@Atlas How does Judge Johnson typically rule?
```

### Analytics
```
@Atlas What states have the highest denial rates?
@Atlas Show me a portfolio summary
@Atlas Which judges approve most cases?
@Atlas Show me all denied cases in Texas
@Atlas Give me judge analytics for California
```

### Excel Editing
```
@Atlas Add a new case: John Doe in Ramsey County, MN — Approved by Judge Smith, case #2025-001, date 2025-03-15
@Atlas Update the county notes for Ramsey, MN to "New courthouse procedures effective 2025"
@Atlas Change the redaction level for Hennepin County, MN to Full Redaction
@Atlas Update case status for Jane Smith in Cook County, IL to Approved
@Atlas Regenerate the HTML
```

---

## How the Knowledge Base Works

When a team member shares an insight in Slack, the bot saves it to `knowledge_entries.json` with:
- The knowledge text
- Who said it and when
- What entity it relates to (county, judge, state, insurance, or general)
- The state code

Every entry gets a short ID (e.g. `[a1b2c3d4]`) that can be used to delete it later.

The file is automatically committed to git after every add/delete, so:
- Knowledge persists across bot restarts
- Changes are version-controlled and auditable
- The team can see who added what and when

When answering questions about a specific county, judge, or state, the bot automatically searches both the structured Excel data AND the knowledge base to give a complete answer.

### Managing Knowledge
```
@Atlas List recent knowledge entries
@Atlas Show me all knowledge about Judge Smith
@Atlas Delete knowledge entry [a1b2c3d4]
```

---

## Architecture

```
slack_bot.py          ← Event handling, Claude tool loop, user name lookup
  │
  ├── data_reader.py  ← Excel read/write, git commit/push
  ├── analytics.py    ← Risk scoring, denial patterns, portfolio analytics
  └── knowledge_store.py ← JSON knowledge base, keyword search, git commit/push
```

Claude (claude-sonnet-4-6) acts as the reasoning layer, deciding which tools to call
based on the user's question. The bot executes the tools and returns results to Claude,
which formulates a final Slack-formatted response.

---

## Troubleshooting

**"Template not found"**
→ Make sure `CBCAssistant_v4.html` is in the same folder

**"Excel file not found"**
→ Make sure `CBC_Settlement_Funding_Master_v4.xlsx` is in the same folder

**"No module named pandas"**
→ Run: `pip install -r requirements.txt`

**Knowledge not persisting after restart**
→ Check that `knowledge_entries.json` exists in the project folder
→ Make sure the bot process has write access to that directory

**User names showing as Slack IDs instead of names**
→ Make sure the `users:read` OAuth scope is added to your Slack app

**Bot not responding to @mentions**
→ Verify `app_mention` is in Event Subscriptions
→ Check that the bot is installed to the workspace
→ Make sure Socket Mode is enabled with a valid `SLACK_APP_TOKEN`

---

**One codebase. Excel + Slack + Knowledge + Analytics. Built for the Court Research team.**
