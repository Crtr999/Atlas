# CBC Assistant Generator

Complete tool for generating the CBC Court Guide from Excel data.

## 📦 What's Included

- **CBC_Settlement_Funding_Master_v4.xlsx** - Master data file
- **CBCAssistant_Complete.html** - Current working application
- **generate_cbc_assistant.py** - ONE-STEP generator script
- **CBCAssistant_v4.html** - Original template (needed for regeneration)
- **slack_bot.py** - Slack bot powered by Claude (query data, edit Excel, regenerate HTML)
- **data_reader.py** - Excel data reader module used by the Slack bot
- **.env.example** - Template for environment variables

## 🚀 How to Regenerate

When you update the Excel file and want to regenerate the HTML:

```bash
python3 generate_cbc_assistant.py
```

That's it! One command does everything:
- ✓ Reads all Excel sheets
- ✓ Processes cases and jurisdictions
- ✓ Adds search bar, approvals, sorting
- ✓ Outputs complete, ready-to-use HTML

## 📝 Updating Data

### Add a New Case
1. Open `CBC_Settlement_Funding_Master_v4.xlsx`
2. Go to **Cases** sheet
3. Add row with: State, County, Judge, Client_Name, Case_Number, Court_Date, Status, Notes
4. Status must contain: "Denied", "Dismissed", or "Approved"
5. Run: `python3 generate_cbc_assistant.py`

### Add County Notes
1. Open `CBC_Settlement_Funding_Master_v4.xlsx`
2. Go to **Jurisdictions** sheet
3. Find the county (e.g., "Ramsey County") or add new row
4. Enter notes in **County_Notes** column
5. Run: `python3 generate_cbc_assistant.py`

**Note:** County names use fuzzy matching, so "Ramsey" in Atlas will match "Ramsey County" in Jurisdictions.

### Update Court Access Info
1. Open `CBC_Settlement_Funding_Master_v4.xlsx`
2. Go to **Court Access** sheet
3. Update: Westlaw_Coverage, Website, Fee_Structure, etc.
4. Run: `python3 generate_cbc_assistant.py`

## ✨ Features in Generated HTML

- 🔍 Search all 3,140 counties with dropdown results
- ✅ Approvals, Dismissals, and Denials tracked
- 📊 Sorted by case count (most active first)
- 📝 Jurisdiction-specific notes displayed
- 🌐 Court access info with Westlaw badges
- 🗺️ State filtering
- 📦 Expandable county and judge cards
- 🏷️ Case details with status badges

## 🛠️ Requirements

```bash
pip install -r requirements.txt
```

Or manually:
```bash
pip install pandas openpyxl slack-bolt slack-sdk anthropic python-dotenv
```

## 💡 Tips

**Testing changes:** After regenerating, open `CBCAssistant_Complete.html` in your browser to test.

**Backup:** The Excel file contains all your data - back it up regularly!

**Template needed:** Keep `CBCAssistant_v4.html` - it's the base template the script uses.

**County names:** If notes don't appear, check that county names match between sheets. The script handles "County" suffix automatically.

## 🐛 Troubleshooting

**"Template not found"**
→ Make sure `CBCAssistant_v4.html` is in the same folder

**"Excel file not found"**
→ Make sure `CBC_Settlement_Funding_Master_v4.xlsx` is in the same folder

**"No module named pandas"**
→ Run: `pip install pandas openpyxl`

**County notes not appearing**
→ Check spelling of county names in Jurisdictions sheet
→ Make sure County_Notes column isn't empty (not just spaces)

**JavaScript errors in browser**
→ Check browser console (F12) for specific error
→ Make sure you regenerated after updating Excel

## 💬 Slack Bot

Talk to `@Claude` in Slack to query data, edit the Excel file, and regenerate the HTML — all from a chat message.

### What You Can Ask

- `@Claude What are the redaction levels in Ramsey, MN?`
- `@Claude Give me feedback on our court process in Minnesota`
- `@Claude How many cases have been approved in Hennepin County?`
- `@Claude Add a new case for John Doe in Ramsey County, MN — Approved by Judge Smith, case #12345`
- `@Claude Update the county notes for Ramsey, MN to "New courthouse procedures effective 2025"`
- `@Claude Regenerate the HTML`

### Slack App Setup

1. Go to [api.slack.com/apps](https://api.slack.com/apps) and click **Create New App** > **From scratch**
2. Name it (e.g., "CBC Atlas Bot") and pick your workspace

3. **Enable Socket Mode:**
   - Go to **Socket Mode** in the sidebar, toggle it on
   - Generate an **App-Level Token** with `connections:write` scope — this is your `SLACK_APP_TOKEN` (starts with `xapp-`)

4. **Set Bot Permissions:**
   - Go to **OAuth & Permissions** > **Bot Token Scopes** and add:
     - `app_mentions:read`
     - `chat:write`
     - `im:history`
     - `im:read`
     - `im:write`

5. **Enable Events:**
   - Go to **Event Subscriptions**, toggle on
   - Under **Subscribe to bot events**, add:
     - `app_mention`
     - `message.im`

6. **Install to Workspace:**
   - Go to **Install App** and install
   - Copy the **Bot User OAuth Token** — this is your `SLACK_BOT_TOKEN` (starts with `xoxb-`)

7. **Get Anthropic API Key:**
   - Go to [console.anthropic.com](https://console.anthropic.com/) and create an API key

### Running the Bot

```bash
# Install dependencies
pip install -r requirements.txt

# Copy and fill in environment variables
cp .env.example .env
# Edit .env with your tokens

# Start the bot
python slack_bot.py
```

The bot connects via Socket Mode (no public URL needed) and listens for @mentions and DMs.

### How It Works

1. You message `@Claude` in Slack with a question
2. The bot sends your question to Claude along with tool definitions for reading/editing the Excel data
3. Claude decides which tools to call (query county, query state, add case, etc.)
4. The bot executes the tool calls against the Excel file and returns results to Claude
5. Claude formulates a response and sends it back to Slack
6. If any edits were made, the HTML is automatically regenerated

## 📞 Using with Claude Code

In terminal:
```bash
cd /path/to/this/folder
claude
```

Then say: "Regenerate the CBC Assistant" or "Add a new case for..."

---

**One script. One command. Complete application. Now with Slack.**
