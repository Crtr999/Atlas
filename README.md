# CBC Assistant Generator

Complete tool for generating the CBC Court Guide from Excel data.

## 📦 What's Included

- **CBC_Settlement_Funding_Master_v4.xlsx** - Master data file
- **CBCAssistant_Complete.html** - Current working application
- **generate_cbc_assistant.py** - ONE-STEP generator script
- **CBCAssistant_v4.html** - Original template (needed for regeneration)

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
pip install pandas openpyxl
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

## 📞 Using with Claude Code

In terminal:
```bash
cd /path/to/this/folder
claude
```

Then say: "Regenerate the CBC Assistant" or "Add a new case for..."

---

**One script. One command. Complete application.**
