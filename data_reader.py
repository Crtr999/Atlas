"""
CBC Atlas Data Reader

Reads the Excel workbook and provides structured access to all sheets.
Used by the Slack bot to answer questions and make edits.
"""

import pandas as pd
import json
import logging
import subprocess
import sys
import time
from pathlib import Path

logger = logging.getLogger("atlas-bot")


class AtlasDataReader:
    """Reads and queries the CBC Settlement Funding Excel workbook."""

    def __init__(self, excel_path="CBC_Settlement_Funding_Master_v4.xlsx"):
        self.excel_path = Path(excel_path)
        if not self.excel_path.exists():
            raise FileNotFoundError(f"Excel file not found: {self.excel_path}")
        self.reload()

    def reload(self):
        """Reload all data from the Excel file."""
        self.df_cases = pd.read_excel(self.excel_path, sheet_name="Cases")
        self.df_jurisdictions = pd.read_excel(self.excel_path, sheet_name="Jurisdictions")
        self.df_state_data = pd.read_excel(self.excel_path, sheet_name="State Data")
        self.df_insurance = pd.read_excel(self.excel_path, sheet_name="Insurance Carriers")
        self.df_court_access = pd.read_excel(self.excel_path, sheet_name="Court Access")

    # ── Query methods ──────────────────────────────────────────────

    def get_all_data_summary(self) -> str:
        """Return a compact summary of all data for Claude context."""
        parts = []

        # Jurisdictions with redaction info
        parts.append("=== JURISDICTIONS (with redaction levels) ===")
        for _, row in self.df_jurisdictions.iterrows():
            state = row.get("State", "")
            county = row.get("County", "")
            redaction = row.get("Redaction_Level", "")
            notes = row.get("Redaction_Notes", "")
            county_notes = row.get("County_Notes", "")
            if pd.isna(redaction):
                redaction = ""
            if pd.isna(notes):
                notes = ""
            if pd.isna(county_notes):
                county_notes = ""
            line = f"  {state} - {county}: Redaction={redaction}"
            if notes:
                line += f" | Notes={notes}"
            if county_notes:
                line += f" | County Notes={county_notes}"
            parts.append(line)

        # Cases
        parts.append("\n=== CASES ===")
        for _, row in self.df_cases.iterrows():
            state = row.get("State", "")
            county = row.get("County", "")
            judge = row.get("Judge", "")
            status = row.get("Status", "")
            client = row.get("Client_Name", "")
            case_num = row.get("Case_Number", "")
            court_date = row.get("Court_Date", "")
            case_notes = row.get("Notes", "")
            if pd.isna(judge):
                judge = "Unknown"
            if pd.isna(status):
                status = ""
            if pd.isna(client):
                client = ""
            if pd.isna(case_num):
                case_num = ""
            if pd.isna(court_date):
                court_date = ""
            if pd.isna(case_notes):
                case_notes = ""
            parts.append(
                f"  {state} {county} | Judge: {judge} | Status: {status} "
                f"| Client: {client} | Case#: {case_num} | Date: {court_date}"
                + (f" | Notes: {case_notes}" if case_notes else "")
            )

        # State Data
        parts.append("\n=== STATE DATA ===")
        for _, row in self.df_state_data.iterrows():
            state = row.get("State", "")
            name = row.get("State_Name", "")
            rate_cap = row.get("Rate_Cap", "")
            ipa = row.get("Requires_IPA", "")
            aff = row.get("Requires_Aff_Dec", "")
            no_poach = row.get("No_Poach_State", "")
            counsel = row.get("Legal_Counsel", "")
            fee = row.get("Expected_Legal_Fee", "")
            add_notes = row.get("Additional_Notes", "")
            parts.append(
                f"  {state} ({name}): Rate Cap={rate_cap}, IPA={ipa}, "
                f"Aff/Dec={aff}, No Poach={no_poach}, Counsel={counsel}, "
                f"Legal Fee={fee}"
                + (f", Notes={add_notes}" if pd.notna(add_notes) and add_notes else "")
            )

        # Insurance Carriers
        parts.append("\n=== INSURANCE CARRIERS ===")
        for _, row in self.df_insurance.iterrows():
            carrier = row.get("Carrier_Name", "")
            admin_fee = row.get("Admin_Fee", "")
            contact = row.get("Contact", "")
            ins_notes = row.get("Notes", "")
            if pd.isna(admin_fee):
                admin_fee = ""
            if pd.isna(contact):
                contact = ""
            if pd.isna(ins_notes):
                ins_notes = ""
            parts.append(
                f"  {carrier}: Admin Fee={admin_fee}, Contact={contact}"
                + (f", Notes={ins_notes}" if ins_notes else "")
            )

        # Court Access (compact - just counties with notable info)
        parts.append("\n=== COURT ACCESS ===")
        for _, row in self.df_court_access.iterrows():
            state = row.get("State", "")
            county = row.get("County", "")
            westlaw = row.get("Westlaw_Coverage", "")
            website = row.get("Website", "")
            fees = row.get("Fee_Structure", "")
            sub_login = row.get("Subscription_Login", "")
            search_notes = row.get("Search_Notes", "")
            court_type = row.get("Court_Type", "")
            if pd.isna(westlaw):
                westlaw = ""
            if pd.isna(website):
                website = ""
            if pd.isna(fees):
                fees = ""
            if pd.isna(sub_login):
                sub_login = ""
            if pd.isna(search_notes):
                search_notes = ""
            if pd.isna(court_type):
                court_type = ""
            line = f"  {state} - {county}:"
            info_parts = []
            if westlaw:
                info_parts.append(f"Westlaw={westlaw}")
            if website:
                info_parts.append(f"Website={website}")
            if fees:
                info_parts.append(f"Fees={fees}")
            if sub_login:
                info_parts.append(f"Login={sub_login}")
            if search_notes:
                info_parts.append(f"Search={search_notes}")
            if court_type:
                info_parts.append(f"Type={court_type}")
            if info_parts:
                line += " " + ", ".join(info_parts)
            parts.append(line)

        return "\n".join(parts)

    def query_county(self, county_name: str, state_code: str = None) -> str:
        """Get all information about a specific county."""
        county_lower = county_name.lower().replace(" county", "").strip()
        parts = []

        # Search jurisdictions
        for _, row in self.df_jurisdictions.iterrows():
            row_county = str(row.get("County", "")).lower().replace(" county", "").strip()
            row_state = str(row.get("State", ""))
            if county_lower in row_county or row_county in county_lower:
                if state_code and row_state.upper() != state_code.upper():
                    continue
                parts.append(f"Jurisdiction: {row_state} - {row.get('County')}")
                redaction = row.get("Redaction_Level", "")
                if pd.notna(redaction) and redaction:
                    parts.append(f"  Redaction Level: {redaction}")
                red_notes = row.get("Redaction_Notes", "")
                if pd.notna(red_notes) and red_notes:
                    parts.append(f"  Redaction Notes: {red_notes}")
                county_notes = row.get("County_Notes", "")
                if pd.notna(county_notes) and county_notes:
                    parts.append(f"  County Notes: {county_notes}")

        # Search court access
        for _, row in self.df_court_access.iterrows():
            row_county = str(row.get("County", "")).lower().replace(" county", "").strip()
            row_state = str(row.get("State", ""))
            if county_lower in row_county or row_county in county_lower:
                if state_code and row_state.upper() != state_code.upper():
                    continue
                parts.append(f"\nCourt Access: {row_state} - {row.get('County')}")
                for col in ["Westlaw_Coverage", "Website", "Fee_Structure",
                            "Subscription_Login", "Search_Notes", "Court_Type"]:
                    val = row.get(col, "")
                    if pd.notna(val) and val:
                        parts.append(f"  {col}: {val}")

        # Search cases
        matching_cases = []
        for _, row in self.df_cases.iterrows():
            row_county = str(row.get("County", "")).lower().replace(" county", "").strip()
            row_state = str(row.get("State", ""))
            if county_lower in row_county or row_county in county_lower:
                if state_code and row_state.upper() != state_code.upper():
                    continue
                matching_cases.append(row)

        if matching_cases:
            parts.append(f"\nCases ({len(matching_cases)} total):")
            for row in matching_cases:
                judge = row.get("Judge", "Unknown")
                if pd.isna(judge):
                    judge = "Unknown"
                status = row.get("Status", "")
                if pd.isna(status):
                    status = ""
                client = row.get("Client_Name", "")
                if pd.isna(client):
                    client = ""
                parts.append(f"  Judge: {judge} | Status: {status} | Client: {client}")

        if not parts:
            return f"No data found for county '{county_name}'" + (
                f" in state '{state_code}'" if state_code else ""
            )

        return "\n".join(parts)

    def query_state(self, state_code: str) -> str:
        """Get all information about a specific state."""
        state_upper = state_code.upper().strip()
        parts = []

        # State data
        for _, row in self.df_state_data.iterrows():
            if str(row.get("State", "")).upper() == state_upper:
                parts.append(f"State: {row.get('State_Name', state_upper)} ({state_upper})")
                for col in ["Rate_Cap", "Requires_IPA", "Requires_Aff_Dec",
                            "No_Poach_State", "Legal_Counsel", "Expected_Legal_Fee",
                            "Additional_Notes"]:
                    val = row.get(col, "")
                    if pd.notna(val) and val:
                        parts.append(f"  {col}: {val}")

        # Count cases in this state
        state_cases = self.df_cases[
            self.df_cases["State"].str.upper() == state_upper
        ]
        if len(state_cases) > 0:
            parts.append(f"\nCases in {state_upper}: {len(state_cases)} total")
            approved = len(state_cases[state_cases["Status"].str.lower().str.contains("approved|granted", na=False)])
            dismissed = len(state_cases[state_cases["Status"].str.lower().str.contains("dismiss", na=False)])
            denied = len(state_cases[state_cases["Status"].str.lower().str.contains("denied", na=False)])
            parts.append(f"  Approved: {approved}, Dismissed: {dismissed}, Denied: {denied}")

            # List counties with cases
            counties = state_cases["County"].unique()
            parts.append(f"  Counties with cases: {', '.join(str(c) for c in counties)}")

        # Count jurisdictions
        state_jurisdictions = self.df_court_access[
            self.df_court_access["State"].str.upper() == state_upper
        ]
        if len(state_jurisdictions) > 0:
            parts.append(f"\nCourt Access entries: {len(state_jurisdictions)} counties")

        if not parts:
            return f"No data found for state '{state_code}'"

        return "\n".join(parts)

    # ── Edit methods ───────────────────────────────────────────────

    def add_case(self, state: str, county: str, judge: str, client_name: str,
                 case_number: str, court_date: str, status: str,
                 notes: str = "") -> str:
        """Add a new case to the Cases sheet."""
        new_row = pd.DataFrame([{
            "State": state,
            "County": county,
            "Judge": judge,
            "Client_Name": client_name,
            "Case_Number": case_number,
            "Court_Date": court_date,
            "Status": status,
            "Notes": notes,
        }])
        self.df_cases = pd.concat([self.df_cases, new_row], ignore_index=True)
        self._save_sheet("Cases", self.df_cases)
        return f"Added case: {client_name} in {county}, {state} (Status: {status})"

    def update_county_notes(self, state: str, county: str, notes: str) -> str:
        """Update or add county notes in the Jurisdictions sheet."""
        county_lower = county.lower().replace(" county", "").strip()
        found = False

        for idx, row in self.df_jurisdictions.iterrows():
            row_county = str(row.get("County", "")).lower().replace(" county", "").strip()
            row_state = str(row.get("State", "")).upper()
            if county_lower == row_county and row_state == state.upper():
                self.df_jurisdictions.at[idx, "County_Notes"] = notes
                found = True
                break

        if not found:
            new_row = pd.DataFrame([{
                "State": state.upper(),
                "County": county,
                "County_Notes": notes,
            }])
            self.df_jurisdictions = pd.concat(
                [self.df_jurisdictions, new_row], ignore_index=True
            )

        self._save_sheet("Jurisdictions", self.df_jurisdictions)
        action = "Updated" if found else "Added"
        return f"{action} county notes for {county}, {state}"

    def update_redaction_level(self, state: str, county: str,
                               redaction_level: str,
                               redaction_notes: str = "") -> str:
        """Update or add redaction level for a county."""
        county_lower = county.lower().replace(" county", "").strip()
        found = False

        for idx, row in self.df_jurisdictions.iterrows():
            row_county = str(row.get("County", "")).lower().replace(" county", "").strip()
            row_state = str(row.get("State", "")).upper()
            if county_lower == row_county and row_state == state.upper():
                self.df_jurisdictions.at[idx, "Redaction_Level"] = redaction_level
                if redaction_notes:
                    self.df_jurisdictions.at[idx, "Redaction_Notes"] = redaction_notes
                found = True
                break

        if not found:
            new_row_data = {
                "State": state.upper(),
                "County": county,
                "Redaction_Level": redaction_level,
            }
            if redaction_notes:
                new_row_data["Redaction_Notes"] = redaction_notes
            new_row = pd.DataFrame([new_row_data])
            self.df_jurisdictions = pd.concat(
                [self.df_jurisdictions, new_row], ignore_index=True
            )

        self._save_sheet("Jurisdictions", self.df_jurisdictions)
        action = "Updated" if found else "Added"
        return f"{action} redaction level for {county}, {state}: {redaction_level}"

    def update_case_status(self, state: str, county: str, client_name: str,
                           new_status: str) -> str:
        """Update the status of an existing case."""
        client_lower = client_name.lower().strip()
        for idx, row in self.df_cases.iterrows():
            row_client = str(row.get("Client_Name", "")).lower().strip()
            row_state = str(row.get("State", "")).upper()
            row_county = str(row.get("County", "")).lower().replace(" county", "").strip()
            county_lower = county.lower().replace(" county", "").strip()
            if (client_lower in row_client or row_client in client_lower) \
                    and row_state == state.upper() \
                    and county_lower == row_county:
                old_status = row.get("Status", "")
                self.df_cases.at[idx, "Status"] = new_status
                self._save_sheet("Cases", self.df_cases)
                return (f"Updated case for {client_name} in {county}, {state}: "
                        f"{old_status} -> {new_status}")

        return f"Case not found for {client_name} in {county}, {state}"

    def _save_sheet(self, sheet_name: str, df: pd.DataFrame):
        """Save a single sheet back to the Excel file, preserving other sheets."""
        with pd.ExcelFile(self.excel_path) as xls:
            all_sheets = {}
            for name in xls.sheet_names:
                if name == sheet_name:
                    all_sheets[name] = df
                else:
                    all_sheets[name] = pd.read_excel(xls, sheet_name=name)

        with pd.ExcelWriter(self.excel_path, engine="openpyxl") as writer:
            for name, data in all_sheets.items():
                data.to_excel(writer, sheet_name=name, index=False)

        self.reload()
        self._git_commit_and_push(f"Update {sheet_name} sheet via CarterBot")

    def _git_commit_and_push(self, commit_message: str):
        """Stage the Excel file, commit, and push to GitHub with retry."""
        repo_dir = str(self.excel_path.parent)
        excel_filename = self.excel_path.name

        try:
            # Stage the changed Excel file
            subprocess.run(
                ["git", "add", excel_filename],
                cwd=repo_dir, capture_output=True, text=True, check=True,
            )

            # Also stage the HTML file if it was regenerated
            html_path = self.excel_path.parent / "CBCAssistant_Complete.html"
            if html_path.exists():
                subprocess.run(
                    ["git", "add", html_path.name],
                    cwd=repo_dir, capture_output=True, text=True,
                )

            # Check if there is anything to commit
            status = subprocess.run(
                ["git", "diff", "--cached", "--quiet"],
                cwd=repo_dir, capture_output=True, text=True,
            )
            if status.returncode == 0:
                logger.info("Git: no staged changes to commit")
                return

            # Commit
            subprocess.run(
                ["git", "commit", "-m", commit_message],
                cwd=repo_dir, capture_output=True, text=True, check=True,
            )
            logger.info(f"Git: committed - {commit_message}")

            # Push with retry (exponential backoff)
            for attempt, wait in enumerate([0, 2, 4, 8, 16], start=1):
                if wait:
                    time.sleep(wait)
                push_result = subprocess.run(
                    ["git", "push"],
                    cwd=repo_dir, capture_output=True, text=True,
                )
                if push_result.returncode == 0:
                    logger.info(f"Git: pushed to remote (attempt {attempt})")
                    return
                logger.warning(
                    f"Git push attempt {attempt} failed: {push_result.stderr.strip()}"
                )

            logger.error("Git: push failed after all retries")

        except subprocess.CalledProcessError as e:
            logger.error(f"Git error: {e.stderr.strip() if e.stderr else e}")
        except Exception as e:
            logger.error(f"Git commit/push error: {e}")

    # ── HTML regeneration ──────────────────────────────────────────

    def regenerate_html(self) -> str:
        """Run the generate_cbc_assistant.py script to regenerate HTML."""
        script_path = self.excel_path.parent / "generate_cbc_assistant.py"
        if not script_path.exists():
            return "Error: generate_cbc_assistant.py not found"

        try:
            result = subprocess.run(
                [sys.executable, str(script_path)],
                capture_output=True,
                text=True,
                cwd=str(self.excel_path.parent),
                timeout=60,
            )
            if result.returncode == 0:
                self._git_commit_and_push("Regenerate HTML via CarterBot")
                return "HTML regenerated successfully.\n" + result.stdout[-500:]
            else:
                return f"HTML generation failed:\n{result.stderr[-500:]}"
        except subprocess.TimeoutExpired:
            return "Error: HTML generation timed out"
        except Exception as e:
            return f"Error running generator: {e}"
