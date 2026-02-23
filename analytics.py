"""
CBC Atlas Analytics Engine

Provides risk scoring and analytical queries over the historical case data
stored in the Excel workbook.  All methods return Slack-formatted strings
ready to send directly to the user.

Designed to work on top of an AtlasDataReader instance so it always reflects
the current state of the Excel file without needing a separate load step.
"""

import logging
from typing import Optional

import pandas as pd

logger = logging.getLogger("atlas-bot")

# Minimum cases needed to report a "reliable" risk score
MIN_RELIABLE_SAMPLE = 5


def _normalize_status(series: pd.Series) -> pd.Series:
    """
    Map raw Status strings to canonical values.

    Returns a Series with values: "Approved" | "Denied" | "Dismissed" | "Other"
    """
    s = series.fillna("").astype(str).str.lower()
    result = pd.Series("Other", index=series.index)
    result[s.str.contains("approved|granted")] = "Approved"
    result[s.str.contains("denied")] = "Denied"
    result[s.str.contains("dismiss")] = "Dismissed"
    return result


class AtlasAnalytics:
    """
    Analytical queries and risk scoring for CBC Atlas case data.

    All public methods accept optional filter arguments and return
    pre-formatted Slack text strings.
    """

    def __init__(self, data_reader):
        """
        Args:
            data_reader: An AtlasDataReader instance.  The analytics engine
                         reads directly from data_reader.df_cases (and other
                         dataframes) so it always has fresh data.
        """
        self.reader = data_reader

    # ── Public API ─────────────────────────────────────────────────

    def get_risk_score(
        self,
        county: Optional[str] = None,
        state: Optional[str] = None,
        judge: Optional[str] = None,
    ) -> str:
        """
        Return a Slack-formatted risk assessment for a county / state / judge
        combination based on historical case outcomes.

        At least one of county, state, or judge should be provided.
        If none are provided, the score covers all cases in the dataset.
        """
        df = self._filter_cases(county=county, state=state, judge=judge)
        label = self._build_label(county=county, state=state, judge=judge)

        if df.empty:
            return f"No case history found for {label}."

        total = len(df)
        statuses = _normalize_status(df["Status"])
        approved = int((statuses == "Approved").sum())
        denied = int((statuses == "Denied").sum())
        dismissed = int((statuses == "Dismissed").sum())
        other = total - approved - denied - dismissed

        approval_rate = approved / total
        denial_rate = denied / total
        risk_tier = self._risk_tier(approval_rate)

        lines = [
            f"*Risk Score — {label}*",
            f"  Risk Level: *{risk_tier}*",
            f"  Total Cases: {total}",
            f"  Approved:  {approved} ({approval_rate:.0%})",
            f"  Denied:    {denied} ({denial_rate:.0%})",
            f"  Dismissed: {dismissed} ({dismissed / total:.0%})",
        ]
        if other > 0:
            lines.append(f"  Other/Unknown: {other}")
        if total < MIN_RELIABLE_SAMPLE:
            lines.append(
                f"  ⚠️  Small sample ({total} case{'s' if total != 1 else ''}) "
                "— interpret with caution"
            )
        return "\n".join(lines)

    def get_denial_patterns(self, top_n: int = 10) -> str:
        """
        Identify the counties and judges with the highest denial rates.

        Only includes entities with at least 3 cases to filter out noise.
        """
        df = self.reader.df_cases.copy()
        if df.empty:
            return "No case data available."

        df["_status"] = _normalize_status(df["Status"])

        # --- By county ---
        county_rows = []
        for (state, county), grp in df.groupby(["State", "County"]):
            if len(grp) < 3:
                continue
            denied = int((grp["_status"] == "Denied").sum())
            if denied == 0:
                continue
            county_rows.append(
                {
                    "label": f"{county}, {state}",
                    "total": len(grp),
                    "denied": denied,
                    "rate": denied / len(grp),
                }
            )
        county_rows.sort(key=lambda x: x["rate"], reverse=True)

        # --- By judge ---
        judge_rows = []
        for judge, grp in df.groupby("Judge"):
            judge_str = str(judge).strip()
            if not judge_str or judge_str.lower() == "nan" or len(grp) < 3:
                continue
            denied = int((grp["_status"] == "Denied").sum())
            if denied == 0:
                continue
            judge_rows.append(
                {
                    "label": f"Judge {judge_str}",
                    "total": len(grp),
                    "denied": denied,
                    "rate": denied / len(grp),
                }
            )
        judge_rows.sort(key=lambda x: x["rate"], reverse=True)

        lines = ["*Denial Patterns (minimum 3 cases)*"]

        if county_rows:
            lines.append("\n*Top Counties by Denial Rate:*")
            for row in county_rows[:top_n]:
                lines.append(
                    f"  • {row['label']}: "
                    f"{row['denied']}/{row['total']} denied "
                    f"({row['rate']:.0%})"
                )
        else:
            lines.append("\nNo counties with 3+ cases and at least one denial found.")

        if judge_rows:
            lines.append("\n*Top Judges by Denial Rate:*")
            for row in judge_rows[:top_n]:
                lines.append(
                    f"  • {row['label']}: "
                    f"{row['denied']}/{row['total']} denied "
                    f"({row['rate']:.0%})"
                )
        else:
            lines.append("\nNo judges with 3+ cases and at least one denial found.")

        return "\n".join(lines)

    def get_judge_analytics(
        self,
        judge_name: Optional[str] = None,
        state: Optional[str] = None,
    ) -> str:
        """
        Return a per-judge performance breakdown.

        If judge_name is provided, only that judge (fuzzy match) is shown.
        If state is provided, only cases from that state are included.
        """
        df = self.reader.df_cases.copy()
        if df.empty:
            return "No case data available."

        if state:
            df = df[df["State"].str.upper() == state.upper()]
        if judge_name:
            df = df[df["Judge"].str.lower().str.contains(judge_name.lower(), na=False)]

        if df.empty:
            parts = []
            if judge_name:
                parts.append(judge_name)
            if state:
                parts.append(state.upper())
            return "No cases found for " + (", ".join(parts) if parts else "the given filters") + "."

        df["_status"] = _normalize_status(df["Status"])

        rows = []
        for judge, grp in df.groupby("Judge"):
            judge_str = str(judge).strip()
            if not judge_str or judge_str.lower() == "nan":
                continue
            total = len(grp)
            approved = int((grp["_status"] == "Approved").sum())
            denied = int((grp["_status"] == "Denied").sum())
            dismissed = int((grp["_status"] == "Dismissed").sum())
            states = ", ".join(sorted(grp["State"].dropna().astype(str).unique()))
            counties = ", ".join(sorted(grp["County"].dropna().astype(str).unique()))
            rows.append(
                {
                    "judge": judge_str,
                    "total": total,
                    "approved": approved,
                    "denied": denied,
                    "dismissed": dismissed,
                    "approval_rate": approved / total,
                    "states": states,
                    "counties": counties,
                }
            )

        rows.sort(key=lambda x: x["total"], reverse=True)

        if judge_name:
            header = judge_name
        elif state:
            header = f"Judges in {state.upper()}"
        else:
            header = "All Judges"

        lines = [f"*Judge Analytics — {header}*"]
        for r in rows:
            lines.append(f"\n  *{r['judge']}*  ({r['states']})")
            lines.append(
                f"    {r['total']} cases | "
                f"Approved: {r['approved']} ({r['approval_rate']:.0%}) | "
                f"Denied: {r['denied']} | "
                f"Dismissed: {r['dismissed']}"
            )
            if r["counties"]:
                lines.append(f"    Counties: {r['counties']}")

        return "\n".join(lines)

    def get_state_analytics(self, state_code: Optional[str] = None) -> str:
        """
        Return case outcome statistics broken down by state.

        If state_code is provided, only that state is shown (with extra detail).
        """
        df = self.reader.df_cases.copy()
        if df.empty:
            return "No case data available."

        if state_code:
            df = df[df["State"].str.upper() == state_code.upper()]
            if df.empty:
                return f"No cases found for state {state_code.upper()}."

        df["_status"] = _normalize_status(df["Status"])

        rows = []
        for state, grp in df.groupby("State"):
            total = len(grp)
            approved = int((grp["_status"] == "Approved").sum())
            denied = int((grp["_status"] == "Denied").sum())
            dismissed = int((grp["_status"] == "Dismissed").sum())
            rows.append(
                {
                    "state": str(state),
                    "total": total,
                    "approved": approved,
                    "denied": denied,
                    "dismissed": dismissed,
                    "approval_rate": approved / total if total else 0,
                }
            )

        rows.sort(key=lambda x: x["total"], reverse=True)

        header = state_code.upper() if state_code else "All States"
        lines = [f"*State Analytics — {header}*"]
        for r in rows:
            lines.append(
                f"  *{r['state']}*: {r['total']} cases | "
                f"Approved: {r['approved']} ({r['approval_rate']:.0%}) | "
                f"Denied: {r['denied']} | "
                f"Dismissed: {r['dismissed']}"
            )

        return "\n".join(lines)

    def get_portfolio_summary(self) -> str:
        """
        Return a high-level snapshot of the full case portfolio.

        Includes totals, outcome rates, top states by volume, and
        the most active judges.
        """
        df = self.reader.df_cases.copy()
        if df.empty:
            return "No case data available."

        df["_status"] = _normalize_status(df["Status"])
        total = len(df)
        approved = int((df["_status"] == "Approved").sum())
        denied = int((df["_status"] == "Denied").sum())
        dismissed = int((df["_status"] == "Dismissed").sum())
        other = total - approved - denied - dismissed

        state_counts = df.groupby("State").size().sort_values(ascending=False)
        judge_counts = (
            df[df["Judge"].notna()]
            .groupby("Judge")
            .size()
            .sort_values(ascending=False)
        )

        lines = [
            "*CBC Settlement Funding — Portfolio Summary*",
            f"  Total Cases:  {total}",
            f"  Approved:     {approved} ({approved / total:.0%})",
            f"  Denied:       {denied} ({denied / total:.0%})",
            f"  Dismissed:    {dismissed} ({dismissed / total:.0%})",
        ]
        if other > 0:
            lines.append(f"  Other/Unknown: {other}")

        lines.append("\n  *States by Case Volume (top 5):*")
        for state, count in state_counts.head(5).items():
            lines.append(f"    • {state}: {count} cases")

        lines.append("\n  *Most Active Judges (top 5):*")
        shown = 0
        for judge, count in judge_counts.head(10).items():
            judge_str = str(judge).strip()
            if not judge_str or judge_str.lower() == "nan":
                continue
            lines.append(f"    • {judge_str}: {count} cases")
            shown += 1
            if shown >= 5:
                break

        return "\n".join(lines)

    def run_analytics_query(
        self,
        query_type: str,
        state: Optional[str] = None,
        judge: Optional[str] = None,
        status: Optional[str] = None,
    ) -> str:
        """
        Dispatch an analytics query by type string.

        query_type options:
            "denial_patterns"   — Counties and judges with highest denial rates
            "judge_analytics"   — Per-judge outcome breakdown
            "state_analytics"   — Per-state outcome breakdown
            "portfolio_summary" — Full portfolio health overview
            "cases_by_status"   — List cases filtered by status (and optionally state)
        """
        dispatch = {
            "denial_patterns": lambda: self.get_denial_patterns(),
            "judge_analytics": lambda: self.get_judge_analytics(
                judge_name=judge, state=state
            ),
            "state_analytics": lambda: self.get_state_analytics(state_code=state),
            "portfolio_summary": lambda: self.get_portfolio_summary(),
            "cases_by_status": lambda: self._cases_by_status(
                status=status, state=state
            ),
        }

        if query_type not in dispatch:
            available = ", ".join(sorted(dispatch.keys()))
            return (
                f"Unknown query type '{query_type}'. "
                f"Available options: {available}"
            )

        return dispatch[query_type]()

    # ── Internal helpers ───────────────────────────────────────────

    def _filter_cases(
        self,
        county: Optional[str],
        state: Optional[str],
        judge: Optional[str],
    ) -> pd.DataFrame:
        df = self.reader.df_cases.copy()
        if state:
            df = df[df["State"].str.upper() == state.upper()]
        if county:
            county_lower = county.lower().replace(" county", "").strip()
            df = df[
                df["County"]
                .str.lower()
                .str.replace(" county", "", regex=False)
                .str.strip()
                == county_lower
            ]
        if judge:
            df = df[df["Judge"].str.lower().str.contains(judge.lower(), na=False)]
        return df

    @staticmethod
    def _build_label(
        county: Optional[str],
        state: Optional[str],
        judge: Optional[str],
    ) -> str:
        parts = []
        if judge:
            parts.append(f"Judge {judge}")
        if county:
            parts.append(county)
        if state:
            parts.append(state.upper())
        return ", ".join(parts) if parts else "All Cases"

    @staticmethod
    def _risk_tier(approval_rate: float) -> str:
        if approval_rate >= 0.70:
            return "LOW RISK ✅"
        elif approval_rate >= 0.50:
            return "MODERATE RISK 🟡"
        elif approval_rate >= 0.30:
            return "HIGH RISK 🔴"
        else:
            return "VERY HIGH RISK 🚫"

    def _cases_by_status(
        self,
        status: Optional[str] = None,
        state: Optional[str] = None,
    ) -> str:
        df = self.reader.df_cases.copy()
        df["_status"] = _normalize_status(df["Status"])

        if state:
            df = df[df["State"].str.upper() == state.upper()]
        if status:
            df = df[df["_status"].str.lower() == status.lower()]

        label_parts = []
        if status:
            label_parts.append(status)
        if state:
            label_parts.append(state.upper())
        label_str = " — " + ", ".join(label_parts) if label_parts else ""

        if df.empty:
            return f"No cases found{label_str}."

        lines = [f"*Cases{label_str}*", f"Total: {len(df)}"]
        for _, row in df.iterrows():
            judge = str(row.get("Judge", "") or "Unknown")
            if judge.lower() == "nan":
                judge = "Unknown"
            client = str(row.get("Client_Name", "") or "")
            if client.lower() == "nan":
                client = ""
            case_num = str(row.get("Case_Number", "") or "")
            if case_num.lower() == "nan":
                case_num = ""
            st = str(row.get("State", "") or "")
            county = str(row.get("County", "") or "")
            lines.append(
                f"  • {county}, {st} | Judge: {judge} | "
                f"Client: {client} | Case#: {case_num} | Status: {row['_status']}"
            )

        return "\n".join(lines)
