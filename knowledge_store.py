"""
CBC Atlas Knowledge Store

A persistent, git-backed knowledge base for storing and retrieving
institutional insights about courts, judges, and jurisdictions.

Knowledge is stored as a plain JSON file (knowledge_entries.json) so it is:
  - Human-readable and editable
  - Version-controlled in git
  - Zero-dependency (no vector DB required)

Retrieval uses keyword pre-filtering to surface the most relevant entries;
the calling LLM (Claude) then reasons over the results to pick the best ones.
"""

import json
import logging
import subprocess
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

logger = logging.getLogger("atlas-bot")

VALID_ENTITY_TYPES = {"county", "judge", "state", "insurance", "general"}

# Maximum entries returned to Claude per search (keeps token cost reasonable)
SEARCH_RESULT_LIMIT = 30


class KnowledgeStore:
    """
    Persistent knowledge base for CBC Atlas institutional insights.

    Each entry has:
        id           - 8-char UUID prefix (e.g. "a1b2c3d4")
        content      - The knowledge text
        author       - Who contributed it (Slack display name)
        entity_type  - "county" | "judge" | "state" | "insurance" | "general"
        entity_name  - Specific entity, e.g. "Judge Williams" or "Tarrant County"
        state        - Two-letter state code, e.g. "TX" (empty string if n/a)
        timestamp    - ISO-8601 UTC timestamp
    """

    FILE_NAME = "knowledge_entries.json"

    def __init__(self, repo_dir: str = "."):
        self.repo_dir = Path(repo_dir)
        self.file_path = self.repo_dir / self.FILE_NAME
        self._entries: list = []
        self._load()
        logger.info(
            f"KnowledgeStore ready: {len(self._entries)} entries in {self.file_path}"
        )

    # ── Persistence ────────────────────────────────────────────────

    def _load(self):
        """Load entries from JSON file, creating it if absent."""
        if self.file_path.exists():
            try:
                with open(self.file_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, list):
                    self._entries = data
                else:
                    logger.warning("knowledge_entries.json has unexpected format; resetting")
                    self._entries = []
            except (json.JSONDecodeError, OSError) as e:
                logger.error(f"Failed to load knowledge entries: {e}")
                self._entries = []
        else:
            self._entries = []

    def _save(self):
        """Write current entries to JSON file."""
        try:
            with open(self.file_path, "w", encoding="utf-8") as f:
                json.dump(self._entries, f, indent=2, ensure_ascii=False)
        except OSError as e:
            logger.error(f"Failed to save knowledge entries: {e}")
            raise

    def _git_commit_and_push(self, message: str):
        """Stage knowledge_entries.json, commit, and push with exponential backoff."""
        try:
            subprocess.run(
                ["git", "add", self.FILE_NAME],
                cwd=str(self.repo_dir),
                capture_output=True,
                text=True,
                check=True,
            )

            # Check for staged changes
            status = subprocess.run(
                ["git", "diff", "--cached", "--quiet"],
                cwd=str(self.repo_dir),
                capture_output=True,
                text=True,
            )
            if status.returncode == 0:
                logger.info("Git: no staged changes for knowledge_entries.json")
                return

            subprocess.run(
                ["git", "commit", "-m", message],
                cwd=str(self.repo_dir),
                capture_output=True,
                text=True,
                check=True,
            )
            logger.info(f"Git: committed — {message}")

            for attempt, wait in enumerate([0, 2, 4, 8, 16], start=1):
                if wait:
                    time.sleep(wait)
                result = subprocess.run(
                    ["git", "push"],
                    cwd=str(self.repo_dir),
                    capture_output=True,
                    text=True,
                )
                if result.returncode == 0:
                    logger.info(f"Git: pushed (attempt {attempt})")
                    return
                logger.warning(
                    f"Git push attempt {attempt} failed: {result.stderr.strip()}"
                )

            logger.error("Git: push failed after all retries")

        except subprocess.CalledProcessError as e:
            logger.error(f"Git error: {e.stderr.strip() if e.stderr else e}")
        except Exception as e:
            logger.error(f"Git commit/push error: {e}")

    # ── Write ──────────────────────────────────────────────────────

    def add_knowledge(
        self,
        content: str,
        author: str,
        entity_type: str = "general",
        entity_name: str = "",
        state: str = "",
    ) -> str:
        """
        Store a knowledge snippet.

        Returns the short ID assigned to the entry (e.g. "a1b2c3d4").
        The file is saved and committed to git automatically.
        """
        if entity_type not in VALID_ENTITY_TYPES:
            logger.warning(f"Unknown entity_type '{entity_type}'; defaulting to 'general'")
            entity_type = "general"

        knowledge_id = str(uuid.uuid4())[:8]
        entry = {
            "id": knowledge_id,
            "content": content.strip(),
            "author": author.strip() or "Team Member",
            "entity_type": entity_type,
            "entity_name": entity_name.strip(),
            "state": state.upper().strip() if state else "",
            "timestamp": datetime.now(timezone.utc).isoformat(),
        }
        self._entries.append(entry)
        self._save()
        self._git_commit_and_push(
            f"Add knowledge entry [{knowledge_id}] via CarterBot"
        )
        logger.info(f"Knowledge added [{knowledge_id}] by {author}: {content[:80]}")
        return knowledge_id

    def delete_knowledge(self, knowledge_id: str) -> bool:
        """
        Remove a knowledge entry by its short ID.

        Returns True if found and deleted, False otherwise.
        """
        before = len(self._entries)
        self._entries = [e for e in self._entries if e["id"] != knowledge_id]
        if len(self._entries) < before:
            self._save()
            self._git_commit_and_push(
                f"Delete knowledge entry [{knowledge_id}] via CarterBot"
            )
            logger.info(f"Knowledge deleted [{knowledge_id}]")
            return True
        logger.warning(f"Knowledge entry [{knowledge_id}] not found for deletion")
        return False

    # ── Read ───────────────────────────────────────────────────────

    def search_knowledge(
        self,
        query: str,
        entity_type: Optional[str] = None,
        state: Optional[str] = None,
        limit: int = SEARCH_RESULT_LIMIT,
    ) -> list:
        """
        Return the most relevant knowledge entries for a query.

        Filters by entity_type and state when provided, then scores
        entries by keyword overlap and sorts by (score DESC, recency DESC).
        The calling LLM reasons over the returned list to pick the best fits.
        """
        results = list(self._entries)

        if entity_type:
            results = [e for e in results if e["entity_type"] == entity_type]
        if state:
            results = [e for e in results if e["state"] == state.upper().strip()]

        # Keyword relevance scoring
        query_words = set(query.lower().split())

        def relevance_score(entry: dict) -> int:
            text = " ".join([
                entry.get("content", ""),
                entry.get("entity_name", ""),
                entry.get("state", ""),
                entry.get("entity_type", ""),
            ]).lower()
            return sum(1 for w in query_words if w in text)

        results.sort(
            key=lambda e: (relevance_score(e), e.get("timestamp", "")),
            reverse=True,
        )
        return results[:limit]

    def get_by_entity(
        self,
        entity_type: str,
        entity_name: str,
        state: Optional[str] = None,
    ) -> list:
        """
        Retrieve all knowledge entries linked to a specific entity.

        Matching is case-insensitive on entity_name.
        """
        name_lower = entity_name.strip().lower()
        results = [
            e for e in self._entries
            if e["entity_type"] == entity_type
            and e["entity_name"].lower() == name_lower
        ]
        if state:
            results = [e for e in results if e["state"] == state.upper().strip()]
        # Sort newest first
        results.sort(key=lambda e: e.get("timestamp", ""), reverse=True)
        return results

    def list_recent(self, limit: int = 10) -> list:
        """Return the most recently added knowledge entries."""
        return sorted(
            self._entries,
            key=lambda e: e.get("timestamp", ""),
            reverse=True,
        )[:limit]

    def get_stats(self) -> dict:
        """Return total count and breakdown by entity_type."""
        total = len(self._entries)
        by_type: dict = {}
        for entry in self._entries:
            t = entry.get("entity_type", "general")
            by_type[t] = by_type.get(t, 0) + 1
        return {"total": total, "by_type": by_type}

    # ── Formatting ─────────────────────────────────────────────────

    @staticmethod
    def format_for_slack(items: list) -> str:
        """Format a list of knowledge entries for Slack display."""
        if not items:
            return "No knowledge entries found."

        lines = []
        for item in items:
            etype = item.get("entity_type", "general")
            ename = item.get("entity_name", "")
            state = item.get("state", "")
            ts = (item.get("timestamp") or "")[:10] or "Unknown date"
            kid = item.get("id", "?")

            if ename and state and etype != "general":
                tag = f"[{etype.title()}: {ename}, {state}]"
            elif ename and etype != "general":
                tag = f"[{etype.title()}: {ename}]"
            elif etype != "general":
                tag = f"[{etype.title()}: {state}]" if state else f"[{etype.title()}]"
            else:
                tag = "[General]"

            lines.append(f"• *[{kid}]* {item['content']}  {tag}")
            lines.append(f"  _by {item.get('author', 'Unknown')} on {ts}_")

        return "\n".join(lines)

    @staticmethod
    def format_as_context(items: list) -> str:
        """
        Format knowledge entries as a compact context block for Claude.

        Used when injecting knowledge into the assistant's reasoning context
        rather than displaying it directly to the user.
        """
        if not items:
            return "(no relevant knowledge entries)"

        lines = []
        for item in items:
            ename = item.get("entity_name", "")
            state = item.get("state", "")
            ts = (item.get("timestamp") or "")[:10]
            entity_tag = ""
            if ename:
                entity_tag = f" [{ename}{', ' + state if state else ''}]"
            lines.append(f"- [{item['id']}]{entity_tag} {item['content']} (by {item.get('author', '?')} on {ts})")

        return "\n".join(lines)
