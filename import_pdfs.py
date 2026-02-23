#!/usr/bin/env python3
"""
CBC Atlas PDF Importer

Parses the CBC "SSC Funded Files" PDFs and appends the records to the
Cases sheet in the Excel workbook.  All imported records receive
Status = "Approved".

Fields extracted from the PDFs:
    Client_Name, State, Court_Date (= Funded File Date), Status, Notes
    (County, Judge, Case_Number are left blank — not present in the PDFs)

Usage:
    python import_pdfs.py                         # auto-discovers PDF files
    python import_pdfs.py "CBC 1-1000 PDF.pdf" "CBC 1000-2000 PDF.pdf"
    python import_pdfs.py --dry-run               # parse only, no Excel write
"""

import argparse
import re
import sys
import time
import subprocess
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd
from dotenv import load_dotenv
import os

load_dotenv()

EXCEL_PATH = os.environ.get("EXCEL_PATH", "./CBC_Settlement_Funding_Master_v4.xlsx")

# ── Regex anchors ───────────────────────────────────────────────────

DATE_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")
DATE_EMB_RE = re.compile(r"(\d{1,2}/\d{1,2}/\d{4})")
STATE_RE = re.compile(r"^[A-Z]{2}$")
DOLLAR_RE = re.compile(r"^\$[\d,]")
NUM_LINE_RE = re.compile(r"^[\d,]+\.\d{2}")

# Lines that are headers, footers, or column labels — skip entirely
SKIP_PATTERNS = [
    re.compile(r"^Number of records$"),
    re.compile(r"^\d{4}$"),               # record counts "1000", "2000"
    re.compile(r"^CBC:"),
    re.compile(r"^Displaying records"),
    re.compile(r"^list view criteria"),
    re.compile(r"^This page can"),
    re.compile(r"^To show fewer"),
    re.compile(r"^report and then click"),
    re.compile(r"^Transaction Name"),
    re.compile(r"^Type$"),
    re.compile(r"^Funded File Date"),
    re.compile(r"^Payments Remaining"),
    re.compile(r"^Last Activity"),
    re.compile(r"^State from Lead"),
    re.compile(r"^Insurer$"),
    re.compile(r"^Net Spread"),
    re.compile(r"^Swap Float"),
    re.compile(r"^Corresponding Swap"),
    re.compile(r"^Representative"),
    re.compile(r"^Alia$"),
    re.compile(r"^Stage$"),
    re.compile(r"https?://"),
    re.compile(r"^\d{1,2}/\d{2}/\d{2},\s+\d{1,2}:\d{2}\s+[AP]M"),  # "2/20/26, 3:35 PM"
    re.compile(r"^\d+/\d+$"),             # page numbers "1/45"
    re.compile(r"^Copyright"),
]

# State-machine states
(S_NAME, S_TYPE, S_DATE2, S_STATE_CODE, S_INSURER, S_AMOUNTS) = range(6)


# ── Text extraction ─────────────────────────────────────────────────

def _should_skip(raw_line: str) -> bool:
    """Return True if this PDF line is a header, footer, or label to ignore."""
    s = raw_line.strip()
    if not s:
        return True
    # Lines that are ONLY non-breaking spaces (e.g. "\xa0 \xa0")
    if not s.replace("\xa0", "").strip():
        return True
    for pat in SKIP_PATTERNS:
        if pat.search(s):
            return True
    return False


def _get_all_lines(pdf_paths: list) -> list:
    """
    Extract cleaned text lines from all pages of all PDFs.
    Skips headers, footers, and column labels.

    We use rstrip() (not strip()) so that leading \xa0 characters are
    preserved.  Alias lines like '\xa0 jrairigh' rely on the leading
    \xa0 to be recognised as financial/alias lines by _is_financial().
    If we strip() them, the \xa0 disappears and they bleed into the
    next client name.
    """
    lines = []
    for path in pdf_paths:
        doc = fitz.open(str(path))
        for page in doc:
            for raw_line in page.get_text().split("\n"):
                if not _should_skip(raw_line):
                    lines.append(raw_line.rstrip())   # preserve leading \xa0
        doc.close()
    return lines


def _normalize_inline_funded(lines: list) -> list:
    """
    Some records have the client name and 'Funded' on the same line,
    e.g. 'Mackenzie Jones Funded'.  Split these so the state machine
    sees them as two separate lines: 'Mackenzie Jones' then 'Funded'.
    """
    result = []
    for line in lines:
        if line.endswith(" Funded"):
            prefix = line[:-7].strip()
            if prefix:
                result.append(prefix)
            result.append("Funded")
        else:
            result.append(line)
    return result


def _is_financial(line: str) -> bool:
    """
    Return True if this line is an amount, swap rate, or alias line
    that follows the dollar amount — NOT part of the next client name.

    Patterns seen in the PDFs:
        "$5,317.96 \xa0"         → dollar amount
        "\xa0 jrairigh"          → alias with non-breaking space prefix
        "29,518.01 ktadd"        → swap rate + alias
        "4.37 ChrisS"            → swap rate continuation + alias
    """
    s = line.strip()
    return bool(
        DOLLAR_RE.match(s)
        or NUM_LINE_RE.match(s)
        or re.match(r"^\xa0", line)
        or re.match(r"^[\d,\.]+\s+\w+$", s)  # "29,518.01 ktadd"
        or not s
    )


# ── State-machine parser ────────────────────────────────────────────

def parse_records(pdf_paths: list) -> list:
    """
    Walk through all PDF lines with a state machine and extract one
    dict per funded-file record.

    Returns a list of dicts ready to be appended to the Cases sheet:
        Client_Name, State, County, Judge, Case_Number,
        Court_Date, Status, Notes
    """
    raw = _get_all_lines(pdf_paths)
    lines = _normalize_inline_funded(raw)
    n = len(lines)

    records = []
    state = S_NAME
    cur: dict = {}
    name_acc: list = []
    i = 0

    while i < n:
        line = lines[i]

        # ── Collect client name until "Funded" + "File" ────────────
        if state == S_NAME:
            if line == "Funded" and i + 1 < n and lines[i + 1] == "File":
                cur = {"name": " ".join(name_acc).strip()}
                name_acc = []
                state = S_TYPE
                i += 2   # skip both "Funded" and "File"
                continue
            elif re.match(r"^\xa0", line):
                pass   # stray alias line — skip defensively
            else:
                name_acc.append(line)

        # ── Skip type tokens until first date ──────────────────────
        elif state == S_TYPE:
            if DATE_RE.match(line):
                cur["funded_date"] = line
                state = S_DATE2
            else:
                m = DATE_EMB_RE.search(line)
                if m:
                    # e.g. "(Unhedged) 2/18/2026" — type suffix + date on one line
                    cur["funded_date"] = m.group(1)
                    state = S_DATE2
                # else: pure type token (SS, SPIA, Lottery …) — ignore

        # ── Consume and discard last-activity date ─────────────────
        elif state == S_DATE2:
            state = S_STATE_CODE   # whatever this line is, move on

        # ── Read two-letter state code ─────────────────────────────
        elif state == S_STATE_CODE:
            if STATE_RE.match(line):
                cur["state"] = line
                cur["insurer_parts"] = []
                state = S_INSURER
            elif DOLLAR_RE.match(line) or NUM_LINE_RE.match(line) or re.match(r"^\xa0", line):
                # No state code found (rare edge case) — skip to amounts
                cur.setdefault("state", "")
                cur["insurer_parts"] = []
                state = S_AMOUNTS
                i -= 1   # reprocess current line inside S_AMOUNTS

        # ── Collect insurer name lines until a financial line ──────
        elif state == S_INSURER:
            if DOLLAR_RE.match(line) or NUM_LINE_RE.match(line) or re.match(r"^\xa0", line):
                cur["insurer"] = " ".join(cur.pop("insurer_parts", [])).strip()
                state = S_AMOUNTS
                i -= 1   # reprocess current line inside S_AMOUNTS
            else:
                cur["insurer_parts"].append(line)

        # ── Skip financial / alias lines; detect next record name ──
        elif state == S_AMOUNTS:
            if _is_financial(line):
                pass   # amount or alias line — discard
            else:
                # Non-financial: either the start of the next client name
                # OR a stand-alone alias word (e.g. "Gary", "ChrisS").
                # Distinguish by looking ahead: if "Funded" + "File" appears
                # within the next 8 lines, this line IS the start of a name.
                found_funded = False
                for j in range(1, 9):
                    if i + j >= n:
                        break
                    if lines[i + j] == "Funded" and i + j + 1 < n and lines[i + j + 1] == "File":
                        found_funded = True
                        break

                if found_funded or i == n - 1:
                    if cur.get("name"):
                        records.append(_build_record(cur))
                    cur = {}
                    name_acc = [line]
                    state = S_NAME
                # else: standalone alias word — ignore

        i += 1

    # Flush any in-progress record at end of stream
    if cur.get("name"):
        records.append(_build_record(cur))

    return records


def _build_record(cur: dict) -> dict:
    insurer = cur.get("insurer", "").strip()
    funded = cur.get("funded_date", "")
    note_parts = ["PDF Import"]
    if funded:
        note_parts.append(f"Funded {funded}")
    if insurer:
        note_parts.append(f"Insurer: {insurer}")
    return {
        "Client_Name": cur.get("name", ""),
        "State":       cur.get("state", ""),
        "County":      "",
        "Judge":       "",
        "Case_Number": "",
        "Court_Date":  funded,
        "Status":      "Approved",
        "Notes":       ". ".join(note_parts) + ".",
    }


# ── Excel import ────────────────────────────────────────────────────

def import_to_excel(records: list, excel_path: str, dry_run: bool = False) -> str:
    """
    Append extracted records to the Cases sheet, skipping duplicates.

    Deduplication key: (Client_Name.lower, State.upper, Court_Date)
    """
    path = Path(excel_path)
    if not path.exists():
        return f"Excel file not found: {path}"

    # Load all sheets
    with pd.ExcelFile(path) as xls:
        all_sheets = {name: pd.read_excel(xls, sheet_name=name) for name in xls.sheet_names}

    df_cases = all_sheets.get("Cases", pd.DataFrame())

    # Build dedup key set from existing rows
    existing_keys: set = set()
    for _, row in df_cases.iterrows():
        key = (
            str(row.get("Client_Name", "")).strip().lower(),
            str(row.get("State", "")).strip().upper(),
            str(row.get("Court_Date", "")).strip(),
        )
        existing_keys.add(key)

    # Filter out duplicates
    new_rows = []
    skipped = 0
    for rec in records:
        key = (
            rec["Client_Name"].strip().lower(),
            rec["State"].strip().upper(),
            rec["Court_Date"].strip(),
        )
        if key in existing_keys:
            skipped += 1
        else:
            new_rows.append(rec)
            existing_keys.add(key)   # prevent same record appearing twice in the new batch

    summary = (
        f"Parsed {len(records)} records from PDFs. "
        f"New: {len(new_rows)}, duplicates skipped: {skipped}."
    )

    if dry_run:
        print(f"[DRY RUN] {summary}")
        for r in new_rows[:10]:
            print(f"  {r['Client_Name']} | {r['State']} | {r['Court_Date']}")
        if len(new_rows) > 10:
            print(f"  ... and {len(new_rows) - 10} more")
        return summary

    if not new_rows:
        return summary

    df_new = pd.DataFrame(new_rows)
    df_cases = pd.concat([df_cases, df_new], ignore_index=True)
    all_sheets["Cases"] = df_cases

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in all_sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)

    summary += f" Cases sheet now has {len(df_cases)} total rows."

    # Git commit
    _git_commit(path, f"Import {len(new_rows)} approved records from PDFs")

    return summary


def _git_commit(excel_path: Path, message: str):
    repo_dir = str(excel_path.parent)
    try:
        subprocess.run(["git", "add", excel_path.name], cwd=repo_dir,
                       capture_output=True, check=True)
        status = subprocess.run(["git", "diff", "--cached", "--quiet"],
                                cwd=repo_dir, capture_output=True)
        if status.returncode == 0:
            return   # nothing staged
        subprocess.run(["git", "commit", "-m", message], cwd=repo_dir,
                       capture_output=True, check=True)
        for wait in [0, 2, 4, 8]:
            if wait:
                time.sleep(wait)
            r = subprocess.run(["git", "push"], cwd=repo_dir, capture_output=True)
            if r.returncode == 0:
                break
    except Exception as e:
        print(f"[git] {e}")


# ── CLI entry point ─────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Import CBC funded-file PDFs into Excel")
    parser.add_argument("pdfs", nargs="*", help="PDF file paths (auto-discovered if omitted)")
    parser.add_argument("--excel", default=EXCEL_PATH, help="Path to Excel workbook")
    parser.add_argument("--dry-run", action="store_true",
                        help="Parse only — print results without writing to Excel")
    args = parser.parse_args()

    # Auto-discover PDFs in current directory if none specified
    if args.pdfs:
        pdf_paths = [Path(p) for p in args.pdfs]
    else:
        pdf_paths = sorted(Path(".").glob("CBC*.pdf"))
        if not pdf_paths:
            pdf_paths = sorted(Path(".").glob("*.pdf"))
        if not pdf_paths:
            print("No PDF files found. Pass file paths as arguments.")
            sys.exit(1)

    # Validate
    missing = [p for p in pdf_paths if not p.exists()]
    if missing:
        for p in missing:
            print(f"File not found: {p}")
        sys.exit(1)

    print(f"PDFs to parse: {[str(p) for p in pdf_paths]}")
    print("Extracting records...")

    records = parse_records(pdf_paths)
    print(f"Extracted {len(records)} records")

    if records:
        print(f"\nSample (first 5):")
        for r in records[:5]:
            print(f"  {r['Client_Name']} | {r['State']} | {r['Court_Date']} | {r['Status']}")
        print()

    result = import_to_excel(records, args.excel, dry_run=args.dry_run)
    print(result)


if __name__ == "__main__":
    main()
