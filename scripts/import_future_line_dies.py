"""
Bluebook Manager — Import Future Line Outsourced Dies

Reads die numbers and descriptions from an Excel sheet and:
  1. Ensures the "Future Line" outsource record exists (creates it if not).
  2. For each die number, looks up the bluebook — creates one if missing.
     If the bluebook already exists and has no description, it is updated.
  3. Links the die to "Future Line" via the outsource_bluebooks join table.
     Already-linked dies are skipped cleanly.

Usage:
    python scripts/import_future_line_dies.py

Excel format:
    - First sheet is used.
    - Column A → Description
    - Column B → Die number  (3–6 digits, optional trailing letter)
    - A header row is fine — rows without a valid die number in column B are skipped.
"""

import os
import re
import sys

# Ensure project root is importable and config.BASE_DIR resolves correctly.
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)
sys.argv[0] = os.path.join(PROJECT_ROOT, "main.py")

try:
    import openpyxl
except ImportError:
    print("ERROR: 'openpyxl' is required.  Run:  pip install openpyxl")
    sys.exit(1)

from dal.database import init_db
from dal import dal

# ── Configuration ──────────────────────────────────────────────────────────────
OUTSOURCE_NAME = "Future Line"
XLSX_PATH = r"C:\Users\hmai\Desktop\opencv\bluebook-project\scripts\excel\run.xlsx"

# All databases the script can import into.
# Add new entries as: ("Label shown in menu", r"C:\path\to\database.db")
KNOWN_DATABASES = [
    ("Main database    (data/bluebook.db)",
     r"C:\Users\hmai\Desktop\opencv\bluebook-project\data\bluebook.db"),
     #("Secondary DB", r"C:\path\to\other\bluebook.db"),
     
]

# A plausible die number: 3–6 digits, optionally followed by a single letter.
DIE_PATTERN = re.compile(r"^\d{3,6}[A-Za-z]?$")


# ── Helpers ────────────────────────────────────────────────────────────────────

def normalise_die(raw) -> str | None:
    """Convert a cell value to a clean base die-number string, or None if invalid.

    Trailing letters are stripped so that variants like 12345A, 12345B all
    resolve to the same bluebook (12345).
    """
    if raw is None:
        return None
    text = str(raw).strip()
    # Remove accidental trailing .0 from numeric Excel cells (e.g. 16059.0 → 16059)
    text = re.sub(r"\.0+$", "", text)
    if not DIE_PATTERN.match(text):
        return None
    # Strip trailing alpha suffix (e.g. 12345A → 12345)
    return re.sub(r"[A-Za-z]+$", "", text)


def load_rows(xlsx_path: str) -> list[tuple[str, str]]:
    """Open the first sheet and return deduplicated (die_number, description) pairs.

    Column A → description (any text; empty string if blank).
    Column B → die number; trailing letters are stripped (12345A → 12345).

    When the same base die number appears more than once (e.g. 12345A and
    12345B), only the first occurrence is kept and a warning is printed.
    """
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.worksheets[0]

    rows: list[tuple[str, str]] = []
    seen: set[str] = set()

    for row in ws.iter_rows(min_col=1, max_col=2, values_only=True):
        col_a, col_b = (row[0], row[1]) if len(row) >= 2 else (None, None)
        die = normalise_die(col_b)
        if not die:
            continue  # skip header / blank / non-die rows
        if die in seen:
            print(f"  ⚠  Duplicate (skipped): {str(col_b).strip()} → already queued as Die# {die}")
            continue
        seen.add(die)
        description = str(col_a).strip() if col_a is not None else ""
        rows.append((die, description))

    wb.close()
    return rows


def ensure_outsource(name: str) -> int:
    """Return the id of the outsource record, creating it if necessary."""
    existing = dal.get_outsource_by_name(name)
    if existing:
        return existing.id
    print(f"  +  Outsource '{name}' not found — creating it.")
    return dal.add_outsource(name)


def ensure_bluebook(die_number: str, description: str) -> tuple[int, bool]:
    """Return (bluebook_id, was_created).

    Creates the bluebook with *description* if it doesn't exist.
    If it already exists but has no description, the description is applied.
    """
    bb = dal.get_bluebook_by_die(die_number)
    if bb:
        if description and not bb.description:
            dal.update_bluebook(bb.id, die_number, description)
        return bb.id, False
    bb_id = dal.add_bluebook(die_number, description)
    return bb_id, True



def already_linked(outsource_id: int, bluebook_id: int) -> bool:
    """Check whether the outsource↔bluebook link already exists."""
    existing = dal.get_outsources_for_bluebook(bluebook_id)
    return any(o.id == outsource_id for o in existing)


# ── Interactive helpers ────────────────────────────────────────────────────────

def pick_database() -> str:
    """Present a numbered menu and return the chosen DB path."""
    div = "=" * 60
    print(f"\n{div}")
    print("  STEP 1 — Select target database")
    print(div)
    for i, (label, _) in enumerate(KNOWN_DATABASES, 1):
        print(f"  [{i}] {label}")
    print("  [C] Custom path (type manually)")
    print(div)

    while True:
        choice = input("  Enter choice: ").strip()
        if choice.upper() == "C":
            path = input("  Enter full path to .db file: ").strip().strip('"')
            if not path.lower().endswith(".db"):
                print("  ✗  Path must end in .db — try again.")
                continue
            return path
        if choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(KNOWN_DATABASES):
                return KNOWN_DATABASES[idx][1]
        print("  ✗  Invalid choice — try again.")


def confirm(prompt: str) -> bool:
    """Ask a yes/no question; return True for yes."""
    return input(f"{prompt} [y/N]: ").strip().lower() in ("y", "yes")


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    import config                    # noqa: F401 — imported to patch its DB_PATH
    import dal.database as _dal_db   # alias avoids shadowing the top-level 'dal' module

    div = "=" * 60

    # ── Step 1: choose database ──────────────────────────────────────────────
    db_path = pick_database()
    if not os.path.isfile(db_path):
        print(f"\n  ✗  Database not found: {db_path}")
        sys.exit(1)

    # Redirect the DAL to the chosen database.
    # dal.database uses `from config import DB_PATH` which binds a local copy
    # at import time — so we must patch dal.database.DB_PATH directly.
    config.DB_PATH = db_path
    _dal_db.DB_PATH = db_path

    # ── Step 2: load Excel ───────────────────────────────────────────────────
    print(f"\n{div}")
    print("  STEP 2 — Load Excel file")
    print(div)
    print(f"  File : {XLSX_PATH}")

    if not os.path.isfile(XLSX_PATH):
        print(f"  ✗  Excel file not found: {XLSX_PATH}")
        sys.exit(1)

    try:
        rows = load_rows(XLSX_PATH)
    except Exception as e:
        print(f"  ✗  Could not read Excel file: {e}")
        sys.exit(1)

    if not rows:
        print("  ✗  No valid die numbers found in column B.  Nothing to import.")
        return

    print(f"  Found {len(rows)} valid die number(s).")

    # ── Step 3: confirm ──────────────────────────────────────────────────────
    print(f"\n{div}")
    print("  STEP 3 — Confirm import")
    print(div)
    print(f"  Database  : {db_path}")
    print(f"  Excel     : {XLSX_PATH}")
    print(f"  Outsource : {OUTSOURCE_NAME}")
    print(f"  Dies      : {len(rows)}")
    print(div)

    if not confirm("  Proceed with import?"):
        print("  Aborted.")
        return

    # ── Step 4: run import ───────────────────────────────────────────────────
    print(f"\n{div}")
    print("  STEP 4 — Importing")
    print(div)

    init_db()
    outsource_id = ensure_outsource(OUTSOURCE_NAME)

    created_bbs: list[str] = []
    linked_new:  list[str] = []
    skipped:     list[str] = []

    for die, description in rows:
        bb_id, was_created = ensure_bluebook(die, description)

        if was_created:
            created_bbs.append(die)
            desc_note = f"  [{description}]" if description else ""
            print(f"  +  Created new bluebook: Die# {die}{desc_note}")

        if already_linked(outsource_id, bb_id):
            skipped.append(die)
            print(f"  ⊘  Die# {die}  →  '{OUTSOURCE_NAME}' (already linked)")
            continue

        dal.link_outsource_bluebook(outsource_id, bb_id)
        linked_new.append(die)
        print(f"  ✓  Die# {die}  →  linked to '{OUTSOURCE_NAME}'")

    # ── Step 5: summary ──────────────────────────────────────────────────────
    print(f"\n{div}")
    print("  STEP 5 — Summary")
    print(div)
    print(f"  Linked (new):        {len(linked_new)}")
    print(f"  Bluebooks created:   {len(created_bbs)}")
    print(f"  Already linked:      {len(skipped)}")
    print(f"  Total processed:     {len(rows)}")
    print(f"{div}\n")


if __name__ == "__main__":
    main()
