"""
Bluebook Manager — Import Customer↔Die Links

Reads customer-to-die mappings from an Excel sheet and:
  1. Ensures each customer record exists (creates it if not).
  2. For each die number, looks up the bluebook — creates one if missing.
  3. Links the customer to the die via the customer_bluebooks join table.
     Already-linked pairs are skipped cleanly.

Usage:
    python scripts/import_customer_dies.py

Excel format:
    - First sheet is used.
    - Column A → Customer name
    - Column B → Die number  (3–6 digits, optional trailing letter stripped)
    - Column C → Description (optional; used when creating a new bluebook)
    - A header row is fine — rows without a valid die number in column B are skipped.
    - Multiple different customers in the same file are supported.

Rules (same as import_future_line_dies.py):
    - Trailing letters stripped: 12345A → 12345
    - Duplicate (customer, die) pairs in the file are skipped after the first.
    - Script is fully idempotent — safe to run multiple times.
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
XLSX_PATH = r"C:\Users\hmai\Desktop\opencv\bluebook-project\scripts\excel\customers.xlsx"

# All databases the script can import into.
# Add new entries as: ("Label shown in menu", r"C:\path\to\database.db")
KNOWN_DATABASES = [
    ("Main database    (data/bluebook.db)",
     r"C:\Users\hmai\Desktop\opencv\bluebook-project\data\bluebook.db"),
    # ("Secondary DB", r"C:\path\to\other\bluebook.db"),
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


def normalise_customer(raw) -> str | None:
    """Return a stripped customer name, or None if blank."""
    if raw is None:
        return None
    text = str(raw).strip()
    return text if text else None


def load_rows(xlsx_path: str) -> list[tuple[str, str, str]]:
    """Open the first sheet and return deduplicated (customer_name, die_number, description) triples.

    Column A → customer name (any non-empty text).
    Column B → die number; trailing letters are stripped (12345A → 12345).
    Column C → description (optional; empty string if blank).

    Rows missing either a customer name or a valid die number are skipped.
    Duplicate (customer, die) pairs — after normalization — are skipped with a warning.
    The same die with a different customer is NOT a duplicate.
    """
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.worksheets[0]

    rows: list[tuple[str, str, str]] = []
    seen: set[tuple[str, str]] = set()

    for row in ws.iter_rows(min_col=1, max_col=3, values_only=True):
        col_a = row[0] if len(row) > 0 else None
        col_b = row[1] if len(row) > 1 else None
        col_c = row[2] if len(row) > 2 else None

        customer = normalise_customer(col_a)
        die = normalise_die(col_b)

        if not customer or not die:
            continue  # skip header / blank / invalid rows

        key = (customer, die)
        if key in seen:
            print(f"  ⚠  Duplicate (skipped): {customer!r} + {str(col_b).strip()} → already queued")
            continue
        seen.add(key)
        description = str(col_c).strip() if col_c is not None else ""
        rows.append((customer, die, description))

    wb.close()
    return rows


def ensure_customer(name: str) -> int:
    """Return the id of the customer record, creating it if necessary."""
    existing = dal.get_customer_by_name(name)
    if existing:
        return existing.id
    print(f"  +  Customer '{name}' not found — creating it.")
    return dal.add_customer(name)


def ensure_bluebook(die_number: str, description: str) -> tuple[int, str]:
    """Return (bluebook_id, action_taken).

    action_taken is one of: 'created', 'updated' (description added), or 'exists'.
    """
    bb = dal.get_bluebook_by_die(die_number)
    if bb:
        if description and not bb.description:
            dal.update_bluebook(bb.id, die_number, description)
            return bb.id, 'updated'
        return bb.id, 'exists'
    bb_id = dal.add_bluebook(die_number, description)
    return bb_id, 'created'


def already_linked(customer_id: int, bluebook_id: int) -> bool:
    """Check whether the customer↔bluebook link already exists."""
    existing = dal.get_customers_for_bluebook(bluebook_id)
    return any(c.id == customer_id for c in existing)


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
    import config                     # noqa: F401 — imported to patch its DB_PATH
    import dal.database as _dal_db    # alias avoids shadowing the top-level 'dal' module

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
        print("  ✗  No valid rows found (need customer in col A, die in col B).")
        return

    unique_customers = sorted({c for c, _, _d in rows})
    print(f"  Found {len(rows)} valid pair(s) across {len(unique_customers)} customer(s).")
    for name in unique_customers:
        count = sum(1 for c, _, _d in rows if c == name)
        print(f"    • {name}  ({count} die{'s' if count != 1 else ''})")

    # ── Step 3: confirm ──────────────────────────────────────────────────────
    print(f"\n{div}")
    print("  STEP 3 — Confirm import")
    print(div)
    print(f"  Database   : {db_path}")
    print(f"  Excel      : {XLSX_PATH}")
    print(f"  Pairs      : {len(rows)}")
    print(f"  Customers  : {len(unique_customers)}")
    print(div)

    if not confirm("  Proceed with import?"):
        print("  Aborted.")
        return

    # ── Step 4: run import ───────────────────────────────────────────────────
    print(f"\n{div}")
    print("  STEP 4 — Importing")
    print(div)

    init_db()

    # Cache customer ids within this run to avoid redundant DB lookups
    customer_id_cache: dict[str, int] = {}

    created_bbs:  list[str] = []
    updated_bbs:  list[str] = []
    linked_new:   list[tuple[str, str]] = []
    skipped:      list[tuple[str, str]] = []

    for customer_name, die, description in rows:
        # Resolve customer
        if customer_name not in customer_id_cache:
            existing_cust = dal.get_customer_by_name(customer_name)
            if existing_cust:
                customer_id_cache[customer_name] = existing_cust.id
            else:
                new_id = dal.add_customer(customer_name)
                customer_id_cache[customer_name] = new_id
                print(f"  +  Created new customer: '{customer_name}'")
        customer_id = customer_id_cache[customer_name]

        # Resolve bluebook
        bb_id, bb_action = ensure_bluebook(die, description)
        if bb_action == 'created':
            created_bbs.append(die)
            desc_note = f"  [{description}]" if description else ""
            print(f"  +  Created new bluebook: Die# {die}{desc_note}")
        elif bb_action == 'updated':
            updated_bbs.append(die)
            print(f"  ~  Updated bluebook desc : Die# {die}  [{description}]")

        # Link
        if already_linked(customer_id, bb_id):
            skipped.append((customer_name, die))
            print(f"  ⊘  Skipped link: Die# {die} → '{customer_name}' (already linked)")
            continue

        dal.link_customer_bluebook(customer_id, bb_id)
        linked_new.append((customer_name, die))
        print(f"  ✓  Linked:       Die# {die} → '{customer_name}'")

    # ── Step 5: summary ──────────────────────────────────────────────────────
    print(f"\n{div}")
    print("  STEP 5 — Summary")
    print(div)
    print(f"  Customers created:   {len(customer_id_cache) - sum(1 for c in customer_id_cache if dal.get_customer_by_name(c))}") # rough approx for log
    print(f"  Bluebooks created:   {len(created_bbs)}")
    print(f"  Bluebooks updated:   {len(updated_bbs)}")
    print(f"  Links established:   {len(linked_new)}")
    print(f"  Links skipped:       {len(skipped)} (already existed)")
    print(f"  Total processed:     {len(rows)}")
    print(f"{div}\n")


if __name__ == "__main__":
    main()
