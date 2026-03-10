"""
Import Packing Instruction files.

Scans the Packing source folder (and subfolders) for files with
extensions: .jpg, .png, .jpeg, .docx

Files are named by die number (e.g. 12345.jpg). The script extracts the
die number from the filename and attaches the file to the matching
bluebook's packing_instruction section.

Usage:
    python scripts/import_packing.py
    python scripts/import_packing.py --dry-run   (preview without importing)
"""

import re

import os
import sys

# Ensure project root is importable and config.BASE_DIR resolves correctly.
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)
sys.argv[0] = os.path.join(PROJECT_ROOT, "main.py")

from dal.database import init_db
from dal import dal
from services import file_service

# ── Configuration ──
SCAN_DIR = r"S:\QCLAB\Bluebooks Related\Packing"
SECTION_TYPE = "packing_instruction"
ALLOWED_EXTENSIONS = {".jpg", ".png", ".jpeg", ".docx"}


def main():
    dry_run = "--dry-run" in sys.argv

    if not os.path.isdir(SCAN_DIR):
        print(f"ERROR: Scan directory not found: {SCAN_DIR}")
        print(f"Make sure the network drive is accessible.")
        sys.exit(1)

    init_db()

    # Build a set of known die numbers for fast lookup
    all_bluebooks = dal.list_bluebooks(limit=10000)
    die_to_bb = {bb.die_number: bb for bb in all_bluebooks}
    print(f"Loaded {len(die_to_bb)} bluebook(s) from database.\n")

    # Collect matching files
    target_files = []
    for root, _dirs, files in os.walk(SCAN_DIR):
        for f in files:
            if f.startswith("~$"):
                continue  # Skip temp Word files
            ext = os.path.splitext(f)[1].lower()
            if ext in ALLOWED_EXTENSIONS:
                target_files.append(os.path.join(root, f))

    if not target_files:
        print(f"No matching files found in {SCAN_DIR}")
        return

    print(f"Found {len(target_files)} file(s) to process.\n")

    if dry_run:
        print("=== DRY RUN MODE (no changes will be made) ===\n")

    linked = []
    skipped = []
    failed = []

    for filepath in target_files:
        filename = os.path.basename(filepath)
        # Die number = first group of digits in filename (e.g. "15010-1.docx" -> "15010")
        stem = os.path.splitext(filename)[0]
        m = re.match(r"(\d+)", stem)
        if not m:
            failed.append((filepath, f"No digits found in filename: {filename}"))
            continue
        die_number = m.group(1)

        if die_number not in die_to_bb:
            failed.append((filepath, f"No bluebook found for Die# {die_number}"))
            continue

        bb = die_to_bb[die_number]

        # Check if file is already attached (by matching filename)
        existing_files = dal.get_files_for_bluebook(bb.id, SECTION_TYPE)
        already_attached = any(
            os.path.basename(ef.file_path) == filename for ef in existing_files
        )
        if already_attached:
            skipped.append((filepath, die_number))
            print(f"  ⊘  {filename}  →  Die# {die_number} (already attached)")
            continue

        if dry_run:
            linked.append((filepath, die_number, None))
            print(f"  ○  {filename}  →  Die# {die_number} (would attach)")
            continue

        # Attach the file
        try:
            bf = file_service.attach_file(bb.id, SECTION_TYPE, filepath)
            linked.append((filepath, die_number, bf.id))
            print(f"  ✓  {filename}  →  Die# {die_number}")
        except Exception as e:
            failed.append((filepath, f"Error attaching to Die# {die_number}: {e}"))

    # ── Summary ──
    print(f"\n{'='*60}")
    print(f"  {'Would attach' if dry_run else 'Attached'}:  {len(linked)}")
    print(f"  Skipped (already attached):  {len(skipped)}")
    print(f"  Failed:   {len(failed)}")
    print(f"{'='*60}")

    if failed:
        print("\n⚠  Files that could NOT be linked:\n")
        for path, reason in failed:
            print(f"  ✗  {path}")
            print(f"     Reason: {reason}\n")


if __name__ == "__main__":
    main()
