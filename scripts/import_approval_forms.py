"""
Bluebook Manager — Import Approval Forms

Scans a folder (and subfolders) for .docx and .pdf files,
extracts numbers from filenames, and links them to the
matching bluebook's approval section.

Example:
    "Approval Form - 16059.docx"  →  links to bluebook die# 16059
"""

import os
import re
import sys

# Ensure project root is importable and config.BASE_DIR resolves correctly.
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)
sys.argv[0] = os.path.join(PROJECT_ROOT, "main.py")

from dal.database import init_db
from dal import dal
from services import file_service

# ── Configuration ──────────────────────────────────────────
# Set this to the folder containing approval form files
SCAN_DIR = r"S:\QCLAB\Bluebooks Related\Sample Approval"
SECTION_TYPE = "approval"
ALLOWED_EXTENSIONS = {".docx", ".pdf"}

# Regex: extract all digit sequences of 3+ digits from a filename,
# optionally followed by trailing letters which are ignored.
# This handles patterns like:
#   "Approval Form - 16059.docx"       → 16059
#   "Approval Form - 16059C.pdf"       → 16059
#   "16059 Approval.pdf"               → 16059
#   "Approval Form - 16059 rev2.docx"  → 16059
NUMBER_PATTERN = re.compile(r"(\d{3,})[A-Za-z]*")


def extract_die_numbers_from_filename(filename: str) -> list[str]:
    """Extract all plausible die numbers (3+ digit sequences) from a filename.

    Trailing letters are stripped (e.g. '16059C' → '16059').
    Returns them in order of appearance. The extension is stripped first.
    """
    name_without_ext = os.path.splitext(filename)[0]
    matches = NUMBER_PATTERN.findall(name_without_ext)
    return matches


def main():
    if not os.path.isdir(SCAN_DIR):
        print(f"ERROR: Scan directory not found: {SCAN_DIR}")
        sys.exit(1)

    init_db()

    # Collect all .docx and .pdf files (skip temp files starting with ~$)
    approval_files = []
    for root, _dirs, files in os.walk(SCAN_DIR):
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            if ext in ALLOWED_EXTENSIONS and not f.startswith("~$"):
                approval_files.append(os.path.join(root, f))

    if not approval_files:
        print(f"No .docx or .pdf files found in {SCAN_DIR}")
        return

    print(f"Found {len(approval_files)} approval file(s) to process.\n")

    linked = []
    created = []
    failed = []

    for filepath in approval_files:
        filename = os.path.basename(filepath)
        die_numbers = extract_die_numbers_from_filename(filename)

        if not die_numbers:
            failed.append((filepath, "No number found in filename"))
            continue

        file_linked = False
        for die_number in die_numbers:
            # Look up the bluebook by exact die number, create if missing
            bb = dal.get_bluebook_by_die(die_number)
            if not bb:
                bb_id = dal.add_bluebook(die_number)
                bb = dal.get_bluebook(bb_id)
                created.append(die_number)
                print(f"  +  Created new bluebook: Die# {die_number}")

            # Check if file is already attached (by matching filename)
            existing_files = dal.get_files_for_bluebook(bb.id, SECTION_TYPE)
            already_attached = any(
                os.path.basename(ef.file_path) == filename for ef in existing_files
            )
            if already_attached:
                print(f"  ⊘  {filename}  →  Die# {die_number} (already attached)")
                file_linked = True
                continue

            # Attach the file
            try:
                bf = file_service.attach_file(bb.id, SECTION_TYPE, filepath)
                linked.append((filepath, die_number, bf.id))
                file_linked = True
                print(f"  ✓  {filename}  →  Die# {die_number}")
            except Exception as e:
                failed.append((filepath, f"Error attaching to Die# {die_number}: {e}"))

        if not file_linked and (filepath, "No number found in filename") not in failed:
            failed.append((filepath, "Could not attach file"))

    # ── Summary ──
    print(f"\n{'='*60}")
    print(f"  Linked:           {len(linked)}")
    print(f"  Bluebooks created: {len(created)}")
    print(f"  Failed:           {len(failed)}")
    print(f"{'='*60}")

    if failed:
        print("\n⚠  Files that could NOT be linked:\n")
        for path, reason in failed:
            print(f"  ✗  {os.path.basename(path)}")
            print(f"     Path:   {path}")
            print(f"     Reason: {reason}\n")


if __name__ == "__main__":
    main()
