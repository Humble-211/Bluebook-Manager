"""
Bluebook Manager — Auto-Link Files Script

Scans specific folders and links all files that contain the given die number
in their filename to the corresponding bluebook.

Usage:
    python auto_link_files.py <die_number>
"""

import os
import sys

# Ensure project root is importable
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

from dal.database import init_db
from dal import dal
from services import file_service
from services.log_service import log

def main():
    if len(sys.argv) < 2:
        print("Usage: python auto_link_files.py <die_number>")
        sys.exit(1)

    die_number = sys.argv[1].strip()
    
    init_db()
    bb = dal.get_bluebook_by_die(die_number)
    if not bb:
        print(f"Bluebook for die number {die_number} not found in database.")
        sys.exit(1)

    bluebook_id = bb.id

    paths_to_search = {
        "cover": r"S:\QCLAB\Bluebooks Related\Blue Book",
        "master_drawings": r"X:\\",
        "qc_drawings": r"S:\QCLAB\Bluebooks Related\QC PDF Drawing",
        "approval": r"S:\QCLAB\Bluebooks Related\Sample Approval",
    }

    def search_for_files(base_dir: str, match_func) -> list[str]:
        found_files = []
        if not base_dir or not os.path.isdir(base_dir):
            return found_files
        for root, _, files in os.walk(base_dir):
            for f in files:
                if match_func(f):
                    found_files.append(os.path.join(root, f))
        return found_files

    def match_any(f, expected_exts=None):
        name, ext = os.path.splitext(f)
        if expected_exts and ext.lower() not in expected_exts:
            return False
        return die_number.lower() in name.lower()

    def process_section(section_type, expected_exts):
        base_dir = paths_to_search.get(section_type)
        if not base_dir:
            return

        # Fetch existing files to prevent duplicate attachments
        existing_files = dal.get_files_for_bluebook(bluebook_id, section_type)
        existing_filenames = {os.path.basename(ef.file_path).lower() for ef in existing_files}

        for path in search_for_files(base_dir, lambda f: match_any(f, expected_exts)):
            filename = os.path.basename(path).lower()
            if filename in existing_filenames:
                continue # Already attached

            try:
                # master_drawings and qc_drawings auto-create shortcuts based on file_service logic
                file_service.attach_file(bluebook_id, section_type, path)
                print(f"Linked {section_type}: {path}")
            except Exception as e:
                log("AUTO_LINK_ERROR", f"Failed to link {section_type} {path}: {e}")

    process_section("cover", [".docx", ".pdf"])
    process_section("master_drawings", [".pdf"])
    process_section("qc_drawings", [".pdf"])
    process_section("approval", [".pdf", ".docx"])

if __name__ == "__main__":
    main()
