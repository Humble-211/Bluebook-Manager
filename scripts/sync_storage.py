"""
Bluebook Manager — Sync Storage Script

Compares the physical files in the `storage/` directory against the `bluebook_files` 
table in the database. Any files found on disk that are not registered in the 
database are deleted to free up space.

Usage:
    python scripts/sync_storage.py
"""

import os
import sys

# Ensure project root is importable
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

from config import STORAGE_ROOT
from dal.database import get_connection, init_db

def main():
    print("Starting Storage Synchronization...")
    init_db()

    # 1. Fetch all known file paths from the database
    conn = get_connection()
    try:
        rows = conn.execute("SELECT file_path FROM bluebook_files").fetchall()
        # file_path in DB is stored relative to STORAGE_ROOT
        db_files_relative = {row["file_path"].replace("/", "\\").lower() for row in rows}
    finally:
        conn.close()

    print(f"Found {len(db_files_relative)} files referenced in the database.")

    if not os.path.exists(STORAGE_ROOT):
        print(f"Storage directory does not exist: {STORAGE_ROOT}")
        return

    # 2. Recursively scan the storage/ directory for all physical files
    physical_files_relative = set()
    for root, dirs, files in os.walk(STORAGE_ROOT):
        for f in files:
            abs_path = os.path.join(root, f)
            rel_path = os.path.relpath(abs_path, STORAGE_ROOT)
            physical_files_relative.add(rel_path.replace("/", "\\").lower())

    print(f"Found {len(physical_files_relative)} files physically on disk.")

    # 3. Compare the two lists
    orphaned_files = physical_files_relative - db_files_relative

    if not orphaned_files:
        print("Storage is perfectly in sync with the database. No orphaned files found.")
        return

    print(f"\nFound {len(orphaned_files)} orphaned file(s) on disk. Deleting now...")

    # 4. Delete the orphaned files
    deleted_count = 0
    failed_count = 0
    for rel_path in orphaned_files:
        # Reconstruct the absolute path securely
        abs_path = os.path.join(STORAGE_ROOT, rel_path)
        if os.path.exists(abs_path):
            try:
                os.remove(abs_path)
                print(f"  [DELETED] {rel_path}")
                deleted_count += 1
            except Exception as e:
                print(f"  [FAILED]  Could not delete {rel_path}: {e}")
                failed_count += 1
        else:
            print(f"  [MISSING] {rel_path} was already deleted or moved.")

    print("\n--- Synchronization Summary ---")
    print(f"Deleted: {deleted_count}")
    print(f"Failed:  {failed_count}")
    print("-------------------------------")

if __name__ == "__main__":
    main()
