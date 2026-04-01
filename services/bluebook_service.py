"""
Bluebook Manager — Bluebook service (business logic).
"""

import os
import shutil

from config import SECTION_FOLDERS, STORAGE_ROOT
from dal import dal
from dal.models import Bluebook
from services.log_service import log


def create_bluebook(die_number: str, description: str = "") -> Bluebook:
    """Create a new bluebook: DB record + disk folders."""
    bid = dal.add_bluebook(die_number, description)

    # Create folder structure on disk
    base = os.path.join(STORAGE_ROOT, die_number)
    for folder in SECTION_FOLDERS.values():
        os.makedirs(os.path.join(base, folder), exist_ok=True)

    log("CREATE_BLUEBOOK", f"Die# {die_number} (id={bid})")
    return dal.get_bluebook(bid)


def get_bluebook(bluebook_id: int) -> Bluebook:
    return dal.get_bluebook(bluebook_id)


def get_bluebook_by_die(die_number: str) -> Bluebook:
    return dal.get_bluebook_by_die(die_number)


def search_bluebooks(search: str = "", customer_id: int = None,
                     search_description: bool = False,
                     search_qa: bool = False) -> list[Bluebook]:
    return dal.list_bluebooks(search=search, customer_id=customer_id,
                              search_description=search_description,
                              search_qa=search_qa)


def delete_bluebook(bluebook_id: int, delete_files: bool = False):
    """Delete a bluebook. Optionally remove disk files."""
    bb = dal.get_bluebook(bluebook_id)
    if not bb:
        return

    if delete_files:
        folder = os.path.join(STORAGE_ROOT, bb.die_number)
        if os.path.isdir(folder):
            shutil.rmtree(folder, ignore_errors=True)

    dal.delete_bluebook(bluebook_id)
    log("DELETE_BLUEBOOK", f"Die# {bb.die_number} (id={bluebook_id}), files_deleted={delete_files}")


def update_bluebook(bluebook_id: int, die_number: str, description: str = ""):
    dal.update_bluebook(bluebook_id, die_number, description)
    log("UPDATE_BLUEBOOK", f"Die# {die_number} (id={bluebook_id})")


def get_storage_path(die_number: str, section_type: str = None) -> str:
    """Get the disk path for a bluebook or section."""
    base = os.path.join(STORAGE_ROOT, die_number)
    if section_type:
        return os.path.join(base, SECTION_FOLDERS[section_type])
    return base
