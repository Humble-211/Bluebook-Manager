"""
Bluebook Manager — Print service.

Uses Windows shell for printing. Falls back to opening the file if
silent printing is not available.
"""

import os
import subprocess
import sys

from config import SECTION_TYPES, STORAGE_ROOT
from dal import dal
from services.log_service import log


def _print_file_win32(abs_path: str) -> bool:
    """Attempt to print a file using Windows ShellExecute."""
    try:
        import win32api
        win32api.ShellExecute(0, "print", abs_path, None, ".", 0)
        return True
    except ImportError:
        pass
    except Exception as e:
        log("PRINT_ERROR", f"win32api print failed for {abs_path}: {e}")

    # Fallback: try subprocess with 'print' verb via PowerShell
    try:
        subprocess.Popen(
            ["powershell", "-Command", f'Start-Process -FilePath "{abs_path}" -Verb Print'],
            shell=False,
        )
        return True
    except Exception as e:
        log("PRINT_ERROR", f"PowerShell print failed for {abs_path}: {e}")

    return False


def print_file(file_path: str) -> bool:
    """Print a single file. file_path may be relative or absolute."""
    from services.file_service import get_absolute_path, resolve_shortcut
    abs_path = get_absolute_path(file_path)
    if not os.path.isfile(abs_path):
        log("PRINT_ERROR", f"File not found: {abs_path}")
        return False

    # Resolve shortcut to its target
    abs_path = resolve_shortcut(abs_path)
    if not os.path.isfile(abs_path):
        log("PRINT_ERROR", f"File not found: {abs_path}")
        return False

    success = _print_file_win32(abs_path)
    if not success:
        # Ultimate fallback: just open the file
        log("PRINT_FALLBACK", f"Opening file for manual print: {abs_path}")
        try:
            os.startfile(abs_path)
            success = True
        except Exception as e:
            log("PRINT_ERROR", f"Failed to open file: {e}")

    if success:
        log("PRINT_FILE", f"path={file_path}")
    return success


def print_section(bluebook_id: int, section_type: str) -> int:
    """Print all files in a section. Returns count of files sent to print."""
    files = dal.get_files_for_bluebook(bluebook_id, section_type=section_type)
    printed = 0
    for bf in files:
        if print_file(bf.file_path):
            printed += 1
    log("PRINT_SECTION", f"bluebook_id={bluebook_id}, section={section_type}, "
                         f"printed={printed}/{len(files)}")
    return printed


def print_all(bluebook_id: int) -> int:
    """Print all files across all sections in order. Returns count printed."""
    bb = dal.get_bluebook(bluebook_id)
    if not bb:
        return 0

    total_printed = 0
    for section in SECTION_TYPES:
        files = dal.get_files_for_bluebook(bluebook_id, section_type=section)
        for bf in files:
            if print_file(bf.file_path):
                total_printed += 1

    log("PRINT_ALL", f"Bluebook Die# {bb.die_number}, total_printed={total_printed}")
    return total_printed
