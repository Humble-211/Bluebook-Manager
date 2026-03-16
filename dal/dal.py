"""
Bluebook Manager — Data Access Layer (CRUD operations).
"""

from __future__ import annotations

import sqlite3
from typing import Optional

from dal.database import get_connection
from dal.models import (
    ActionLog,
    Bluebook,
    BluebookFile,
    Customer,
    Outsource,
    SharedFileMap,
)


# ──────────────────────────────────────
# Customers
# ──────────────────────────────────────

def add_customer(name: str, contact_info: str = "") -> int:
    conn = get_connection()
    cur = conn.execute(
        "INSERT INTO customers (name, contact_info) VALUES (?, ?)",
        (name, contact_info),
    )
    conn.commit()
    cid = cur.lastrowid
    conn.close()
    return cid


def get_customer(customer_id: int) -> Optional[Customer]:
    conn = get_connection()
    row = conn.execute("SELECT * FROM customers WHERE id = ?", (customer_id,)).fetchone()
    conn.close()
    if row:
        return Customer(id=row["id"], name=row["name"],
                        contact_info=row["contact_info"],
                        created_at=row["created_at"])
    return None


def get_customer_by_name(name: str) -> Optional[Customer]:
    conn = get_connection()
    row = conn.execute("SELECT * FROM customers WHERE name = ?", (name,)).fetchone()
    conn.close()
    if row:
        return Customer(id=row["id"], name=row["name"],
                        contact_info=row["contact_info"],
                        created_at=row["created_at"])
    return None


def list_customers() -> list[Customer]:
    conn = get_connection()
    rows = conn.execute("SELECT * FROM customers ORDER BY name").fetchall()
    conn.close()
    return [Customer(id=r["id"], name=r["name"],
                     contact_info=r["contact_info"],
                     created_at=r["created_at"]) for r in rows]


def update_customer(customer_id: int, name: str, contact_info: str = ""):
    conn = get_connection()
    conn.execute(
        "UPDATE customers SET name = ?, contact_info = ? WHERE id = ?",
        (name, contact_info, customer_id),
    )
    conn.commit()
    conn.close()


def delete_customer(customer_id: int):
    conn = get_connection()
    conn.execute("DELETE FROM customer_bluebooks WHERE customer_id = ?", (customer_id,))
    conn.execute("DELETE FROM customers WHERE id = ?", (customer_id,))
    conn.commit()
    conn.close()


# ──────────────────────────────────────
# Outsources
# ──────────────────────────────────────

def add_outsource(name: str, contact_info: str = "") -> int:
    conn = get_connection()
    cur = conn.execute(
        "INSERT INTO outsources (name, contact_info) VALUES (?, ?)",
        (name, contact_info),
    )
    conn.commit()
    oid = cur.lastrowid
    conn.close()
    return oid


def get_outsource(outsource_id: int) -> Optional[Outsource]:
    conn = get_connection()
    row = conn.execute("SELECT * FROM outsources WHERE id = ?", (outsource_id,)).fetchone()
    conn.close()
    if row:
        return Outsource(id=row["id"], name=row["name"],
                         contact_info=row["contact_info"],
                         created_at=row["created_at"])
    return None


def get_outsource_by_name(name: str) -> Optional[Outsource]:
    conn = get_connection()
    row = conn.execute("SELECT * FROM outsources WHERE name = ?", (name,)).fetchone()
    conn.close()
    if row:
        return Outsource(id=row["id"], name=row["name"],
                         contact_info=row["contact_info"],
                         created_at=row["created_at"])
    return None


def list_outsources() -> list[Outsource]:
    conn = get_connection()
    rows = conn.execute("SELECT * FROM outsources ORDER BY name").fetchall()
    conn.close()
    return [Outsource(id=r["id"], name=r["name"],
                      contact_info=r["contact_info"],
                      created_at=r["created_at"]) for r in rows]


def update_outsource(outsource_id: int, name: str, contact_info: str = ""):
    conn = get_connection()
    conn.execute(
        "UPDATE outsources SET name = ?, contact_info = ? WHERE id = ?",
        (name, contact_info, outsource_id),
    )
    conn.commit()
    conn.close()


def delete_outsource(outsource_id: int):
    conn = get_connection()
    conn.execute("DELETE FROM outsource_bluebooks WHERE outsource_id = ?", (outsource_id,))
    conn.execute("DELETE FROM outsources WHERE id = ?", (outsource_id,))
    conn.commit()
    conn.close()


# ──────────────────────────────────────
# Bluebooks
# ──────────────────────────────────────

def add_bluebook(die_number: str, description: str = "") -> int:
    conn = get_connection()
    cur = conn.execute(
        "INSERT INTO bluebooks (die_number, description) VALUES (?, ?)",
        (die_number, description),
    )
    conn.commit()
    bid = cur.lastrowid
    conn.close()
    return bid


def get_bluebook(bluebook_id: int) -> Optional[Bluebook]:
    conn = get_connection()
    row = conn.execute("SELECT * FROM bluebooks WHERE id = ?", (bluebook_id,)).fetchone()
    if not row:
        conn.close()
        return None
    bb = Bluebook(id=row["id"], die_number=row["die_number"],
                  description=row["description"],
                  created_at=row["created_at"],
                  updated_at=row["updated_at"])
    # Fetch linked customer names
    crows = conn.execute("""
        SELECT c.name FROM customers c
        JOIN customer_bluebooks cb ON c.id = cb.customer_id
        WHERE cb.bluebook_id = ?
        ORDER BY c.name
    """, (bluebook_id,)).fetchall()
    bb.customer_names = [cr["name"] for cr in crows]
    # Fetch linked outsource names
    orows = conn.execute("""
        SELECT o.name FROM outsources o
        JOIN outsource_bluebooks ob ON o.id = ob.outsource_id
        WHERE ob.bluebook_id = ?
        ORDER BY o.name
    """, (bluebook_id,)).fetchall()
    bb.outsource_names = [orow["name"] for orow in orows]
    conn.close()
    return bb


def get_bluebook_by_die(die_number: str) -> Optional[Bluebook]:
    conn = get_connection()
    row = conn.execute("SELECT * FROM bluebooks WHERE die_number = ?",
                       (die_number,)).fetchone()
    if not row:
        conn.close()
        return None
    bb = Bluebook(id=row["id"], die_number=row["die_number"],
                  description=row["description"],
                  created_at=row["created_at"],
                  updated_at=row["updated_at"])
    conn.close()
    return bb


def list_bluebooks(search: str = "", customer_id: Optional[int] = None,
                   limit: int = 200,
                   search_description: bool = False) -> list[Bluebook]:
    """List bluebooks with optional die-number or description search and customer filter."""
    conn = get_connection()

    search_col = "b.description" if search_description else "b.die_number"
    search_col_noalias = "description" if search_description else "die_number"

    if customer_id:
        query = """
            SELECT DISTINCT b.* FROM bluebooks b
            JOIN customer_bluebooks cb ON b.id = cb.bluebook_id
            WHERE cb.customer_id = ?
        """
        params: list = [customer_id]
        if search:
            query += f" AND {search_col} LIKE ?"
            params.append(f"%{search}%")
        query += " ORDER BY b.die_number"
        if limit:
            query += f" LIMIT {limit}"
        rows = conn.execute(query, params).fetchall()
    else:
        if search:
            rows = conn.execute(
                f"SELECT * FROM bluebooks WHERE {search_col_noalias} LIKE ? ORDER BY die_number LIMIT {limit}",
                (f"%{search}%",)
            ).fetchall()
        else:
            rows = conn.execute(
                f"SELECT * FROM bluebooks ORDER BY die_number LIMIT {limit}"
            ).fetchall()

    # Batch-fetch all customer names in ONE query instead of N+1
    bb_ids = [r["id"] for r in rows]
    customer_map: dict[int, list[str]] = {}
    if bb_ids:
        placeholders = ",".join("?" * len(bb_ids))
        crows = conn.execute(f"""
            SELECT cb.bluebook_id, c.name
            FROM customers c
            JOIN customer_bluebooks cb ON c.id = cb.customer_id
            WHERE cb.bluebook_id IN ({placeholders})
            ORDER BY c.name
        """, bb_ids).fetchall()
        for cr in crows:
            customer_map.setdefault(cr["bluebook_id"], []).append(cr["name"])

    # Batch-fetch all outsource names in ONE query
    outsource_map: dict[int, list[str]] = {}
    if bb_ids:
        orows = conn.execute(f"""
            SELECT ob.bluebook_id, o.name
            FROM outsources o
            JOIN outsource_bluebooks ob ON o.id = ob.outsource_id
            WHERE ob.bluebook_id IN ({placeholders})
            ORDER BY o.name
        """, bb_ids).fetchall()
        for orow in orows:
            outsource_map.setdefault(orow["bluebook_id"], []).append(orow["name"])

    bluebooks = []
    for r in rows:
        bb = Bluebook(id=r["id"], die_number=r["die_number"],
                      description=r["description"],
                      created_at=r["created_at"],
                      updated_at=r["updated_at"])
        bb.customer_names = customer_map.get(bb.id, [])
        bb.outsource_names = outsource_map.get(bb.id, [])
        bluebooks.append(bb)

    conn.close()
    return bluebooks


def update_bluebook(bluebook_id: int, die_number: str, description: str = ""):
    conn = get_connection()
    conn.execute(
        "UPDATE bluebooks SET die_number = ?, description = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
        (die_number, description, bluebook_id),
    )
    conn.commit()
    conn.close()


def delete_bluebook(bluebook_id: int):
    conn = get_connection()
    conn.execute("DELETE FROM shared_files_map WHERE linked_bluebook_id = ?", (bluebook_id,))
    # Delete shared refs where original file belongs to this bluebook
    conn.execute("""
        DELETE FROM shared_files_map WHERE original_file_id IN (
            SELECT id FROM bluebook_files WHERE bluebook_id = ?
        )
    """, (bluebook_id,))
    conn.execute("DELETE FROM bluebook_files WHERE bluebook_id = ?", (bluebook_id,))
    conn.execute("DELETE FROM customer_bluebooks WHERE bluebook_id = ?", (bluebook_id,))
    conn.execute("DELETE FROM outsource_bluebooks WHERE bluebook_id = ?", (bluebook_id,))
    conn.execute("DELETE FROM bluebooks WHERE id = ?", (bluebook_id,))
    conn.commit()
    conn.close()


# ──────────────────────────────────────
# Customer ↔ Bluebook Links
# ──────────────────────────────────────

def link_customer_bluebook(customer_id: int, bluebook_id: int):
    conn = get_connection()
    try:
        conn.execute(
            "INSERT OR IGNORE INTO customer_bluebooks (customer_id, bluebook_id) VALUES (?, ?)",
            (customer_id, bluebook_id),
        )
        conn.commit()
    finally:
        conn.close()


def unlink_customer_bluebook(customer_id: int, bluebook_id: int):
    conn = get_connection()
    conn.execute(
        "DELETE FROM customer_bluebooks WHERE customer_id = ? AND bluebook_id = ?",
        (customer_id, bluebook_id),
    )
    conn.commit()
    conn.close()


def get_bluebooks_for_customer(customer_id: int) -> list[Bluebook]:
    return list_bluebooks(customer_id=customer_id)


def get_customers_for_bluebook(bluebook_id: int) -> list[Customer]:
    conn = get_connection()
    rows = conn.execute("""
        SELECT c.* FROM customers c
        JOIN customer_bluebooks cb ON c.id = cb.customer_id
        WHERE cb.bluebook_id = ?
        ORDER BY c.name
    """, (bluebook_id,)).fetchall()
    conn.close()
    return [Customer(id=r["id"], name=r["name"],
                     contact_info=r["contact_info"],
                     created_at=r["created_at"]) for r in rows]


# ──────────────────────────────────────
# Outsource ↔ Bluebook Links
# ──────────────────────────────────────

def link_outsource_bluebook(outsource_id: int, bluebook_id: int):
    conn = get_connection()
    try:
        conn.execute(
            "INSERT OR IGNORE INTO outsource_bluebooks (outsource_id, bluebook_id) VALUES (?, ?)",
            (outsource_id, bluebook_id),
        )
        conn.commit()
    finally:
        conn.close()


def unlink_outsource_bluebook(outsource_id: int, bluebook_id: int):
    conn = get_connection()
    conn.execute(
        "DELETE FROM outsource_bluebooks WHERE outsource_id = ? AND bluebook_id = ?",
        (outsource_id, bluebook_id),
    )
    conn.commit()
    conn.close()


def get_outsources_for_bluebook(bluebook_id: int) -> list[Outsource]:
    conn = get_connection()
    rows = conn.execute("""
        SELECT o.* FROM outsources o
        JOIN outsource_bluebooks ob ON o.id = ob.outsource_id
        WHERE ob.bluebook_id = ?
        ORDER BY o.name
    """, (bluebook_id,)).fetchall()
    conn.close()
    return [Outsource(id=r["id"], name=r["name"],
                      contact_info=r["contact_info"],
                      created_at=r["created_at"]) for r in rows]


# ──────────────────────────────────────
# Bluebook Files
# ──────────────────────────────────────

def add_bluebook_file(bluebook_id: int, section_type: str,
                      file_path: str, display_order: int = 0) -> int:
    conn = get_connection()
    cur = conn.execute(
        "INSERT INTO bluebook_files (bluebook_id, section_type, file_path, display_order) "
        "VALUES (?, ?, ?, ?)",
        (bluebook_id, section_type, file_path, display_order),
    )
    conn.commit()
    fid = cur.lastrowid
    conn.close()
    return fid


def get_bluebook_file(file_id: int) -> Optional[BluebookFile]:
    conn = get_connection()
    row = conn.execute("SELECT * FROM bluebook_files WHERE id = ?", (file_id,)).fetchone()
    conn.close()
    if row:
        return BluebookFile(
            id=row["id"], bluebook_id=row["bluebook_id"],
            section_type=row["section_type"], file_path=row["file_path"],
            display_order=row["display_order"], created_at=row["created_at"],
        )
    return None


def get_files_for_bluebook(bluebook_id: int,
                           section_type: Optional[str] = None) -> list[BluebookFile]:
    """Get own files + shared files for a bluebook, optionally filtered by section."""
    conn = get_connection()

    # Own files
    if section_type:
        own_rows = conn.execute(
            "SELECT * FROM bluebook_files WHERE bluebook_id = ? AND section_type = ? "
            "ORDER BY display_order, id",
            (bluebook_id, section_type),
        ).fetchall()
    else:
        own_rows = conn.execute(
            "SELECT * FROM bluebook_files WHERE bluebook_id = ? "
            "ORDER BY section_type, display_order, id",
            (bluebook_id,),
        ).fetchall()

    files = []
    for r in own_rows:
        bf = BluebookFile(
            id=r["id"], bluebook_id=r["bluebook_id"],
            section_type=r["section_type"], file_path=r["file_path"],
            display_order=r["display_order"], created_at=r["created_at"],
            is_shared=False,
        )
        files.append(bf)

    # Shared files (files from other bluebooks shared into this one)
    if section_type:
        shared_rows = conn.execute("""
            SELECT bf.*, sfm.id as sfm_id, b.die_number as source_die
            FROM shared_files_map sfm
            JOIN bluebook_files bf ON sfm.original_file_id = bf.id
            JOIN bluebooks b ON bf.bluebook_id = b.id
            WHERE sfm.linked_bluebook_id = ? AND bf.section_type = ?
            ORDER BY bf.display_order, bf.id
        """, (bluebook_id, section_type)).fetchall()
    else:
        shared_rows = conn.execute("""
            SELECT bf.*, sfm.id as sfm_id, b.die_number as source_die
            FROM shared_files_map sfm
            JOIN bluebook_files bf ON sfm.original_file_id = bf.id
            JOIN bluebooks b ON bf.bluebook_id = b.id
            WHERE sfm.linked_bluebook_id = ?
            ORDER BY bf.section_type, bf.display_order, bf.id
        """, (bluebook_id,)).fetchall()

    for r in shared_rows:
        bf = BluebookFile(
            id=r["id"], bluebook_id=r["bluebook_id"],
            section_type=r["section_type"], file_path=r["file_path"],
            display_order=r["display_order"], created_at=r["created_at"],
            is_shared=True,
            shared_from_die_number=r["source_die"],
        )
        files.append(bf)

    conn.close()
    return files


def delete_bluebook_file(file_id: int):
    conn = get_connection()
    # Remove sharing refs first
    conn.execute("DELETE FROM shared_files_map WHERE original_file_id = ?", (file_id,))
    conn.execute("DELETE FROM bluebook_files WHERE id = ?", (file_id,))
    conn.commit()
    conn.close()


def update_bluebook_file_path(file_id: int, new_path: str):
    conn = get_connection()
    conn.execute(
        "UPDATE bluebook_files SET file_path = ? WHERE id = ?",
        (new_path, file_id),
    )
    conn.commit()
    conn.close()

def get_file_counts_batch(bb_ids: list[int]) -> dict[int, int]:
    """Get file counts for multiple bluebooks in one query."""
    if not bb_ids:
        return {}
    conn = get_connection()
    placeholders = ",".join("?" * len(bb_ids))
    rows = conn.execute(f"""
        SELECT bluebook_id, COUNT(*) as cnt
        FROM bluebook_files
        WHERE bluebook_id IN ({placeholders})
        GROUP BY bluebook_id
    """, bb_ids).fetchall()
    conn.close()
    return {r["bluebook_id"]: r["cnt"] for r in rows}


def get_next_display_order(bluebook_id: int, section_type: str) -> int:
    conn = get_connection()
    row = conn.execute(
        "SELECT MAX(display_order) as mx FROM bluebook_files "
        "WHERE bluebook_id = ? AND section_type = ?",
        (bluebook_id, section_type),
    ).fetchone()
    conn.close()
    return (row["mx"] or 0) + 1


# ──────────────────────────────────────
# Shared Files
# ──────────────────────────────────────

def add_shared_file(original_file_id: int, linked_bluebook_id: int) -> int:
    conn = get_connection()
    cur = conn.execute(
        "INSERT OR IGNORE INTO shared_files_map (original_file_id, linked_bluebook_id) "
        "VALUES (?, ?)",
        (original_file_id, linked_bluebook_id),
    )
    conn.commit()
    sid = cur.lastrowid
    conn.close()
    return sid


def remove_shared_file(original_file_id: int, linked_bluebook_id: int):
    conn = get_connection()
    conn.execute(
        "DELETE FROM shared_files_map WHERE original_file_id = ? AND linked_bluebook_id = ?",
        (original_file_id, linked_bluebook_id),
    )
    conn.commit()
    conn.close()


def remove_all_shared_refs(original_file_id: int):
    conn = get_connection()
    conn.execute("DELETE FROM shared_files_map WHERE original_file_id = ?",
                 (original_file_id,))
    conn.commit()
    conn.close()


def get_shared_targets(original_file_id: int) -> list[Bluebook]:
    """Get all bluebooks that a file is shared to."""
    conn = get_connection()
    rows = conn.execute("""
        SELECT b.* FROM bluebooks b
        JOIN shared_files_map sfm ON b.id = sfm.linked_bluebook_id
        WHERE sfm.original_file_id = ?
        ORDER BY b.die_number
    """, (original_file_id,)).fetchall()
    conn.close()
    return [Bluebook(id=r["id"], die_number=r["die_number"],
                     description=r["description"]) for r in rows]


def is_file_shared(file_id: int) -> bool:
    conn = get_connection()
    row = conn.execute(
        "SELECT COUNT(*) as cnt FROM shared_files_map WHERE original_file_id = ?",
        (file_id,),
    ).fetchone()
    conn.close()
    return row["cnt"] > 0


# ──────────────────────────────────────
# Action Log
# ──────────────────────────────────────

def log_action(action: str, details: str = ""):
    conn = get_connection()
    conn.execute(
        "INSERT INTO action_log (action, details) VALUES (?, ?)",
        (action, details),
    )
    conn.commit()
    conn.close()


def get_recent_logs(limit: int = 100) -> list[ActionLog]:
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM action_log ORDER BY timestamp DESC LIMIT ?", (limit,)
    ).fetchall()
    conn.close()
    return [ActionLog(id=r["id"], action=r["action"],
                      details=r["details"],
                      timestamp=r["timestamp"]) for r in rows]


# ──────────────────────────────────────
# QA Counter
# ──────────────────────────────────────

def get_next_qa_number(year: int) -> int:
    """Find the lowest available QA number for a given year by scanning existing files.

    Looks at all quality_alerts filenames matching QA-YY-NNN-* to collect
    used numbers, then returns the smallest gap starting from 1.
    Deleting a QA file automatically frees its number for reuse.
    """
    import re
    conn = get_connection()
    rows = conn.execute(
        "SELECT file_path FROM bluebook_files WHERE section_type = 'quality_alerts'"
    ).fetchall()
    conn.close()

    pattern = re.compile(rf"QA-{year:02d}-(\d{{3}})", re.IGNORECASE)
    used: set[int] = set()
    for r in rows:
        filename = r["file_path"].rsplit("\\", 1)[-1].rsplit("/", 1)[-1]
        m = pattern.match(filename)
        if m:
            used.add(int(m.group(1)))

    # Find lowest available number starting from 1
    num = 1
    while num in used:
        num += 1
    return num


def get_qa_counter(year: int) -> int:
    """Get the current QA counter for a year (without incrementing)."""
    conn = get_connection()
    row = conn.execute(
        "SELECT last_number FROM qa_counter WHERE year = ?", (year,)).fetchone()
    conn.close()
    return row["last_number"] if row else 0


def set_qa_counter(year: int, number: int):
    """Set the QA counter for a year to a specific value."""
    conn = get_connection()
    conn.execute(
        "INSERT INTO qa_counter (year, last_number) VALUES (?, ?) "
        "ON CONFLICT(year) DO UPDATE SET last_number = ?",
        (year, number, number))
    conn.commit()
    conn.close()
