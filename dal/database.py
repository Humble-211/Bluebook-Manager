"""
Bluebook Manager — Database initialization and connection management.
"""

import os
import sqlite3

from config import DB_PATH


def get_connection() -> sqlite3.Connection:
    """Return a connection to the SQLite database, creating it if needed."""
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    """Create all tables if they don't exist."""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.executescript("""
        CREATE TABLE IF NOT EXISTS customers (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            name          TEXT NOT NULL UNIQUE,
            contact_info  TEXT,
            created_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS bluebooks (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            die_number    TEXT NOT NULL UNIQUE,
            description   TEXT,
            created_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS customer_bluebooks (
            customer_id   INTEGER REFERENCES customers(id) ON DELETE CASCADE,
            bluebook_id   INTEGER REFERENCES bluebooks(id) ON DELETE CASCADE,
            PRIMARY KEY (customer_id, bluebook_id)
        );

        CREATE TABLE IF NOT EXISTS bluebook_files (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            bluebook_id   INTEGER REFERENCES bluebooks(id) ON DELETE CASCADE,
            section_type  TEXT NOT NULL,
            file_path     TEXT NOT NULL,
            display_order INTEGER DEFAULT 0,
            created_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS shared_files_map (
            id                 INTEGER PRIMARY KEY AUTOINCREMENT,
            original_file_id   INTEGER REFERENCES bluebook_files(id) ON DELETE CASCADE,
            linked_bluebook_id INTEGER REFERENCES bluebooks(id) ON DELETE CASCADE,
            shared_at          TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(original_file_id, linked_bluebook_id)
        );

        CREATE TABLE IF NOT EXISTS action_log (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            action     TEXT NOT NULL,
            details    TEXT,
            timestamp  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS qa_counter (
            year         INTEGER PRIMARY KEY,
            last_number  INTEGER NOT NULL DEFAULT 0
        );

        CREATE INDEX IF NOT EXISTS idx_bluebooks_die_number
            ON bluebooks(die_number);
        CREATE INDEX IF NOT EXISTS idx_customer_bluebooks_bluebook
            ON customer_bluebooks(bluebook_id);
        CREATE INDEX IF NOT EXISTS idx_bluebook_files_bluebook
            ON bluebook_files(bluebook_id);
        CREATE INDEX IF NOT EXISTS idx_bluebook_files_section
            ON bluebook_files(bluebook_id, section_type);
        CREATE INDEX IF NOT EXISTS idx_shared_files_original
            ON shared_files_map(original_file_id);
        CREATE INDEX IF NOT EXISTS idx_shared_files_linked
            ON shared_files_map(linked_bluebook_id);
    """)

    conn.commit()
    conn.close()
