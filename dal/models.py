"""
Bluebook Manager — Data models (dataclasses).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


@dataclass
class Customer:
    id: Optional[int] = None
    name: str = ""
    contact_info: str = ""
    created_at: Optional[datetime] = None


@dataclass
class Bluebook:
    id: Optional[int] = None
    die_number: str = ""
    description: str = ""
    created_at: Optional[datetime] = None
    updated_at: Optional[datetime] = None
    # Populated by queries — not stored directly
    customer_names: list[str] = field(default_factory=list)


@dataclass
class BluebookFile:
    id: Optional[int] = None
    bluebook_id: Optional[int] = None
    section_type: str = ""
    file_path: str = ""
    display_order: int = 0
    created_at: Optional[datetime] = None
    # Runtime attributes (not in DB directly)
    is_shared: bool = False
    shared_from_die_number: str = ""


@dataclass
class SharedFileMap:
    id: Optional[int] = None
    original_file_id: Optional[int] = None
    linked_bluebook_id: Optional[int] = None
    shared_at: Optional[datetime] = None


@dataclass
class ActionLog:
    id: Optional[int] = None
    action: str = ""
    details: str = ""
    timestamp: Optional[datetime] = None
