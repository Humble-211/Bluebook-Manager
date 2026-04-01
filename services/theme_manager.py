"""
Bluebook Manager — Theme Manager.

Manages application themes. Themes are QSS files stored in ui/resources/.
The selected theme is persisted in settings.json at the project root.

Usage:
    from services.theme_manager import ThemeManager
    tm = ThemeManager(app)
    tm.apply_theme("midnight")  # or "oceanic", "arctic", "ember"
"""

from __future__ import annotations

import json
import os
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from PySide6.QtWidgets import QApplication

from config import BUNDLE_DIR, BASE_DIR

SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")

# Theme registry: name -> (label, QSS filename)
THEMES: dict[str, tuple[str, str]] = {
    "midnight": ("Midnight", "theme_midnight.qss"),
    "oceanic":  ("Oceanic",  "theme_oceanic.qss"),
    "ember":    ("Ember",    "theme_ember.qss"),
    "arctic":   ("Arctic",   "theme_arctic.qss"),
}

DEFAULT_THEME = "midnight"


class ThemeManager:
    """Application-level theme controller."""

    def __init__(self, app: "QApplication"):
        self._app = app
        self._current = DEFAULT_THEME

    # ── Public API ──────────────────────────────────────────────

    @property
    def current(self) -> str:
        return self._current

    @property
    def theme_names(self) -> list[str]:
        return list(THEMES.keys())

    def label(self, name: str) -> str:
        return THEMES.get(name, (name, ""))[0]

    def load_saved(self):
        """Apply the last saved theme (falls back to default)."""
        saved = self._load_settings().get("theme", DEFAULT_THEME)
        if saved not in THEMES:
            saved = DEFAULT_THEME
        self.apply_theme(saved)

    def apply_theme(self, name: str):
        """Apply a theme by name and persist the choice."""
        if name not in THEMES:
            name = DEFAULT_THEME

        _, qss_file = THEMES[name]
        qss_path = os.path.join(BUNDLE_DIR, "ui", "resources", qss_file)

        if os.path.isfile(qss_path):
            with open(qss_path, "r", encoding="utf-8") as f:
                self._app.setStyleSheet(f.read())
            self._current = name
            self._save_settings({"theme": name})

    def next_theme(self) -> str:
        """Cycle to the next theme and apply it. Returns the new theme name."""
        keys = list(THEMES.keys())
        idx = (keys.index(self._current) + 1) % len(keys)
        next_name = keys[idx]
        self.apply_theme(next_name)
        return next_name

    # ── Persistence ─────────────────────────────────────────────

    def _load_settings(self) -> dict:
        try:
            if os.path.isfile(SETTINGS_FILE):
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def _save_settings(self, data: dict):
        try:
            existing = self._load_settings()
            existing.update(data)
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(existing, f, indent=2)
        except Exception:
            pass
