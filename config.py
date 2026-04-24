"""
Bluebook Manager — Configuration & Constants
"""

import os
import sys

# When running as a PyInstaller bundle, bundled data (templates, QSS) are in
# sys._MEIPASS, while user data (database, storage, logs) should be next to
# the .exe.  When running from source, both are the project root.
if getattr(sys, 'frozen', False):
    # Frozen exe: bundled assets extracted here
    BUNDLE_DIR = sys._MEIPASS
    # Persistent / user data lives next to the exe
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = BUNDLE_DIR

# Paths — user / runtime data (persist next to exe)
STORAGE_ROOT = os.path.join(BASE_DIR, "storage")
DB_PATH = os.path.join(BASE_DIR, "data", "bluebook.db")
LOG_DIR = os.path.join(BASE_DIR, "logs")
LOG_FILE = os.path.join(LOG_DIR, "bluebook_manager.log")

# Paths — bundled assets (inside the exe payload)
TEMPLATE_DIR = os.path.join(BUNDLE_DIR, "templates")
RESOURCE_DIR = os.path.join(BUNDLE_DIR, "resources")
LOGO_PATH = os.path.join(RESOURCE_DIR, "logo.png")


# Section type constants (order matters for printing)
SECTION_TYPES = [
    "cover",
    "master_drawings",
    "qc_drawings",
    "approval",
    "quality_alerts",
    "quality_notes",
    "packing_instruction",
    "fit_and_functions",
]

# Human-readable labels for sections
SECTION_LABELS = {
    "cover": "1. Front Page (Cover)",
    "master_drawings": "2. Master Drawing(s)",
    "qc_drawings": "3. QC Drawing(s)",
    "approval": "4. Approval Form",
    "quality_alerts": "5. Quality Alerts",
    "quality_notes": "6. Quality Notes / Requirements",
    "packing_instruction": "7. Packing Instruction",
    "fit_and_functions": "8. Fit and Functions",
}

# Disk subfolder names per section
SECTION_FOLDERS = {
    "cover": "01_Cover",
    "master_drawings": "02_MasterDrawings",
    "qc_drawings": "03_QCDrawings",
    "approval": "04_Approval",
    "quality_alerts": "05_QualityAlerts",
    "quality_notes": "06_QualityNotes",
    "packing_instruction": "07_PackingInstruction",
    "fit_and_functions": "08_FitAndFunctions",
}

# Allowed file types per section
SECTION_FILE_TYPES = {
    "cover": [".docx"],
    "master_drawings": [".pdf"],
    "qc_drawings": [".pdf"],
    "approval": [".pdf", ".docx"],
    "quality_alerts": [".docx"],
    "quality_notes": [".docx", ".pdf"],
    "packing_instruction": [".docx", ".jpg", ".png", ".jpeg", ".pdf"],
    "fit_and_functions": [".docx"],
}

# Sections that support creating from template
TEMPLATE_SECTIONS = ["cover", "quality_alerts", "quality_notes", "packing_instruction", "fit_and_functions"]

# Sections that support sharing
SHAREABLE_SECTIONS = ["quality_alerts", "quality_notes", "fit_and_functions", "packing_instruction"]

# Max files per section (0 = unlimited, but practically limited)
SECTION_MAX_FILES = {
    "cover": 0,
    "master_drawings": 0,
    "qc_drawings": 0,
    "approval": 0,
    "quality_alerts": 0,
    "quality_notes": 0,
    "packing_instruction": 0,
    "fit_and_functions": 0,
}
