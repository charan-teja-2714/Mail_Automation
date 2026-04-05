"""
utils.py - Shared utility functions for the Weekly Sanity Report project.
"""

import os
import logging

# Configure logging once for the entire project
logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(message)s"
)
logger = logging.getLogger(__name__)


# Status normalization map (case-insensitive)
STATUS_VARIANTS = {
    "done":    "Done",
    "ok":      "Done",
    "pass":    "Done",
    "passed":  "Done",
    "warning": "Warning",
    "warn":    "Warning",
    "caution": "Warning",
    "failed":  "Failed",
    "fail":    "Failed",
    "error":   "Failed",
    "na":      "N/A",
    "n/a":     "N/A",
}

# Status → CSS inline color (for HTML)
STATUS_COLORS_HTML = {
    "Done":    "#1a7a1a",   # dark green
    "Warning": "#b35900",   # dark orange
    "Failed":  "#cc0000",   # dark red
    "N/A":     "#555555",   # grey
}

# Status → background badge color (for HTML cells)
STATUS_BG_HTML = {
    "Done":    "#d4edda",
    "Warning": "#fff3cd",
    "Failed":  "#f8d7da",
    "N/A":     "#e9ecef",
}

# Status → python-docx RGBColor tuple
STATUS_COLORS_DOCX = {
    "Done":    (0,   128,  0),   # green
    "Warning": (204, 102,  0),   # orange
    "Failed":  (204,   0,  0),   # red
    "N/A":     (100, 100, 100),  # grey
}


def normalize_status(raw_status: str) -> str:
    """
    Normalize a raw status string to one of: Done, Warning, Failed, N/A.
    Defaults to 'N/A' if unrecognized.
    """
    if not isinstance(raw_status, str):
        return "N/A"
    key = raw_status.strip().lower()
    return STATUS_VARIANTS.get(key, raw_status.strip() or "N/A")


def resolve_path(path: str, base_dir: str = None) -> str:
    """
    Resolve a file path.  If the path is already absolute, return it as-is.
    Otherwise resolve it relative to base_dir (defaults to the project root,
    i.e. the directory containing this file).
    """
    if os.path.isabs(path):
        return path
    root = base_dir or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(root, path)


def image_exists(image_path: str, base_dir: str = None) -> str | None:
    """
    Return the absolute path if the image file exists, otherwise None.
    Accepts both absolute paths and paths relative to base_dir.
    """
    full = resolve_path(image_path, base_dir)
    if os.path.isfile(full):
        return full
    logger.warning("Image not found, skipping: %s", full)
    return None
