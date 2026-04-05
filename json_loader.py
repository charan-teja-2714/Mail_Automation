"""
json_loader.py - Load and validate activity data from a JSON file.
"""

import json
import os
from utils import logger, normalize_status

# Fields that every activity record must contain
REQUIRED_FIELDS = {"sno", "activity", "status"}

# Fields used by the DOCX generator (optional but validated when present)
OPTIONAL_FIELDS = {"doc_title", "doc_description", "image"}


def load_activities(json_path: str) -> list[dict]:
    """
    Load the JSON file at *json_path*, validate each record, and return
    a list of cleaned activity dictionaries.

    Raises:
        FileNotFoundError  – if the file does not exist.
        ValueError         – if the top-level JSON is not a list.
    """
    if not os.path.isfile(json_path):
        raise FileNotFoundError(f"JSON file not found: {json_path}")

    logger.info("Loading JSON data from: %s", json_path)

    with open(json_path, "r", encoding="utf-8") as fh:
        try:
            raw = json.load(fh)
        except json.JSONDecodeError as exc:
            raise ValueError(f"Invalid JSON in '{json_path}': {exc}") from exc

    if not isinstance(raw, list):
        raise ValueError(
            f"Expected a JSON array at the top level, got {type(raw).__name__}."
        )

    activities = []
    for idx, record in enumerate(raw, start=1):
        cleaned = _validate_record(record, idx)
        if cleaned is not None:
            activities.append(cleaned)

    logger.info("Loaded %d valid activity record(s).", len(activities))
    return activities


def _validate_record(record: dict, position: int) -> dict | None:
    """
    Validate a single record dict.  Returns the cleaned dict, or None if the
    record should be skipped (missing required fields).
    """
    if not isinstance(record, dict):
        logger.warning("Record #%d is not an object – skipping.", position)
        return None

    # Check required fields
    missing = REQUIRED_FIELDS - record.keys()
    if missing:
        logger.warning(
            "Record #%d is missing required field(s) %s – skipping.",
            position, missing
        )
        return None

    # Normalize status
    record["status"] = normalize_status(record["status"])

    # Fill optional fields with sensible defaults
    record.setdefault("doc_title", str(record.get("activity", f"Activity {position}")))
    record.setdefault("doc_description", "No description provided.")
    record.setdefault("image", None)

    # Ensure sno is an integer
    try:
        record["sno"] = int(record["sno"])
    except (TypeError, ValueError):
        record["sno"] = position

    return record
