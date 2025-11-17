
"""Utility helpers for ExcelRepairTool."""

import shutil
import json
from datetime import datetime
from pathlib import Path


def ensure_dir(path: Path):
    """Ensure directory exists."""
    path.mkdir(parents=True, exist_ok=True)


def safe_copy(src: Path, dst: Path):
    """Safely copy a file to destination."""
    shutil.copy2(src, dst)


def timestamp():
    """Return current timestamp string."""
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


def save_json_log(data, output_path):
    """Save structured data to a JSON file."""
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def extract_text_from_shared_strings(shared_path):
    """
    Return list of all string values from sharedStrings.xml
    """
    from xml.etree import ElementTree as ET
    if not shared_path.exists():
        return []

    try:
        tree = ET.parse(shared_path)
        root = tree.getroot()
        texts = [t.text for t in root.iter() if t.tag.endswith('t') and t.text]
        return texts
    except ET.ParseError:
        return []
