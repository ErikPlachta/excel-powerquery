
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
