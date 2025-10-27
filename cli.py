
"""Command-line interface for ExcelRepairTool."""

import argparse
from .core import ExcelRepairTool


def main():
    parser = argparse.ArgumentParser(
        description="ExcelRepairTool - Detect and repair corrupted Excel (.xlsx) files."
    )
    parser.add_argument("file", help="Path to the Excel file to analyze.")
    parser.add_argument("--no-repair", action="store_true", help="Only analyze, do not attempt repair.")
    parser.add_argument("--json-log", action="store_true", help="Output a JSON report.")
    args = parser.parse_args()

    tool = ExcelRepairTool(args.file)
    tool.run(repair=not args.no_repair, json_log=args.json_log)


if __name__ == "__main__":
    main()
