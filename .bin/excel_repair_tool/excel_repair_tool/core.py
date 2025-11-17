
"""
ExcelRepairTool Core Module
Enhanced: Checks for named ranges, slicers, data sources, and sharedStrings column references.
"""

import zipfile
import shutil
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET
from .utils import safe_copy, ensure_dir, timestamp, save_json_log, extract_text_from_shared_strings
from .constants import XML_NAMESPACES


class ExcelRepairTool:
    def __init__(self, xlsx_path: str):
        self.source_path = Path(xlsx_path).resolve()
        if not self.source_path.exists():
            raise FileNotFoundError(f"Excel file not found: {xlsx_path}")

        self.temp_dir = Path(tempfile.mkdtemp(prefix="excel_repair_"))
        ensure_dir(self.temp_dir)

        self.working_copy = self.temp_dir / f"{self.source_path.stem}_copy.xlsx"
        safe_copy(self.source_path, self.working_copy)

        self.extracted_dir = self.temp_dir / "unzipped"
        self.corrupt_files = []
        self.log_entries = []
        self.json_report = {
            "summary": {},
            "corrupt_files": [],
            "named_ranges": [],
            "slicers": [],
            "data_connections": [],
            "media_issues": [],
            "shared_string_columns": [],
            "pivot_fields": [],
        }

    def log(self, message: str, to_console: bool = True):
        self.log_entries.append(message)
        if to_console:
            print(message)

    def verify_corruption(self) -> bool:
        self.log("🔍 Verifying file integrity...")
        try:
            with zipfile.ZipFile(self.working_copy, 'r') as zip_ref:
                corrupt_member = zip_ref.testzip()
                if corrupt_member:
                    self.log(f"[❌] Corrupt ZIP member found: {corrupt_member}")
                    return True
        except zipfile.BadZipFile:
            self.log("[❌] File is not a valid zip (fatal corruption).")
            return True
        self.log("[✅] File is a valid ZIP container.")
        return False

    def unzip_xlsx(self):
        self.log(f"📦 Extracting workbook contents to: {self.extracted_dir}")
        with zipfile.ZipFile(self.working_copy, 'r') as zip_ref:
            zip_ref.extractall(self.extracted_dir)

    def scan_for_corruption(self):
        self.log("🧪 Scanning for XML corruption...")
        for file_path in self.extracted_dir.rglob("*.xml"):
            try:
                ET.parse(file_path)
            except ET.ParseError as e:
                self.corrupt_files.append(file_path)
                self.json_report["corrupt_files"].append(str(file_path.relative_to(self.extracted_dir)))
                self.log(f"[❌] {file_path.relative_to(self.extracted_dir)} — {e}")
        self.json_report["summary"]["corrupt_count"] = len(self.corrupt_files)

    def analyze_dependencies(self):
        self.log("🔗 Analyzing named ranges, slicers, connections, pivots...")
        workbook_path = self.extracted_dir / "xl" / "workbook.xml"
        if workbook_path.exists():
            try:
                tree = ET.parse(workbook_path)
                root = tree.getroot()
                defined_names = root.findall(".//a:definedNames/a:definedName", XML_NAMESPACES)
                for dn in defined_names:
                    self.json_report["named_ranges"].append({
                        "name": dn.attrib.get("name"),
                        "value": dn.text
                    })
            except ET.ParseError:
                self.log("⚠️ workbook.xml could not be parsed.")

        conn_path = self.extracted_dir / "xl" / "connections.xml"
        if conn_path.exists():
            try:
                tree = ET.parse(conn_path)
                root = tree.getroot()
                for conn in root.findall(".//a:connection", XML_NAMESPACES):
                    self.json_report["data_connections"].append(conn.attrib)
            except ET.ParseError:
                self.log("⚠️ connections.xml not parseable.")

        slicer_dir = self.extracted_dir / "xl" / "slicerCaches"
        if slicer_dir.exists():
            for slicer in slicer_dir.glob("*.xml"):
                self.json_report["slicers"].append(str(slicer.relative_to(self.extracted_dir)))

        pivot_dir = self.extracted_dir / "xl" / "pivotCache"
        if pivot_dir.exists():
            for pivot in pivot_dir.glob("pivotCacheDefinition*.xml"):
                try:
                    tree = ET.parse(pivot)
                    fields = [f.attrib.get("name") for f in tree.getroot().findall(".//a:cacheField", XML_NAMESPACES)]
                    self.json_report["pivot_fields"].append({
                        "file": str(pivot.relative_to(self.extracted_dir)),
                        "fields": fields
                    })
                except ET.ParseError:
                    self.log(f"[⚠️] Could not parse {pivot.name}")

    def validate_shared_strings(self):
        shared_path = self.extracted_dir / "xl" / "sharedStrings.xml"
        strings = extract_text_from_shared_strings(shared_path)
        self.json_report["shared_string_columns"] = strings[:100]  # Limit output for large files

    def validate_media(self):
        media_dir = self.extracted_dir / "xl" / "media"
        if media_dir.exists():
            for media_file in media_dir.glob("*"):
                if media_file.stat().st_size == 0:
                    self.json_report["media_issues"].append(str(media_file.name))
                    self.log(f"[⚠️] Empty media file: {media_file.name}")

    def repair_corruption(self):
        self.log("🛠 Repairing corrupted XML files...")
        for file_path in self.corrupt_files:
            repaired_rows = []
            inside_row = False
            with file_path.open("r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    if "<row" in line:
                        inside_row = True
                        repaired_rows.append(line)
                    elif "</row>" in line:
                        inside_row = False
                        repaired_rows.append(line)
                    elif inside_row:
                        repaired_rows.append(line)
            content = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
                "<sheetData>\n" + "".join(repaired_rows) + "</sheetData>\n</worksheet>"
            )
            file_path.write_text(content, encoding="utf-8")
            self.log(f"[✅] Repaired: {file_path.name}")

    def repackage_xlsx(self, output_path=None):
        output_path = Path(output_path or (self.source_path.parent / f"{self.source_path.stem}_repaired.xlsx"))
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            for file in self.extracted_dir.rglob("*"):
                if file.is_file():
                    zipf.write(file, file.relative_to(self.extracted_dir))
        self.log(f"📦 Repackaged workbook: {output_path}")
        self.json_report["summary"]["output_file"] = str(output_path)

    def run(self, repair=True, json_log=False):
        try:
            if self.verify_corruption():
                self.log("⚠️ ZIP corruption detected.")
            self.unzip_xlsx()
            self.scan_for_corruption()
            self.analyze_dependencies()
            self.validate_shared_strings()
            self.validate_media()
            if repair and self.corrupt_files:
                self.repair_corruption()
            self.repackage_xlsx()
            if json_log:
                log_path = self.source_path.with_name(self.source_path.stem + "_diagnostic.json")
                save_json_log(self.json_report, log_path)
                self.log(f"📝 JSON report saved to: {log_path}")
        finally:
            self.log(f"🧹 Cleaning up: {self.temp_dir}")
            shutil.rmtree(self.temp_dir, ignore_errors=True)
