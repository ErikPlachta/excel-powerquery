from excel_repair_tool import ExcelRepairTool

# Path to your test .xlsx file
test_file = "test_files/CorruptWorkbook.xlsx"

# Initialize the tool
tool = ExcelRepairTool(test_file)

# Run full analysis and repair
tool.run(repair=True, json_log=True)