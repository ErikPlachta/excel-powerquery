
from setuptools import setup, find_packages

setup(
    name="excel_repair_tool",
    version="1.0.0",
    author="Your Name",
    description="Utility for analyzing and repairing corrupted Excel (.xlsx) files.",
    packages=find_packages(),
    python_requires=">=3.8",
    entry_points={
        "console_scripts": [
            "excel-repair=excel_repair_tool.cli:main",
        ],
    },
)
