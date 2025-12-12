# Utility Scripts Collection

A collection of independent Python utility scripts for data processing and conversion tasks.

## Scripts Overview

### 1. Confluence HTML to Markdown Converter

**Location:** `Conf. HTML to MD converter/recursive_html_converter.py`

A powerful tool for converting exported Confluence HTML documentation to clean Markdown format.

**Features:**
- Recursive conversion starting from an index.html file
- Two output modes:
  - Single combined Markdown file with hierarchical sections
  - Multiple Markdown files with hierarchical numbering (e.g., `01_guide.md`, `01.01_installation.md`)
- Automatic link updates to point to converted .md files
- Customizable HTML preprocessing to remove Confluence footers and breadcrumbs
- Smart header detection and Markdown postprocessing
- Timestamped output directories

**Usage:**
```bash
# Interactive mode
python recursive_html_converter.py

# Convert to multiple files
python recursive_html_converter.py index.html

# Convert to single combined file
python recursive_html_converter.py index.html -s

# Custom output with depth limit
python recursive_html_converter.py index.html -o output_docs -d 5
```

**Dependencies:**
- html2text
- Standard library: pathlib, argparse, re, json, datetime

---

### 2. Excel Metadata Extraction Tool

**Location:** `Extract sheet headers/app.py`

A comprehensive tool for extracting metadata from Excel files, CSV files, and other spreadsheet formats.

**Features:**
- Supports multiple file formats: `.xlsx`, `.xls`, `.xlsm`, `.csv`, `.txt`, `.ttx`, `.xml`
- Two extraction modes:
  - **Simple mode:** Extract file/sheet/column names
  - **Full mode:** Extract detailed metadata including headers, data types, formulas, and statistics
- Smart header detection with configurable scan depth
- Handles merged cells automatically
- Recursive folder scanning with filtering options
- File pattern matching (glob-style)
- YAML configuration file support for automated processing
- Outputs to both JSON and Excel formats

**Usage:**

Interactive mode:
```bash
python app.py
```

With configuration file:
```bash
# Edit config.yaml with your settings, then run:
python app.py
```

**Configuration:**
The tool uses `config.yaml` for automated processing. Key settings include:
- Target folder path and scanning mode
- File filtering patterns
- Extraction mode (simple/full)
- Output format preferences (JSON/Excel)

**Dependencies:**
- pandas
- openpyxl
- xlrd
- PyYAML
- Standard library: pathlib, json, csv, xml

---

## Installation

Each script can be used independently. Install the required dependencies for the script you want to use:

For the HTML to Markdown converter:
```bash
pip install html2text
```

For the Excel metadata extraction tool:
```bash
pip install pandas openpyxl xlrd pyyaml
```

## General Notes

- All scripts are standalone and can be moved/used independently
- Each script creates timestamped output directories to prevent overwriting previous results
- Both scripts support interactive mode and command-line arguments
- Configuration files and preprocessing functions can be customized for specific use cases

## Author

ralphxiaoz

## License

These scripts are provided as-is for utility purposes.
