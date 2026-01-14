# Power-BI-Project-Formatter

This application is an automated tool to convert Microsoft Power BI PBIX files into PBIP project files. PBIP (Power BI Project) format stores reports and semantic models as human-readable files suitable for version control.

## Prerequisites

### Power BI Desktop
Power BI Desktop must be installed on your system. Download from:
https://powerbi.microsoft.com/desktop/

### Enable PBIP Preview Feature
The PBIP save option is a preview feature that must be enabled:

1. Open Power BI Desktop
2. Go to **File > Options and settings > Options**
3. Select **Preview features**
4. Check **Power BI Project (.pbip) save option**
5. Restart Power BI Desktop

### Python Dependencies
Install the required Python packages:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage
Navigate to a directory containing PBIX files and run:

```bash
python pbix_converter.py
```

### Specify Directory
You can also specify a directory path:

```bash
python pbix_converter.py "C:\Path\To\PBIX\Files"
```

### Interactive Mode
The tool will:
1. Scan for PBIX files in the directory
2. Display found files with sizes
3. Prompt you to select which files to convert
4. Create a `pbip_output/` subfolder with converted projects

### Example Session
```
==================================================
  Power BI PBIX to PBIP Converter
==================================================

Working directory: C:\Reports

Found 3 PBIX file(s):

  1. SalesReport.pbix (15.2 MB)
  2. InventoryDashboard.pbix (8.7 MB)
  3. FinanceMetrics.pbix (22.1 MB)

Options:
  [A] Convert all files
  [S] Select specific files (enter numbers separated by commas)
  [Q] Quit

Your choice: A
```

## Output Structure

For each PBIX file, the converter creates a project folder:

```
pbip_output/
├── SalesReport/
│   ├── SalesReport.pbip
│   ├── SalesReport.Report/
│   │   ├── definition.pbir
│   │   ├── report.json
│   │   └── ...
│   └── SalesReport.SemanticModel/
│       ├── definition.pbism
│       ├── definition/
│       │   └── (TMDL files)
│       └── ...
└── ...
```

## Important Notes

- **Do not use mouse/keyboard** during conversion - the tool automates Power BI Desktop's UI
- Only one PBIX file is converted at a time
- Close Power BI Desktop before starting
- Conversion time depends on file size and complexity

## Troubleshooting

### Power BI Desktop not found
Ensure Power BI Desktop is installed in a standard location or add it to your PATH.

### PBIP option not available
Enable the preview feature as described in Prerequisites.

### Conversion fails
- Ensure no other Power BI files are open
- Check that the PBIX file is not corrupted
- Verify sufficient disk space for output
