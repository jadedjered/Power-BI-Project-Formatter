# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Python console application that converts Microsoft Power BI PBIX files to PBIP (Power BI Project) format by automating Power BI Desktop's UI.

**Key constraint**: There is no API or library for direct PBIX→PBIP conversion. The only method is automating Power BI Desktop's "File > Save As" dialog.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run the converter (current directory)
python pbix_converter.py

# Run with specific directory
python pbix_converter.py "C:\Path\To\PBIX\Files"
```

## Architecture

### File Structure
- `pbix_converter.py` - Main CLI entry point, handles user interaction and file discovery
- `pbi_automation.py` - Core UI automation module using pywinauto/pyautogui
- `requirements.txt` - Python dependencies (pywinauto, pyautogui, psutil)

### Key Functions in `pbi_automation.py`
- `find_pbi_desktop()` - Locates Power BI Desktop installation
- `convert_pbix_to_pbip()` - Main orchestration function for conversion
- `open_pbix()` - Launches PBI Desktop with a file
- `wait_for_pbi_ready()` - Waits for application to fully load (pywinauto window detection)
- `save_as_pbip()` - Automates the Save As dialog sequence
- `close_pbi_desktop()` - Graceful shutdown with force-kill fallback

### Automation Approach
Uses Windows UI Automation (pywinauto with UIA backend) for window detection and pyautogui for keyboard/mouse simulation:
1. Launch PBI Desktop with PBIX as argument
2. Wait for main window to contain filename in title
3. Send Alt+F for File menu, navigate to Save As
4. In Save dialog: set path, change file type to PBIP, click Save
5. Close PBI Desktop

## File Formats

- **PBIX**: Binary ZIP archive with compressed DataModel, Report, DataMashup components
- **PBIP**: Folder structure with `.pbip` pointer file, `.Report/` folder, `.SemanticModel/` folder containing TMDL files
- **TMDL**: Tabular Model Definition Language - human-readable model definitions

## Dependencies

- **pywinauto**: Windows UI automation for dialog detection
- **pyautogui**: Keyboard/mouse simulation for menu navigation
- **psutil**: Process detection and management
