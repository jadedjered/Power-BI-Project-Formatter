#!/usr/bin/env python
"""
Power BI PBIX to PBIP Converter

A console application that converts Microsoft Power BI PBIX files to PBIP
(Power BI Project) format by automating Power BI Desktop.

Usage:
    python pbix_converter.py [directory]

If no directory is specified, the current working directory is used.
"""

import os
import sys
from pathlib import Path
from typing import List, Optional

from pbi_automation import (
    find_pbi_desktop,
    is_pbi_desktop_running,
    convert_pbix_to_pbip,
    kill_pbi_desktop,
)


# Output folder name
OUTPUT_FOLDER_NAME = "pbip_output"


def print_header():
    """Print the application header."""
    print()
    print("=" * 50)
    print("  Power BI PBIX to PBIP Converter")
    print("=" * 50)
    print()


def find_pbix_files(directory: str) -> List[Path]:
    """
    Find all PBIX files in the specified directory.

    Args:
        directory: Directory to search.

    Returns:
        List of Path objects for found PBIX files.
    """
    dir_path = Path(directory)
    return sorted(dir_path.glob("*.pbix"))


def display_files(files: List[Path]) -> None:
    """Display the list of found PBIX files."""
    print(f"Found {len(files)} PBIX file(s):")
    print()
    for i, f in enumerate(files, 1):
        size_mb = f.stat().st_size / (1024 * 1024)
        print(f"  {i}. {f.name} ({size_mb:.1f} MB)")
    print()


def get_user_selection(files: List[Path]) -> Optional[List[Path]]:
    """
    Prompt user to select which files to convert.

    Args:
        files: List of available PBIX files.

    Returns:
        List of selected files, or None to quit.
    """
    print("Options:")
    print("  [A] Convert all files")
    print("  [S] Select specific files (enter numbers separated by commas)")
    print("  [Q] Quit")
    print()

    while True:
        choice = input("Your choice: ").strip().upper()

        if choice == 'Q':
            return None

        if choice == 'A':
            return files

        if choice == 'S':
            selection = input("Enter file numbers (e.g., 1,3,5): ").strip()
            try:
                indices = [int(x.strip()) for x in selection.split(',')]
                selected = []
                for idx in indices:
                    if 1 <= idx <= len(files):
                        selected.append(files[idx - 1])
                    else:
                        print(f"  Warning: Ignoring invalid number {idx}")

                if selected:
                    return selected
                else:
                    print("  No valid files selected. Please try again.")
            except ValueError:
                print("  Invalid input. Please enter numbers separated by commas.")
            continue

        print("  Invalid choice. Please enter A, S, or Q.")


def check_prerequisites() -> bool:
    """
    Check that all prerequisites are met.

    Returns:
        True if all prerequisites are met, False otherwise.
    """
    print("Checking prerequisites...")

    # Check for Power BI Desktop
    pbi_path = find_pbi_desktop()
    if not pbi_path:
        print()
        print("ERROR: Power BI Desktop not found!")
        print()
        print("Please install Power BI Desktop from:")
        print("  https://powerbi.microsoft.com/desktop/")
        print()
        print("After installation, ensure the PBIP preview feature is enabled:")
        print("  1. Open Power BI Desktop")
        print("  2. Go to File > Options and settings > Options")
        print("  3. Select 'Preview features'")
        print("  4. Check 'Power BI Project (.pbip) save option'")
        print("  5. Restart Power BI Desktop")
        return False

    print(f"  Power BI Desktop found: {pbi_path}")

    # Check if PBI Desktop is running
    if is_pbi_desktop_running():
        print()
        print("WARNING: Power BI Desktop is currently running.")
        print()
        response = input("Close Power BI Desktop to continue? [Y/N]: ").strip().upper()
        if response == 'Y':
            print("  Closing Power BI Desktop...")
            kill_pbi_desktop()
            print("  Done.")
        else:
            print()
            print("Please close Power BI Desktop manually and try again.")
            return False

    print("  Prerequisites OK")
    print()
    return True


def convert_files(files: List[Path], output_base: Path) -> dict:
    """
    Convert a list of PBIX files to PBIP format.

    Args:
        files: List of PBIX files to convert.
        output_base: Base output directory.

    Returns:
        Dictionary with conversion results.
    """
    results = {
        'success': [],
        'failed': [],
    }

    total = len(files)

    for i, pbix_file in enumerate(files, 1):
        project_name = pbix_file.stem
        output_folder = output_base / project_name

        print(f"\nConverting {pbix_file.name}... [{i}/{total}]")
        print(f"  Output: {output_folder}")

        print("  Opening Power BI Desktop...")
        success, message = convert_pbix_to_pbip(
            str(pbix_file),
            str(output_folder),
            project_name
        )

        if success:
            print(f"  SUCCESS: {message}")
            results['success'].append(pbix_file.name)
        else:
            print(f"  FAILED: {message}")
            results['failed'].append((pbix_file.name, message))

    return results


def print_summary(results: dict) -> None:
    """Print the conversion summary."""
    print()
    print("=" * 50)
    print("  Conversion Summary")
    print("=" * 50)
    print()
    print(f"  Successful: {len(results['success'])}")
    print(f"  Failed:     {len(results['failed'])}")
    print()

    if results['success']:
        print("Successful conversions:")
        for name in results['success']:
            print(f"    {name}")
        print()

    if results['failed']:
        print("Failed conversions:")
        for name, reason in results['failed']:
            print(f"    {name}")
            print(f"      Reason: {reason}")
        print()


def main():
    """Main entry point."""
    print_header()

    # Determine working directory
    if len(sys.argv) > 1:
        work_dir = sys.argv[1]
        if not os.path.isdir(work_dir):
            print(f"ERROR: '{work_dir}' is not a valid directory.")
            sys.exit(1)
    else:
        work_dir = os.getcwd()

    print(f"Working directory: {work_dir}")
    print()

    # Find PBIX files
    pbix_files = find_pbix_files(work_dir)

    if not pbix_files:
        print("No PBIX files found in the current directory.")
        print()
        print("Usage: python pbix_converter.py [directory]")
        sys.exit(0)

    # Display found files
    display_files(pbix_files)

    # Get user selection
    selected_files = get_user_selection(pbix_files)

    if selected_files is None:
        print()
        print("Conversion cancelled.")
        sys.exit(0)

    print()
    print(f"Selected {len(selected_files)} file(s) for conversion.")

    # Confirm before proceeding
    print()
    print("IMPORTANT: During conversion, do not use the mouse or keyboard.")
    print("The automation needs to control Power BI Desktop.")
    print()
    confirm = input("Ready to begin? [Y/N]: ").strip().upper()

    if confirm != 'Y':
        print()
        print("Conversion cancelled.")
        sys.exit(0)

    # Check prerequisites
    print()
    if not check_prerequisites():
        sys.exit(1)

    # Create output directory
    output_base = Path(work_dir) / OUTPUT_FOLDER_NAME
    output_base.mkdir(exist_ok=True)
    print(f"Output folder: {output_base}")

    # Perform conversions
    results = convert_files(selected_files, output_base)

    # Print summary
    print_summary(results)

    # Exit code based on results
    if results['failed']:
        sys.exit(1)
    sys.exit(0)


if __name__ == "__main__":
    main()
