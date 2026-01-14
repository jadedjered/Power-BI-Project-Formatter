"""
Power BI Desktop UI Automation Module

This module provides functions to automate Power BI Desktop for converting
PBIX files to PBIP project format using Windows UI automation.
"""

import os
import time
import subprocess
from pathlib import Path
from typing import Optional, Tuple

import psutil
import pyautogui
from pywinauto import Application, Desktop
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.timings import TimeoutError as PywinautoTimeoutError


# Common Power BI Desktop installation paths
PBI_DESKTOP_PATHS = [
    r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe",
    r"C:\Program Files (x86)\Microsoft Power BI Desktop\bin\PBIDesktop.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WindowsApps\PBIDesktopStore.exe"),
    os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WindowsApps\Microsoft.MicrosoftPowerBIDesktop_8wekyb3d8bbwe\PBIDesktop.exe"),
]

# Timeouts (in seconds)
STARTUP_TIMEOUT = 120  # Max time to wait for PBI Desktop to start
DIALOG_TIMEOUT = 30    # Max time to wait for dialogs
SAVE_TIMEOUT = 60      # Max time to wait for save operation


class PBIAutomationError(Exception):
    """Custom exception for Power BI automation errors."""
    pass


def find_pbi_desktop() -> Optional[str]:
    """
    Locate Power BI Desktop executable on the system.

    Returns:
        Path to PBIDesktop.exe if found, None otherwise.
    """
    for path in PBI_DESKTOP_PATHS:
        expanded_path = os.path.expandvars(path)
        if os.path.exists(expanded_path):
            return expanded_path

    # Try to find via registry or PATH
    try:
        result = subprocess.run(
            ["where", "PBIDesktop.exe"],
            capture_output=True,
            text=True,
            timeout=10
        )
        if result.returncode == 0:
            return result.stdout.strip().split('\n')[0]
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pass

    return None


def is_pbi_desktop_running() -> bool:
    """Check if Power BI Desktop is currently running."""
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] and 'pbidesktop' in proc.info['name'].lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False


def kill_pbi_desktop() -> bool:
    """
    Forcefully terminate all Power BI Desktop processes.

    Returns:
        True if any processes were killed, False otherwise.
    """
    killed = False
    for proc in psutil.process_iter(['name', 'pid']):
        try:
            if proc.info['name'] and 'pbidesktop' in proc.info['name'].lower():
                proc.kill()
                killed = True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    if killed:
        time.sleep(2)  # Wait for processes to fully terminate

    return killed


def open_pbix(pbix_path: str, pbi_exe_path: str) -> subprocess.Popen:
    """
    Open a PBIX file in Power BI Desktop.

    Args:
        pbix_path: Full path to the PBIX file.
        pbi_exe_path: Path to Power BI Desktop executable.

    Returns:
        Popen object for the started process.
    """
    abs_path = os.path.abspath(pbix_path)
    if not os.path.exists(abs_path):
        raise PBIAutomationError(f"PBIX file not found: {abs_path}")

    process = subprocess.Popen([pbi_exe_path, abs_path])
    return process


def wait_for_pbi_ready(filename: str, timeout: int = STARTUP_TIMEOUT) -> Application:
    """
    Wait for Power BI Desktop to fully load a file.

    Args:
        filename: Name of the PBIX file (without path) to look for in title.
        timeout: Maximum seconds to wait.

    Returns:
        pywinauto Application object connected to PBI Desktop.

    Raises:
        PBIAutomationError: If timeout is reached.
    """
    base_name = Path(filename).stem
    start_time = time.time()

    while time.time() - start_time < timeout:
        try:
            # Look for window with the file name in title
            app = Application(backend='uia').connect(
                title_re=f".*{base_name}.*Power BI Desktop.*",
                timeout=5
            )

            main_window = app.window(title_re=f".*{base_name}.*Power BI Desktop.*")

            # Wait for window to be ready (not showing loading)
            if main_window.exists() and main_window.is_visible():
                # Give it a bit more time to fully initialize
                time.sleep(3)
                return app

        except (ElementNotFoundError, PywinautoTimeoutError):
            pass

        time.sleep(2)

    raise PBIAutomationError(
        f"Timeout waiting for Power BI Desktop to load '{filename}'. "
        f"Waited {timeout} seconds."
    )


def save_as_pbip(app: Application, output_folder: str, project_name: str) -> Tuple[bool, str]:
    """
    Automate the Save As dialog to save current file as PBIP.

    Args:
        app: pywinauto Application connected to Power BI Desktop.
        output_folder: Directory where the PBIP project should be saved.
        project_name: Name for the PBIP project (without extension).

    Returns:
        Tuple of (success: bool, message: str).
    """
    try:
        # Get the main window
        main_window = app.top_window()
        main_window.set_focus()
        time.sleep(0.5)

        # Open File menu using keyboard shortcut
        pyautogui.hotkey('alt', 'f')
        time.sleep(1)

        # Navigate to "Save as" - it's typically accessed via the File menu
        # In Power BI Desktop, we need to click "Save as" or use keyboard
        pyautogui.press('a')  # 'a' for Save As in many versions
        time.sleep(1)

        # If that didn't work, try clicking directly or using arrow keys
        # Try pressing down arrow to navigate menu and enter
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('down')  # Navigate past Save to Save As
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(1.5)

        # Wait for Save As dialog to appear
        save_dialog = None
        for _ in range(DIALOG_TIMEOUT):
            try:
                # Try to find the Save As dialog
                save_dialog = app.window(title_re=".*Save.*", control_type="Window")
                if save_dialog.exists():
                    break
            except ElementNotFoundError:
                pass

            # Also check for a file dialog
            try:
                desktop = Desktop(backend='uia')
                save_dialog = desktop.window(title_re=".*Save.*")
                if save_dialog.exists():
                    break
            except ElementNotFoundError:
                pass

            time.sleep(1)

        if not save_dialog or not save_dialog.exists():
            # Fallback: try keyboard shortcut for Save As
            main_window.set_focus()
            pyautogui.hotkey('ctrl', 'shift', 's')
            time.sleep(2)

        # Prepare the output path
        output_path = os.path.join(os.path.abspath(output_folder), project_name)

        # Type the file path in the filename field
        # First, try to find and clear the filename field
        time.sleep(1)

        # Use keyboard to navigate - Ctrl+L or Tab to filename field, then type path
        pyautogui.hotkey('alt', 'n')  # Common shortcut for filename field
        time.sleep(0.5)

        # Clear existing text and type new path
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.2)
        pyautogui.typewrite(output_path, interval=0.02)
        time.sleep(0.5)

        # Now change the file type to PBIP
        # Tab to file type dropdown or use Alt+T
        pyautogui.hotkey('alt', 't')  # File type dropdown
        time.sleep(0.5)

        # Type to search for PBIP in dropdown
        pyautogui.typewrite('pbip', interval=0.05)
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(0.5)

        # Click Save button
        pyautogui.hotkey('alt', 's')  # Save button
        time.sleep(SAVE_TIMEOUT)

        # Verify the output was created
        pbip_file = os.path.join(output_folder, f"{project_name}.pbip")
        if os.path.exists(pbip_file):
            return True, f"Successfully saved to {output_folder}"
        else:
            # Check if the folder structure was created
            report_folder = os.path.join(output_folder, f"{project_name}.Report")
            model_folder = os.path.join(output_folder, f"{project_name}.SemanticModel")
            if os.path.exists(report_folder) or os.path.exists(model_folder):
                return True, f"Successfully saved to {output_folder}"

            return False, "Save operation may have failed - output files not found"

    except Exception as e:
        return False, f"Error during save operation: {str(e)}"


def close_pbi_desktop(app: Application = None, force: bool = False) -> bool:
    """
    Close Power BI Desktop gracefully.

    Args:
        app: pywinauto Application object (optional).
        force: If True, force kill the process.

    Returns:
        True if closed successfully.
    """
    if force:
        return kill_pbi_desktop()

    try:
        if app:
            main_window = app.top_window()
            main_window.set_focus()
            time.sleep(0.3)

        # Try Alt+F4
        pyautogui.hotkey('alt', 'F4')
        time.sleep(2)

        # Handle "Don't Save" dialog if it appears
        try:
            desktop = Desktop(backend='uia')
            # Look for save prompt dialog
            for _ in range(5):
                try:
                    dialog = desktop.window(title_re=".*Power BI Desktop.*")
                    if dialog.exists():
                        # Try to find "Don't Save" or "No" button
                        pyautogui.press('tab')  # Navigate to Don't Save
                        pyautogui.press('enter')
                        break
                except ElementNotFoundError:
                    break
                time.sleep(0.5)
        except Exception:
            pass

        time.sleep(2)

        # Verify it's closed
        if not is_pbi_desktop_running():
            return True

        # Force kill if still running
        return kill_pbi_desktop()

    except Exception:
        return kill_pbi_desktop()


def convert_pbix_to_pbip(
    pbix_path: str,
    output_folder: str,
    project_name: Optional[str] = None
) -> Tuple[bool, str]:
    """
    Convert a PBIX file to PBIP format.

    This is the main function that orchestrates the entire conversion process.

    Args:
        pbix_path: Path to the source PBIX file.
        output_folder: Directory where the PBIP project should be saved.
        project_name: Optional name for the project. If None, uses PBIX filename.

    Returns:
        Tuple of (success: bool, message: str).
    """
    # Find Power BI Desktop
    pbi_exe = find_pbi_desktop()
    if not pbi_exe:
        return False, (
            "Power BI Desktop not found. Please ensure it is installed.\n"
            "Download from: https://powerbi.microsoft.com/desktop/"
        )

    # Check if PBI Desktop is already running
    if is_pbi_desktop_running():
        return False, (
            "Power BI Desktop is already running. "
            "Please close it before starting the conversion."
        )

    # Determine project name
    if project_name is None:
        project_name = Path(pbix_path).stem

    # Create output folder if needed
    os.makedirs(output_folder, exist_ok=True)

    app = None
    process = None

    try:
        # Open the PBIX file
        process = open_pbix(pbix_path, pbi_exe)

        # Wait for PBI Desktop to fully load
        app = wait_for_pbi_ready(pbix_path)

        # Perform the Save As operation
        success, message = save_as_pbip(app, output_folder, project_name)

        return success, message

    except PBIAutomationError as e:
        return False, str(e)
    except Exception as e:
        return False, f"Unexpected error: {str(e)}"
    finally:
        # Always try to close PBI Desktop
        close_pbi_desktop(app, force=True)

        if process:
            try:
                process.terminate()
            except Exception:
                pass
