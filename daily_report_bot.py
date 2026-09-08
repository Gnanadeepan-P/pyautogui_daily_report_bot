import pyautogui
import time
import datetime
import re
import tkinter as tk
import subprocess
import os
import ctypes
from ctypes import wintypes

# ==========================================================
# SETTINGS
# ==========================================================

pyautogui.PAUSE = 0.4

CITY = "Chennai"
WEATHER_URL = "https://www.google.com/search?q=weather+in+Chennai"

# ==========================================================
# WINDOWS HELPERS
# ==========================================================

user32 = ctypes.windll.user32
SW_RESTORE = 9


def get_visible_windows(keyword):
    """Return visible Windows whose title contains keyword."""
    matches = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def enum_callback(hwnd, lparam):
        if not user32.IsWindowVisible(hwnd):
            return True

        length = user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return True

        title_buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, title_buffer, length + 1)
        title = title_buffer.value

        if keyword.lower() in title.lower():
            matches.append((hwnd, title))

        return True

    user32.EnumWindows(enum_callback, 0)
    return matches


def get_foreground_title():
    """Get the title of the currently active Windows window."""
    hwnd = user32.GetForegroundWindow()

    if not hwnd:
        return ""

    length = user32.GetWindowTextLengthW(hwnd)
    if length == 0:
        return ""

    buffer = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buffer, length + 1)
    return buffer.value


def activate_window(keyword, attempts=20, wait_seconds=1):
    """
    Find and activate a window.

    IMPORTANT:
    PyAutoGUI keyboard input is allowed only after this function
    confirms that the requested application is in the foreground.
    """
    print(f"Searching for {keyword} window...")

    for attempt in range(1, attempts + 1):
        windows = get_visible_windows(keyword)

        if windows:
            hwnd, title = windows[0]

            try:
                user32.ShowWindow(hwnd, SW_RESTORE)
                user32.BringWindowToTop(hwnd)
                user32.SetForegroundWindow(hwnd)
            except Exception:
                pass

            time.sleep(wait_seconds)

            foreground = get_foreground_title()
            if keyword.lower() in foreground.lower():
                print(f"{keyword} activated successfully.")
                return hwnd

            # Windows can sometimes block SetForegroundWindow.
            # Try Alt+Tab only if the requested application is
            # definitely NOT already in the foreground.
            if keyword.lower() not in get_foreground_title().lower():
                pyautogui.hotkey("alt", "tab")
                time.sleep(1)

                if keyword.lower() in get_foreground_title().lower():
                    print(f"{keyword} activated successfully using Alt+Tab.")
                    return hwnd

        print(f"Activation attempt {attempt}/{attempts}...")
        time.sleep(wait_seconds)

    print(f"ERROR: {keyword} window could not be activated.")
    return None


def get_window_rect(hwnd):
    """Return left, top, right, bottom coordinates of a window."""
    rect = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return rect.left, rect.top, rect.right, rect.bottom


def is_foreground(keyword):
    return keyword.lower() in get_foreground_title().lower()


# ==========================================================
# CLIPBOARD
# ==========================================================

def get_clipboard_text():
    root = tk.Tk()
    root.withdraw()

    try:
        text = root.clipboard_get()
    except tk.TclError:
        text = ""

    root.destroy()
    return text


# ==========================================================
# APPLICATION LOCATORS
# ==========================================================

def find_chrome():
    """Find Chrome using common Windows installation locations."""
    candidates = [
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
    ]

    for path in candidates:
        if os.path.isfile(path):
            return path

    return None


def find_excel():
    """Find Microsoft Excel using common Office installation locations."""
    candidates = [
        os.path.expandvars(r"%ProgramFiles%\Microsoft Office\root\Office16\EXCEL.EXE"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft Office\root\Office16\EXCEL.EXE"),
        os.path.expandvars(r"%ProgramFiles%\Microsoft Office\Office16\EXCEL.EXE"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft Office\Office16\EXCEL.EXE"),
    ]

    for path in candidates:
        if os.path.isfile(path):
            return path

    return None


def launch_chrome(url):
    chrome_path = find_chrome()

    if chrome_path:
        print("Chrome executable found.")
        subprocess.Popen([chrome_path, url])
        return

    # Fallback: Windows can resolve Chrome through its registered app.
    subprocess.Popen(
        ["cmd", "/c", "start", "", "chrome.exe", url],
        creationflags=subprocess.CREATE_NO_WINDOW
    )


def launch_excel():
    excel_path = find_excel()

    if excel_path:
        print("Excel executable found.")
        subprocess.Popen([excel_path])
        return

    # Fallback: Windows can resolve Excel through App Paths.
    subprocess.Popen(
        ["cmd", "/c", "start", "", "excel.exe"],
        creationflags=subprocess.CREATE_NO_WINDOW
    )


# ==========================================================
# EXCEL SAVE-AS HELPER
# ==========================================================

def save_workbook_as(excel_path):
    """
    Save the current Excel workbook using Excel's F12 Save As shortcut.

    F12 is used instead of Ctrl+Shift+S + Alt+N because Alt+N is
    the keyboard accelerator for the Insert tab in modern Excel
    when the Save As dialog has not opened/focused correctly.
    """
    if os.path.isfile(excel_path):
        try:
            os.remove(excel_path)
            print("Existing report removed so overwrite is not required.")
        except OSError as error:
            print("Could not remove existing report:", error)

    # Make absolutely sure the Excel window has focus.
    if not is_foreground("Excel"):
        if not activate_window("Excel", attempts=15, wait_seconds=1):
            return False

    # Make sure the last cell entry has been committed.
    pyautogui.press("enter")
    time.sleep(2)

    print("Opening Excel Save As dialog with F12...")
    pyautogui.press("f12")
    time.sleep(5)

    # F12 should open the Windows Save As dialog. Do NOT use Alt+N
    # here; Alt+N can select the Insert tab in Excel.
    dialog_title = get_foreground_title()
    print("Current window:", dialog_title)

    # The filename box is normally focused automatically when the
    # Save As dialog opens. Ctrl+A selects the existing filename.
    pyautogui.hotkey("ctrl", "a")
    time.sleep(1)
    pyautogui.write(excel_path, interval=0.01)
    time.sleep(1)
    pyautogui.press("enter")

    # Wait for Excel to complete the save.
    for attempt in range(1, 21):
        if os.path.isfile(excel_path):
            return True

        # Excel may display a confirmation dialog after Save As.
        # Press Enter once after a short wait, then continue checking.
        if attempt in (5, 10):
            pyautogui.press("enter")

        print(f"Waiting for Excel to save... {attempt}/20")
        time.sleep(1)

    return os.path.isfile(excel_path)

# ==========================================================
# START
# ==========================================================

now = datetime.datetime.now()
date_time = now.strftime("%Y-%m-%d %H:%M:%S")
today = now.strftime("%Y-%m-%d")

# Always save beside this Python script.
project_folder = os.path.dirname(os.path.abspath(__file__))

excel_path = os.path.join(
    project_folder,
    f"daily_report_{today}.xlsx"
)

screenshot_path = os.path.join(
    project_folder,
    f"daily_report_{today}.png"
)

print()
print("========================================")
print("       DAILY REPORT AUTOMATION")
print("========================================")
print("Date & Time:", date_time)
print("Project folder:", project_folder)

# ==========================================================
# STEP 1 - OPEN CHROME
# ==========================================================

print()
print("1. Opening Chrome...")

if not get_visible_windows("Chrome"):
    launch_chrome(WEATHER_URL)
else:
    print("Chrome is already open.")

time.sleep(8)

# ==========================================================
# STEP 2 - ACTIVATE CHROME
# ==========================================================

print("2. Activating Chrome...")

if not activate_window("Chrome", attempts=15, wait_seconds=1):
    print("STOPPING AUTOMATION: Chrome could not be activated.")
    raise SystemExit(1)

# ==========================================================
# STEP 3 - OPEN WEATHER PAGE
# ==========================================================

print("3. Opening weather page...")

if not is_foreground("Chrome"):
    print("STOPPING AUTOMATION: Chrome is not in the foreground.")
    raise SystemExit(1)

pyautogui.hotkey("ctrl", "l")
time.sleep(1)
pyautogui.write(WEATHER_URL, interval=0.02)
pyautogui.press("enter")

time.sleep(8)

# ==========================================================
# STEP 4 - COPY PAGE INFORMATION
# ==========================================================

print("4. Copying webpage information...")

if not is_foreground("Chrome"):
    if not activate_window("Chrome", attempts=5, wait_seconds=1):
        print("STOPPING AUTOMATION: Chrome is not active.")
        raise SystemExit(1)

pyautogui.hotkey("ctrl", "a")
time.sleep(2)
pyautogui.hotkey("ctrl", "c")
time.sleep(3)

page_text = get_clipboard_text()

if not page_text:
    print("ERROR: Could not copy webpage information.")
    raise SystemExit(1)

print("Webpage information copied successfully.")

# ==========================================================
# STEP 5 - FIND TEMPERATURE
# ==========================================================

print("5. Extracting temperature...")

temperature_match = re.search(
    r"(-?\d+(?:\.\d+)?)\s*°?\s*([CF])",
    page_text,
    re.IGNORECASE
)

if temperature_match:
    temperature = temperature_match.group(1)
    unit = temperature_match.group(2).upper()
    temperature_text = f"{temperature}°{unit}"
else:
    temperature_text = "Temperature not found"

print("Fetched Data:", temperature_text)

# ==========================================================
# STEP 6 - CREATE COMMENT
# ==========================================================

if temperature_match:
    temp_value = float(temperature_match.group(1))

    if unit == "C":
        if temp_value >= 35:
            comment = "Very hot day; stay hydrated"
        elif temp_value >= 30:
            comment = "Warm day; suitable with precautions"
        else:
            comment = "Good for outdoor activities"
    else:
        comment = "Weather information available"
else:
    comment = "Weather information unavailable"

print("Comment:", comment)

# ==========================================================
# STEP 7 - CLOSE CHROME
# ==========================================================
# This is critical. Excel data must never be sent to Chrome.

print()
print("6. Closing Chrome before opening Excel...")

if is_foreground("Chrome"):
    pyautogui.hotkey("alt", "f4")
    time.sleep(4)

# ==========================================================
# STEP 8 - OPEN EXCEL
# ==========================================================

print("7. Opening Microsoft Excel...")

if not get_visible_windows("Excel"):
    launch_excel()
else:
    print("Excel is already open.")

print("Waiting for Excel to start...")
time.sleep(10)

# ==========================================================
# STEP 9 - ACTIVATE EXCEL WITH RETRIES
# ==========================================================

print("8. Activating Excel...")

excel_hwnd = activate_window(
    "Excel",
    attempts=25,
    wait_seconds=1
)

if not excel_hwnd:
    print()
    print("STOPPING AUTOMATION.")
    print("Excel could not be activated.")
    print("Close Excel and run the script again.")
    raise SystemExit(1)

# Extra startup time for slow Excel installations.
time.sleep(5)

# ==========================================================
# STEP 10 - CREATE NEW WORKBOOK
# ==========================================================

print("9. Creating new workbook...")

if not is_foreground("Excel"):
    if not activate_window("Excel", attempts=10, wait_seconds=1):
        print("STOPPING AUTOMATION: Excel is not active.")
        raise SystemExit(1)

pyautogui.hotkey("ctrl", "n")
time.sleep(7)

# Verify again before ANY data is typed.
if not is_foreground("Excel"):
    if not activate_window("Excel", attempts=10, wait_seconds=1):
        print("STOPPING AUTOMATION: Excel lost focus.")
        raise SystemExit(1)

print("Excel is confirmed as the active application.")

# ==========================================================
# STEP 11 - ENTER HEADERS AND DATA
# ==========================================================

print("10. Entering report data...")

# Final safety check immediately before typing.
if not is_foreground("Excel"):
    print("ERROR: Excel is not the foreground application.")
    print("No data will be typed.")
    raise SystemExit(1)

# Header row
pyautogui.write("Date & Time", interval=0.03)
pyautogui.press("tab")

pyautogui.write("Fetched Data", interval=0.03)
pyautogui.press("tab")

pyautogui.write("Comment", interval=0.03)

# Move from C1 to A2.
pyautogui.press("home")
pyautogui.press("down")

# Data row
pyautogui.write(date_time, interval=0.02)
pyautogui.press("tab")

pyautogui.write(
    f"{CITY}: {temperature_text}",
    interval=0.02
)

pyautogui.press("tab")

pyautogui.write(
    comment,
    interval=0.02
)

time.sleep(3)

# ==========================================================
# STEP 12 - SAVE EXCEL FILE
# ==========================================================

print()
print("11. Saving Excel file...")
print("Save location:")
print(excel_path)

if not is_foreground("Excel"):
    if not activate_window("Excel", attempts=10, wait_seconds=1):
        print("STOPPING AUTOMATION: Excel is not active before Save As.")
        raise SystemExit(1)

time.sleep(2)

if save_workbook_as(excel_path):
    print("SUCCESS: Excel file saved.")
else:
    print()
    print("ERROR: Excel file was NOT saved.")
    print("Expected location:")
    print(excel_path)
    print()
    print("Excel may still be showing a Save As or confirmation dialog.")
    raise SystemExit(1)

# ==========================================================
# STEP 13 - TAKE SCREENSHOT OF EXCEL
# ==========================================================

print()
print("12. Taking screenshot of final Excel sheet...")

excel_hwnd = activate_window(
    "Excel",
    attempts=10,
    wait_seconds=1
)

if not excel_hwnd:
    print("ERROR: Could not activate Excel for screenshot.")
    raise SystemExit(1)

time.sleep(2)

try:
    left, top, right, bottom = get_window_rect(excel_hwnd)

    width = right - left
    height = bottom - top

    if width > 0 and height > 0:
        pyautogui.screenshot(
            screenshot_path,
            region=(left, top, width, height)
        )
    else:
        pyautogui.screenshot(screenshot_path)

except Exception as error:
    print("Window screenshot failed:", error)
    print("Taking full-screen screenshot instead.")
    pyautogui.screenshot(screenshot_path)

print("Screenshot saved:")
print(screenshot_path)

# ==========================================================
# FINAL RESULT
# ==========================================================

print()
print("========================================")
print("       AUTOMATION COMPLETED")
print("========================================")
print()
print("Excel file:")
print(excel_path)
print()
print("Screenshot:")
print(screenshot_path)
print()
print("Fetched data:")
print(temperature_text)
print()
print("Comment:")
print(comment)
print()
print("========================================")
