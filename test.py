import pyautogui
from docx import Document
from docx.shared import Inches
import time
import os
import tkinter as tk
from tkinter import simpledialog
from datetime import datetime
try:
    # Windows-only imports for window capture
    import win32gui
    import win32ui
    import win32con
    import win32api
    from PIL import Image
    WIN32_AVAILABLE = True
except Exception:
    WIN32_AVAILABLE = False

# --- 1. Configuration (Change to your coordinates) ---
# Regions are (X start, Y start, Width, Height)
# Adjust according to your screen layout

GRAPH_REGION = (632, 292, 1269, 583)  # (X, Y, W, H) for the graph region

TEMP_GRAPH_FILE = "temp_graph.png"
FINAL_REPORT_NAME = "Simulation_Report_Final.docx"

# --- 2. Proses Tangkap Layar ---
def take_screenshot(region, filename):
    # Small pause to ensure the target window is focused
    time.sleep(1)
    print(f"Taking screenshot: {filename}...")

    # Capture the specified region and save
    im = pyautogui.screenshot(region=region)
    im.save(filename)
    print(f"Screenshot {filename} saved.")


def find_edge_window():
    """Return the HWND of a Microsoft Edge window (first match) or None."""
    if not WIN32_AVAILABLE:
        return None

    def _enum(hwnd, results):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd) or ""
        cls = win32gui.GetClassName(hwnd) or ""
        # Match common Edge window title/class (Edge is Chromium-based)
        if 'edge' in title.lower() or 'microsoft edge' in title.lower() or 'msedge' in title.lower():
            results.append(hwnd)
        elif 'chrome' in cls.lower() or 'chrome_widget_win' in cls.lower():
            # Edge uses Chromium class names in some versions; also check title for Edge
            if 'edge' in title.lower() or 'microsoft edge' in title.lower():
                results.append(hwnd)

    matches = []
    win32gui.EnumWindows(_enum, matches)
    return matches[0] if matches else None


def capture_window_only(hwnd, filename):
    """Capture the given window's image (hwnd) to filename using PrintWindow.
    Falls back to region screenshot if PrintWindow fails or win32 not available.
    """
    if not WIN32_AVAILABLE or hwnd is None:
        return False

    try:
        # Get window rectangle (including borders)
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bottom - top

        # Create device context
        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        # Create bitmap to save
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
        saveDC.SelectObject(saveBitMap)

        # Try PrintWindow (1 to include layered windows)
        result = win32gui.PrintWindow(hwnd, saveDC.GetSafeHdc(), 1)

        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)

        # Convert raw data to PIL Image
        im = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1
        )
        im.save(filename)

        # Cleanup
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)

        return bool(result)
    except Exception:
        return False

# --- 3. Pembuatan Laporan Word ---
def create_word_report(filename):
    """Create and save a Word report with the given filename."""
    print("Starting Word document creation...")
    document = Document()

    document.add_heading('STANDARD SIMULATION TEST REPORT', 0)

    # Insert current date
    document.add_heading('Test date:', level=3)
    current_date = datetime.now().strftime("%B %d, %Y %H:%M:%S")
    document.add_paragraph(current_date)

    document.add_paragraph('\n')  # blank line

    # Insert graph
    document.add_heading('Graph:', level=3)
    document.add_picture(TEMP_GRAPH_FILE, width=Inches(6))

    # Save document
    document.save(filename)
    print(f"\n✅ Final Word report saved as: {filename}")

# --- 4. Fungsi Utama ---
def main():
    # Make sure the target window is active before running this!

    # Try to capture Microsoft Edge window only (preferred)
    final_capture_ok = False
    if WIN32_AVAILABLE:
        hwnd = find_edge_window()
        if hwnd:
            print("Found Edge window, capturing only that window...")
            final_capture_ok = capture_window_only(hwnd, TEMP_GRAPH_FILE)
            if not final_capture_ok:
                # If PrintWindow failed, fall back to a region capture of the window rect
                try:
                    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                    region = (left, top, right - left, bottom - top)
                    take_screenshot(region, TEMP_GRAPH_FILE)
                    final_capture_ok = True
                except Exception:
                    final_capture_ok = False

    if not final_capture_ok:
        # Fallback: capture preconfigured screen region
        print("Edge window not found or capture failed — using configured region.")
        take_screenshot(GRAPH_REGION, TEMP_GRAPH_FILE)

    # Ask the user for the report filename via popup (without extension recommended)
    root = tk.Tk()
    root.withdraw()  # hide main window
    default_name = os.path.splitext(FINAL_REPORT_NAME)[0]
    user_input = simpledialog.askstring("Save Report", "Enter report file name (without extension):", initialvalue=default_name)
    root.destroy()

    if user_input is None:
        # If user cancels, use the default name
        final_filename = FINAL_REPORT_NAME
    else:
        user_input = user_input.strip()
        if user_input == "":
            final_filename = FINAL_REPORT_NAME
        else:
            if not user_input.lower().endswith('.docx'):
                user_input = user_input + '.docx'
            final_filename = user_input

    # Create report
    create_word_report(final_filename)

    # Clean up temporary files
    os.remove(TEMP_GRAPH_FILE)
    print("Temporary files cleaned up.")

if __name__ == '__main__':
    main()