import pyautogui
from docx import Document
from docx.shared import Inches
import time
import os
import tkinter as tk
from tkinter import simpledialog

# --- 1. Configuration (Change to your coordinates) ---
# Regions are (X start, Y start, Width, Height)
# Adjust according to your screen layout

GRAPH_REGION = (100, 100, 600, 400)  # (X, Y, W, H) for the graph region
DATE_REGION = (800, 50, 200, 50)     # (X, Y, W, H) for the date/timestamp region

TEMP_GRAPH_FILE = "temp_graph.png"
TEMP_DATE_FILE = "temp_date.png"
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

# --- 3. Pembuatan Laporan Word ---
def create_word_report(filename):
    """Create and save a Word report with the given filename."""
    print("Starting Word document creation...")
    document = Document()

    document.add_heading('STANDARD SIMULATION TEST REPORT', 0)

    # Insert date image (as a small header image)
    document.add_heading('Test date:', level=3)
    document.add_picture(TEMP_DATE_FILE, width=Inches(2))

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

    # Take graph screenshot
    take_screenshot(GRAPH_REGION, TEMP_GRAPH_FILE)

    # Take date screenshot
    take_screenshot(DATE_REGION, TEMP_DATE_FILE)

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
    os.remove(TEMP_DATE_FILE)
    print("Temporary files cleaned up.")

if __name__ == '__main__':
    main()