import sys
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QVBoxLayout,
    QPushButton,
    QMessageBox,
    QDialog,
    QLineEdit,
    QInputDialog,
)
from PyQt5.QtGui import QPixmap, QPainter, QPen, QColor, QImage
from PyQt5.QtCore import Qt, QPoint, QRect
from PIL import ImageGrab, Image  # Pillow for screenshotting
from docx import Document
from docx.shared import Inches
from datetime import date  # New: Import date for current date
import os
import time

# --- Configuration (Report Name will be prompted) ---
FINAL_PDF_FILE_BASE = "Interactive_Test_Report_Final.pdf"  # Base name, will be modified with user input
TEMPLATES_PATH = os.path.join(os.path.dirname(__file__), "templates")  # Directory for Word template
if not os.path.exists(TEMPLATES_PATH):
    os.makedirs(TEMPLATES_PATH)
WORD_TEMPLATE_REPORT = os.path.join(TEMPLATES_PATH, "report_template.docx")

# Create a dummy template if it doesn't exist
if not os.path.exists(WORD_TEMPLATE_REPORT):
    doc = Document()
    doc.add_heading("TEST REPORT", 0)
    doc.add_heading("Test Results Graph:", level=3)
    doc.add_paragraph(" [Area for Graph] ")
    doc.add_heading("Testing Date:", level=3)
    doc.add_paragraph(" [Area for Date - Replaced by Live Date] ")
    doc.save(WORD_TEMPLATE_REPORT)
    print(f"Dummy Word template '{WORD_TEMPLATE_REPORT}' created.")


class ScreenshotSelector(QDialog):
    """A frameless dialog to allow the user to select an area of the screen for cropping."""

    def __init__(self, full_screen_image, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Screenshot Area")
        # Set the dialog size to cover the entire screen
        self.setGeometry(0, 0, full_screen_image.width(), full_screen_image.height())
        # Frameless window, always on top
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setCursor(Qt.CrossCursor)  # Cursor becomes crosshair

        self.full_screen_pixmap = QPixmap.fromImage(full_screen_image)
        self.start_point = QPoint()
        self.end_point = QPoint()
        self.selection_rect = QRect()
        self.is_drawing = False

    def paintEvent(self, event):
        painter = QPainter(self)
        # Draw the full screen background image
        painter.drawPixmap(0, 0, self.full_screen_pixmap)

        # Draw the selection rectangle if currently drawing
        if self.is_drawing:
            pen = QPen(QColor(255, 0, 0))  # Red color
            pen.setWidth(2)
            painter.setPen(pen)
            painter.drawRect(self.selection_rect)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.start_point = event.pos()
            self.end_point = event.pos()
            self.is_drawing = True
            self.selection_rect = QRect(self.start_point, self.end_point).normalized()
            self.update()

    def mouseMoveEvent(self, event):
        if self.is_drawing:
            self.end_point = event.pos()
            self.selection_rect = QRect(self.start_point, self.end_point).normalized()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.is_drawing = False
            self.close()  # Close window after selection is complete

class App:
    """Main application logic for capturing screenshots and generating the report."""

    def __init__(self):
        self.graph_path = "graph_temp.png"
        # self.date_path is no longer needed as the date is generated programmatically
        self.final_report_name = ""
        self.final_pdf_name = ""

    def get_report_name(self):
        """Prompts the user for the final report name."""
        text, ok = QInputDialog.getText(
            None,
            "Report Name Input",
            'Enter the desired name for the final DOCX report (e.g., "System_Test_A"):',
            QLineEdit.Normal,
            f"Test_Report_{date.today().strftime('%Y%m%d')}",
        )

        if ok and text:
            # Ensure it ends with .docx
            self.final_report_name = text.strip()
            if not self.final_report_name.lower().endswith(".docx"):
                self.final_report_name += ".docx"

            # Set the PDF name based on the DOCX name
            base_name = os.path.splitext(self.final_report_name)[0]
            self.final_pdf_name = base_name + ".pdf"
            return True
        else:
            return False

    def take_interactive_screenshot(self, prompt_text="Select area"):
        """Initiates full-screen capture and interactive selection."""
        QMessageBox.information(
            None,
            "Ready for Screenshot",
            f"Click OK, then **{prompt_text}** using Drag-and-Drop on the screen.",
        )

        # Temporarily hide PyQt application window before taking the screenshot
        app = QApplication.instance()
        if app:
            for widget in app.topLevelWidgets():
                widget.hide()

        time.sleep(0.5)  # Give time for the window to hide

        full_screen_image_pil = ImageGrab.grab()  # Capture full screen screenshot

        if app:
            # Show PyQt application window again after screenshot capture
            for widget in app.topLevelWidgets():
                widget.show()

        # Convert PIL Image to Qt-compatible format (QImage)
        img_data = full_screen_image_pil.tobytes("raw", "RGB")
        full_screen_image_qt = QImage(
            img_data,
            full_screen_image_pil.size[0],
            full_screen_image_pil.size[1],
            full_screen_image_pil.size[0] * 3,
            QImage.Format_RGB888,
        )

        selector = ScreenshotSelector(full_screen_image_qt)
        # Run dialog and wait until finished
        selector.exec()

        # After the user finishes selection and the selector window is closed,
        # we get the selection coordinates
        x = selector.selection_rect.x()
        y = selector.selection_rect.y()
        w = selector.selection_rect.width()
        h = selector.selection_rect.height()

        if w == 0 or h == 0:
            QMessageBox.warning(None, "Warning", "Area not selected or too small. Please try again.")
            return None

        # Crop image according to selection
        cropped_image_pil = full_screen_image_pil.crop((x, y, x + w, y + h))
        return cropped_image_pil

    def find_and_replace_text(self, document, old_text, new_text):
        """Finds and replaces text in all paragraphs of a Word document."""
        for paragraph in document.paragraphs:
            if old_text in paragraph.text:
                paragraph.text = paragraph.text.replace(old_text, new_text)
        return document

    def run(self):
        # 1. Get Report Name from User
        if not self.get_report_name():
            QMessageBox.critical(None, "Error", "Report name not provided. Exiting.")
            return

        QMessageBox.information(
            None,
            "Start Application",
            f"Report name set to: **{self.final_report_name}**. Click OK to start screenshot capture.",
        )

        # 2. Capture Graph Screenshot
        graph_image = self.take_interactive_screenshot("select the **GRAPH** area")
        if graph_image:
            graph_image.save(self.graph_path)
            print("Graph screenshot captured successfully.")
        else:
            return  # Exit if graph capture fails

        # The date screenshot capture is REMOVED as per the requirement for automated date

        # 3. Create Word Report
        print("Starting Word document creation...")
        document = Document(WORD_TEMPLATE_REPORT)

        # Get the current date and format it
        current_date_str = date.today().strftime("%Y-%m-%d")  # e.g., 2023-11-04

        # Find and replace the Date placeholder
        self.find_and_replace_text(document, "[Area for Date - Replaced by Live Date]", current_date_str)
        self.find_and_replace_text(document, "[Area for Date]", current_date_str)

        # Insert Graph (replacing the placeholder text by inserting the image after the paragraph)
        for paragraph in document.paragraphs:
            if "[Area for Graph]" in paragraph.text:
                paragraph.text = paragraph.text.replace("[Area for Graph]", "")  # Clear the placeholder text
                paragraph.add_run().add_picture(self.graph_path, width=Inches(6))  # Insert image
                break  # Assuming only one placeholder for the graph

        document.save(self.final_report_name)
        print(f"\n✅ Final Word report saved as: {self.final_report_name}")

        # 4. Convert to PDF (Mostly for Windows environments)
        try:
            import docxtopdf

            # Provide an explicit output filename to docxtopdf
            docxtopdf.convert(self.final_report_name, self.final_pdf_name)
            print(f"✅ Final PDF report saved as: {self.final_pdf_name}")
        except ImportError:
            print("Warning: docxtopdf is not installed or does not support this OS. PDF not created.")
        except Exception as e:
            print(f"Error converting to PDF: {e}. Make sure MS Word is installed.")

        # 5. Clean up temporary files
        if os.path.exists(self.graph_path):
            os.remove(self.graph_path)
        print("Temporary files cleaned up.")
        QMessageBox.information(
            None,
            "Finished",
            f"Report '{self.final_report_name}' and '{self.final_pdf_name}' have been created.",
        )


if __name__ == '__main__':
    # Initialize the QApplication instance before creating any Qt widgets
    app = QApplication(sys.argv)
    main_app = App()
    main_app.run()
    # Execute the application's event loop
    sys.exit(app.exec_())