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
    QFileDialog, # New: For file selection
)
from PyQt5.QtGui import QPixmap, QPainter, QPen, QColor, QImage
from PyQt5.QtCore import Qt, QPoint, QRect, QDir # New: QDir
from PIL import ImageGrab, Image
from docx import Document
from docx.shared import Inches
from datetime import datetime
import os
import time

# --- Configuration & Template Management ---

def create_dummy_template(filepath, placeholder_text):
    """Creates a basic template with the required placeholder if the file does not exist."""
    print(f"Template '{filepath}' not found. Creating a dummy template with placeholder.")
    doc = Document()
    doc.add_heading("INTERACTIVE TEST REPORT", 0)
    doc.add_paragraph("This report was generated using the screenshot tool.")
    doc.add_heading("Results Insertion Area:", level=3)
    doc.add_paragraph(f"Placeholder for content: **{placeholder_text}**")
    doc.add_paragraph(placeholder_text) # The actual placeholder string
    doc.save(filepath)
    print(f"Dummy Word template '{filepath}' created.")
    return filepath


class ScreenshotSelector(QDialog):
    """A frameless dialog to allow the user to select an area of the screen for cropping."""

    def __init__(self, full_screen_image, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Screenshot Area")
        # Set the dialog size to cover the entire screen
        desktop = QApplication.instance().desktop()
        self.setGeometry(0, 0, desktop.width(), desktop.height())
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
        self.input_template_path = ""
        self.final_report_name = ""
        self.final_pdf_name = ""
        self.placeholder_text = "[INSERT_CONTENT_HERE]" # Default placeholder

    def get_report_details(self):
        """Prompts the user for all necessary report details."""
        
        # 1. Get Placeholder Text
        ph_text, ok = QInputDialog.getText(
            None,
            "Placeholder Input",
            'Enter the **UNIQUE TEXT** in your document where the image and date should be inserted. (e.g., "[GRAPH_LOC_1]")',
            QLineEdit.Normal,
            self.placeholder_text,
        )
        if ok and ph_text:
            self.placeholder_text = ph_text.strip()
        else:
            return False

        # 2. Get Input Template Path (using QFileDialog for better UX)
        input_template_path, _ = QFileDialog.getOpenFileName(
            None,
            "Select Input Word Template (.docx)",
            QDir.currentPath(),
            "Word Documents (*.docx)",
        )
        if input_template_path:
            self.input_template_path = input_template_path
        else:
            QMessageBox.warning(None, "Warning", "No template file selected. Using 'default_template.docx'.")
            self.input_template_path = "default_template.docx"

        # If the selected file doesn't exist, create a dummy one
        if not os.path.exists(self.input_template_path):
             create_dummy_template(self.input_template_path, self.placeholder_text)


        # 3. Get Final Output Report Name
        base_template_name = os.path.splitext(os.path.basename(self.input_template_path))[0]
        default_output_name = f"{base_template_name}_Output_{datetime.now().strftime('%Y%m%d')}.docx"

        output_name, ok = QInputDialog.getText(
            None,
            "Output Name Input",
            'Enter the desired name for the final DOCX report:',
            QLineEdit.Normal,
            default_output_name,
        )

        if ok and output_name:
            self.final_report_name = output_name.strip()
            if not self.final_report_name.lower().endswith(".docx"):
                self.final_report_name += ".docx"
            
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

        app = QApplication.instance()
        if app:
            for widget in app.topLevelWidgets():
                # Check if it's the main application window and hide it
                if isinstance(widget, QWidget) and widget.windowTitle() == "PyQt Report Generator":
                    widget.hide()
                    break # Assuming only one main window to hide

        time.sleep(0.5) 

        full_screen_image_pil = ImageGrab.grab()

        if app:
            for widget in app.topLevelWidgets():
                if isinstance(widget, QWidget) and widget.windowTitle() == "PyQt Report Generator":
                    widget.show()
                    break

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
        selector.exec()

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

    def insert_content_at_placeholder(self, document, placeholder, image_path, date_str):
        """Finds the placeholder and replaces it with the date and image."""
        
        insertion_successful = False
        
        for paragraph in document.paragraphs:
            if placeholder in paragraph.text:
                # 1. Clear the placeholder text
                paragraph.text = paragraph.text.replace(placeholder, "")

                # 2. Get custom text from user and insert with date
                custom_text, ok = QInputDialog.getText(
                    None,
                    "Custom Text Input",
                    "Enter the text you want to add before the date:",
                    QLineEdit.Normal,
                    "Report Content Inserted"
                )
                if ok:
                    run_date = paragraph.add_run()
                    run_date.add_text(f"{custom_text}: {date_str}\n")
                    run_date.bold = True
                
                # 3. Insert the Image
                paragraph.add_run().add_picture(image_path, width=Inches(6))
                
                insertion_successful = True
                break  # Assuming only one insertion per report

        if not insertion_successful:
            QMessageBox.critical(None, "Error", f"Could not find the unique placeholder text: '{placeholder}' in the document.")
        
        return document, insertion_successful

    def run(self):
        # 0. Set up a dummy main window title for hiding logic
        main_window = QWidget()
        main_window.setWindowTitle("PyQt Report Generator")
        main_window.setGeometry(100, 100, 300, 100)
        main_window.show()


        # 1. Get Report Details from User
        if not self.get_report_details():
            QMessageBox.critical(None, "Error", "Report details not provided. Exiting.")
            main_window.close()
            return

        QMessageBox.information(
            None,
            "Start Application",
            f"Input Template: **{os.path.basename(self.input_template_path)}**\nOutput Report: **{self.final_report_name}**\nInsertion Placeholder: **{self.placeholder_text}**\n\nClick OK to start screenshot capture.",
        )
        main_window.hide()

        # 2. Capture Graph Screenshot
        graph_image = self.take_interactive_screenshot("select the **GRAPH** area")
        
        if graph_image:
            graph_image.save(self.graph_path)
            print("Graph screenshot captured successfully.")
        else:
            main_window.show()
            return 

        # 3. Create Word Report
        print("Starting Word document creation...")
        try:
            document = Document(self.input_template_path)
        except Exception as e:
            QMessageBox.critical(None, "Error Loading Document", f"Failed to load document '{self.input_template_path}'. Error: {e}")
            main_window.show()
            return

        # Get the current date
        current_date_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Find and replace the placeholder with content
        document, success = self.insert_content_at_placeholder(
            document, 
            self.placeholder_text, 
            self.graph_path, 
            current_date_str
        )
        
        if not success:
            main_window.show()
            return

        document.save(self.final_report_name)
        print(f"\n✅ Final Word report saved as: {self.final_report_name}")

        # 4. Convert to PDF (Optional/Non-Essential)
        try:
            import docxtopdf
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
        
        main_window.show()
        QMessageBox.information(
            None,
            "Finished",
            f"Report '{self.final_report_name}' has been created from '{os.path.basename(self.input_template_path)}'.",
        )


if __name__ == '__main__':
    # Initialize the QApplication instance before creating any Qt widgets
    app = QApplication(sys.argv)
    main_app = App()
    main_app.run()
    # Execute the application's event loop
    sys.exit(app.exec_())