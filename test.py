import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QPushButton, QMessageBox, QDialog
from PyQt5.QtGui import QPixmap, QPainter, QPen, QColor, QImage
from PyQt5.QtCore import Qt, QPoint, QRect
from PIL import ImageGrab, Image # Pillow untuk screenshot
from docx import Document
from docx.shared import Inches
import os
import time

# --- Konfigurasi ---
NAMA_LAPORAN_FINAL = "Laporan_Uji_Interaktif_Final.docx"
FILE_FINAL_PDF = "Laporan_Uji_Interaktif_Final.pdf" # Hanya berlaku jika docxtopdf berfungsi
PATH_TEMPLATES = os.path.join(os.path.dirname(__file__), "templates") # Direktori untuk template Word
if not os.path.exists(PATH_TEMPLATES):
    os.makedirs(PATH_TEMPLATES)
TEMPLATE_LAPORAN_WORD = os.path.join(PATH_TEMPLATES, "template_laporan.docx")

# Buat template dummy jika belum ada
if not os.path.exists(TEMPLATE_LAPORAN_WORD):
    doc = Document()
    doc.add_heading('LAPORAN UJI', 0)
    doc.add_heading('Grafik Hasil Uji:', level=3)
    doc.add_paragraph(' [Area untuk Grafik] ')
    doc.add_heading('Tanggal Pengujian:', level=3)
    doc.add_paragraph(' [Area untuk Tanggal] ')
    doc.save(TEMPLATE_LAPORAN_WORD)
    print(f"Template Word dummy '{TEMPLATE_LAPORAN_WORD}' dibuat.")


class ScreenshotSelector(QDialog):
    def __init__(self, full_screen_image, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Pilih Area Screenshot")
        self.setGeometry(0, 0, full_screen_image.width(), full_screen_image.height())
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint) # Tanpa bingkai, selalu di atas
        self.setCursor(Qt.CrossCursor) # Kursor jadi tanda plus
        
        self.full_screen_pixmap = QPixmap.fromImage(full_screen_image)
        self.start_point = QPoint()
        self.end_point = QPoint()
        self.selection_rect = QRect()
        self.is_drawing = False

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.drawPixmap(0, 0, self.full_screen_pixmap)
        
        if self.is_drawing:
            pen = QPen(QColor(255, 0, 0)) # Warna merah
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
            self.close() # Tutup jendela setelah seleksi selesai

class App:
    def __init__(self):
        self.grafik_path = "grafik_temp.png"
        self.tanggal_path = "tanggal_temp.png"

    def ambil_screenshot_interaktif(self, prompt_text="Pilih area"):
        QMessageBox.information(None, "Siap Untuk Screenshot", f"Klik OK, lalu {prompt_text} dengan Drag-and-Drop di layar.")
        
        # Sembunyikan aplikasi PyQt sementara untuk mengambil screenshot layar penuh
        # ini penting agar jendela aplikasi kita tidak ikut di-screenshot
        app = QApplication.instance()
        if app:
            for widget in app.topLevelWidgets():
                widget.hide()
        
        time.sleep(0.5) # Beri waktu untuk menyembunyikan jendela

        full_screen_image_pil = ImageGrab.grab() # Ambil screenshot seluruh layar
        
        if app:
            for widget in app.topLevelWidgets():
                widget.show() # Tampilkan kembali jendela aplikasi PyQt setelah screenshot
        
        # Konversi PIL Image ke format yang bisa digunakan oleh Qt
        img_data = full_screen_image_pil.tobytes('raw', 'RGB')
        full_screen_image_qt = QImage(img_data, full_screen_image_pil.size[0], full_screen_image_pil.size[1], full_screen_image_pil.size[0] * 3, QImage.Format_RGB888)
        
        selector = ScreenshotSelector(full_screen_image_qt)
        # Jalankan dialog dan tunggu sampai selesai
        selector.exec()
        
        # Setelah user selesai menyeleksi dan jendela selector ditutup,
        # kita bisa mendapatkan koordinat seleksi
        x, y, w, h = selector.selection_rect.x(), selector.selection_rect.y(), selector.selection_rect.width(), selector.selection_rect.height()
        
        if w == 0 or h == 0:
            QMessageBox.warning(None, "Peringatan", "Area tidak dipilih atau terlalu kecil. Silakan coba lagi.")
            return None
            
        # Potong gambar sesuai seleksi
        cropped_image_pil = full_screen_image_pil.crop((x, y, x + w, y + h))
        return cropped_image_pil

    def run(self):
        QMessageBox.information(None, "Mulai Aplikasi", "Aplikasi siap mengambil screenshot untuk laporan. Klik OK.")
        
        # 1. Ambil Screenshot Grafik
        grafik_image = self.ambil_screenshot_interaktif("pilih area GRAFIK")
        if grafik_image:
            grafik_image.save(self.grafik_path)
            print("Grafik berhasil di-screenshot.")
        else:
            return

        # 2. Ambil Screenshot Tanggal
        tanggal_image = self.ambil_screenshot_interaktif("pilih area TANGGAL")
        if tanggal_image:
            tanggal_image.save(self.tanggal_path)
            print("Tanggal berhasil di-screenshot.")
        else:
            # Jika user tidak memilih tanggal, kita bisa skip atau pakai tanggal sistem
            QMessageBox.information(None, "Info", "Pengambilan tanggal dilewati.")
            # Misalnya, gunakan tanggal sistem jika tidak ada screenshot tanggal
            # document.add_paragraph(f"Tanggal Pengujian: {time.strftime('%Y-%m-%d')}")


        # 3. Buat Laporan Word
        print("Memulai pembuatan dokumen Word...")
        document = Document(TEMPLATE_LAPORAN_WORD)
        
        # Cari dan ganti placeholder atau sisipkan di akhir
        # Untuk kasus ini, kita sisipkan di akhir dokumen, atau Anda bisa menargetkan placeholder spesifik
        
        # Sisipkan Tanggal (jika ada)
        if os.path.exists(self.tanggal_path):
            document.add_picture(self.tanggal_path, width=Inches(2)) # Sesuaikan ukuran
            document.add_paragraph('\n')
        
        # Sisipkan Grafik (jika ada)
        if os.path.exists(self.grafik_path):
            document.add_picture(self.grafik_path, width=Inches(6)) # Sesuaikan ukuran
            document.add_paragraph('\n')
        
        document.save(NAMA_LAPORAN_FINAL)
        print(f"\n✅ Laporan Word final disimpan sebagai: {NAMA_LAPORAN_FINAL}")

        # 4. Konversi ke PDF (Hanya untuk Windows)
        try:
            import docxtopdf
            docxtopdf.convert(NAMA_LAPORAN_FINAL)
            print(f"✅ Laporan PDF final disimpan sebagai: {FILE_FINAL_PDF}")
        except ImportError:
            print("Peringatan: docxtopdf tidak terinstal atau tidak mendukung OS ini. PDF tidak dibuat.")
        except Exception as e:
            print(f"Error saat mengkonversi ke PDF: {e}. Pastikan MS Word terinstal.")


        # 5. Bersihkan file temporer
        if os.path.exists(self.grafik_path):
            os.remove(self.grafik_path)
        if os.path.exists(self.tanggal_path):
            os.remove(self.tanggal_path)
        print("File temporer dibersihkan.")
        QMessageBox.information(None, "Selesai", f"Laporan '{NAMA_LAPORAN_FINAL}' dan '{FILE_FINAL_PDF}' telah dibuat.")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_app = App()
    main_app.run()
    sys.exit(app.exec_())