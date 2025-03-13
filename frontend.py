import sys
import os
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                           QWidget, QFileDialog, QProgressBar, QLabel, QTextEdit)
from PyQt5.QtCore import QThread, pyqtSignal
from comtypes import client

import win32com.client


class ConversionWorker(QThread):
    progress = pyqtSignal(int)
    log = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, files):
        super().__init__()
        self.files = files
        self.is_running = True

    def convert_single_file(self, ppt_path):
        try:
            powerpoint = win32com.client.Dispatch("Powerpoint.Application")
            abs_path = os.path.abspath(ppt_path)
            deck = powerpoint.Presentations.Open(abs_path)
            
            pdf_path = os.path.abspath(str(Path(ppt_path).with_suffix('.pdf')))
            deck.SaveAs(pdf_path, 32)  # ppSaveAsPDF = 32
            
            deck.Close()
            powerpoint.Quit()
            return True, f"Successfully converted: {ppt_path}"
        except Exception as e:
            try:
                deck.Close()
                powerpoint.Quit()
            except:
                pass
            return False, f"Error converting {ppt_path}: {str(e)}"
    def run(self):
        with ThreadPoolExecutor(max_workers=4) as executor:
            for i, result in enumerate(executor.map(self.convert_single_file, self.files)):
                if not self.is_running:
                    break
                success, message = result
                self.log.emit(message)
                self.progress.emit(int((i + 1) / len(self.files) * 100))
        self.finished.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bulk PPT to PDF Converter")
        self.setMinimumSize(600, 400)
        
        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # UI Elements
        self.select_btn = QPushButton("Select PPT Files")
        self.convert_btn = QPushButton("Convert to PDF")
        self.convert_btn.setEnabled(False)
        
        self.status_label = QLabel("No files selected")
        self.progress_bar = QProgressBar()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        
        # Add widgets to layout
        layout.addWidget(self.select_btn)
        layout.addWidget(self.status_label)
        layout.addWidget(self.convert_btn)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.log_text)
        
        # Connect signals
        self.select_btn.clicked.connect(self.select_files)
        self.convert_btn.clicked.connect(self.start_conversion)
        
        self.files = []
        self.worker = None

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select PowerPoint Files", "",
            "PowerPoint Files (*.ppt *.pptx)")
        
        if files:
            self.files = files
            self.status_label.setText(f"Selected {len(files)} files")
            self.convert_btn.setEnabled(True)
            self.log_text.clear()

    def start_conversion(self):
        if not self.files:
            return
            
        self.convert_btn.setEnabled(False)
        self.select_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        
        self.worker = ConversionWorker(self.files)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log.connect(self.log_message)
        self.worker.finished.connect(self.conversion_finished)
        self.worker.start()

    def log_message(self, message):
        self.log_text.append(message)

    def conversion_finished(self):
        self.convert_btn.setEnabled(True)
        self.select_btn.setEnabled(True)
        self.status_label.setText("Conversion completed")

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()