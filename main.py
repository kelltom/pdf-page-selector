import os
import sys
from pathlib import Path

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QLineEdit, 
                             QFileDialog, QMessageBox, QGroupBox)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyPDF2 import PdfReader, PdfWriter


# Stylesheet constants
DARK_THEME = """
    QMainWindow, QWidget {
        background-color: #2b2b2b;
        color: #ffffff;
    }
    QGroupBox {
        border: 1px solid #555555;
        border-radius: 5px;
        margin-top: 10px;
        padding-top: 10px;
        font-weight: bold;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 5px;
    }
    QLineEdit {
        background-color: #3c3c3c;
        color: #ffffff;
        border: 1px solid #555555;
        border-radius: 3px;
        padding: 5px;
    }
    QPushButton {
        background-color: #0d47a1;
        color: #ffffff;
        border: none;
        border-radius: 3px;
        padding: 8px;
        font-weight: bold;
    }
    QPushButton:hover {
        background-color: #1565c0;
    }
    QPushButton:pressed {
        background-color: #0a3d91;
    }
    QPushButton:disabled {
        background-color: #555555;
        color: #888888;
    }
"""


def parse_page_ranges(page_input):
    """Parse page ranges like '1-3,5,6-9,11' into a sorted list of page numbers."""
    pages = set()
    
    for part in page_input.replace(" ", "").split(","):
        if not part:
            continue
            
        if "-" in part:
            start, end = part.split("-", 1)
            start_page, end_page = int(start), int(end)
            if start_page > end_page:
                raise ValueError(f"Invalid range {part} (start > end)")
            pages.update(range(start_page, end_page + 1))
        else:
            pages.add(int(part))
    
    return sorted(pages)


def trim_pdf(input_path, page_numbers, output_path):
    """Create a new PDF with only the specified pages."""
    reader = PdfReader(input_path)
    writer = PdfWriter()
    total_pages = len(reader.pages)
    
    valid_pages = []
    invalid_pages = []
    
    for page_num in page_numbers:
        if 1 <= page_num <= total_pages:
            writer.add_page(reader.pages[page_num - 1])
            valid_pages.append(page_num)
        else:
            invalid_pages.append(page_num)
    
    if not valid_pages:
        raise ValueError("No valid pages to include in the output PDF.")
    
    with open(output_path, "wb") as f:
        writer.write(f)
    
    message = f"Successfully created PDF with {len(valid_pages)} pages"
    if invalid_pages:
        message += f"\n\nSkipped invalid pages: {invalid_pages} (PDF has {total_pages} pages)"
    
    return message


class PDFPageSelectorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.input_path = None
        self.output_path = None
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("PDF Page Selector")
        self.setMinimumSize(400, 500)
        self.resize(400, 500)
        self.setStyleSheet(DARK_THEME)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Title
        title = QLabel("PDF Page Selector")
        title.setFont(QFont("", 18, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        layout.addSpacing(10)
        
        # Input section
        layout.addWidget(self._create_input_section())
        
        # Page ranges section
        layout.addWidget(self._create_page_section())
        
        # Output section
        layout.addWidget(self._create_output_section())
        
        # Process button
        self.process_button = QPushButton("Create Trimmed PDF")
        self.process_button.setMinimumHeight(40)
        self.process_button.setFont(QFont("", 11, QFont.Weight.Bold))
        self.process_button.clicked.connect(self.process_pdf)
        self.process_button.setEnabled(False)
        layout.addWidget(self.process_button)
        
        # Status label
        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)
        
        layout.addStretch()
    
    def _create_input_section(self):
        group = QGroupBox("Input PDF")
        layout = QVBoxLayout()
        
        row = QHBoxLayout()
        self.input_label = QLabel("No file selected")
        self.input_label.setStyleSheet("color: #888888; padding: 5px;")
        self.input_label.setWordWrap(True)
        row.addWidget(self.input_label, 1)
        
        browse_btn = QPushButton("Browse...")
        browse_btn.setFixedWidth(100)
        browse_btn.clicked.connect(self.browse_input)
        row.addWidget(browse_btn)
        
        layout.addLayout(row)
        
        self.page_info_label = QLabel("")
        self.page_info_label.setStyleSheet("color: #aaaaaa; font-size: 10pt;")
        layout.addWidget(self.page_info_label)
        
        group.setLayout(layout)
        return group
    
    def _create_page_section(self):
        group = QGroupBox("Page Ranges")
        layout = QVBoxLayout()
        
        self.page_entry = QLineEdit()
        self.page_entry.setPlaceholderText("e.g., 1-3,5,6-9,11")
        layout.addWidget(self.page_entry)
        
        help_label = QLabel("Examples: 1-3,5,6-9,11  or  1,3,5  or  10-20")
        help_label.setStyleSheet("color: #aaaaaa; font-size: 9pt;")
        layout.addWidget(help_label)
        
        group.setLayout(layout)
        return group
    
    def _create_output_section(self):
        group = QGroupBox("Output PDF")
        layout = QVBoxLayout()
        
        row = QHBoxLayout()
        self.output_label = QLabel("No output location selected")
        self.output_label.setStyleSheet("color: #888888; padding: 5px;")
        self.output_label.setWordWrap(True)
        row.addWidget(self.output_label, 1)
        
        browse_btn = QPushButton("Browse...")
        browse_btn.setFixedWidth(100)
        browse_btn.clicked.connect(self.browse_output)
        row.addWidget(browse_btn)
        
        layout.addLayout(row)
        group.setLayout(layout)
        return group
    
    def _update_label(self, label, text, active=False):
        """Update a label with truncated text and appropriate styling."""
        display_text = text if len(text) < 80 else "..." + text[-77:]
        label.setText(display_text)
        color = "#ffffff" if active else "#888888"
        label.setStyleSheet(f"color: {color}; padding: 5px;")
    
    def browse_input(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Input PDF", "", "PDF files (*.pdf);;All files (*.*)"
        )
        
        if not filename:
            return
        
        try:
            reader = PdfReader(filename)
            total_pages = len(reader.pages)
            
            self.input_path = filename
            self._update_label(self.input_label, filename, active=True)
            self.page_info_label.setText(f"Total pages: {total_pages}")
            
            # Auto-suggest output path
            input_file = Path(filename)
            self.output_path = str(input_file.parent / f"{input_file.stem}_trimmed.pdf")
            self._update_label(self.output_label, self.output_path, active=True)
            
            self.process_button.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not read PDF: {str(e)}")
            self.input_path = None
            self._update_label(self.input_label, "No file selected")
            self.page_info_label.setText("")
    
    def browse_output(self):
        if self.input_path:
            input_file = Path(self.input_path)
            initial = f"{input_file.parent}/{input_file.stem}_trimmed.pdf"
        else:
            initial = str(Path.home() / "output_trimmed.pdf")
        
        filename, _ = QFileDialog.getSaveFileName(
            self, "Save Trimmed PDF As", initial, "PDF files (*.pdf);;All files (*.*)"
        )
        
        if filename:
            self.output_path = filename
            self._update_label(self.output_label, filename, active=True)
    
    def process_pdf(self):
        page_input = self.page_entry.text().strip()
        if not page_input:
            QMessageBox.warning(self, "Error", "Please enter page ranges.")
            return
        
        try:
            page_numbers = parse_page_ranges(page_input)
        except (ValueError, AttributeError) as e:
            QMessageBox.warning(self, "Error", f"Invalid page format:\n{str(e)}")
            return
        
        if self.output_path and os.path.exists(self.output_path):
            reply = QMessageBox.question(
                self, "Confirm Overwrite",
                f"File already exists:\n{self.output_path}\n\nOverwrite?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return
        
        self.status_label.setText("Processing...")
        self.status_label.setStyleSheet("color: #aaaaaa;")
        QApplication.processEvents()
        
        try:
            message = trim_pdf(self.input_path, page_numbers, self.output_path)
            self.status_label.setText(f"Success! Saved to: {self.output_path}")
            self.status_label.setStyleSheet("color: #4caf50;")
            QMessageBox.information(self, "Success", message)
        except Exception as e:
            self.status_label.setText("Failed")
            self.status_label.setStyleSheet("color: #f44336;")
            QMessageBox.critical(self, "Error", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFPageSelectorApp()
    window.show()
    sys.exit(app.exec())
