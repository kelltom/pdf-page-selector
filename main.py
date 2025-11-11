import os
import sys
from pathlib import Path

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QLineEdit, 
                             QFileDialog, QMessageBox, QGroupBox, QRadioButton, 
                             QButtonGroup)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyPDF2 import PdfReader, PdfWriter


class Mode:
    """Represents a PDF processing mode with all associated metadata."""
    
    def __init__(self, name, display_name, section_title, placeholder, help_text, 
                 process_func, output_suffix):
        self.name = name
        self.display_name = display_name
        self.section_title = section_title
        self.placeholder = placeholder
        self.help_text = help_text
        self.process_func = process_func
        self.output_suffix = output_suffix


def load_stylesheet():
    """Load the QSS stylesheet from a file."""
    try:
        with open("style.qss", "r") as f:
            return f.read()
    except FileNotFoundError:
        print("Warning: style.qss not found. Using default styles.")
        return ""


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


def split_pdf(input_path, chunk_size, output_path):
    """Split a PDF into multiple files with specified chunk size."""
    reader = PdfReader(input_path)
    total_pages = len(reader.pages)
    
    if chunk_size <= 0:
        raise ValueError("Chunk size must be a positive integer.")
    
    # Determine output naming
    output_file = Path(output_path)
    base_name = output_file.stem
    output_dir = output_file.parent
    
    created_files = []
    chunk_num = 1
    
    for start_page in range(0, total_pages, chunk_size):
        writer = PdfWriter()
        end_page = min(start_page + chunk_size, total_pages)
        
        for page_idx in range(start_page, end_page):
            writer.add_page(reader.pages[page_idx])
        
        # Generate output filename
        output_filename = output_dir / f"{base_name}_part{chunk_num}.pdf"
        with open(output_filename, "wb") as f:
            writer.write(f)
        
        created_files.append(str(output_filename))
        chunk_num += 1
    
    num_chunks = len(created_files)
    message = f"Successfully split PDF into {num_chunks} file{'s' if num_chunks > 1 else ''}"
    message += f"\n\nCreated {num_chunks} PDF{'s' if num_chunks > 1 else ''} in:\n{output_dir}"
    
    return message


class PDFPageSelectorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.input_path = None
        self.output_path = None
        
        # Define available modes
        self.modes = [
            Mode(
                name="selection",
                display_name="Selection",
                section_title="Page Ranges",
                placeholder="e.g., 1-3,5,6-9,11",
                help_text="Examples: 1-3,5,6-9,11  or  1,3,5  or  10-20",
                process_func=self._process_selection,
                output_suffix="_trimmed"
            ),
            Mode(
                name="split",
                display_name="Split",
                section_title="Chunk Size",
                placeholder="e.g., 10",
                help_text="Number of pages per file (greater than 0)",
                process_func=self._process_split,
                output_suffix="_split"
            )
        ]
        
        # Default to first mode in list
        self.current_mode = self.modes[0]
        
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("PDF Page Selector")
        self.setMinimumSize(400, 580)
        self.resize(400, 580)
        self.setStyleSheet(load_stylesheet())
        
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
        
        # Mode selection
        layout.addWidget(self._create_mode_section())
        
        # Page ranges section (dynamic based on mode)
        self.page_section = self._create_page_section()
        layout.addWidget(self.page_section)
        
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
    
    def _create_mode_section(self):
        group = QGroupBox("Mode")
        layout = QHBoxLayout()
        
        self.mode_group = QButtonGroup(self)
        
        # Dynamically create radio buttons for each mode
        for i, mode in enumerate(self.modes):
            radio = QRadioButton(mode.display_name)
            if i == 0:  # Select first mode by default
                radio.setChecked(True)
            radio.toggled.connect(lambda checked, m=mode: self._on_mode_changed(checked, m))
            self.mode_group.addButton(radio)
            layout.addWidget(radio)
        
        layout.addStretch()
        
        group.setLayout(layout)
        return group
    
    def _create_page_section(self):
        group = QGroupBox("Page Ranges")
        layout = QVBoxLayout()
        
        self.page_entry = QLineEdit()
        self.page_entry.setPlaceholderText("e.g., 1-3,5,6-9,11")
        layout.addWidget(self.page_entry)
        
        self.help_label = QLabel("Examples: 1-3,5,6-9,11  or  1,3,5  or  10-20")
        self.help_label.setStyleSheet("color: #aaaaaa; font-size: 9pt;")
        layout.addWidget(self.help_label)
        
        group.setLayout(layout)
        return group
    
    def _update_page_section_for_mode(self):
        """Update the page section UI based on current mode."""
        self.page_section.setTitle(self.current_mode.section_title)
        self.page_entry.setPlaceholderText(self.current_mode.placeholder)
        self.help_label.setText(self.current_mode.help_text)
        self.page_entry.clear()
    
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
    
    def _on_mode_changed(self, checked, mode):
        """Handle mode radio button changes."""
        if checked:
            self.current_mode = mode
            self._update_page_section_for_mode()
            self._update_output_suggestion()
    
    def _update_output_suggestion(self):
        """Update the output path suggestion based on mode and input."""
        if not self.input_path:
            return
        
        input_file = Path(self.input_path)
        self.output_path = str(input_file.parent / f"{input_file.stem}{self.current_mode.output_suffix}.pdf")
        
        self._update_label(self.output_label, self.output_path, active=True)
    
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
            self._update_output_suggestion()
            
            self.process_button.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not read PDF: {str(e)}")
            self.input_path = None
            self._update_label(self.input_label, "No file selected")
            self.page_info_label.setText("")
    
    def browse_output(self):
        if self.input_path:
            input_file = Path(self.input_path)
            initial = f"{input_file.parent}/{input_file.stem}{self.current_mode.output_suffix}.pdf"
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
            # Determine appropriate input type based on current mode
            input_type = self.current_mode.section_title.lower()
            QMessageBox.warning(self, "Error", f"Please enter {input_type}.")
            return
        
        self.status_label.setText("Processing...")
        self.status_label.setStyleSheet("color: #aaaaaa;")
        QApplication.processEvents()
        
        try:
            # Use the current mode's process function
            self.current_mode.process_func(page_input)
        except Exception as e:
            self.status_label.setText("Failed")
            self.status_label.setStyleSheet("color: #f44336;")
            QMessageBox.critical(self, "Error", str(e))
    
    def _process_selection(self, page_input):
        """Process PDF in selection mode."""
        self._ensure_output_path()
        try:
            page_numbers = parse_page_ranges(page_input)
        except (ValueError, AttributeError) as e:
            raise ValueError(f"Invalid page format:\n{str(e)}")
        
        if self.output_path and os.path.exists(self.output_path):
            if not self._confirm_overwrite():
                return
        
        message = trim_pdf(self.input_path, page_numbers, self.output_path)
        self._show_success(message)
    
    def _process_split(self, chunk_input):
        """Process PDF in split mode."""
        self._ensure_output_path()
        try:
            chunk_size = int(chunk_input)
            if chunk_size <= 0:
                raise ValueError("Chunk size must be a positive integer.")
        except ValueError:
            raise ValueError("Chunk size must be a positive integer.")
        
        # Check if any output files would be overwritten
        if self.output_path:
            output_file = Path(self.output_path)
            base_name = output_file.stem
            output_dir = output_file.parent
            
            # Check if part1 exists as a simple overwrite check
            first_file = output_dir / f"{base_name}_part1.pdf"
            if first_file.exists():
                if not self._confirm_overwrite():
                    return
        
        message = split_pdf(self.input_path, chunk_size, self.output_path)
        self._show_success(message)
    
    def _confirm_overwrite(self):
        """Ask user to confirm file overwrite."""
        reply = QMessageBox.question(
            self, "Confirm Overwrite",
            f"Output file(s) already exist.\n\nOverwrite?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        return reply == QMessageBox.StandardButton.Yes
    
    def _show_success(self, message):
        """Display success message."""
        output_dir = Path(self.output_path).parent if self.output_path else ""
        self.status_label.setText(f"Success! Saved to: {output_dir}")
        self.status_label.setStyleSheet("color: #4caf50;")
        QMessageBox.information(self, "Success", message)

    def _ensure_output_path(self):
        """Validate the output path before writing files."""
        if not self.output_path:
            raise ValueError("Please choose an output location before processing.")

        if not self.input_path:
            return

        input_path = Path(self.input_path).resolve()
        output_path = Path(self.output_path).resolve()

        # Prevent overwriting the source file, which can corrupt the PDF while reading.
        if input_path == output_path:
            raise ValueError("Output file must be different from the input file.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFPageSelectorApp()
    window.show()
    sys.exit(app.exec())
