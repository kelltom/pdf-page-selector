# ChiselPDF

A simple desktop application for working with PDF files. ChiselPDF allows you to extract specific pages from a PDF document by specifying page ranges (e.g., 1-3,5,7-10), or split a large PDF into smaller chunks of a specified size. The application features a clean, dark-themed interface built with PyQt6 and supports batch page selection with an intuitive range syntax.

## Getting Started

### Clone the Repository

```bash
git clone https://github.com/kelltom/pdf-page-selector.git
cd pdf-page-selector
```

### Set Up Virtual Environment

**Windows (PowerShell):**
```powershell
python -m venv env
.\env\Scripts\Activate.ps1
```

**macOS/Linux:**
```bash
python3 -m venv env
source env/bin/activate
```

### Install Dependencies

```bash
pip install -e .
```

### Run the Application

```bash
python main.py
```

## Building an Executable

### Windows

```powershell
pyinstaller --onefile --windowed --name "ChiselPDF" --add-data "style.qss;." main.py
```

The executable will be created in the `dist` folder.

### macOS

```bash
pyinstaller --onefile --windowed --name "ChiselPDF" --add-data "style.qss:." main.py
```

The application bundle will be created in the `dist` folder.

**Note:** On macOS, use a colon (`:`) as the separator in `--add-data`, while Windows uses a semicolon (`;`).
