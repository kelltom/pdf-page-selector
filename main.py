import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
try:
    from PyPDF2 import PdfReader, PdfWriter
except ImportError:
    print("PyPDF2 library is required. Installing...")
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "PyPDF2"])
    from PyPDF2 import PdfReader, PdfWriter


def parse_page_ranges(page_input):
    """
    Parse page ranges like '1-3,5,6-9,11' into a list of page numbers.
    Returns a sorted list of unique page numbers (1-indexed).
    """
    pages = set()
    parts = page_input.replace(" ", "").split(",")
    
    for part in parts:
        if not part:  # Skip empty parts
            continue
        if "-" in part:
            # Handle range like "1-3"
            start, end = part.split("-", 1)
            try:
                start_page = int(start)
                end_page = int(end)
                if start_page > end_page:
                    raise ValueError(f"Invalid range {part} (start > end)")
                pages.update(range(start_page, end_page + 1))
            except ValueError as e:
                raise ValueError(f"Invalid range format '{part}': {e}")
        else:
            # Handle single page number
            try:
                pages.add(int(part))
            except ValueError:
                raise ValueError(f"Invalid page number '{part}'")
    
    return sorted(pages)


def trim_pdf(input_path, page_numbers, output_path):
    """
    Create a new PDF with only the specified pages.
    
    Args:
        input_path: Path to the input PDF file
        page_numbers: List of page numbers to include (1-indexed)
        output_path: Path where the output PDF will be saved
    
    Returns:
        tuple: (success: bool, message: str, valid_pages: list)
    """
    try:
        reader = PdfReader(input_path)
        writer = PdfWriter()
        
        total_pages = len(reader.pages)
        
        # Validate and add pages
        valid_pages = []
        invalid_pages = []
        for page_num in page_numbers:
            if 1 <= page_num <= total_pages:
                # Convert to 0-indexed for PyPDF2
                writer.add_page(reader.pages[page_num - 1])
                valid_pages.append(page_num)
            else:
                invalid_pages.append(page_num)
        
        if not valid_pages:
            return False, "No valid pages to include in the output PDF.", []
        
        # Write the output PDF
        with open(output_path, "wb") as output_file:
            writer.write(output_file)
        
        message = f"Successfully created PDF with {len(valid_pages)} pages"
        if invalid_pages:
            message += f"\n\nSkipped invalid pages: {invalid_pages} (PDF has {total_pages} pages)"
        
        return True, message, valid_pages
        
    except Exception as e:
        return False, f"Error processing PDF: {str(e)}", []


class PDFPageSelectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Page Selector")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        self.input_path = None
        self.total_pages = 0
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="PDF Page Selector", 
                                font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Input PDF section
        ttk.Label(main_frame, text="Input PDF:").grid(row=1, column=0, sticky=tk.W, pady=5)
        
        self.input_label = ttk.Label(main_frame, text="No file selected", 
                                     foreground="gray", relief="sunken", padding=5)
        self.input_label.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Button(main_frame, text="Browse...", 
                  command=self.browse_input).grid(row=1, column=2, pady=5)
        
        # Page info label
        self.page_info_label = ttk.Label(main_frame, text="", foreground="blue")
        self.page_info_label.grid(row=2, column=1, sticky=tk.W, padx=5)
        
        # Page ranges section
        ttk.Label(main_frame, text="Page Ranges:").grid(row=3, column=0, sticky=tk.W, pady=5)
        
        self.page_entry = ttk.Entry(main_frame, width=40)
        self.page_entry.grid(row=3, column=1, sticky="ew", padx=5, pady=5)
        
        # Help text
        help_frame = ttk.Frame(main_frame)
        help_frame.grid(row=4, column=1, sticky="ew", padx=5, pady=5)
        
        help_text = ttk.Label(help_frame, 
                             text="Examples: 1-3,5,6-9,11  or  1,3,5  or  10-20",
                             foreground="gray", font=("Arial", 9))
        help_text.pack(anchor=tk.W)
        
        # Output section
        ttk.Label(main_frame, text="Output PDF:").grid(row=5, column=0, sticky=tk.W, pady=5)
        
        self.output_label = ttk.Label(main_frame, text="No output location selected", 
                                      foreground="gray", relief="sunken", padding=5)
        self.output_label.grid(row=5, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Button(main_frame, text="Browse...", 
                  command=self.browse_output).grid(row=5, column=2, pady=5)
        
        # Process button
        self.process_button = ttk.Button(main_frame, text="Create Trimmed PDF", 
                                        command=self.process_pdf, state=tk.DISABLED)
        self.process_button.grid(row=6, column=0, columnspan=3, pady=20)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="", wraplength=550)
        self.status_label.grid(row=7, column=0, columnspan=3, pady=10)
        
    def browse_input(self):
        filename = filedialog.askopenfilename(
            title="Select Input PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if filename:
            self.input_path = filename
            # Truncate long paths for display
            display_name = filename if len(filename) < 60 else "..." + filename[-57:]
            self.input_label.config(text=display_name, foreground="black")
            
            # Get page count
            try:
                reader = PdfReader(filename)
                self.total_pages = len(reader.pages)
                self.page_info_label.config(text=f"(Total pages: {self.total_pages})")
                
                # Suggest default output
                input_file = Path(filename)
                default_output = input_file.parent / f"{input_file.stem}_trimmed{input_file.suffix}"
                self.output_path = str(default_output)
                display_output = str(default_output) if len(str(default_output)) < 60 else "..." + str(default_output)[-57:]
                self.output_label.config(text=display_output, foreground="black")
                
                self.update_process_button()
            except Exception as e:
                messagebox.showerror("Error", f"Could not read PDF: {str(e)}")
                self.input_path = None
                self.total_pages = 0
                self.input_label.config(text="No file selected", foreground="gray")
                self.page_info_label.config(text="")
    
    def browse_output(self):
        if self.input_path:
            input_file = Path(self.input_path)
            initial_file = f"{input_file.stem}_trimmed{input_file.suffix}"
            initial_dir = input_file.parent
        else:
            initial_file = "output_trimmed.pdf"
            initial_dir = Path.home()
        
        filename = filedialog.asksaveasfilename(
            title="Save Trimmed PDF As",
            initialfile=initial_file,
            initialdir=initial_dir,
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if filename:
            self.output_path = filename
            display_name = filename if len(filename) < 60 else "..." + filename[-57:]
            self.output_label.config(text=display_name, foreground="black")
            self.update_process_button()
    
    def update_process_button(self):
        if self.input_path and hasattr(self, 'output_path'):
            self.process_button.config(state=tk.NORMAL)
        else:
            self.process_button.config(state=tk.DISABLED)
    
    def process_pdf(self):
        # Validate page input
        page_input = self.page_entry.get().strip()
        if not page_input:
            messagebox.showerror("Error", "Please enter page ranges.")
            return
        
        try:
            page_numbers = parse_page_ranges(page_input)
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid page format:\n{str(e)}")
            return
        
        if not page_numbers:
            messagebox.showerror("Error", "No valid page numbers found.")
            return
        
        # Check if output file exists
        if os.path.exists(self.output_path):
            if not messagebox.askyesno("Confirm Overwrite", 
                                       f"File already exists:\n{self.output_path}\n\nOverwrite?"):
                return
        
        # Process the PDF
        self.status_label.config(text="Processing...", foreground="blue")
        self.root.update()
        
        success, message, valid_pages = trim_pdf(self.input_path, page_numbers, self.output_path)
        
        if success:
            self.status_label.config(text=f"Success! Saved to:\n{self.output_path}", 
                                   foreground="green")
            messagebox.showinfo("Success", message)
        else:
            self.status_label.config(text="Failed", foreground="red")
            messagebox.showerror("Error", message)


def main():
    root = tk.Tk()
    app = PDFPageSelectorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
