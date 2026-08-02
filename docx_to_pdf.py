"""
DOCX to PDF Converter using Microsoft Word COM automation.

This module provides:
- convert(src_path, out_dir) - importable conversion function
- GUI app when run as __main__
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, ttk
from pathlib import Path
import threading
import traceback

# COM imports
try:
    import win32com.client
    import pythoncom
except ImportError:
    print("ERROR: pywin32 is required. Install it with: pip install pywin32")
    sys.exit(1)


def convert(src_path, out_dir):
    """
    Convert a single DOCX/DOC file to PDF using Microsoft Word COM automation.

    Args:
        src_path: Absolute Windows path to .docx/.doc file
        out_dir: Absolute Windows path to output directory

    Returns:
        str: Absolute path to the created PDF file

    Raises:
        FileNotFoundError: If source file does not exist or cannot be read
        PermissionError: If output file is locked or out_dir is not writable
        RuntimeError: If Word is not installed or conversion fails
    """
    # Normalize to absolute paths
    src_path = os.path.abspath(str(src_path))
    out_dir = os.path.abspath(str(out_dir))

    # Validate inputs
    if not os.path.exists(src_path):
        raise FileNotFoundError(f"Source file not found: {src_path}")

    if not os.path.isfile(src_path):
        raise FileNotFoundError(f"Path is not a file: {src_path}")

    if not os.path.exists(out_dir):
        raise FileNotFoundError(f"Output directory does not exist: {out_dir}")

    if not os.access(out_dir, os.W_OK):
        raise PermissionError(f"Output directory is not writable: {out_dir}")

    # Compute output filename
    stem = Path(src_path).stem
    pdf_path = os.path.join(out_dir, stem + ".pdf")

    # Initialize COM for this thread
    pythoncom.CoInitialize()

    word = None
    doc = None

    try:
        # Create a separate Word process (DispatchEx, not Dispatch)
        word = win32com.client.DispatchEx("Word.Application")

        # Suppress UI interactions
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        # Open the document read-only
        doc = word.Documents.Open(
            src_path,
            ReadOnly=True,
            AddToRecentFiles=False,
            ConfirmConversions=False
        )

        # Export as PDF (format 17 = wdFormatPDF)
        doc.ExportAsFixedFormat(
            OutputFileName=pdf_path,
            ExportFormat=17
        )

        return pdf_path

    except FileNotFoundError:
        raise
    except PermissionError:
        raise
    except Exception as e:
        error_msg = str(e)

        # Distinguish common COM errors
        if "Word.Application" in error_msg or "dispatch" in error_msg.lower():
            raise RuntimeError("Microsoft Word is required but wasn't found.")
        elif "Permission denied" in error_msg or "is locked" in error_msg.lower():
            raise PermissionError(
                "Could not write the PDF. Close it if it's open in another program and try again."
            )
        else:
            raise RuntimeError(f"Conversion failed: {error_msg}")

    finally:
        # Always close the document
        if doc is not None:
            try:
                doc.Close(False)  # False = don't save changes
            except Exception:
                pass

        # Always quit Word
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

        # Uninitialize COM
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


class DocxToPdfApp:
    """Tkinter GUI for DOCX to PDF conversion."""

    def __init__(self, root):
        self.root = root
        self.root.title("DOCX -> PDF Converter")
        self.root.geometry("500x250")
        self.root.resizable(False, False)

        self.src_path = tk.StringVar()
        self.out_dir = tk.StringVar()
        self.is_converting = False

        self._build_ui()

    def _build_ui(self):
        """Build the GUI layout."""
        # Title
        title = tk.Label(self.root, text="DOCX -> PDF Converter", font=("Arial", 14, "bold"))
        title.grid(row=0, column=0, columnspan=3, pady=10)

        # Source file row
        tk.Label(self.root, text="Word file:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        src_entry = tk.Entry(self.root, textvariable=self.src_path, state="readonly", width=40)
        src_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        self.browse_src_btn = tk.Button(self.root, text="Browse", command=self._browse_source)
        self.browse_src_btn.grid(row=1, column=2, padx=5, pady=5)

        # Output directory row
        tk.Label(self.root, text="Save to:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        out_entry = tk.Entry(self.root, textvariable=self.out_dir, state="readonly", width=40)
        out_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        self.browse_out_btn = tk.Button(self.root, text="Browse", command=self._browse_output)
        self.browse_out_btn.grid(row=2, column=2, padx=5, pady=5)

        # Convert button
        self.convert_btn = tk.Button(
            self.root, text="Convert", command=self._on_convert, state="disabled"
        )
        self.convert_btn.grid(row=3, column=0, columnspan=3, pady=15)

        # Status label
        self.status_label = tk.Label(
            self.root, text="", wraplength=480, justify="left", fg="black"
        )
        self.status_label.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

        # Configure grid weights
        self.root.columnconfigure(1, weight=1)

    def _browse_source(self):
        """Browse for source DOCX/DOC file."""
        path = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[
                ("Word Documents", "*.docx;*.doc"),
                ("All files", "*.*")
            ]
        )
        if path:
            self.src_path.set(path)
            # Default output dir to the source file's directory
            self.out_dir.set(os.path.dirname(path))
            self.convert_btn.config(state="normal")
            self.status_label.config(text="")

    def _browse_output(self):
        """Browse for output directory."""
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self.out_dir.set(path)

    def _on_convert(self):
        """Handle Convert button click (start worker thread)."""
        if not self.src_path.get() or not self.out_dir.get():
            self._set_status("Please select both a source file and output folder.", "error")
            return

        # Disable buttons during conversion
        self.convert_btn.config(state="disabled")
        self.browse_src_btn.config(state="disabled")
        self.browse_out_btn.config(state="disabled")
        self._set_status("Converting...", "")
        self.is_converting = True

        # Start worker thread
        thread = threading.Thread(target=self._convert_worker, daemon=True)
        thread.start()

    def _convert_worker(self):
        """Worker thread for conversion (runs COM code)."""
        try:
            pdf_path = convert(self.src_path.get(), self.out_dir.get())
            pdf_filename = os.path.basename(pdf_path)
            self.root.after(0, self._set_status, f"Created {pdf_filename}", "success")
        except FileNotFoundError as e:
            self.root.after(0, self._set_status, str(e), "error")
        except PermissionError as e:
            self.root.after(0, self._set_status, str(e), "error")
        except RuntimeError as e:
            self.root.after(0, self._set_status, str(e), "error")
        except Exception as e:
            msg = f"Unexpected error: {str(e)}"
            self.root.after(0, self._set_status, msg, "error")
        finally:
            self.is_converting = False
            self.root.after(0, self._enable_buttons)

    def _set_status(self, message, status_type=""):
        """Update status label."""
        if status_type == "error":
            self.status_label.config(text=message, fg="red")
        elif status_type == "success":
            self.status_label.config(text=message, fg="green")
        else:
            self.status_label.config(text=message, fg="black")

    def _enable_buttons(self):
        """Re-enable buttons after conversion."""
        if self.src_path.get():
            self.convert_btn.config(state="normal")
        self.browse_src_btn.config(state="normal")
        self.browse_out_btn.config(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    app = DocxToPdfApp(root)
    root.mainloop()
