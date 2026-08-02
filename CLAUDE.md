# Claude Instructions for docx-to-pdf

## Project Overview

A Windows Python/tkinter utility that converts Microsoft Word documents (.docx, .doc) to PDF by automating Word itself via COM. Built for a pastor preparing sermons, lessons, and articles. The tool achieves pixel-perfect formatting fidelity because it uses Word's native PDF export, not a pure-Python renderer.

**Hard constraints:** Windows-only. Requires Microsoft Word installed. Single-file scope is intentional (no batch mode). Do not add these features without being explicitly asked.

## COM Automation Rules

The `convert()` function uses Microsoft Word's COM interface. The following rules **MUST be preserved** when editing conversion code. Each exists for a critical reason:

### 1. DispatchEx (not Dispatch)
```python
word = win32com.client.DispatchEx("Word.Application")
```
- `DispatchEx` creates a separate, independent Word process.
- `Dispatch` would reuse an existing Word instance if one is open.
- **Why:** If the user has Word open with their own documents, a failed conversion or crash must never affect or be killed by their existing Word window. DispatchEx isolates the conversion.

### 2. Visible=False and DisplayAlerts=0
```python
word.Visible = False
word.DisplayAlerts = 0  # wdAlertsNone
```
- **Why:** A document with tracked changes, a broken linked image, or missing fonts would trigger a modal dialog asking the user to intervene. This dialog is invisible to the automation layer, causing a silent hang.
- Setting these to False/0 suppresses all UI prompts.

### 3. Documents.Open with ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False
```python
doc = word.Documents.Open(
    src_path,
    ReadOnly=True,
    AddToRecentFiles=False,
    ConfirmConversions=False
)
```
- **ReadOnly=True:** Never mutate the user's source document. Even if there are unsaved changes in the Word model, they are discarded on close.
- **AddToRecentFiles=False:** Opening a document for conversion is not a "user action" and should not pollute the Word recent-files list.
- **ConfirmConversions=False:** Suppresses format-confirmation dialogs (e.g., "This file is .doc; open as .doc or try to convert to .docx?").

### 4. ExportAsFixedFormat with ExportFormat=17
```python
doc.ExportAsFixedFormat(
    OutputFileName=pdf_path,
    ExportFormat=17  # wdFormatPDF
)
```
- Format 17 is the Word constant for PDF export. Do not hardcode a different value.
- This is Word's native PDF export, preserving all formatting, fonts, images, and layout exactly as the user sees it.

### 5. try/finally with Guaranteed Cleanup
```python
finally:
    if doc is not None:
        try:
            doc.Close(False)  # False = don't save changes
        except Exception:
            pass
    
    if word is not None:
        try:
            word.Quit()
        except Exception:
            pass
    
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass
```
- **doc.Close(False):** Close without saving. The `False` is essential; never omit it.
- **word.Quit():** Always quit the Word process. A failed run must not leave an orphaned `WINWORD.EXE` in Task Manager.
- **pythoncom.CoUninitialize():** Clean up the COM context.
- **Always** use try/except in finally blocks. Even cleanup can fail; do not let that crash the error handler.
- **Why:** A single failed run that doesn't clean up will leave Word.Application running indefinitely, potentially exhausting resources or preventing future conversions.

### 6. Absolute Paths Only
```python
src_path = os.path.abspath(str(src_path))
out_dir = os.path.abspath(str(out_dir))
```
- Word's COM interface does not accept relative paths; it requires absolute Windows paths.
- Normalize on entry; do not pass relative paths to any Word method.

## Threading Model

- Conversion runs on a **worker thread** (`threading.Thread`, daemon=True).
- The worker calls `convert()`, which handles COM initialization on that thread (`pythoncom.CoInitialize()` per-thread).
- **CRITICAL:** Never touch tkinter widgets directly from the worker thread. All GUI updates must be marshaled back to the main thread via `root.after(0, callback, *args)`.
- Example (from the code):
  ```python
  self.root.after(0, self._set_status, f"Created {pdf_filename}", "success")
  ```
  This schedules `_set_status` to run on the main event loop.

## Error Handling

The `convert()` function raises three exception types, each with a plain-English message:

1. **FileNotFoundError** — source file missing, is not a file, or output directory missing/not writable.
2. **PermissionError** — output directory not writable, or output PDF is locked by another program.
3. **RuntimeError** — Word not installed or conversion failed (e.g., corrupted document, unsupported format).

The GUI layer catches each and displays the exception message to the user as status text (red for errors, green for success). **Never let a raw traceback reach the user.**

## Testing

- No automated test suite currently exists (known gap).
- Manual verification:
  1. Import `convert()` and call it directly with a real .docx file into a scratch output directory.
  2. Verify the PDF exists, is > 1 KB, and its first bytes are `%PDF`.
  3. Check for orphaned processes: `tasklist /FI "IMAGENAME eq WINWORD.EXE"` should show none after a successful run.
  4. Error path: call `convert()` with a nonexistent file; confirm a clean `FileNotFoundError` rather than a traceback.
  5. **Caveat:** Always use a scratch output directory for testing, never write over production sermon PDFs.

## Code Style

- Plain, functional tkinter. Simple over pretty.
- Clear variable names and comments on non-obvious logic.
- No optimization premature to correctness.

## When Editing

- Do not remove or weaken any of the COM rules above without explicit approval.
- Do not add batch mode, recursive scanning, or skip-if-newer logic without being asked.
- Do not change the exception types or their messages without considering how they appear in the GUI.
- Do not refactor threading in a way that violates the worker-thread / main-thread separation.
