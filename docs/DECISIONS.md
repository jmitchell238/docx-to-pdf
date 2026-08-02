# Architecture Decision Records

## ADR-001: Use Microsoft Word COM Automation for Conversion

**Context:**
Converting a Word document to PDF requires preserving pixel-perfect formatting, including fonts, colors, images, embedded charts, tables, and complex styles. Three approaches were considered:

1. **Microsoft Word COM automation** — launch Word, open the document, export to PDF.
2. **LibreOffice headless** — `libreoffice --headless --convert-to pdf document.docx`
3. **Pure-Python DOCX parsing** — libraries like `python-docx` parse the .docx file and render to PDF using a PDF library (e.g., reportlab).

**Decision:**
Use Microsoft Word COM automation (pywin32).

**Consequences:**
- ✓ **Fidelity:** Word's PDF export is byte-for-byte identical to what the user sees in Word. No formatting drift, no missing fonts, no image corruption.
- ✓ **Reliability:** Word handles all .docx quirks, corrupted images, broken hyperlinks, etc.
- ✗ **Platform constraint:** Windows only. Word is not available on Linux or macOS.
- ✗ **Dependency:** Requires Microsoft Word to be installed. LibreOffice and pure-Python approaches do not.
- ✗ **Performance:** Slower than headless tools (spawns a full Word process), but acceptable for single-document use.

**Why not LibreOffice?**
- LibreOffice headless is multi-platform and lightweight, but formatting often drifts from the original Word document (fonts, spacing, complex tables).
- For a pastor preparing sermons, this drift is unacceptable. A sermon must look identical in PDF as in Word.

**Why not pure-Python?**
- Pure-Python DOCX parsing (python-docx + reportlab) is fast and lightweight, but the fidelity is even lower than LibreOffice.
- Many formatting features in Word have no equivalent in these libraries.

**Trade-off:**
We accept the Windows + Word dependency in exchange for pixel-perfect fidelity. This is the only way to guarantee the PDF matches what the author prepared.

---

## ADR-002: Use tkinter for the GUI

**Context:**
The GUI needs to be simple and functional. The owner uses Python on Windows and prefers minimal external dependencies.

**Decision:**
Use tkinter (Python's standard GUI framework).

**Consequences:**
- ✓ **No extra dependencies:** tkinter comes with Python.
- ✓ **Minimal code:** The GUI is ~120 lines (browse buttons, text fields, one convert button, a status label).
- ✓ **Matches existing tools:** The owner uses tkinter for other projects; familiar patterns.
- ✗ **Basic aesthetics:** tkinter GUIs look functional but not polished. No modern styling.
- ✗ **Threading complexity:** tkinter is not thread-safe; conversion must run on a worker thread.

**Why not PyQt/PySide?**
- Would require installing and maintaining another large dependency.
- Overkill for a simple file-dialog + button + label app.

**Why not web (Flask + HTML)?**
- Unnecessary complexity; a desktop app is simpler and faster to iterate.
- Would require a running web server.

---

## ADR-003: Single-File Scope Only (No Batch Mode)

**Context:**
The owner requested a tool to convert one document at a time: a sermon, a Sunday School lesson, or a newspaper article. The workflow is:
1. Prepare the document in Word.
2. Click the desktop shortcut to launch the converter.
3. Select the document.
4. Get a PDF.
5. Close the app.

There is no production workflow that requires batch conversion (e.g., "convert all .docx files in this folder to PDFs").

**Decision:**
Support only single-file-at-a-time conversion. No batch mode, no recursive folder scanning, no "skip if newer" optimization.

**Consequences:**
- ✓ **Simplicity:** Code is shorter, fewer edge cases.
- ✓ **Clarity of intent:** Each run is explicit; no hidden conversions.
- ✓ **Fault isolation:** If one document fails, it does not affect others.
- ✗ **Not suitable for bulk workflows:** If a user later needs to convert 100 documents, this tool is not the solution.

**What if batch is needed later?**
- Call the `convert()` function in a loop from a separate Python script.
- Or use this tool as a library: `from docx_to_pdf import convert; for doc in docs: convert(doc, out_dir)`.

**This is intentional.** Do not add batch mode without explicit request.

---

## ADR-004: Keep Conversion Logic Separate and Importable

**Context:**
The code needs to be both:
1. A standalone GUI application (launched from the Desktop).
2. Reusable as a library (importable and callable from other Python code for testing, batch loops, integrations).

**Decision:**
- Define the core logic as a standalone `convert(src_path, out_dir)` function.
- Wrap it in a tkinter GUI class only when run as `__main__`.
- This allows both use cases without code duplication.

**Consequences:**
- ✓ **Testability:** The conversion logic can be tested without a GUI or user interaction.
- ✓ **Reusability:** Other Python projects can import and call `convert()` directly.
- ✓ **Separation of concerns:** Business logic is independent of presentation (GUI).
- ✓ **Code clarity:** The conversion algorithm is easy to understand; the GUI is a thin wrapper.
- ✗ **Slight overhead:** Must handle both modes (importable + GUI).

**Example use as a library:**
```python
from docx_to_pdf import convert

pdf_path = convert(r"C:\Documents\sermon.docx", r"C:\Output")
print(f"PDF created: {pdf_path}")
```

**Example use as a GUI:**
```bash
python docx_to_pdf.py
```

---

## ADR-005: COM Cleanup in Finally Block (Never Orphan Word)

**Context:**
Microsoft Word.Application is a heavy resource. Each `DispatchEx` call spawns a new WINWORD.EXE process. If a conversion fails and Word is not quit, the process remains running indefinitely, wasting memory and potentially blocking subsequent conversions.

**Decision:**
Always wrap conversion in try/finally. The finally block must:
1. Close the document (without saving).
2. Quit the Word process.
3. Uninitialize COM.
4. Swallow exceptions in finally (do not let cleanup failure mask the original error).

**Consequences:**
- ✓ **Resource safety:** No orphaned Word processes after a failed conversion.
- ✓ **Repeatability:** Subsequent conversions work correctly even after an error.
- ✓ **Fault tolerance:** If cleanup fails, the error is logged silently, and the main error is returned to the user.
- ✗ **Debugging difficulty:** If cleanup does fail, the user may not know (silent catch).

**Debugging an orphaned process:**
```bash
tasklist /FI "IMAGENAME eq WINWORD.EXE"
```
If this shows WINWORD.EXE, something went wrong in a previous run and cleanup failed. Kill it manually:
```bash
taskkill /IM WINWORD.EXE /F
```
Then file a bug report.

---

## Summary Table

| Decision | Choice | Key Benefit | Key Trade-off |
|----------|--------|-------------|---------------|
| **Conversion Engine** | Word COM | Pixel-perfect fidelity | Windows + Word required |
| **GUI Framework** | tkinter | No extra dependencies | Basic aesthetics |
| **Scope** | Single file | Simplicity | Not suitable for batch |
| **Code Architecture** | Importable function + GUI wrapper | Testable + reusable | Slight complexity |
| **Resource Cleanup** | try/finally | No orphaned processes | Silent cleanup failures |
