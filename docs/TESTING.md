# Testing

## Current State

**There is no automated test suite.** This is a known gap. Testing is currently manual and verification-based.

## Manual Verification Procedure

The code has been tested manually with the following steps:

### 1. Import and Call `convert()` Directly

Create a test script (e.g., `test_convert.py` in the repo root):

```python
import os
from docx_to_pdf import convert

# Use a real .docx file for testing
src_file = r"C:\Path\To\Test\Document.docx"
output_dir = r"C:\Temp\_test_out"  # Scratch directory

# Create the scratch directory if it doesn't exist
os.makedirs(output_dir, exist_ok=True)

try:
    pdf_path = convert(src_file, output_dir)
    print(f"✓ PDF created: {pdf_path}")
    
    # Verify the file exists
    if os.path.isfile(pdf_path):
        print(f"✓ File exists: {os.path.getsize(pdf_path)} bytes")
    
    # Check the PDF signature
    with open(pdf_path, "rb") as f:
        header = f.read(4)
        if header == b"%PDF":
            print("✓ PDF signature valid")
        else:
            print("✗ Not a valid PDF (wrong signature)")
    
except FileNotFoundError as e:
    print(f"✗ FileNotFoundError: {e}")
except PermissionError as e:
    print(f"✗ PermissionError: {e}")
except RuntimeError as e:
    print(f"✗ RuntimeError: {e}")
except Exception as e:
    print(f"✗ Unexpected error: {e}")
```

Run it:
```bash
cd C:\path\to\docx-to-pdf
python test_convert.py
```

**Expected output:**
```
✓ PDF created: C:\Temp\_test_out\Document.pdf
✓ File exists: 45678 bytes
✓ PDF signature valid
```

### 2. Verify No Orphaned Word Processes

After a successful conversion, check that no WINWORD.EXE process is left running:

```bash
tasklist /FI "IMAGENAME eq WINWORD.EXE"
```

**Expected output:**
```
INFO: No tasks are running which match the specified criteria.
```

If WINWORD.EXE is listed, cleanup failed and a Word process is orphaned.

### 3. Test Error Paths

#### Test: Source File Missing

```python
from docx_to_pdf import convert

try:
    pdf_path = convert(r"C:\Nonexistent\file.docx", r"C:\Temp")
except FileNotFoundError as e:
    print(f"✓ Caught FileNotFoundError: {e}")
except Exception as e:
    print(f"✗ Wrong exception type: {type(e).__name__}: {e}")
```

**Expected output:**
```
✓ Caught FileNotFoundError: Source file not found: C:\Nonexistent\file.docx
```

#### Test: Output Directory Not Writable

```python
from docx_to_pdf import convert

try:
    # Try to write to a system directory with no permissions
    pdf_path = convert(r"C:\Some\Document.docx", r"C:\Windows\System32")
except PermissionError as e:
    print(f"✓ Caught PermissionError: {e}")
except Exception as e:
    print(f"✗ Wrong exception type: {type(e).__name__}: {e}")
```

**Expected output:**
```
✓ Caught PermissionError: Output directory is not writable: C:\Windows\System32
```

#### Test: PDF Locked by Another Program

```python
from docx_to_pdf import convert
import subprocess
import os

src_file = r"C:\Path\To\Document.docx"
output_dir = r"C:\Temp\_test_out"

# First conversion (succeeds)
pdf_path = convert(src_file, output_dir)
print(f"✓ First conversion succeeded: {pdf_path}")

# Open the PDF in Notepad (or any program)
subprocess.Popen(["notepad", pdf_path])

# Wait a moment
import time
time.sleep(1)

# Try to convert again (should fail because PDF is open)
try:
    pdf_path2 = convert(src_file, output_dir)
    print("✗ Second conversion should have failed")
except PermissionError as e:
    print(f"✓ Caught PermissionError: {e}")
except Exception as e:
    print(f"✗ Wrong exception type: {type(e).__name__}: {e}")
finally:
    # Close Notepad
    os.system("taskkill /IM notepad.exe /F")
```

**Expected output:**
```
✓ First conversion succeeded: C:\Temp\_test_out\Document.pdf
✓ Caught PermissionError: Could not write the PDF. Close it if it's open in another program and try again.
```

### 4. Test GUI (Source Code)

Launch the GUI and verify basic functionality:

```bash
python docx_to_pdf.py
```

**User interaction:**
1. Click "Browse" next to "Word file:" and select a .docx file.
   - Verify the path appears in the field.
   - Verify the output directory defaults to the same folder.
   - Verify the "Convert" button becomes enabled.

2. (Optional) Click "Browse" next to "Save to:" and select a different output directory.

3. Click "Convert".
   - Verify status changes to "Converting..." in black text.
   - Verify the buttons (Browse, Convert) are disabled during conversion.
   - Wait for the PDF to be created (this may take several seconds).
   - Verify status changes to "Created *filename*.pdf" in green text.
   - Verify the buttons are re-enabled.

4. Open the output PDF in your default PDF viewer and verify formatting matches the original Word document.

5. Close the GUI window.

### 4b. Testing a Frozen Build (dist\DocxToPdf.exe)

After running `build.bat`, the frozen executable must be tested end-to-end to verify that `pywin32` and COM automation work correctly when bundled.

**Verification Checklist:**

#### 1. GUI Launch
```bash
dist\DocxToPdf.exe
```

- Verify the GUI window appears.
- Verify **no console window** is visible (this is the purpose of `--windowed`).
- Close the window.

#### 2. End-to-End Conversion (Frozen Binary)

Create a test script *outside* the project directory (to simulate an independent user):

```python
# test_frozen_build.py (in C:\temp or similar)

import sys
sys.path.insert(0, r"C:\path\to\docx-to-pdf")
from docx_to_pdf import convert
import os

src_file = r"C:\path\to\test\document.docx"
output_dir = r"C:\temp\_frozen_test_out"

os.makedirs(output_dir, exist_ok=True)

try:
    pdf_path = convert(src_file, output_dir)
    print(f"✓ PDF created: {pdf_path}")
    
    # Verify file size and signature
    file_size = os.path.getsize(pdf_path)
    print(f"✓ File size: {file_size} bytes")
    
    with open(pdf_path, "rb") as f:
        header = f.read(4)
        if header == b"%PDF":
            print("✓ PDF signature valid")
        else:
            print("✗ Not a valid PDF")
    
    # Check for orphaned processes
    import subprocess
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE"],
        capture_output=True, text=True
    )
    if "WINWORD.EXE" not in result.stdout:
        print("✓ No orphaned Word processes")
    else:
        print("✗ Orphaned Word process found")
        
except Exception as e:
    print(f"✗ Test failed: {type(e).__name__}: {e}")
```

Run it:
```bash
python C:\temp\test_frozen_build.py
```

**Expected output:**
```
✓ PDF created: C:\temp\_frozen_test_out\document.pdf
✓ File size: 196542 bytes
✓ PDF signature valid
✓ No orphaned Word processes
```

**Why this is essential:**
- Bundling `pywin32` in PyInstaller does not guarantee COM works when frozen. The temporary unpacking environment must be exercised with a real conversion.
- A failed end-to-end test (e.g., "RuntimeError: Microsoft Word is not installed") indicates the frozen binary is broken and must not be distributed.
- Do not rely on the GUI launching or the executable existing as proof that the build works. The conversion logic must actually run and succeed.

#### 3. Verify File Integrity

- PDF file size should be reasonable (typically >100 KB for a sermon or article; >200 KB for multi-page documents with images).
- Opening the PDF in Adobe Reader or a web browser should display correctly formatted text, images, and layout matching the original Word document.

### 5. Test Complex Document

Optionally, test with a complex Word document to verify formatting fidelity:

- Multi-page document
- Embedded images
- Tables with merged cells
- Headers and footers
- Multiple styles (bold, italic, colors)
- Page breaks
- Footnotes/endnotes

**Expected result:** The PDF output is visually identical to the Word document when opened in Word or previewed on-screen.

## Caveats

### Scratch Output Directory

**Always** test with a scratch output directory (e.g., `C:\Temp\_test_out`), never with a production directory of sermon PDFs. A test run that fails could corrupt or overwrite real files.

### Test Files

Use test .docx files that you don't mind being converted multiple times or overwritten. Do not use the only copy of an important document.

### Environment Dependencies

- **Windows** (Linux/macOS will not work; Windows-specific COM and file paths)
- **Microsoft Word** installed (Home, Pro, or Microsoft 365; not OpenOffice)
- **Python 3.7+** (tested with Python 3.12.6)
- **pywin32** installed (`pip install pywin32`)

Verify your environment before testing:

```bash
python --version
pip show pywin32
# Should show pywin32 installed
```

## Future: Automated Testing

A proper test suite would:
- Create temporary .docx files programmatically (using python-docx or a fixture).
- Call `convert()` in an isolated test process.
- Verify output PDFs without opening them (parse PDF metadata, check file size, signature).
- Run in CI/CD (e.g., GitHub Actions, but only on Windows runners).
- Test all three exception types (FileNotFoundError, PermissionError, RuntimeError).

This is a known gap but not urgent for a single-file utility used by one person.

## Debugging

If something goes wrong:

1. **Check for orphaned processes:**
   ```bash
   tasklist /FI "IMAGENAME eq WINWORD.EXE"
   ```
   Kill any orphaned WINWORD.EXE:
   ```bash
   taskkill /IM WINWORD.EXE /F
   ```

2. **Enable Python exception tracebacks:**
   Remove try/except blocks temporarily to see the full error. (For development only; restore them before committing.)

3. **Test with a simple document:**
   Use a minimal .docx file (single page, no images, basic text) to isolate formatting issues.

4. **Restart Word:**
   If COM behaves strangely, restart Word and reboot if necessary.
