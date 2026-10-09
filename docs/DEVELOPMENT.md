# Development

## Running from source

```bash
pip install pywin32
python docx_to_pdf.py
```

## Building the .exe

```bash
pip install --user pyinstaller pillow
build.bat
```

This creates `dist\DocxToPdf.exe` (about 15 MB), which runs on its own without Python. If `pip` gives a permissions error, `--user` is what fixes it.

## Using it as a library

`convert()` can be called from other Python code:

```python
from docx_to_pdf import convert

pdf_path = convert(
    src_path=r"C:\Documents\MyLetter.docx",
    out_dir=r"C:\Output",
)
```

```python
def convert(src_path, out_dir) -> str
```

- `src_path`: absolute path to the .docx or .doc file
- `out_dir`: absolute path to the output folder
- Returns the absolute path of the PDF.
- Raises `FileNotFoundError` if the document or folder doesn't exist, `PermissionError` if the folder isn't writable or the PDF is open somewhere, and `RuntimeError` if Word isn't installed or the export fails.

It works on any thread, since it sets up COM for the thread it runs on. For several files, call it in a loop.

## Tests

See [TESTING.md](TESTING.md).
