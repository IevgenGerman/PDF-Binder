# PDF Binder

Combines PDF, Word, Excel, PowerPoint, images and text files into one PDF.

## Install and run (Windows)

1. Install Python 3.10+ from python.org (tick "Add python.exe to PATH").
2. In this folder: `pip install -r requirements.txt`
3. Start: `python pdf_binder.py`

To get a single `PDF_Binder.exe` that runs without Python, double-click `build_exe.bat`.
The exe is in the `dist` folder.

## Converting Office files

- **With MS Office:** Word, Excel and PowerPoint are converted through Office.
- **Without MS Office:** install LibreOffice (free). The program finds it automatically.
- **No extra software needed:** PDF, images (jpg, png, tif, ...) and txt.

## How to use

- **Add documents:** use **Add files...** or drag files or folders into the window.
- **Order:** select a row and use **Up**, **Down** or **Remove**.
- **In contents:** ticked means the **Title** goes into the table of contents.
  - The title defaults to the file name. Double-click it to edit.
  - **Level:** 1 = section, 2 = subsection, 3 = sub-subsection.
- **Preview:** the right side shows the pages of the selected document. Double-click a page for a large view.
- **Rotate:**
  - **Document -90 / +90** rotates every page of the document.
  - **Selected pages -90 / +90** rotates only the pages you selected (click, Ctrl+click or Shift+click).
  - You can rotate as many times as you need.
- **Save PDF...** asks about:
  - the table of contents (its entries are clickable, and also become PDF bookmarks);
  - page numbers: position, format, and whether the contents pages are counted;
  - header and footer. You can use these placeholders in them:
    - `{page}`, `{pages}`
    - `{title}` (title of the current document)
    - `{date}`
    - `{file}` (name of the output file)

Rotations are written into the saved PDF itself, so it opens rotated in any viewer.
Headers, footers and numbers are placed upright on rotated pages as well.
