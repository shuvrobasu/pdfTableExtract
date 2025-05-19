# pdfTableExtract
# PDF Table Selector and Exporter

A professional Tkinter GUI utility to:
- Preview all PDF files in a folder (with **sharp, scrollable preview**)
- **Filter and jump to files instantly** by typing in the filter box
- **Auto-detect tables** in any PDF, on any page, and select/export them to Excel
- Export one, multiple, or all tables per PDF (each to a separate sheet)
- **Select tables visually** (with interactive rectangles/handles)
- No command line needed

## Features

- **Folder PDF browser**: Open any folder to see all PDFs.
- **Live filename filter**: As you type in the filter box, only matching PDFs are listed.
- **PDF preview with scrollbars**: Shows each page in true-to-original sharpness, including large pages with scroll.
- **Table detection**: Tables are detected on every page and outlined with colored rectangles and handles.
- **Interactive selection**: Click rectangles or handles to select or deselect tables.
- **Multi-selection**: Select one, several, or all tables on a page.
- **Export**: Export selected or all tables to Excel, with each table as a separate sheet.
- **Navigation**: Quickly go to any page or PDF, and use the exit button or window close to quit.
- **Handles**: Large, easy-to-click handles for table selection.

## How to Use

1. **Install dependencies:**
    ```sh
    pip install pdfplumber Pillow openpyxl
    ```

2. **Run the script:**
    ```sh
    python your_script_name.py
    ```

3. **Open a folder:**
   - Click **Select PDF Folder** and pick the directory containing your PDFs.

4. **Filter files:**
   - Type part of a filename (case-insensitive) in the filter box above the file list. Only matching files are shown.

5. **Preview and select:**
   - Click a filename to preview it.
   - Use **Prev/Next** to move between pages.
   - Detected tables are outlined. Click inside rectangles or on their handles to select/deselect.
   - Use **Select All Tables** to select all on the current page.

6. **Export:**
   - Click **Export Selected Tables** to export selected tables to Excel (`.xlsx`), with one sheet per table.
   - Click **Export All Tables** to export all tables in the PDF (each table on a separate sheet).

7. **Exit:**
   - Click the **Exit** button (bottom right) or close the window.

## Requirements

- Python 3.8 or newer (Tkinter included)
- `pdfplumber`
- `Pillow`
- `openpyxl`

Install all requirements with:
```sh
pip install pdfplumber Pillow openpyxl

