# DWG-AutoFill

**DWG-AutoFill** is a desktop application for Windows, built on Python, designed to automate the process of filling data into DWG/DXF templates from external data sources (Excel, CSV, JSON).

## Key Features

*   **Offline Operation:** Works without an internet connection or server.
*   **Batch Generation:** Generates multiple DWG files from a single data source.
*   **Interactive 2D Viewer:** Uses `matplotlib` for vector-based preview with zoom and pan.
*   **Highlighting:** Overlays to highlight fields that have been automatically filled.
*   **AutoCAD PRO Mode (Optional):** If AutoCAD is installed, it uses the COM API for 100% WYSIWYG PDF export and preview.

## Technology Stack

*   **Language:** Python 3.11
*   **DWG/DXF Handling:** `ezdxf`
*   **Data Handling:** `pandas`, `openpyxl`, `rapidfuzz`
*   **GUI & Viewer:** `PySimpleGUI` or `PyQt` with `matplotlib`
*   **Packaging:** `PyInstaller`
*   **Windows Integration:** `pywin32` (for AutoCAD COM)
