# AnyConvert

**AnyConvert** is an offline file converter for Windows 11 (x64) that integrates directly into the Windows File Explorer context menu and provides a compact progress window for quick, local document and image conversions.

---

## Key Features

* **100% Offline Processing:** All conversions run locally with no internet connection or cloud uploads required.
* **File Explorer Integration:** Convert supported files directly via a dedicated right-click submenu.
* **Lightweight UI:** Displays a compact progress window during active conversions.
* **Broad Format Support:** Handles PDFs, modern and legacy Office documents, OpenDocument formats, CSVs, and common raster images.

---

## Installation

1. Download the latest **`AnyConvertSetup.exe`** installer from the [**Releases**](../../releases) page.
2. Run `AnyConvertSetup.exe` to install AnyConvert and register the File Explorer context menu.
3. Right-click any supported file in File Explorer and select your target format from the **AnyConvert** submenu.

> **Optional Dependency (LibreOffice):** To convert Word, PowerPoint, OpenDocument, or legacy Office files, install [LibreOffice](https://www.libreoffice.org/). AnyConvert automatically detects `soffice.exe` on the system `PATH` or in the default LibreOffice installation directory. Python and C++ build tools are **not** required to run the installed application.

---

## Supported Conversions

| Input Formats | Available Output Formats |
| :--- | :--- |
| **PDF** | DOCX (layout-preserved or fast text), XLSX (tables), PPTX (visual slides), TXT, PNG, JPEG, BMP, TIFF, WebP |
| **DOCX, DOC, ODT, RTF, PPTX, PPT, ODP, XLSX, XLS, ODS** | A4 PDF, DOCX, PPTX, XLSX, TXT, PNG, JPEG, BMP, TIFF, WebP *(excludes same-format output)* |
| **PNG, JPG/JPEG, BMP, TIFF/TIF, WebP, GIF** | A4 PDF, DOCX, PPTX, PNG, JPEG, BMP, TIFF, WebP *(excludes same-format output)* |
| **CSV** | XLSX *(UTF-8 comma-separated text)* |
| **XLSX** | CSV *(exported as one file per worksheet, in addition to the Office outputs above)* |

### Conversion Behavior & Limitations

* **Output Location:** Converted files are written to the same directory as the source file. Multi-page image exports generate one image file per page.
* **Presentations:** Cross-format conversions to PowerPoint produce slides containing rendered page images.
* **Tables & Spreadsheets:** PDF-to-Excel conversions (and conversions routed through PDF) require detectable table structures. Native XLSX-to-PDF and XLSX-to-Word conversions rely on a cell-value exporter, so charts and complex formulas may have limited fidelity.
* **Text & Images:** TXT exports extract plain document text. Image-to-Word conversions embed the visual content directly and do not perform OCR.
* **Legacy Office Formats:** Upgrades from legacy Word, Excel, and PowerPoint files are processed directly through LibreOffice.

---

## Building from Source

### Prerequisites

Install the following tools on your build machine:

* **Python (x64)**
* **Microsoft C++ Build Tools** *(include the **Desktop development with C++** workload and the **Windows 11 SDK**)*
* **Inno Setup**

### Build Steps

Open PowerShell in the root directory of the repository and run:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build-setup.ps1
