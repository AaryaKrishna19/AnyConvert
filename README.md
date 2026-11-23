# AnyConvert

Offline Windows 11 x64 file converter with a File Explorer submenu and a compact progress window.

## Build

Install Python x64, Microsoft C++ Build Tools (Desktop development with C++ and Windows 11 SDK), and Inno Setup. From PowerShell in the extracted repository:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build-setup.ps1
```

Run dist/AnyConvertSetup.exe and right-click a supported file. The build needs internet and WinGet to stage the x64 LibreOffice engine and Microsoft VC++ runtime. Setup embeds both along with the frozen Python runtime and libraries. Allow several GB for staging. Build tools are only needed on the build PC. Test the finished installer offline on a clean x64 Windows 11 PC before publishing a release.

## Everyday conversions

PDF offers Word, PNG, JPEG and Excel tables. Word offers A4 PDF. PowerPoint offers PDF and slide PNGs. Excel offers PDF and CSV (one file per sheet). CSV opens as an Excel workbook. Images offer A4 PDF and common image formats. Legacy Office and OpenDocument files also offer migration to their modern Office equivalent. Table extraction requires detectable tables; scanned PDFs need an OCR workflow for editable text.

## Publish the download website

Commit the extracted files at the repository root. In Settings → Pages, choose Deploy from a branch, main, /(root). Use the standard Jekyll Pages build with the included _config.yml and index.html. Upload dist/AnyConvertSetup.exe as a release asset named AnyConvertSetup.exe and publish it as the latest release. The download and release notes links use this repository's GitHub metadata, including on a custom domain.

## Source repository

Upload the extracted source files to GitHub. Keep generated EXEs and installers in Releases. The build uses a local development signing certificate for Explorer registration. Keep its private key on the build machine.
