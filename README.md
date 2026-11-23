<div align="center">

<img src="AnyConvert150.png" alt="AnyConvert logo" width="110">

# AnyConvert

**Right-click. Convert. Done.**

Offline file conversion in the Windows 11 right-click menu.<br>
PDF to Word, Word to PDF, slides to images and more, all on your own PC.

[![Latest release](https://img.shields.io/github/v/release/AaryaKrishna19/AnyConvert?style=for-the-badge&color=4a63ff)](https://github.com/AaryaKrishna19/AnyConvert/releases/latest)
![Windows 11 x64](https://img.shields.io/badge/Windows-11%20x64-0078D4?style=for-the-badge)
![Offline](https://img.shields.io/badge/Works-offline-17b26a?style=for-the-badge)

[**Download for Windows**](https://github.com/AaryaKrishna19/AnyConvert/releases/latest) · [Website](https://aaryakrishna19.github.io/AnyConvert/) · [Report an issue](https://github.com/AaryaKrishna19/AnyConvert/issues)

</div>

---

## What is AnyConvert?

AnyConvert adds a **Convert** option to the main Windows 11 right-click menu. Select a file, choose the format you want, and the converted copy is saved right next to the original. There is no app to open and no website to visit.

It was built for people who regularly need things like PDF to Word or PDF to image, and who don't want to upload private documents to an online converter.

## Download

1. Go to the [latest release](https://github.com/AaryaKrishna19/AnyConvert/releases/latest).
2. Download **AnyConvertSetup.exe**.
3. Run it and follow the installer.

The installer bundles the conversion engines and the runtime they need, so there is nothing else to install.

**Requirements:** Windows 11, 64-bit (x64).

## How to use it

1. Right-click a supported file in File Explorer.
2. Hover **AnyConvert** in the main context menu. You don't need to open "Show more options".
3. Choose a format, for example **Convert to Word**.
4. A small progress window appears in the centre of the screen and closes when the job is done.
5. Find the converted file beside the original.

## Supported conversions

| From | Convert to | Notes |
| --- | --- | --- |
| PDF | Word (.docx) | Rebuilds the page in Word format |
| PDF | Excel (.xlsx) | Detectable tables only |
| PDF | PNG or JPEG | Image output |
| Word (.docx) | PDF | A4 page size |
| PowerPoint (.pptx) | PDF | A4 page size |
| PowerPoint (.pptx) | PNG | One image per slide |
| Excel (.xlsx) | PDF | A4 page size |
| Excel (.xlsx) | CSV | One file per sheet |
| CSV | Excel (.xlsx) | |
| JPG or PNG | PDF | A4 page size |
| JPG or PNG | The other image format | JPG to PNG, or PNG to JPG |

## Privacy

AnyConvert works entirely on your computer. Files are read from your disk, converted locally and saved next to the original. Nothing is uploaded.

## Good to know

- **Layouts can shift.** Converting a PDF to Word rebuilds the page, so complex layouts may need a quick touch-up. AnyConvert does not promise a pixel-perfect copy.
- **Best with text-based PDFs.** Scanned pages are not run through text recognition (OCR), so they stay as images.
- **PDF to Excel finds tables, not magic.** It picks up the tables it can detect. Pages without clear tables give little back.
- **Not every format is supported.** AnyConvert covers the conversions listed above.

## Feedback

Found a bug or have a file that converts badly? [Open an issue](https://github.com/AaryaKrishna19/AnyConvert/issues) and describe the file type and what went wrong. Please don't attach private documents.

---

<div align="center">

Made by **Aarya Softwares**

</div>
