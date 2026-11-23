"""Additional offline conversion routes for AnyConvert."""
import csv
import os
import shutil
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

import pymupdf
from PIL import Image, ImageOps

from anyconvert_formats import IMAGES


def office_export(source, target, folder):
    app_dir = Path(sys.executable if getattr(sys, "frozen", False) else __file__).parent
    packaged = app_dir / "OfficeEngine/program/soffice.exe"
    executable = str(packaged) if packaged.is_file() else shutil.which("soffice") or next(
        (str(path) for key in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA")
         if os.getenv(key)
         for path in [Path(os.environ[key]) / "LibreOffice/program/soffice.exe"]
         if path.is_file()), None)
    if not executable:
        raise ValueError("Install LibreOffice for Word, PowerPoint and legacy Office rendering.")
    profile = (folder / "profile").resolve().as_uri()
    command = [executable, "-env:UserInstallation=" + profile, "--headless",
               "--convert-to", target, "--outdir", str(folder), str(source)]
    result = subprocess.run(command, capture_output=True, timeout=600,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    output = folder / f"{source.stem}.{target}"
    if result.returncode or not output.is_file():
        raise ValueError("LibreOffice could not export this file. Check whether it is password protected.")
    return output


def image_frames(source):
    with Image.open(source) as image:
        for index in range(getattr(image, "n_frames", 1)):
            image.seek(index)
            yield ImageOps.exif_transpose(image).convert("RGBA")


def save_image(image, path):
    if path.suffix in (".jpg", ".bmp"):
        flat = Image.new("RGB", image.size, "white")
        flat.paste(image, mask=image.getchannel("A"))
        flat.save(path, quality=95)
    else:
        image.save(path, lossless=True)


def fit_a4(pdf, path):
    with pymupdf.open() as output:
        for index, page in enumerate(pdf):
            width, height = (210, 297) if page.rect.width <= page.rect.height else (297, 210)
            dest = output.new_page(width=width * 72 / 25.4, height=height * 72 / 25.4)
            if page.get_contents():
                dest.show_pdf_page(dest.rect, pdf, index)
        output.save(path, deflate=True)


def page_slides(pdf, path, dpi):
    from io import BytesIO
    from pptx import Presentation
    from pptx.util import Inches
    deck = Presentation()
    first = pdf[0].rect
    deck.slide_width = Inches(10)
    deck.slide_height = Inches(max(1, min(56, 10 * first.height / first.width)))
    for number, page in enumerate(pdf, 1):
        print(f"Rendering page {number}/{len(pdf)}...", flush=True)
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        scale = min(deck.slide_width / page.rect.width, deck.slide_height / page.rect.height)
        width, height = int(page.rect.width * scale), int(page.rect.height * scale)
        picture = page.get_pixmap(dpi=dpi).tobytes("png")
        slide.shapes.add_picture(BytesIO(picture), (deck.slide_width-width)//2,
                                 (deck.slide_height-height)//2, width, height)
    deck.save(path)


def convert_extra(source, target, dpi):
    from converter import available_path, convert_pdf, convert_excel, image_to_pdf
    from openpyxl import Workbook, load_workbook
    ext = source.suffix.lower()
    with TemporaryDirectory(prefix="anyconvert-", dir=source.parent) as temp:
        folder = Path(temp)
        work = folder / "work"
        work.mkdir()
        out = folder / f"{source.stem}.{target}"
        outputs = []
        if ext == ".csv" and target == "xlsx":
            book = Workbook()
            with source.open(encoding="utf-8-sig", newline="") as stream:
                for r, row in enumerate(csv.reader(stream), 1):
                    for c, value in enumerate(row, 1):
                        cell = book.active.cell(r, c, value)
                        cell.data_type = "s"
            book.save(out)
        elif ext == ".xlsx" and target == "csv":
            book = load_workbook(source, read_only=True, data_only=True)
            try:
                for number, sheet in enumerate(book.worksheets, 1):
                    path = folder / f"{source.stem}-sheet-{number:03d}.csv"
                    with path.open("w", encoding="utf-8-sig", newline="") as stream:
                        csv.writer(stream).writerows(sheet.values)
                    outputs.append(path)
            finally:
                book.close()
        elif ext in IMAGES and target in IMAGES.values():
            for index, frame in enumerate(image_frames(source), 1):
                path = out if index == 1 else folder / f"{source.stem}-frame-{index:03d}.{target}"
                save_image(frame, path)
                outputs.append(path)
        elif ext in (".ppt", ".odp") and target == "pptx" or ext in (".doc", ".odt", ".rtf") and target == "docx" or ext in (".xls", ".ods") and target == "xlsx":
            out = office_export(source, target, work)
        else:
            if ext == ".pdf":
                intermediate = source
            elif ext == ".xlsx":
                local = work / source.name
                shutil.copyfile(source, local)
                convert_excel(local, "pdf")
                intermediate = local.with_suffix(".pdf")
            elif ext in IMAGES:
                intermediate = work / f"{source.stem}.pdf"
                with pymupdf.open() as merged:
                    for index, frame in enumerate(image_frames(source)):
                        part = work / f"frame-{index}.png"
                        frame.save(part)
                        image_to_pdf(part)
                        with pymupdf.open(part.with_suffix(".pdf")) as page:
                            merged.insert_pdf(page)
                    merged.save(intermediate)
            else:
                intermediate = office_export(source, "pdf", work)
            if target in ("xlsx", "docx"):
                local = folder / f"{source.stem}.pdf"
                shutil.copyfile(intermediate, local)
                convert_pdf(local, target, dpi)
            else:
                with pymupdf.open(intermediate) as pdf:
                    if pdf.needs_pass:
                        raise ValueError("Unlock the PDF before using this conversion.")
                    if target == "pdf":
                        fit_a4(pdf, out)
                    elif target == "pptx":
                        page_slides(pdf, out, dpi)
                    elif target == "txt":
                        out.write_text("\n\f\n".join(p.get_text(sort=True) for p in pdf), encoding="utf-8")
                    else:
                        for index, page in enumerate(pdf, 1):
                            print(f"Rendering page {index}/{len(pdf)}...", flush=True)
                            pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csRGB, alpha=False)
                            image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                            path = folder / f"{source.stem}-page-{index:03d}.{target}"
                            save_image(image.convert("RGBA"), path)
                            outputs.append(path)
        for path in outputs or [out]:
            destination = available_path(source.parent / path.name)
            path.rename(destination)
            print(f"Saved {destination.name}", flush=True)
