import argparse
import sys
from getpass import getpass
from io import BytesIO
from pathlib import Path

import pymupdf


from anyconvert_formats import FORMATS
from extra_formats import convert_extra


def report(message):
    print(message, flush=True)


def available_path(path):
    output = path
    number = 2

    while output.exists():
        output = path.with_name(f"{path.stem} ({number}){path.suffix}")
        number += 1

    return output


def pdf_to_images(pdf, source, target, dpi):
    for number, page in enumerate(pdf, start=1):
        report(f"Rendering page {number}/{len(pdf)}...")
        output = available_path(
            source.with_name(f"{source.stem}-page-{number:03d}.{target}")
        )
        pixels = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csRGB, alpha=False)

        if target == "jpg":
            pixels.save(str(output), jpg_quality=95)
        else:
            pixels.save(str(output))

        report(f"Saved {output.name}")


def pdf_to_word_fast(pdf, source):
    from docx import Document

    document = Document()
    pages_written = 0

    for number, page in enumerate(pdf, start=1):
        report(f"Extracting page {number}/{len(pdf)}...")
        lines = [
            line
            for line in page.get_text("text", sort=True).splitlines()
            if line.strip()
        ]

        if not lines:
            report(f"Page {number} has no extractable text. Skipping.")
            continue

        if pages_written:
            document.add_page_break()

        for line in lines:
            document.add_paragraph(line)

        pages_written += 1

    if not pages_written:
        raise ValueError("No extractable text found. Scanned pages require OCR.")

    output = available_path(source.with_name(f"{source.stem}-fast.docx"))
    document.save(output)
    report(f"Saved {output.name}")


def pdf_to_word(source, password):
    from pdf2docx import Converter

    output = available_path(source.with_suffix(".docx"))
    report(f"Reconstructing the layout of {source.name}...")

    converter = Converter(str(source), password=password)
    try:
        converter.convert(str(output))
    finally:
        converter.close()

    if not output.is_file():
        raise RuntimeError("The Word conversion engine produced no output.")

    report(f"Saved {output.name}")


def pdf_to_excel(pdf, source):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    workbook = Workbook()
    workbook.remove(workbook.active)

    line = Side(style="thin", color="FF505050")
    box = Border(left=line, right=line, top=line, bottom=line)
    alignment = Alignment(vertical="top", wrap_text=True)
    tables_written = 0

    for page_number, page in enumerate(pdf, start=1):
        report(f"Detecting tables on page {page_number}/{len(pdf)}...")
        finder = page.find_tables()

        if not finder.tables:
            finder = page.find_tables(strategy="text")

        for table in finder.tables:
            rows = table.extract()
            if not any(value and value.strip() for row in rows for value in row):
                continue

            tables_written += 1
            sheet = workbook.create_sheet(f"Page {page_number} Table {tables_written}")
            column_count = max(len(row) for row in rows)
            widths = [10] * column_count

            for row_number, values in enumerate(rows, start=1):
                for column in range(1, column_count + 1):
                    cell = sheet.cell(row=row_number, column=column)
                    cell.border = box
                    cell.alignment = alignment

                    value = values[column - 1] if column <= len(values) else None
                    if value is not None:
                        cell.value = value
                        # Keep source text literal, including values starting with "=".
                        cell.data_type = "s"
                        length = max(map(len, value.splitlines()), default=0)
                        widths[column - 1] = max(widths[column - 1], length + 2)

            for column, width in enumerate(widths, start=1):
                sheet.column_dimensions[get_column_letter(column)].width = min(50, width)

    if not tables_written:
        raise ValueError("No tables with extractable text found in this PDF.")

    output = available_path(source.with_suffix(".xlsx"))
    workbook.save(output)
    report(f"Saved {output.name} with {tables_written} bordered table(s)")


def convert_pdf(source, target, dpi):
    password = ""

    with pymupdf.open(str(source)) as pdf:
        if pdf.needs_pass:
            if not sys.stdin.isatty():
                raise ValueError("Run converter.py in PowerShell to enter the PDF password.")

            password = getpass("PDF password: ")
            if not pdf.authenticate(password):
                raise ValueError("Incorrect PDF password.")

        if target in ("png", "jpg"):
            pdf_to_images(pdf, source, target, dpi)
            return

        if target == "docx-fast":
            pdf_to_word_fast(pdf, source)
            return

        if target == "xlsx":
            pdf_to_excel(pdf, source)
            return

    pdf_to_word(source, password)


def image_to_pdf(source):
    from PIL import Image, ImageOps

    output = available_path(source.with_suffix(".pdf"))
    report(f"Fitting {source.name} onto A4...")

    with Image.open(source) as original:
        orientation = original.getexif().get(274, 1)

        if original.format == "JPEG" and orientation == 1:
            image_bytes = source.read_bytes()
            width, height = original.size
        else:
            image = ImageOps.exif_transpose(original).convert("RGBA")
            width, height = image.size
            buffer = BytesIO()
            image.save(buffer, format="PNG")
            image_bytes = buffer.getvalue()

    page_width = 210 * 72 / 25.4
    page_height = 297 * 72 / 25.4

    if width > height:
        page_width, page_height = page_height, page_width

    margin = 10 * 72 / 25.4
    scale = min(
        (page_width - 2 * margin) / width,
        (page_height - 2 * margin) / height,
    )
    fitted_width = width * scale
    fitted_height = height * scale
    left = (page_width - fitted_width) / 2
    top = (page_height - fitted_height) / 2

    with pymupdf.open() as pdf:
        page = pdf.new_page(width=page_width, height=page_height)
        page.insert_image(
            pymupdf.Rect(left, top, left + fitted_width, top + fitted_height),
            stream=image_bytes,
        )
        pdf.save(str(output), deflate=True)

    report(f"Saved {output.name}")


def convert_image_format(source, target):
    from PIL import Image, ImageOps

    output = available_path(source.with_suffix(f".{target}"))
    report(f"Converting {source.name} to {target.upper()}...")

    with Image.open(source) as original:
        image = ImageOps.exif_transpose(original)

        if target == "jpg":
            rgba = image.convert("RGBA")
            rgb = Image.new("RGB", rgba.size, "white")
            rgb.paste(rgba, mask=rgba.getchannel("A"))
            rgb.save(output, format="JPEG", quality=95, subsampling=0)
        else:
            if image.mode == "CMYK":
                image = image.convert("RGB")
            image.save(output, format="PNG")

    report(f"Saved {output.name}")


def worksheet_rows(sheet, include_borders=False):
    """Keep cell positions and trim unused space at the right and bottom."""
    rows = []
    columns = 0
    last_used_row = 0

    for number, values in enumerate(sheet.iter_rows(), start=1):
        row = list(values)

        while row:
            cell = row[-1]
            if cell.value is not None:
                break

            border = getattr(cell, "border", None)
            if include_borders and border is not None and any(
                getattr(border, edge, None) is not None and getattr(border, edge).style
                for edge in ("left", "right", "top", "bottom")
            ):
                break

            row.pop()

        if row:
            columns = max(columns, len(row))
            last_used_row = number

        rows.append(row)

    return rows[:last_used_row], columns


def cell_text(value):
    return "" if value is None else str(value)


def excel_to_word(workbook, source):
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.shared import Mm

    document = Document()
    section = document.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Mm(297)
    section.page_height = Mm(210)
    section.left_margin = Mm(12)
    section.right_margin = Mm(12)
    tables_written = 0

    for sheet in workbook.worksheets:
        report(f"Reading sheet: {sheet.title}")
        rows, columns = worksheet_rows(sheet)
        if not rows:
            continue

        if tables_written:
            document.add_page_break()

        document.add_heading(sheet.title, level=1)
        table = document.add_table(rows=0, cols=columns)
        table.style = "Table Grid"

        for number, values in enumerate(rows, start=1):
            for cell, source_cell in zip(table.add_row().cells, values):
                if source_cell.value is not None:
                    cell.text = cell_text(source_cell.value)

            if number % 100 == 0:
                report(f"Writing row {number}/{len(rows)}...")

        tables_written += 1

    if not tables_written:
        raise ValueError("The workbook contains no stored cell values.")

    output = available_path(source.with_suffix(".docx"))
    document.save(output)
    report(f"Saved {output.name}")


def excel_border_commands(rows):
    """Translate explicit Excel cell edges into ReportLab table lines."""
    from openpyxl.styles.colors import COLOR_INDEX
    from reportlab.lib import colors

    commands = []
    directions = {
        "left": "LINEBEFORE",
        "right": "LINEAFTER",
        "top": "LINEABOVE",
        "bottom": "LINEBELOW",
    }
    weights = {"hair": 0.5, "medium": 1.25, "thick": 2, "double": 0.65}

    for row_index, row in enumerate(rows):
        for column_index, cell in enumerate(row):
            border = getattr(cell, "border", None)
            if border is None:
                continue

            for edge, direction in directions.items():
                side = getattr(border, edge, None)
                if side is None or not side.style:
                    continue

                style = side.style
                color = colors.HexColor("#303030")
                if side.color is not None:
                    if side.color.type == "rgb":
                        color = colors.HexColor("#" + side.color.rgb[-6:])
                    elif side.color.type == "indexed":
                        index = side.color.indexed
                        if 0 <= index < len(COLOR_INDEX):
                            color = colors.HexColor("#" + COLOR_INDEX[index][-6:])

                width = weights.get(style, 1.25 if style.startswith("medium") else 0.75)
                pattern = None
                if style == "dotted":
                    pattern = (1, 2)
                elif "dashdotdot" in style.lower():
                    pattern = (4, 2, 1, 2, 1, 2)
                elif "dashdot" in style.lower():
                    pattern = (4, 2, 1, 2)
                elif "dash" in style.lower():
                    pattern = (4, 2)

                position = (column_index, row_index)
                count = 2 if style == "double" else 1
                commands.append(
                    (direction, position, position, width, color, 0, pattern, 0, count)
                )

    return commands


def excel_to_pdf(workbook, source):
    from xml.sax.saxutils import escape

    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import LongTable, PageBreak, Paragraph, SimpleDocTemplate

    output = available_path(source.with_suffix(".pdf"))
    page_size = landscape(A4)
    margin = 28
    usable_width = page_size[0] - 2 * margin - 12
    font_name = "Helvetica"

    arial = Path(r"C:\Windows\Fonts\arial.ttf")
    if arial.is_file():
        pdfmetrics.registerFont(TTFont("ProjectConvertArial", str(arial)))
        font_name = "ProjectConvertArial"

    heading_style = ParagraphStyle(
        "SheetTitle", fontName=font_name, fontSize=13, leading=16, spaceAfter=8
    )
    story = []

    for sheet in workbook.worksheets:
        report(f"Preparing A4 pages for sheet: {sheet.title}")
        rows, columns = worksheet_rows(sheet, include_borders=True)
        if not rows:
            continue

        if story:
            story.append(PageBreak())

        story.append(Paragraph(escape(sheet.title), heading_style))
        column_width = usable_width / columns
        font_size = min(9, max(5, column_width / 5))
        padding = min(4, column_width / 8)
        cell_style = ParagraphStyle(
            "Cell",
            fontName=font_name,
            fontSize=font_size,
            leading=font_size * 1.25,
            splitLongWords=True,
        )

        table_rows = []
        for number, cells in enumerate(rows, start=1):
            row = [
                Paragraph(escape(cell_text(cell.value)).replace("\n", "<br/>"), cell_style)
                if cell.value is not None else ""
                for cell in cells
            ]
            row.extend([""] * (columns - len(row)))
            table_rows.append(row)

            if number % 250 == 0:
                report(f"Preparing row {number}/{len(rows)}...")

        styling = [
            ("GRID", (0, 0), (-1, -1), 0.55, colors.HexColor("#666666")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), padding),
            ("RIGHTPADDING", (0, 0), (-1, -1), padding),
        ]
        styling.extend(excel_border_commands(rows))

        story.append(
            LongTable(
                table_rows,
                colWidths=[column_width] * columns,
                splitInRow=1,
                style=styling,
            )
        )

    if not story:
        raise ValueError("The workbook contains no stored cell values or bordered cells.")

    report("Writing bordered A4 PDF...")
    document = SimpleDocTemplate(
        str(output),
        pagesize=page_size,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=margin,
    )
    document.build(story)
    report(f"Saved {output.name}")


def convert_excel(source, target):
    from openpyxl import load_workbook

    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        if target == "pdf":
            excel_to_pdf(workbook, source)
        else:
            excel_to_word(workbook, source)
    finally:
        workbook.close()


def main():
    parser = argparse.ArgumentParser(description="Convert files locally.")
    parser.add_argument("source", type=Path)
    parser.add_argument(
        "--format",
        required=True,
        choices=sorted(set().union(*FORMATS.values())),
    )
    parser.add_argument("--dpi", type=int, default=200)
    parser.add_argument("--from-explorer", action="store_true")
    args = parser.parse_args()

    if not 72 <= args.dpi <= 600:
        parser.error("--dpi must be between 72 and 600")

    try:
        source = args.source.resolve()
        extension = source.suffix.lower()

        if not source.is_file():
            raise ValueError(f"File does not exist: {source}")

        if args.format not in FORMATS.get(extension, ()):
            raise ValueError(f"Unsupported conversion: {extension} to {args.format}")

        if extension == ".pdf" and args.format in ("png","jpg","docx","docx-fast","xlsx"):
            convert_pdf(source, args.format, args.dpi)
        elif extension == ".xlsx" and args.format in ("pdf","docx"):
            convert_excel(source, args.format)
        elif extension in (".jpg",".jpeg",".png") and args.format == "pdf":
            image_to_pdf(source)
        else:
            convert_extra(source, args.format, args.dpi)

        return 0

    except Exception as error:
        print(f"Conversion failed: {error}", file=sys.stderr, flush=True)
        if args.from_explorer and sys.stdin.isatty():
            input("Press Enter to close...")
        return 1


if __name__ == "__main__":
    sys.exit(main())
