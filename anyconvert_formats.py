"""Formats offered by Explorer and by the desktop picker, in display order."""
IMAGES = {
    ".png": "png", ".jpg": "jpg", ".jpeg": "jpg", ".webp": "webp",
    ".bmp": "bmp", ".gif": "gif", ".tif": "tiff", ".tiff": "tiff",
}

FORMATS = {
    ".pdf": ("docx", "png", "jpg", "xlsx"),
    ".docx": ("pdf",),
    ".doc": ("pdf", "docx"),
    ".odt": ("pdf", "docx"),
    ".rtf": ("pdf", "docx"),
    ".pptx": ("pdf", "png"),
    ".ppt": ("pdf", "pptx", "png"),
    ".odp": ("pdf", "pptx"),
    ".xlsx": ("pdf", "csv"),
    ".xls": ("pdf", "xlsx"),
    ".ods": ("pdf", "xlsx"),
    ".csv": ("xlsx",),
}
for extension, image_format in IMAGES.items():
    FORMATS[extension] = ("pdf",) + tuple(
        target for target in ("jpg", "png") if target != image_format
    )
