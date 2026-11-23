"""AnyConvert conversion window.

A compact, DPI-aware progress window in the Windows 11 style: it follows the
system light/dark setting, matches the title bar to the window, and draws all
shapes with Pillow at 4x so edges stay smooth on every display.

Started by the Explorer menu as ``AnyConvert <format> <path>``, or by hand as
``AnyConvert [path]`` to pick the file and format in the window.
"""
import base64
import io
import logging
import math
import queue
import re
import sys
import threading
import time
import tkinter as tk
from contextlib import redirect_stderr, redirect_stdout
from functools import lru_cache
from pathlib import Path
from tkinter import filedialog
from tkinter import font as tkfont

from PIL import Image, ImageChops, ImageDraw, ImageFont

try:
    from PIL import ImageTk
except Exception:  # Falls back to PNG data if the Tk bridge is missing.
    ImageTk = None

try:
    from anyconvert_formats import FORMATS
except ModuleNotFoundError:
    from converter import FORMATS


APP_NAME = "AnyConvert"
WIDTH = 480          # logical pixels, scaled by the monitor DPI
PAD = 24
DETAILS_H = 150
SS = 4               # supersampling factor for Pillow drawing
LANCZOS = getattr(Image, "Resampling", Image).LANCZOS

LABELS = {
    "pptx": "PowerPoint presentation", "txt": "Plain text", "csv": "CSV",
    "webp": "WebP", "bmp": "Bitmap", "tiff": "TIFF",
    "pdf": "A4 PDF",
    "png": "PNG",
    "jpg": "JPEG",
    "docx": "Word document",
    "docx-fast": "Word document (fast text)",
    "xlsx": "Excel workbook",
}

THEMES = {
    "dark": {
        "bg": "#1C1F23", "panel": "#252A2F", "line": "#343A40",
        "text": "#F4F7F8", "muted": "#A0ACB2", "faint": "#6F7B82",
        "track": "#30363C", "accent": "#22B8CC", "accent2": "#5EE6C8",
        "ok": "#3DD68C", "err": "#FF6B6B", "on_accent": "#04262C",
        "btn": "#2D3338", "btn_hi": "#383F46", "btn_lo": "#262B2F",
    },
    "light": {
        "bg": "#F4F7F7", "panel": "#FFFFFF", "line": "#D8E2E3",
        "text": "#10262B", "muted": "#566A70", "faint": "#8A9BA1",
        "track": "#DCE6E7", "accent": "#0B6F83", "accent2": "#1FB5A6",
        "ok": "#1F9D5B", "err": "#D3363A", "on_accent": "#FFFFFF",
        "btn": "#FFFFFF", "btn_hi": "#EEF4F4", "btn_lo": "#E2EBEB",
    },
}

KIND_COLORS = {
    "pdf": "#E5484D", "word": "#2F6FDE", "sheet": "#1E9E57",
    "slide": "#E8772E", "image": "#8B5CF6", "text": "#68787F",
}
EXT_KIND = {
    "pdf": "pdf", "doc": "word", "docx": "word", "docx-fast": "word",
    "odt": "word", "rtf": "word", "xls": "sheet", "xlsx": "sheet",
    "ods": "sheet", "csv": "sheet", "ppt": "slide", "pptx": "slide",
    "odp": "slide", "txt": "text",
}
for _ext in ("png", "jpg", "jpeg", "webp", "bmp", "gif", "tif", "tiff"):
    EXT_KIND[_ext] = "image"


def target_label(source_ext, target):
    if target in ("png", "jpg") and source_ext in (".pdf", ".ppt", ".pptx", ".odp"):
        return "PNG images" if target == "png" else "JPEG images"
    return LABELS.get(target, target.upper())


def kind_color(ext):
    return KIND_COLORS[EXT_KIND.get(ext.lstrip(".").lower(), "text")]


def badge_label(ext):
    ext = ext.lstrip(".").lower()
    if ext == "docx-fast":
        ext = "docx"
    return {"jpeg": "JPG", "tiff": "TIF"}.get(ext, ext.upper())[:4]


# --------------------------------------------------------------------------
# Small colour / easing helpers
# --------------------------------------------------------------------------

def rgb(value):
    value = value.lstrip("#")
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


def hexcolor(color):
    return "#%02x%02x%02x" % tuple(color)


def mix(a, b, t):
    return tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def lighten(color, t):
    return mix(color, (255, 255, 255), t)


def darken(color, t):
    return mix(color, (0, 0, 0), t)


def smooth(t):
    t = max(0.0, min(1.0, t))
    return t * t * (3 - 2 * t)


def ease_out_cubic(t):
    return 1 - (1 - max(0.0, min(1.0, t))) ** 3


def ease_out_back(t):
    t = max(0.0, min(1.0, t))
    c1 = 1.70158
    c3 = c1 + 1
    return 1 + c3 * (t - 1) ** 3 + c1 * (t - 1) ** 2


# --------------------------------------------------------------------------
# Pillow drawing (everything is rendered at SS x, then scaled down)
# --------------------------------------------------------------------------

@lru_cache(maxsize=None)
def pil_font(size, bold=True):
    names = (("seguisb.ttf", "segoeuib.ttf", "arialbd.ttf", "DejaVuSans-Bold.ttf")
             if bold else ("segoeui.ttf", "arial.ttf", "DejaVuSans.ttf"))
    for name in names:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    try:
        return ImageFont.load_default(size)
    except TypeError:
        return ImageFont.load_default()


def pill(w, h, fill, radius=None, outline=None, ow=1):
    """Anti-aliased rounded rectangle, w x h pixels."""
    r = h / 2 if radius is None else radius
    img = Image.new("RGBA", (w * SS, h * SS), rgb(fill) + (0,))
    ImageDraw.Draw(img).rounded_rectangle(
        (0, 0, w * SS - 1, h * SS - 1), radius=r * SS, fill=rgb(fill) + (255,),
        outline=rgb(outline) + (255,) if outline else None,
        width=ow * SS if outline else 0,
    )
    return img.resize((w, h), LANCZOS)


@lru_cache(maxsize=64)
def badge_img(label, color, w, h):
    """A document icon with a folded corner and the file type on it."""
    W, H = w * SS, h * SS
    base = rgb(color)
    fold = int(W * 0.34)
    body = Image.new("RGBA", (W, H))
    d = ImageDraw.Draw(body)
    top, bottom = lighten(base, 0.14), darken(base, 0.12)
    for y in range(H):
        d.line([(0, y), (W, y)], fill=mix(top, bottom, y / max(1, H - 1)) + (255,))
    mask = Image.new("L", (W, H), 0)
    md = ImageDraw.Draw(mask)
    md.rounded_rectangle((0, 0, W - 1, H - 1), radius=int(W * 0.15), fill=255)
    md.polygon([(W - fold, 0), (W, 0), (W, fold)], fill=0)
    body.putalpha(mask)
    flap = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    ImageDraw.Draw(flap).polygon(
        [(W - fold, 0), (W - fold, fold), (W, fold)], fill=lighten(base, 0.6) + (255,))
    body.alpha_composite(flap)
    size = int(H * 0.2)
    font = pil_font(size)
    while size > 8 and ImageDraw.Draw(body).textlength(label, font=font) > W * 0.82:
        size -= 2
        font = pil_font(size)
    ImageDraw.Draw(body).text((W / 2, H * 0.67), label, font=font,
                              fill=(255, 255, 255, 255), anchor="mm")
    return body.resize((w, h), LANCZOS)


def status_img(box, color, ring, kind, t=1.0):
    """Round check / alert badge. `t` is the pop-in scale (can overshoot 1)."""
    W = box * SS
    img = Image.new("RGBA", (W, W), rgb(color) + (0,))
    if t <= 0.02:
        return img.resize((box, box), LANCZOS)
    d = ImageDraw.Draw(img)
    c = W / 2
    R = (W / 2) / 1.3 * t
    d.ellipse((c - R, c - R, c + R, c + R), fill=rgb(ring) + (255,))
    u = R - R * 0.2
    d.ellipse((c - u, c - u, c + u, c + u), fill=rgb(color) + (255,))
    white = (255, 255, 255, 255)
    lw = max(2, int(u * 0.26))

    def dot(x, y):
        d.ellipse((x - lw / 2, y - lw / 2, x + lw / 2, y + lw / 2), fill=white)

    if kind == "ok":
        pts = [(c - 0.44 * u, c + 0.02 * u), (c - 0.11 * u, c + 0.36 * u), (c + 0.5 * u, c - 0.32 * u)]
        d.line(pts, fill=white, width=lw, joint="curve")
        for p in (pts[0], pts[2]):
            dot(*p)
    else:
        d.line([(c, c - 0.46 * u), (c, c + 0.08 * u)], fill=white, width=lw)
        dot(c, c - 0.46 * u)
        d.ellipse((c - lw * 0.62, c + 0.42 * u - lw * 0.62, c + lw * 0.62, c + 0.42 * u + lw * 0.62), fill=white)
    return img.resize((box, box), LANCZOS)


def flow_img(w, h, phase, color, alpha=None):
    """Three chevrons that light up in sequence (a flow from source to target)."""
    W, H = w * SS, h * SS
    img = Image.new("RGBA", (W, H), rgb(color) + (0,))
    d = ImageDraw.Draw(img)
    lw = max(3, int(H * 0.12))
    ch, cw = H * 0.27, H * 0.15
    for i in range(3):
        if alpha is None:
            a = 0.22 + 0.78 * max(0.0, math.sin(2 * math.pi * (phase - i * 0.15))) ** 1.5
        else:
            a = alpha
        fill = rgb(color) + (int(255 * a),)
        cx = W * (i + 1) / 4
        pts = [(cx - cw, H / 2 - ch), (cx + cw, H / 2), (cx - cw, H / 2 + ch)]
        d.line(pts, fill=fill, width=lw, joint="curve")
        for p in (pts[0], pts[2]):
            d.ellipse((p[0] - lw / 2, p[1] - lw / 2, p[0] + lw / 2, p[1] + lw / 2), fill=fill)
    return img.resize((w, h), LANCZOS)


def chevron_img(w, h, color, up=False):
    W, H = w * SS, h * SS
    img = Image.new("RGBA", (W, H), rgb(color) + (0,))
    lw = max(3, int(H * 0.2))
    a, b = (H * 0.22, H * 0.78) if not up else (H * 0.78, H * 0.22)
    pts = [(W * 0.12, a), (W * 0.5, b), (W * 0.88, a)]
    d = ImageDraw.Draw(img)
    d.line(pts, fill=rgb(color) + (255,), width=lw, joint="curve")
    for p in (pts[0], pts[2]):
        d.ellipse((p[0] - lw / 2, p[1] - lw / 2, p[0] + lw / 2, p[1] + lw / 2), fill=rgb(color) + (255,))
    return img.resize((w, h), LANCZOS)


def file_icon_img(size, color):
    """Round "+" icon for the empty 'choose a file' card."""
    W = size * SS
    img = Image.new("RGBA", (W, W), rgb(color) + (0,))
    d = ImageDraw.Draw(img)
    d.ellipse((0, 0, W - 1, W - 1), fill=rgb(color) + (255,))
    lw = max(3, int(W * 0.075))
    c = W / 2
    arm = W * 0.2
    white = (255, 255, 255, 255)
    d.line([(c - arm, c), (c + arm, c)], fill=white, width=lw)
    d.line([(c, c - arm), (c, c + arm)], fill=white, width=lw)
    for p in ((c - arm, c), (c + arm, c), (c, c - arm), (c, c + arm)):
        d.ellipse((p[0] - lw / 2, p[1] - lw / 2, p[0] + lw / 2, p[1] + lw / 2), fill=white)
    return img.resize((size, size), LANCZOS)


@lru_cache(maxsize=16)
def _pill_mask(W, H):
    m = Image.new("L", (W, H), 0)
    ImageDraw.Draw(m).rounded_rectangle((0, 0, W - 1, H - 1), radius=H // 2, fill=255)
    return m


@lru_cache(maxsize=16)
def _track(W, H, color):
    t = Image.new("RGBA", (W, H), rgb(color) + (0,))
    t.paste(rgb(color) + (255,), (0, 0, W, H), _pill_mask(W, H))
    return t


@lru_cache(maxsize=16)
def _strip(W, H, c1, c2):
    """Looping horizontal gradient, 2W wide so any W-wide window is seamless."""
    a, b = rgb(c1), rgb(c2)
    total = 2 * W
    row = Image.new("RGB", (total, 1))
    row.putdata([mix(a, b, (1 - math.cos(2 * math.pi * x / W)) / 2) for x in range(total)])
    return row.resize((total, H), Image.NEAREST)


@lru_cache(maxsize=8)
def _comet(seg, H):
    row = Image.new("L", (seg, 1))
    ramp = seg * 0.34
    row.putdata([int(255 * smooth(min(x, seg - 1 - x) / ramp)) for x in range(seg)])
    return row.resize((seg, H), Image.NEAREST)


def bar_img(w, h, frac, flow, comet, track, c1, c2, tint=None, tint_amt=0.0):
    """Progress bar: flowing gradient fill, or a travelling comet when unknown."""
    W, H = w * SS, h * SS
    out = _track(W, H, track).copy()
    off = int((flow % 1.0) * W)
    color = _strip(W, H, c1, c2).crop((off, 0, off + W, H))
    if tint is not None and tint_amt > 0:
        color = Image.blend(color, Image.new("RGB", (W, H), rgb(tint)), tint_amt)
    mask = Image.new("L", (W, H), 0)
    if comet is not None:
        seg = int(W * 0.36)
        x = int(-seg + (W + seg) * smooth(comet))
        mask.paste(_comet(seg, H), (x, 0))
        mask = ImageChops.multiply(mask, _pill_mask(W, H))
    else:
        fw = int(W * max(0.0, min(1.0, frac)))
        if fw > 0:
            fw = max(fw, H)
            ImageDraw.Draw(mask).rounded_rectangle((0, 0, fw - 1, H - 1), radius=H // 2, fill=255)
    out.paste(color, (0, 0), mask)
    return out.resize((w, h), LANCZOS)


_USE_IMAGETK = ImageTk is not None


def to_photo(img, old=None):
    """Tk image from a Pillow image; reuses `old` in place when sizes match.

    Uses Pillow's Tk bridge when it works and quietly falls back to PNG data
    (slower, but identical to look at) if the bridge is unavailable.
    """
    global _USE_IMAGETK
    if _USE_IMAGETK:
        try:
            if old is not None and old.width() == img.width and old.height() == img.height:
                old.paste(img)
                return old
            return ImageTk.PhotoImage(img)
        except Exception:
            _USE_IMAGETK = False
    buffer = io.BytesIO()
    img.save(buffer, "PNG")
    data = base64.b64encode(buffer.getvalue())
    if isinstance(old, tk.PhotoImage):
        old.configure(data=data)
        return old
    return tk.PhotoImage(data=data)


# --------------------------------------------------------------------------
# Windows integration (all optional; every call is guarded)
# --------------------------------------------------------------------------

def system_is_dark():
    if sys.platform != "win32":
        return True
    try:
        import winreg
        path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
            return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
    except OSError:
        return False


def enable_dpi_awareness():
    """Without this Windows stretches the window as a bitmap and it looks blurry."""
    if sys.platform != "win32":
        return
    import ctypes
    try:
        ctypes.WinDLL("shcore").SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.WinDLL("user32").SetProcessDPIAware()
        except Exception:
            pass


def set_app_id():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.WinDLL("shell32").SetCurrentProcessExplicitAppUserModelID("AaryaSoftwares.AnyConvert")
        except Exception:
            pass


def cursor_monitor():
    """Work area (left, top, right, bottom) and DPI of the monitor under the cursor."""
    if sys.platform != "win32":
        return None, None
    try:
        import ctypes
        from ctypes import wintypes as w

        class MonitorInfo(ctypes.Structure):
            _fields_ = [("size", w.DWORD), ("monitor", w.RECT),
                        ("work", w.RECT), ("flags", w.DWORD)]

        user32 = ctypes.WinDLL("user32", use_last_error=True)
        user32.MonitorFromPoint.argtypes = [w.POINT, w.DWORD]
        user32.MonitorFromPoint.restype = w.HANDLE
        user32.GetMonitorInfoW.argtypes = [w.HANDLE, ctypes.POINTER(MonitorInfo)]
        point = w.POINT()
        user32.GetCursorPos(ctypes.byref(point))
        monitor = user32.MonitorFromPoint(point, 2)
        info = MonitorInfo()
        info.size = ctypes.sizeof(info)
        if not monitor or not user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
            return None, None
        area = (info.work.left, info.work.top, info.work.right, info.work.bottom)
        dpi = None
        try:
            shcore = ctypes.WinDLL("shcore")
            shcore.GetDpiForMonitor.argtypes = [
                w.HANDLE, ctypes.c_int, ctypes.POINTER(ctypes.c_uint), ctypes.POINTER(ctypes.c_uint)]
            dx, dy = ctypes.c_uint(), ctypes.c_uint()
            if shcore.GetDpiForMonitor(monitor, 0, ctypes.byref(dx), ctypes.byref(dy)) == 0:
                dpi = dx.value
        except Exception:
            pass
        return area, dpi
    except Exception:
        return None, None


def style_title_bar(root, theme, dark):
    """Windows 11: dark/light caption, caption colour = window colour, round corners."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        root.update_idletasks()
        hwnd = ctypes.c_void_p(int(root.wm_frame(), 0))
        dwm = ctypes.WinDLL("dwmapi")

        def colorref(value):
            r, g, b = rgb(value)
            return r | (g << 8) | (b << 16)

        for attribute, value in (
            (20, int(dark)),                 # immersive dark mode
            (33, 2),                         # rounded corners
            (34, colorref(theme["line"])),   # border colour
            (35, colorref(theme["bg"])),     # caption colour
            (36, colorref(theme["text"])),   # caption text colour
        ):
            number = ctypes.c_int(value)
            dwm.DwmSetWindowAttribute(hwnd, attribute, ctypes.byref(number), 4)
    except Exception:
        pass


# --------------------------------------------------------------------------
# Conversion worker (unchanged behaviour)
# --------------------------------------------------------------------------

class LogStream(io.TextIOBase):
    def __init__(self, events):
        self.events = events

    def write(self, text):
        if text:
            self.events.put(("log", text))
        return len(text)


def convert_file(source, target, events):
    stream = LogStream(events)
    original_args = sys.argv
    original_stdin = sys.stdin
    logger = logging.getLogger()
    previous_handlers = logger.handlers[:]
    previous_level = logger.level
    handler = logging.StreamHandler(stream)
    handler.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    exit_code = 1

    try:
        with redirect_stdout(stream), redirect_stderr(stream):
            sys.stdin = io.StringIO()
            sys.argv = [APP_NAME, "--format", target, str(source)]
            logger.handlers = [handler]
            logger.setLevel(logging.INFO)
            import converter
            exit_code = converter.main()
    except (Exception, SystemExit) as error:
        stream.write(f"\nConversion failed: {error}\n")
    finally:
        logger.handlers = previous_handlers
        logger.setLevel(previous_level)
        handler.close()
        sys.argv = original_args
        sys.stdin = original_stdin
        events.put(("done", exit_code))


# --------------------------------------------------------------------------
# Progress model: turns converter log lines into a status and a fraction
# --------------------------------------------------------------------------

ANSI = re.compile(r"\x1b\[[0-9;?]*[ -/]*[@-~]")

# pdf2docx reports four stages; "analyzing" is usually the slowest by far.
STAGE_WEIGHTS = {4: (0.05, 0.45, 0.30, 0.20)}


def stage_span(k, n):
    weights = STAGE_WEIGHTS.get(n) or (1 / n,) * n
    return sum(weights[:k - 1]), weights[k - 1]


class Model:
    STAGE = re.compile(r"\[(\d+)/(\d+)\]\s*(.+)")
    COUNT = re.compile(r"(?:\bpage\s+|\()(\d+)\s*/\s*(\d+)", re.I)
    PREFIX = re.compile(r"^(?:\[[A-Z]+\]\s*)?(.*?\bpage)\s+\d+\s*/", re.I)

    def __init__(self):
        self.fraction = None      # None means "unknown": show the travelling bar
        self.step = 0.0           # expected size of the next real update
        self.status = "Starting…"
        self.stage = None
        self.stage_text = ""
        self.saved = []
        self.error = None
        self.updated = time.perf_counter()

    def _advance(self, fraction, step):
        self.fraction = max(self.fraction or 0.0, min(fraction, 0.99))
        self.step = step
        self.updated = time.perf_counter()

    def feed(self, line):
        raw = line
        line = ANSI.sub("", line).strip()
        if not line:
            return
        if line.startswith("Saved "):
            self.saved.append(line[6:].split(" with ")[0].strip())
            return
        if line.startswith("Conversion failed"):
            self.error = line.split(":", 1)[1].strip() if ":" in line else line
            return
        stage = self.STAGE.search(line)
        if stage:
            k, n = int(stage[1]), int(stage[2])
            if n >= 1 and 1 <= k <= n:
                self.stage = (k, n)
                self.stage_text = stage[3].strip().rstrip(".… ")
                self.status = self.stage_text
                start, width = stage_span(k, n)
                self._advance(start, width)
            return
        count = self.COUNT.search(line)
        if count:
            i, n = int(count[1]), int(count[2])
            if 1 <= i <= n:
                if self.stage:
                    start, width = stage_span(*self.stage)
                    self._advance(start + width * i / n, width / n)
                    self.status = f"{self.stage_text} · {i} of {n}"
                else:
                    self._advance((i - 1) / n, 1 / n)
                    prefix = self.PREFIX.match(line)
                    self.status = f"{prefix[1].strip()} {i} of {n}" if prefix else f"Page {i} of {n}"
            return
        # Ignore log chatter and library warnings; only show our own messages.
        if (line.startswith("[") or raw[:1].isspace() or line.lower().startswith("terminated")
                or re.search(r"warning|\.py:|traceback", line, re.I)):
            return
        self.status = line.rstrip(".… ")

    def goal(self, now):
        """Where the bar should head: the last real value plus a small creep."""
        if self.fraction is None:
            return None
        idle = now - self.updated
        headroom = 0.55 * self.step * (1 - math.exp(-idle / 5))
        return min(0.99, self.fraction + headroom)


def friendly_error(reason):
    reason = (reason or "").strip()
    if not reason:
        return "Something went wrong while converting this file."
    if "password" in reason.lower():
        return "This file is password protected. Remove the password and try again."
    reason = reason[0].upper() + reason[1:]
    return reason if reason.endswith((".", "!", "?")) else reason + "."


def fit(text, font, width, keep_tail=0):
    """Shorten text with an ellipsis so it fits `width` pixels."""
    if font.measure(text) <= width:
        return text
    tail = text[-keep_tail:] if keep_tail else ""
    head = text[:-keep_tail] if keep_tail else text
    while head and font.measure(head + "…" + tail) > width:
        head = head[:-1]
    return head + "…" + tail


# --------------------------------------------------------------------------
# The window
# --------------------------------------------------------------------------

class ConvertWindow:
    def __init__(self, root):
        self.root = root
        self.dark = system_is_dark()
        self.T = THEMES["dark" if self.dark else "light"]
        self.area, dpi = cursor_monitor()
        dpi = dpi or root.winfo_fpixels("1i")
        self.S = max(0.75, min(4.0, dpi / 96.0))

        root.configure(bg=self.T["bg"])
        root.resizable(False, False)
        self.make_fonts()
        self.canvas = tk.Canvas(root, bg=self.T["bg"], highlightthickness=0, bd=0,
                                width=self.px(WIDTH), height=self.px(190))
        self.canvas.pack(fill="both", expand=True)
        self.make_log()

        self.events = queue.Queue()
        self.model = Model()
        self.state = "empty"          # empty, ready, running, done, failed
        self.source = None
        self.formats = ()
        self.target = None
        self.details_open = False
        self.shown = False
        self.cur_h = None
        self.resize_token = 0
        self.buffer = ""
        self.keep = []
        self.hint_job = None
        self.fail_text = ""
        self.done_status = ""
        self.footer_y = 160
        self.dst_pos = (0, 0, 40, 52)
        self.flow_args = (1, 1, self.T["accent"], None)

        self.started = 0.0
        self.disp = 0.0
        self.flow_phase = 0.0
        self.stripe_phase = 0.0
        self.comet_phase = 0.0
        self.done_at = 0.0
        self.last = time.perf_counter()
        self.last_texts = {}
        self.bar_photo = None
        self.flow_photo = None
        self.overlay_photo = None
        self.overlay_item = None

        root.bind("<Escape>", lambda _e: self.request_close())
        root.bind("<Return>", lambda _e: self.default_action())

    # ---- basics ---------------------------------------------------------
    def px(self, value):
        return int(round(value * self.S))

    def make_fonts(self):
        families = set(tkfont.families(self.root))

        def pick(*names):
            for name in names:
                if name in families:
                    return name
            return tkfont.nametofont("TkDefaultFont").actual("family")

        ui = pick("Segoe UI Variable Text", "Segoe UI", "Inter", "Open Sans", "DejaVu Sans")
        semi_name = "Segoe UI Semibold"
        semi_ok = semi_name in families
        mono = pick("Cascadia Mono", "Consolas", "DejaVu Sans Mono")

        def font(family, size, weight="normal"):
            return tkfont.Font(root=self.root, family=family, size=-self.px(size), weight=weight)

        heavy = (semi_name, "normal") if semi_ok else (ui, "bold")
        self.f_title = font(heavy[0], 15, heavy[1])
        self.f_big = font(heavy[0], 19, heavy[1])
        self.f_btn = font(heavy[0], 13, heavy[1])
        self.f_body = font(ui, 13)
        self.f_small = font(ui, 12)
        self.f_mono = font(mono, 12)

    def make_log(self):
        T = self.T
        self.log = tk.Text(
            self.canvas, wrap="word", state="disabled", relief="flat", bd=0,
            highlightthickness=0, bg=T["panel"], fg=T["muted"], font=self.f_mono,
            padx=self.px(8), pady=self.px(4), cursor="arrow", insertwidth=0,
            selectbackground=T["accent"], selectforeground=T["on_accent"],
        )
        self.log.tag_configure("err", foreground=T["err"])
        self.log.tag_configure("ok", foreground=T["ok"])

    def keepalive(self, photo):
        self.keep.append(photo)
        return photo

    def put(self, img, x, y, anchor="nw", tags=()):
        photo = self.keepalive(to_photo(img))
        item = self.canvas.create_image(self.px(x), self.px(y), image=photo, anchor=anchor, tags=tags)
        return item, photo

    def label(self, x, y, text, font, fill, anchor="nw", width=None, tags=()):
        options = {"width": self.px(width)} if width else {}
        return self.canvas.create_text(self.px(x), self.px(y), text=text, font=font,
                                       fill=fill, anchor=anchor, tags=tags, **options)

    # ---- geometry / showing --------------------------------------------
    def work_area(self):
        if self.area:
            return self.area
        return 0, 0, self.root.winfo_screenwidth(), self.root.winfo_screenheight()

    def apply_geometry(self, h):
        w = self.px(WIDTH)
        left, top, right, bottom = self.work_area()
        caption = max(0, self.root.winfo_rooty() - self.root.winfo_y())
        x = left + (right - left - w) // 2
        y = top + (bottom - top - (h + caption)) // 2
        self.canvas.configure(width=w, height=h)
        self.root.geometry(f"{w}x{h}+{x}+{max(top, y)}")

    def resize(self, logical_h, animate):
        target = self.px(logical_h)
        start = self.cur_h
        self.cur_h = target
        self.resize_token += 1
        token = self.resize_token
        if not animate or start is None or start == target or not self.shown:
            self.apply_geometry(target)
            return
        steps = 9

        def step(i):
            if token != self.resize_token:
                return
            t = ease_out_cubic(i / steps)
            self.apply_geometry(round(start + (target - start) * t))
            if i < steps:
                self.root.after(13, step, i + 1)

        step(1)

    def fade(self, start, end, steps, then=None):
        def step(i):
            try:
                self.root.attributes("-alpha", start + (end - start) * i / steps)
                if i < steps:
                    self.root.after(14, step, i + 1)
                elif then:
                    then()
            except tk.TclError:
                pass

        step(0)

    def show(self):
        root = self.root
        root.attributes("-alpha", 0.0)
        root.deiconify()
        root.update_idletasks()
        style_title_bar(root, self.T, self.dark)
        self.shown = True
        self.apply_geometry(self.cur_h)
        root.lift()
        root.attributes("-topmost", True)
        root.after(300, lambda: root.attributes("-topmost", False))
        try:
            root.focus_force()
        except tk.TclError:
            pass
        self.fade(0.0, 1.0, 7)

    # ---- state entry points --------------------------------------------
    def begin(self):
        args = sys.argv[1:]
        explorer_target = None
        path = None
        if len(args) >= 2:
            explorer_target, path = args[0], args[1]
        elif len(args) == 1:
            path = args[0]

        problem = None
        if path:
            if self.load(path):
                if explorer_target is not None:
                    if explorer_target in self.formats:
                        self.target = explorer_target
                    else:
                        problem = "This conversion isn’t available for this file type."
            else:
                problem = "Choose a supported document, spreadsheet, presentation or image."
        if problem:
            self.fail_with(problem, animate=False)
        elif explorer_target is not None:
            self.state = "running"
            self.prepare_running()
            self.render(animate=False)
        else:
            self.state = "ready" if self.source else "empty"
            self.render(animate=False)
        self.show()
        if self.state == "running":
            self.root.after(90, self.launch)
        self.last = time.perf_counter()
        self.tick()

    def load(self, path):
        candidate = Path(path).resolve()
        options = FORMATS.get(candidate.suffix.lower(), ())
        if not candidate.is_file() or not options:
            return False
        self.source = candidate
        self.formats = options
        self.target = "pdf" if "pdf" in options else options[0]
        return True

    def browse(self):
        path = filedialog.askopenfilename(
            parent=self.root,
            filetypes=[("Supported files", " ".join("*" + ext for ext in FORMATS))],
        )
        if not path:
            return
        if self.load(path):
            self.state = "ready"
            self.render(animate=True)
        else:
            self.fail_with("Choose a supported document, spreadsheet, presentation or image.",
                           animate=True)

    def choose(self, fmt):
        self.target = fmt
        self.render(animate=False)

    def default_action(self):
        if self.state == "ready":
            self.start()
        elif self.state == "failed":
            self.root.destroy()

    def start(self):
        if self.state != "ready" or self.source is None or self.target is None:
            return
        self.state = "running"
        self.prepare_running()
        self.render(animate=True)
        self.root.after(90, self.launch)

    def prepare_running(self):
        self.model = Model()
        self.disp = 0.0
        self.started = time.perf_counter()
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def launch(self):
        threading.Thread(target=convert_file, args=(self.source, self.target, self.events),
                         daemon=True).start()

    def fail_with(self, message, animate):
        """Show an error that happened before any conversion started."""
        self.state = "failed"
        self.fail_text = message
        self.render(animate=animate)

    def request_close(self):
        if self.state == "running":
            self.flash_hint("Almost there — this window closes by itself.")
            return
        self.root.destroy()

    def flash_hint(self, text):
        if self.hint_job:
            self.root.after_cancel(self.hint_job)
        self.canvas.delete("hint")
        self.label(WIDTH / 2, self.footer_y, text, self.f_small, self.T["muted"],
                   anchor="center", tags=("hint",))
        self.hint_job = self.root.after(2200, lambda: self.canvas.delete("hint"))

    # ---- widgets ---------------------------------------------------------
    def add_button(self, x, y, w, h, text, command, primary=False, selected=False,
                   chip=False):
        T, c = self.T, self.canvas
        pw, ph = self.px(w), self.px(h)
        radius = ph / 2 if chip else self.px(8)
        accent = rgb(T["accent"])

        def colors(state):
            if primary:
                fill = hexcolor({"n": accent, "h": lighten(accent, 0.12), "p": darken(accent, 0.12)}[state])
                return fill, fill, T["on_accent"]
            if selected:
                tint = mix(rgb(T["bg"]), accent, {"n": 0.20, "h": 0.28, "p": 0.14}[state])
                return hexcolor(tint), T["accent"], T["text"]
            fill = {"n": T["btn"], "h": T["btn_hi"], "p": T["btn_lo"]}[state]
            return fill, T["line"], T["text"]

        images = {}
        for state in ("n", "h", "p"):
            fill, outline, _ = colors(state)
            images[state] = self.keepalive(to_photo(pill(pw, ph, fill, radius, outline, 1)))
        tag = f"btn{len(self.keep)}"
        item = c.create_image(self.px(x), self.px(y), image=images["n"], anchor="nw", tags=(tag,))
        c.create_text(self.px(x) + pw / 2, self.px(y) + ph / 2, text=text, font=self.f_btn,
                      fill=colors("n")[2], anchor="center", tags=(tag,))
        inside = {"v": False}

        def enter(_e):
            inside["v"] = True
            c.configure(cursor="hand2")
            c.itemconfigure(item, image=images["h"])

        def leave(_e):
            inside["v"] = False
            c.configure(cursor="")
            c.itemconfigure(item, image=images["n"])

        def press(_e):
            c.itemconfigure(item, image=images["p"])

        def release(_e):
            c.itemconfigure(item, image=images["h"] if inside["v"] else images["n"])
            if inside["v"]:
                command()

        c.tag_bind(tag, "<Enter>", enter)
        c.tag_bind(tag, "<Leave>", leave)
        c.tag_bind(tag, "<ButtonPress-1>", press)
        c.tag_bind(tag, "<ButtonRelease-1>", release)

    def add_link(self, x, y_center, text, command, color=None, hover=None, chevron=None):
        """Text link (optionally followed by a chevron) centred on y_center."""
        T, c = self.T, self.canvas
        color = color or T["muted"]
        hover = hover or T["text"]
        tag = f"lnk{len(self.keep)}"
        item = self.label(x, y_center, text, self.f_small, color, anchor="w", tags=(tag,))
        chev = {}
        if chevron is not None:
            x1 = c.bbox(item)[2]
            up = chevron == "up"
            chev["n"] = self.keepalive(to_photo(chevron_img(self.px(10), self.px(7), color, up)))
            chev["h"] = self.keepalive(to_photo(chevron_img(self.px(10), self.px(7), hover, up)))
            chev["item"] = c.create_image(x1 + self.px(6), self.px(y_center), image=chev["n"],
                                          anchor="w", tags=(tag,))

        def enter(_e):
            c.configure(cursor="hand2")
            c.itemconfigure(item, fill=hover)
            if chev:
                c.itemconfigure(chev["item"], image=chev["h"])

        def leave(_e):
            c.configure(cursor="")
            c.itemconfigure(item, fill=color)
            if chev:
                c.itemconfigure(chev["item"], image=chev["n"])

        c.tag_bind(tag, "<Enter>", enter)
        c.tag_bind(tag, "<Leave>", leave)
        c.tag_bind(tag, "<ButtonRelease-1>", lambda _e: command())
        return item

    # ---- rendering -------------------------------------------------------
    def render(self, animate=False):
        self.canvas.delete("all")
        self.keep = []
        self.last_texts = {}
        self.overlay_item = None
        self.bar_photo = None
        self.flow_photo = None
        draw = {"empty": self.draw_empty, "ready": self.draw_ready,
                "running": self.draw_running, "done": self.draw_running,
                "failed": self.draw_failed}[self.state]
        height = draw()
        self.resize(height, animate)
        self.root.title(APP_NAME)

    def draw_header(self, name, subtitle, sub_color, src_ext, dst_ext, flow, link=None):
        T = self.T
        bw, bh, y0, flow_w = 40, 52, 22, 46
        x_dst = PAD + bw + flow_w
        self.put(badge_img(badge_label(src_ext), kind_color(src_ext), self.px(bw), self.px(bh)), PAD, y0)
        self.put(badge_img(badge_label(dst_ext), kind_color(dst_ext), self.px(bw), self.px(bh)), x_dst, y0)
        self.dst_pos = (x_dst, y0, bw, bh)

        if flow == "run":
            color, alpha = T["accent"], None
        elif flow == "ok":
            color, alpha = T["ok"], 1.0
        elif flow == "err":
            color, alpha = T["err"], 0.85
        else:
            color, alpha = T["faint"], 0.7
        self.flow_args = (self.px(flow_w), self.px(24), color, alpha)
        self.flow_item, self.flow_photo = self.put(
            flow_img(self.px(flow_w), self.px(24), 0.0, color, alpha), PAD + bw, y0 + (bh - 24) / 2)

        x_text = x_dst + bw + 20
        text_w = WIDTH - x_text - PAD
        keep_tail = len(Path(name).suffix) + 3
        self.title_item = self.label(x_text, y0 + 4, fit(name, self.f_title, self.px(text_w), keep_tail),
                                     self.f_title, T["text"])
        self.subtitle_item = None
        if link:
            self.add_link(x_text, y0 + 36, subtitle, link, color=sub_color, hover=T["text"])
        else:
            self.subtitle_item = self.label(x_text, y0 + 29, fit(subtitle, self.f_small, self.px(text_w)),
                                            self.f_small, sub_color)

    def draw_details(self, y):
        T = self.T
        w, h = self.px(WIDTH - 2 * PAD), self.px(DETAILS_H)
        self.put(pill(w, h, T["panel"], self.px(10), T["line"], 1), PAD, y)
        self.canvas.create_window(self.px(PAD) + self.px(5), self.px(y) + self.px(5), window=self.log,
                                  anchor="nw", width=w - self.px(10), height=h - self.px(10))
        self.log.see("end")

    def draw_empty(self):
        T, c = self.T, self.canvas
        self.label(PAD, 24, "Convert a file", self.f_big, T["text"])
        self.label(PAD, 54, "Choose a document, spreadsheet, presentation or image.",
                   self.f_small, T["muted"])
        card_y, card_h = 92, 84
        w, h = self.px(WIDTH - 2 * PAD), self.px(card_h)
        tag = "card"
        base = self.keepalive(to_photo(pill(w, h, T["panel"], self.px(12), T["line"], 1)))
        hover = self.keepalive(to_photo(pill(w, h, T["btn_hi"], self.px(12), T["accent"], 1)))
        item = c.create_image(self.px(PAD), self.px(card_y), image=base, anchor="nw", tags=(tag,))
        icon = self.keepalive(to_photo(file_icon_img(self.px(40), T["accent"])))
        c.create_image(self.px(PAD + 22), self.px(card_y + card_h / 2), image=icon, anchor="w", tags=(tag,))
        self.label(PAD + 78, card_y + 24, "Choose a file…", self.f_title, T["text"], tags=(tag,))
        self.label(PAD + 78, card_y + 48, "PDF, Word, Excel, PowerPoint, CSV or images",
                   self.f_small, T["muted"], tags=(tag,))
        c.tag_bind(tag, "<Enter>", lambda _e: (c.configure(cursor="hand2"), c.itemconfigure(item, image=hover)))
        c.tag_bind(tag, "<Leave>", lambda _e: (c.configure(cursor=""), c.itemconfigure(item, image=base)))
        c.tag_bind(tag, "<ButtonRelease-1>", lambda _e: self.browse())
        self.footer_y = card_y + card_h + 10
        return card_y + card_h + 26

    def draw_ready(self):
        T = self.T
        self.draw_header(self.source.name, "Change file", T["accent"], self.source.suffix,
                         self.target, flow="idle", link=self.browse)
        self.label(PAD, 98, "Convert to", self.f_small, T["muted"])
        x, y = PAD, 122
        for fmt in self.formats:
            text = LABELS.get(fmt, fmt.upper())
            w = self.f_btn.measure(text) / self.S + 30
            if x + w > WIDTH - PAD:
                x, y = PAD, y + 42
            self.add_button(x, y, w, 34, text, lambda f=fmt: self.choose(f),
                            selected=(fmt == self.target), chip=True)
            x += w + 8
        by = y + 34 + 22
        self.add_button(WIDTH - PAD - 132, by, 132, 38, "Convert", self.start, primary=True)
        self.footer_y = by + 19
        return by + 38 + 24

    def draw_running(self):
        T = self.T
        src_ext = self.source.suffix
        done = self.state == "done"
        self.draw_header(self.source.name, f"Converting to {target_label(src_ext, self.target)}",
                         T["muted"], src_ext, self.target, flow="ok" if done else "run")
        self.status_item = self.label(PAD, 106, "", self.f_body, T["text"], anchor="w")
        self.percent_item = self.label(WIDTH - PAD, 106, "", self.f_btn, T["muted"], anchor="e")
        bar_w, bar_h = self.px(WIDTH - 2 * PAD), self.px(8)
        self.bar_item, self.bar_photo = self.put(
            bar_img(bar_w, bar_h, 0, 0, None, T["track"], T["accent"], T["accent2"]), PAD, 124)
        self.footer_y = 156
        self.add_link(PAD, 156, "Hide details" if self.details_open else "Details",
                      self.toggle_details, chevron="up" if self.details_open else "down")
        self.elapsed_item = self.label(WIDTH - PAD, 156, "00:00", self.f_small, T["muted"], anchor="e")
        height = 184
        if self.details_open:
            self.draw_details(height - 8)
            height += DETAILS_H + 4
        if done:
            self.place_overlay("ok", 1.0)
        self.refresh_texts(force=True)
        return height

    def draw_failed(self):
        T = self.T
        if self.source is not None and self.target is not None and self.fail_text == "":
            self.draw_header(self.source.name, "Conversion failed", T["err"], self.source.suffix,
                             self.target, flow="err")
            self.place_overlay("err", 1.0)
            reason = friendly_error(self.model.error or self.last_log_line())
            top, with_details = 96, True
        else:
            self.label(PAD, 24, "Can’t convert this file", self.f_big, T["err"])
            reason, top, with_details = self.fail_text, 62, False
        message = self.label(PAD, top, reason, self.f_body, T["text"], width=WIDTH - 2 * PAD)
        y = self.canvas.bbox(message)[3] / self.S + 18
        if with_details:
            self.add_link(PAD, y + 18, "Hide details" if self.details_open else "Details",
                          self.toggle_details, chevron="up" if self.details_open else "down")
            self.add_button(WIDTH - PAD - 104 - 8 - 132, y, 132, 36, "Copy details", self.copy_details)
        self.add_button(WIDTH - PAD - 104, y, 104, 36, "Close", self.root.destroy, primary=True)
        self.footer_y = y + 18
        height = y + 36 + 22
        if self.details_open and with_details:
            self.draw_details(height - 8)
            height += DETAILS_H + 4
        return height

    def last_log_line(self):
        lines = [ln for ln in self.log.get("1.0", "end").splitlines() if ln.strip()]
        return lines[-1] if lines else ""

    def place_overlay(self, kind, t):
        T = self.T
        x, y, bw, bh = self.dst_pos
        box = self.px(28)
        color = T["ok"] if kind == "ok" else T["err"]
        img = status_img(box, color, T["bg"], kind, t)
        if self.overlay_item is None:
            self.overlay_item, self.overlay_photo = self.put(img, 0, 0, anchor="center")
            self.canvas.coords(self.overlay_item, self.px(x + bw - 2), self.px(y + bh + 1))
        else:
            self.overlay_photo = to_photo(img, self.overlay_photo)
            self.canvas.itemconfigure(self.overlay_item, image=self.overlay_photo)

    def copy_details(self):
        text = self.log.get("1.0", "end").strip() or (self.model.error or "")
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.canvas.delete("copied")
        self.label(PAD, self.footer_y - 9, "Copied to clipboard", self.f_small, self.T["ok"],
                   anchor="w", tags=("copied",))
        self.root.after(1800, lambda: self.canvas.delete("copied"))

    def toggle_details(self):
        self.details_open = not self.details_open
        self.render(animate=True)

    # ---- running-state updates ------------------------------------------
    def refresh_texts(self, force=False):
        if self.state not in ("running", "done"):
            return
        pct = None
        if self.state == "done":
            status, percent = self.done_status, "Done"
        else:
            status = self.model.status
            if self.model.fraction is not None:
                pct = min(99, int(self.disp * 100))
            percent = "" if pct is None else f"{pct}%"
        elapsed = int(time.perf_counter() - self.started)
        texts = {
            "status": fit(status, self.f_body, self.px(WIDTH - 2 * PAD - 70)),
            "percent": percent,
            "elapsed": f"{elapsed // 60:02d}:{elapsed % 60:02d}",
        }
        for key, item in (("status", self.status_item), ("percent", self.percent_item),
                          ("elapsed", self.elapsed_item)):
            if force or self.last_texts.get(key) != texts[key]:
                self.canvas.itemconfigure(item, text=texts[key])
        if self.state == "done":
            self.canvas.itemconfigure(self.percent_item, fill=self.T["ok"])
        self.last_texts = texts
        if self.state == "running":
            self.root.title(f"{pct}% · {APP_NAME}" if pct else APP_NAME)

    def log_append(self, line):
        tag = ""
        if re.match(r"\[(ERROR|CRITICAL|WARNING)\]|Conversion failed", line):
            tag = "err"
        elif line.startswith("Saved "):
            tag = "ok"
        self.log.configure(state="normal")
        self.log.insert("end", line + "\n", tag)
        self.log.configure(state="disabled")
        self.log.see("end")

    def pump(self):
        for _ in range(200):
            try:
                kind, value = self.events.get_nowait()
            except queue.Empty:
                return
            if kind == "log":
                self.buffer += value.replace("\r", "\n")
                while "\n" in self.buffer:
                    line, self.buffer = self.buffer.split("\n", 1)
                    if line.strip():
                        self.model.feed(line)
                        self.log_append(ANSI.sub("", line))
            else:
                if self.buffer.strip():
                    self.model.feed(self.buffer)
                    self.log_append(ANSI.sub("", self.buffer).strip())
                self.buffer = ""
                self.finish(value)
                return

    def finish(self, code):
        if code == 0:
            saved = self.model.saved
            if len(saved) == 1:
                self.done_status = f"Saved {saved[0]}"
            elif saved:
                self.done_status = f"Saved {len(saved)} files"
            else:
                self.done_status = "Conversion complete"
            self.state = "done"
            self.done_at = time.perf_counter()
            if self.subtitle_item is not None:
                self.canvas.itemconfigure(self.subtitle_item, text="Finished", fill=self.T["ok"])
            w, h, color, _ = self.flow_args
            self.flow_args = (w, h, self.T["ok"], 1.0)
            self.flow_photo = to_photo(flow_img(w, h, 0.0, self.T["ok"], 1.0), self.flow_photo)
            self.canvas.itemconfigure(self.flow_item, image=self.flow_photo)
            self.place_overlay("ok", 0.0)
            self.refresh_texts(force=True)
            self.root.title(APP_NAME)
        else:
            self.state = "failed"
            self.fail_text = ""
            self.render(animate=True)

    # ---- animation loop --------------------------------------------------
    def tick(self):
        try:
            now = time.perf_counter()
            dt = min(0.1, now - self.last)
            self.last = now
            self.pump()
            if self.state == "running":
                self.tick_running(now, dt)
            elif self.state == "done":
                self.tick_done(now, dt)
            self.root.after(16, self.tick)
        except tk.TclError:
            pass

    def update_flow(self, phase):
        w, h, color, alpha = self.flow_args
        if alpha is not None:
            return
        self.flow_photo = to_photo(flow_img(w, h, phase, color, None), self.flow_photo)
        self.canvas.itemconfigure(self.flow_item, image=self.flow_photo)

    def set_bar(self, img):
        photo = to_photo(img, self.bar_photo)
        if photo is not self.bar_photo:
            self.bar_photo = photo
            self.canvas.itemconfigure(self.bar_item, image=photo)

    def tick_running(self, now, dt):
        T = self.T
        goal = self.model.goal(now)
        self.flow_phase = (self.flow_phase + dt / 1.3) % 1.0
        self.stripe_phase = (self.stripe_phase + dt / 2.6) % 1.0
        comet = None
        if goal is None:
            self.comet_phase = (self.comet_phase + dt / 1.5) % 1.0
            comet = self.comet_phase
        else:
            self.disp = max(self.disp, self.disp + (goal - self.disp) * (1 - math.exp(-dt * 6)))
        self.set_bar(bar_img(self.px(WIDTH - 2 * PAD), self.px(8), self.disp, self.stripe_phase,
                             comet, T["track"], T["accent"], T["accent2"]))
        self.update_flow(self.flow_phase)
        self.refresh_texts()

    def tick_done(self, now, dt):
        T = self.T
        t = now - self.done_at
        self.disp += (1.0 - self.disp) * (1 - math.exp(-dt * 14))
        frac = 1.0 if self.disp > 0.995 else self.disp
        self.set_bar(bar_img(self.px(WIDTH - 2 * PAD), self.px(8), frac, self.stripe_phase, None,
                             T["track"], T["accent"], T["accent2"],
                             tint=T["ok"], tint_amt=smooth(t / 0.4)))
        if t < 0.5:
            self.place_overlay("ok", ease_out_back(t / 0.45))
        hold = 1.0
        if t > hold:
            alpha = 1.0 - (t - hold) / 0.18
            if alpha <= 0:
                self.root.destroy()
                return
            self.root.attributes("-alpha", alpha)
        self.refresh_texts()


def main():
    enable_dpi_awareness()
    set_app_id()
    sys.setswitchinterval(0.002)  # keep the window smooth while converting
    root = tk.Tk()
    root.withdraw()
    root.title(APP_NAME)
    icon = Path(sys.executable if getattr(sys, "frozen", False) else __file__).with_name("AnyConvert.ico")
    if icon.is_file():
        try:
            root.iconbitmap(str(icon))
        except tk.TclError:
            pass
    window = ConvertWindow(root)
    root.protocol("WM_DELETE_WINDOW", window.request_close)
    root.after_idle(window.begin)
    root.mainloop()


if __name__ == "__main__":
    main()
