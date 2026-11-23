import io
import logging
import queue
import sys
import threading
import tkinter as tk
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from anyconvert_formats import FORMATS
except ModuleNotFoundError:
    from converter import FORMATS
import time
import re


APP_NAME = "AnyConvert"
LABELS = {
    "pptx": "PowerPoint slides", "txt": "Plain text", "csv": "CSV",
    "webp": "WebP", "bmp": "Bitmap", "tiff": "TIFF",
    "pdf": "A4 PDF",
    "png": "PNG",
    "jpg": "JPEG",
    "docx": "Word document",
    "docx-fast": "Word document (fast text)",
    "xlsx": "Excel workbook",
}


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


def main():
    root = tk.Tk()
    root.withdraw()
    root.title(APP_NAME)
    icon = Path(sys.executable if getattr(sys, "frozen", False) else __file__).with_name("AnyConvert.ico")
    if icon.is_file():
        root.iconbitmap(str(icon))
    root.resizable(False, False)

    frame = ttk.Frame(root, padding=18)
    frame.grid(sticky="nsew")

    events = queue.Queue()
    source = None
    formats = ()
    busy = False
    info_open = False
    started = 0.0
    active_stage = "Preparing"
    info_lines = []

    filename = tk.StringVar(value="Choose a file to convert")
    status = tk.StringVar(value="")

    ttk.Label(
        frame,
        textvariable=filename,
        font=("Segoe UI", 11, "bold"),
        wraplength=390,
    ).grid(row=0, column=0, sticky="w")

    controls = ttk.Frame(frame)
    controls.grid(row=1, column=0, sticky="ew", pady=(12, 0))

    def select_file(path):
        nonlocal source, formats
        candidate = Path(path).resolve()
        options = FORMATS.get(candidate.suffix.lower(), ())
        if not candidate.is_file() or not options:
            messagebox.showerror(
                APP_NAME,
                "Choose a supported document, spreadsheet, presentation or image.",
                parent=root,
            )
            return

        source = candidate
        formats = options
        filename.set(candidate.name)
        picker.configure(values=[LABELS.get(fmt, fmt.upper()) for fmt in formats])
        picker.current(formats.index("pdf") if "pdf" in formats else 0)
        convert_button.configure(state="normal")

    def browse():
        path = filedialog.askopenfilename(
            parent=root,
            filetypes=[("Supported files", " ".join("*" + ext for ext in FORMATS))],
        )
        if path:
            select_file(path)

    ttk.Button(controls, text="Choose file", command=browse).grid(
        row=0, column=0, sticky="w"
    )
    picker = ttk.Combobox(controls, state="readonly", width=23)
    picker.grid(row=0, column=1, padx=8)

    activity = ttk.Frame(frame)
    activity.grid(row=2, column=0, sticky="w", pady=(12, 0))
    progress_value = tk.DoubleVar(value=0)
    percent = tk.StringVar(value="")
    elapsed = tk.StringVar(value="00:00")
    ttk.Label(activity, textvariable=status).grid(row=0, column=0, sticky="w")
    ttk.Label(activity, textvariable=percent).grid(row=0, column=1, sticky="e")
    progress_bar = ttk.Progressbar(activity, length=360, maximum=100)
    progress_bar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 10))
    ttk.Label(activity, text="Elapsed time").grid(row=2, column=0, sticky="w")
    ttk.Label(activity, textvariable=elapsed).grid(row=2, column=1, sticky="e")
    activity.grid_remove()

    def unknown_progress():
        progress_bar.stop()
        progress_bar.configure(mode="indeterminate")
        progress_bar.start(45)
        percent.set("")

    def update_stage(message):
        nonlocal active_stage
        stage = re.search(r"\[(\d+)/4\]\s*(.+)", message)
        if stage:
            active_stage = stage.group(2).strip().rstrip(".")
            status.set(active_stage)
            unknown_progress()
            return
        count = re.search(r"(?:\bpage\s+|\()(\d+)\s*/\s*(\d+)", message, re.I)
        if count:
            page, total = map(int, count.groups())
            if 1 <= page <= total:
                progress_bar.stop()
                value = 100 * page / total
                progress_bar.configure(mode="determinate", value=value)
                percent.set(f"{value:.0f}%")
                status.set(f"{active_stage}  ·  {page} of {total}")

    def center_convertany_window():
        # Center the outer window, including its title bar, on the cursor's monitor.
        import ctypes
        from ctypes import wintypes as w

        root.update_idletasks()
        api = ctypes.WinDLL("user32", use_last_error=True)
        api.MonitorFromPoint.argtypes = [w.POINT, w.DWORD]
        api.MonitorFromPoint.restype = w.HANDLE
        class MonitorInfo(ctypes.Structure):
            _fields_ = [("size", w.DWORD), ("monitor", w.RECT),
                        ("work", w.RECT), ("flags", w.DWORD)]
        point = w.POINT()
        api.GetCursorPos(ctypes.byref(point))
        monitor = api.MonitorFromPoint(point, 2)
        info = MonitorInfo()
        info.size = ctypes.sizeof(info)
        rect = w.RECT()
        hwnd = w.HWND(int(root.wm_frame(), 0))
        if (api.GetMonitorInfoW(w.HANDLE(monitor), ctypes.byref(info))
                and api.GetWindowRect(hwnd, ctypes.byref(rect))):
            area = info.work
            x = (area.left + area.right - rect.right + rect.left) // 2
            y = (area.top + area.bottom - rect.bottom + rect.top) // 2
            api.SetWindowPos(hwnd, None, x, y, 0, 0, 0x0015)

    details = ttk.Frame(frame)
    details.grid(row=4, column=0, sticky="ew", pady=(8, 0))
    log = tk.Text(
        details, width=50, height=8, wrap="word", state="disabled",
        font=("Consolas", 9), relief="solid", borderwidth=1,
    )
    scroll = ttk.Scrollbar(details, command=log.yview)
    log.configure(yscrollcommand=scroll.set)
    log.pack(side="left", fill="both", expand=True)
    scroll.pack(side="right", fill="y")
    details.grid_remove()

    def toggle_info():
        nonlocal info_open
        info_open = not info_open
        if info_open:
            details.grid()
            info_button.configure(text="Hide info")
            log.configure(state="normal")
            log.delete("1.0", "end")
            log.insert("end", "".join(info_lines))
            log.configure(state="disabled")
            log.see("end")
        else:
            details.grid_remove()
            info_button.configure(text="Show info")
        root.after_idle(center_convertany_window)

    info_button = ttk.Button(frame, text="Show info", command=toggle_info)
    info_button.grid(row=3, column=0, sticky="w", pady=(12, 0))
    info_button.grid_remove()

    def start_conversion():
        nonlocal busy, started, active_stage
        if source is None or picker.current() < 0 or busy:
            return

        target = formats[picker.current()]
        busy = True
        controls.grid_remove()
        status.set(f"Converting to {LABELS.get(target, target.upper())}…")
        activity.grid()
        info_button.grid()
        center_convertany_window()
        started = time.monotonic()
        active_stage = f"Converting to {LABELS[target]}"
        status.set(active_stage)
        unknown_progress()

        threading.Thread(
            target=convert_file, args=(source, target, events), daemon=True
        ).start()

    convert_button = ttk.Button(
        controls, text="Convert", command=start_conversion, state="disabled"
    )
    convert_button.grid(row=0, column=2)

    def update_window():
        nonlocal busy
        if busy:
            seconds = int(time.monotonic() - started)
            elapsed.set(f"{seconds // 60:02d}:{seconds % 60:02d}")
        for _ in range(100):
            try:
                event, value = events.get_nowait()
            except queue.Empty:
                break

            if event == "log":
                update_stage(value)
                info_lines.append(value)
                if info_open:
                    log.configure(state="normal")
                    log.insert("end", value)
                    log.configure(state="disabled")
                    log.see("end")
                continue

            busy = False
            progress_bar.stop()
            if value == 0:
                progress_bar.configure(mode="determinate", value=100)
                percent.set("100%")
                status.set("Conversion complete")
                root.after(450, root.destroy)
            else:
                status.set("Conversion failed. Open Show info for details.")
                ttk.Button(frame, text="Close", command=root.destroy).grid(
                    row=5, column=0, sticky="e", pady=(10, 0)
                )
            return

        root.after(60, update_window)

    def close_window():
        if not busy:
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", close_window)
    root.after(60, update_window)

    if len(sys.argv) == 2:
        root.after_idle(lambda: select_file(sys.argv[1]))

    def show_window():
        root.attributes("-alpha", 0)
        root.deiconify()
        center_convertany_window()
        root.attributes("-alpha", 1)

    root.after_idle(show_window)
    root.mainloop()


if __name__ == "__main__":
    main()