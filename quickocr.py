"""QuickOCR - a portable, offline OCR utility for Windows."""

import sys
import os
import json
import math
import time
import queue
import ctypes
import logging
import tempfile
import threading
import tkinter as tk
from ctypes import wintypes
from logging.handlers import RotatingFileHandler
from tkinter import messagebox
from typing import Any, Callable, Dict, List, NamedTuple, Optional, Tuple

import pytesseract
import pyperclip
from PIL import ImageGrab, Image, ImageOps
from pytesseract import Output

APP_NAME = "QuickOCR"
VERSION = "1.1.0"
CONFIG_FILENAME = "config.json"
LOG_FILENAME = "quickocr.log"

DEFAULT_LANG = "eng+fra"
OCR_CONFIG = "--psm 6"
OCR_TIMEOUT = 30
CONF_ACCEPT = 80.0          # good enough; stop trying other preprocessing variants
CONF_WARN = 60.0            # below this the read is unreliable, so say so
VARIANT_BUDGET = 4.0        # seconds; stop trying more variants past this
UPSCALE_MAX = 3.0
UPSCALE_PIXEL_CAP = 12_000_000
GRAB_DELAY_MS = 80          # let the desktop repaint after hiding the overlay
OVERLAY_POLL_MS = 400       # how often the overlay checks that it still owns the screen
OVERLAY_MAX_SECONDS = 120   # hard ceiling on how long the overlay may cover the desktop
WINDOW_W, WINDOW_H = 300, 190

THEME = {
    "bg_main":    "#1e1e1e",
    "bg_header":  "#252526",
    "fg_text":    "#cccccc",
    "accent":     "#007acc",
    "accent_hov": "#0098ff",
    "close_hov":  "#e81123",
    "min_hov":    "#3e3e42",
    "success":    "#4cc790",
    "warning":    "#e0a030",
    "instruction":"#888888"
}

GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080
SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN = 76, 77
SM_CXVIRTUALSCREEN, SM_CYVIRTUALSCREEN = 78, 79
DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4

log = logging.getLogger(APP_NAME)


def _configure_win32() -> None:
    """Declare prototypes for the user32 calls we make.

    Without argtypes ctypes passes Python ints as 32-bit C ints, which silently
    truncates window handles.
    """
    if os.name != 'nt':
        return
    try:
        user32 = ctypes.windll.user32
        user32.GetParent.argtypes = [wintypes.HWND]
        user32.GetParent.restype = wintypes.HWND
        user32.GetForegroundWindow.restype = wintypes.HWND
        user32.SetForegroundWindow.argtypes = [wintypes.HWND]
        user32.BringWindowToTop.argtypes = [wintypes.HWND]
        user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND,
                                                    ctypes.POINTER(wintypes.DWORD)]
        user32.GetWindowThreadProcessId.restype = wintypes.DWORD
        user32.AttachThreadInput.argtypes = [wintypes.DWORD, wintypes.DWORD, wintypes.BOOL]
        user32.GetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int]
        user32.SetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_long]
        user32.GetDpiForWindow.argtypes = [wintypes.HWND]
        user32.GetDpiForWindow.restype = wintypes.UINT
    except (AttributeError, OSError):
        log.warning("could not declare win32 prototypes")


def toplevel_hwnd(widget: tk.Misc) -> int:
    """The real top-level window handle.

    Tk's winfo_id() gives an inner child window; SetForegroundWindow and the extended
    window styles only work on the wrapper above it.
    """
    hwnd = widget.winfo_id()
    if os.name != 'nt':
        return hwnd
    try:
        return ctypes.windll.user32.GetParent(hwnd) or hwnd
    except (AttributeError, OSError):
        return hwnd


def force_foreground(hwnd: int) -> bool:
    """Make hwnd the foreground window, working around the foreground lock.

    Windows only lets the process that currently owns the foreground change it, so
    attach to that thread's input queue for the duration of the call. Without this the
    overlay never receives keyboard focus and its Escape binding can never fire.
    """
    if os.name != 'nt':
        return False
    try:
        user32 = ctypes.windll.user32
        if user32.SetForegroundWindow(hwnd):
            return True
        current = user32.GetForegroundWindow()
        other = user32.GetWindowThreadProcessId(current, None) if current else 0
        ours = ctypes.windll.kernel32.GetCurrentThreadId()
        attached = bool(other) and other != ours and bool(
            user32.AttachThreadInput(ours, other, True))
        try:
            user32.BringWindowToTop(hwnd)
            return bool(user32.SetForegroundWindow(hwnd))
        finally:
            if attached:
                user32.AttachThreadInput(ours, other, False)
    except (AttributeError, OSError):
        log.exception("could not take the foreground")
        return False


def user_data_dir() -> str:
    base = os.getenv('APPDATA') or os.getenv('LOCALAPPDATA') or os.path.expanduser('~')
    if not base or not os.path.isdir(base):
        base = tempfile.gettempdir()
    folder = os.path.join(base, APP_NAME)
    os.makedirs(folder, exist_ok=True)
    return folder


def setup_logging() -> None:
    log.setLevel(logging.INFO)
    try:
        handler = RotatingFileHandler(
            os.path.join(user_data_dir(), LOG_FILENAME),
            maxBytes=256 * 1024, backupCount=2, encoding='utf-8'
        )
        handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)-7s %(message)s'))
        log.addHandler(handler)
    except OSError:
        log.addHandler(logging.NullHandler())
    log.info("--- %s %s starting (python %s) ---", APP_NAME, VERSION, sys.version.split()[0])


def enable_high_dpi_awareness() -> None:
    """Per-monitor v2, so Tk coordinates stay in real screen pixels on mixed-DPI setups.

    Under the older system-DPI mode Windows bitmap-stretches our windows on any monitor
    whose scaling differs from the primary one, which makes the selection rectangle map
    to the wrong physical pixels and capture the wrong region of the screen.
    """
    if os.name != 'nt':
        return
    try:
        user32 = ctypes.windll.user32
        user32.SetProcessDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        if user32.SetProcessDpiAwarenessContext(
                ctypes.c_void_p(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)):
            log.info("dpi awareness: per-monitor v2")
            return
    except (AttributeError, OSError):
        pass
    try:
        if ctypes.windll.shcore.SetProcessDpiAwareness(2) == 0:      # per-monitor v1
            log.info("dpi awareness: per-monitor v1")
            return
    except (AttributeError, OSError):
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
        log.info("dpi awareness: system")
    except (AttributeError, OSError):
        log.warning("could not set any DPI awareness mode")


def dpi_scale(widget: tk.Misc) -> float:
    if os.name != 'nt':
        return 1.0
    try:
        dpi = ctypes.windll.user32.GetDpiForWindow(toplevel_hwnd(widget))
        if dpi:
            return dpi / 96.0
    except (AttributeError, OSError, tk.TclError):
        pass
    return 1.0


def virtual_screen() -> Tuple[int, int, int, int]:
    """(x, y, width, height) of the whole desktop, in physical pixels."""
    try:
        user32 = ctypes.windll.user32
        return (user32.GetSystemMetrics(SM_XVIRTUALSCREEN),
                user32.GetSystemMetrics(SM_YVIRTUALSCREEN),
                user32.GetSystemMetrics(SM_CXVIRTUALSCREEN),
                user32.GetSystemMetrics(SM_CYVIRTUALSCREEN))
    except (AttributeError, OSError):
        return (0, 0, 1920, 1080)


def get_resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def force_taskbar_visibility(root_window: tk.Tk):
    """Give the borderless window a normal taskbar button."""
    try:
        user32 = ctypes.windll.user32
        hwnd = toplevel_hwnd(root_window)
        style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        wanted = (style & ~WS_EX_TOOLWINDOW) | WS_EX_APPWINDOW
        if wanted == style:
            return                       # already correct; skip the hide/show flicker
        user32.SetWindowLongW(hwnd, GWL_EXSTYLE, wanted)
        root_window.withdraw()
        root_window.after(10, root_window.deiconify)
    except (AttributeError, OSError, tk.TclError):
        log.exception("could not update the taskbar style")


class ConfigManager:
    @staticmethod
    def _get_path() -> str:
        return os.path.join(user_data_dir(), CONFIG_FILENAME)

    @staticmethod
    def load() -> Dict[str, Any]:
        try:
            path = ConfigManager._get_path()
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
                log.warning("config is a %s, not an object - ignoring", type(data).__name__)
        except (OSError, ValueError, TypeError) as exc:
            log.warning("could not read config: %s", exc)
        return {}

    @staticmethod
    def save(data: Dict[str, Any]):
        try:
            with open(ConfigManager._get_path(), 'w', encoding='utf-8') as f:
                json.dump(data, f)
        except (OSError, ValueError, TypeError) as exc:
            log.warning("could not write config: %s", exc)


class OCRResult(NamedTuple):
    text: str
    confidence: float
    variant: str


class OCREngine:
    """Reads text from an image, trying several preprocessing variants.

    No single binarization wins everywhere: a fixed threshold erases low-contrast grey
    text, while Otsu misfires when the text covers only a few percent of the pixels
    (it splits the background instead). So each variant is scored with Tesseract's own
    confidence and the best one is returned.
    """

    def __init__(self, lang: str = DEFAULT_LANG):
        self.tesseract_cmd = get_resource_path(os.path.join('Tesseract-OCR', 'tesseract.exe'))
        self.tessdata_dir = get_resource_path(os.path.join('Tesseract-OCR', 'tessdata'))
        if sys.platform.startswith('win'):
            pytesseract.pytesseract.tesseract_cmd = self.tesseract_cmd
        if not os.path.exists(self.tesseract_cmd):
            log.error("bundled tesseract is missing: %s", self.tesseract_cmd)
        self.available = self.available_languages()
        self.lang = lang if self.supports(lang) else (self.available[0] if self.available else 'eng')
        log.info("languages available=%s using=%s", ','.join(self.available) or 'none', self.lang)

    def available_languages(self) -> List[str]:
        """Whatever .traineddata actually shipped, so adding a language needs no code change."""
        try:
            return sorted(name[:-len('.traineddata')]
                          for name in os.listdir(self.tessdata_dir)
                          if name.endswith('.traineddata') and not name.startswith('osd'))
        except OSError:
            log.exception("could not list %s", self.tessdata_dir)
            return []

    def supports(self, lang: str) -> bool:
        return bool(lang) and all(part in self.available for part in lang.split('+'))

    def language_choices(self) -> List[str]:
        """The default combination first, then each installed language on its own."""
        choices = [DEFAULT_LANG] if self.supports(DEFAULT_LANG) else []
        return choices + [name for name in self.available if name not in choices]

    @staticmethod
    def _upscaled_gray(img: Image.Image) -> Image.Image:
        """Upscale small snips for legibility, capped so a full-screen grab stays quick."""
        w, h = img.size
        factor = max(1.0, min(UPSCALE_MAX, math.sqrt(UPSCALE_PIXEL_CAP / max(1, w * h))))
        if factor > 1.01:
            img = img.resize((max(1, int(w * factor)), max(1, int(h * factor))),
                             Image.Resampling.LANCZOS)
        return img.convert('L')

    @staticmethod
    def _otsu_threshold(gray: Image.Image) -> int:
        hist = gray.histogram()
        total = sum(hist)
        total_sum = sum(i * hist[i] for i in range(256))
        weight_bg = sum_bg = 0
        best_variance, threshold = -1.0, 128
        for i in range(256):
            weight_bg += hist[i]
            if weight_bg == 0:
                continue
            weight_fg = total - weight_bg
            if weight_fg == 0:
                break
            sum_bg += i * hist[i]
            variance = ((sum_bg / weight_bg - (total_sum - sum_bg) / weight_fg) ** 2
                        * weight_bg * weight_fg)
            if variance > best_variance:
                best_variance, threshold = variance, i
        return threshold

    @staticmethod
    def _ink_on_white(mono: Image.Image) -> Image.Image:
        """Tesseract prefers dark text on light; text is normally the minority of pixels."""
        hist = mono.histogram()
        if hist[0] > hist[255]:
            mono = ImageOps.invert(mono)
        return mono.convert('1', dither=Image.Dither.NONE)

    @classmethod
    def _prep_otsu(cls, img: Image.Image) -> Image.Image:
        """Adaptive threshold - reads low-contrast and grey text a fixed cut destroys."""
        gray = cls._upscaled_gray(img)
        threshold = cls._otsu_threshold(gray)
        return cls._ink_on_white(gray.point(lambda x: 0 if x < threshold else 255, 'L'))

    @classmethod
    def _prep_fixed(cls, img: Image.Image) -> Image.Image:
        """High fixed cut - beats Otsu on noisy, colourful art (game menus, trading cards)."""
        gray = cls._upscaled_gray(img)
        return cls._ink_on_white(gray.point(lambda x: 0 if x > 115 else 255, 'L'))

    @classmethod
    def _prep_gray(cls, img: Image.Image) -> Image.Image:
        """No binarization - let Tesseract do its own, which copes well with gradients."""
        return ImageOps.autocontrast(cls._upscaled_gray(img), cutoff=2)

    def _run(self, prepared: Image.Image, variant: str) -> OCRResult:
        data = pytesseract.image_to_data(prepared, lang=self.lang, config=OCR_CONFIG,
                                         output_type=Output.DICT, timeout=OCR_TIMEOUT)
        lines: Dict[Tuple[int, int, int], List[str]] = {}
        confidences: List[float] = []
        for i, raw in enumerate(data['text']):
            word = (raw or '').strip()
            if not word:
                continue
            key = (data['block_num'][i], data['par_num'][i], data['line_num'][i])
            lines.setdefault(key, []).append(word)
            confidence = float(data['conf'][i])
            if confidence >= 0:
                confidences.append(confidence)
        text = '\n'.join(' '.join(words) for _, words in sorted(lines.items()))
        mean = sum(confidences) / len(confidences) if confidences else 0.0
        return OCRResult(text, mean, variant)

    def extract_text(self, img: Image.Image) -> OCRResult:
        best = OCRResult('', 0.0, 'none')
        started = time.monotonic()
        last_took = 0.0
        for variant, prepare in (('otsu', self._prep_otsu),
                                 ('fixed', self._prep_fixed),
                                 ('gray', self._prep_gray)):
            elapsed = time.monotonic() - started
            # variants cost about the same, so use the last one to predict the next
            if elapsed + last_took > VARIANT_BUDGET:
                log.info("stopping after %.2fs, another variant would overrun the budget", elapsed)
                break
            attempt = time.monotonic()
            try:
                result = self._run(prepare(img), variant)
            except (pytesseract.TesseractError, RuntimeError, OSError, ValueError):
                log.exception("variant %s failed", variant)
                continue
            finally:
                last_took = time.monotonic() - attempt
            log.info("variant=%s conf=%.1f chars=%d in %.2fs",
                     variant, result.confidence, len(result.text), last_took)
            if result.confidence > best.confidence:
                best = result
            if best.confidence >= CONF_ACCEPT:
                break
        log.info("best variant=%s conf=%.1f in %.2fs", best.variant, best.confidence,
                 time.monotonic() - started)
        return best


class SnippingOverlay(tk.Toplevel):
    """Full-desktop dimmed overlay used to draw the capture rectangle.

    Every exit path funnels through _finish(), which is idempotent and always both
    destroys the overlay and calls back. An overlay that cannot be dismissed covers
    the whole desktop while the main window is hidden, which looks like a system hang.
    """

    def __init__(self, parent, on_complete: Callable[[Optional[Image.Image]], None]):
        super().__init__(parent)
        self._on_complete: Optional[Callable] = on_complete
        self._finished = False
        self._capturing = False
        self._lost_foreground = 0
        self._watchdog_id: Optional[str] = None
        self._opened_at = time.monotonic()
        self.start_pos: Optional[Tuple[float, float]] = None
        self.rect = None

        self.v_x, self.v_y, self.v_width, self.v_height = virtual_screen()
        self.geometry(f"{self.v_width}x{self.v_height}+{self.v_x}+{self.v_y}")
        self.overrideredirect(True)
        self.attributes('-alpha', 0.3)
        self.attributes('-topmost', True)
        self.configure(bg="black")

        self.canvas = tk.Canvas(self, cursor="cross", bg="grey11", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.create_text(
            self.v_width // 2, 40,
            text="Drag to select text  -  click once, right-click or press Esc to cancel",
            fill="#ffffff", font=("Segoe UI", 14)
        )

        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        # bound on the toplevel, so they fire for events anywhere inside the overlay
        self.bind('<Escape>', self._on_cancel)
        self.bind('<ButtonPress-3>', self._on_cancel)
        self.bind('<Destroy>', self._on_destroy)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        self._claim_focus()
        self._watchdog_id = self.after(OVERLAY_POLL_MS, self._watchdog)

    def _is_foreground(self) -> bool:
        if os.name != 'nt':
            return True
        try:
            return ctypes.windll.user32.GetForegroundWindow() == toplevel_hwnd(self)
        except (AttributeError, OSError, tk.TclError):
            return True

    def _watchdog(self):
        """Never let a topmost, full-desktop window outlive our ownership of the screen.

        If anything else comes to the front - Task Manager after Ctrl+Alt+Del, say - a
        topmost overlay would hide it and the desktop would look frozen with no way out.
        """
        self._watchdog_id = None
        if self._finished or self._capturing:
            return

        if time.monotonic() - self._opened_at > OVERLAY_MAX_SECONDS:
            log.warning("overlay hit its %ds ceiling; cancelling", OVERLAY_MAX_SECONDS)
            self._finish(None)
            return

        if self._is_foreground():
            self._lost_foreground = 0
        else:
            self._lost_foreground += 1
            if self._lost_foreground == 1:
                try:
                    self.attributes('-topmost', False)   # stop covering whatever took over
                except tk.TclError:
                    pass
            elif self._lost_foreground >= 2:
                log.warning("overlay lost the foreground; cancelling so nothing is trapped")
                self._finish(None)
                return

        self._watchdog_id = self.after(OVERLAY_POLL_MS, self._watchdog)

    def _stop_watchdog(self):
        if self._watchdog_id is not None:
            try:
                self.after_cancel(self._watchdog_id)
            except (tk.TclError, ValueError):
                pass
            self._watchdog_id = None

    def _claim_focus(self):
        """Without being the foreground window, key bindings such as Escape never fire."""
        self.update_idletasks()
        try:
            if not force_foreground(toplevel_hwnd(self)):
                log.warning("overlay did not take the foreground; Escape may not respond")
        except tk.TclError:
            pass
        try:
            self.focus_force()
            self.canvas.focus_set()
            self.grab_set()
        except tk.TclError:
            log.exception("could not focus the overlay")

    def _on_press(self, event):
        self.start_pos = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        self.rect = self.canvas.create_rectangle(
            *self.start_pos, *self.start_pos,
            outline=THEME['accent'], width=2
        )

    def _on_drag(self, event):
        if self.rect is None or self.start_pos is None:
            return
        cur_x, cur_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_pos[0], self.start_pos[1], cur_x, cur_y)

    def _on_release(self, event):
        if self.start_pos is None:
            self._finish(None)           # a release with no press must not strand the overlay
            return

        x1, y1 = self.start_pos
        x2, y2 = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        x_min, x_max = min(x1, x2), max(x1, x2)
        y_min, y_max = min(y1, y2), max(y1, y2)

        if (x_max - x_min) < 5 or (y_max - y_min) < 5:
            self._finish(None)
            return

        # overlay coordinates are physical pixels relative to the virtual screen origin
        capture_bbox = (
            int(x_min + self.v_x),
            int(y_min + self.v_y),
            int(x_max + self.v_x),
            int(y_max + self.v_y)
        )
        self._hide_then_grab(capture_bbox)

    def _hide_then_grab(self, capture_bbox: Tuple[int, int, int, int]):
        """Hide first: this overlay is a layered window and is composited into the grab,
        which would darken the snip and skew the OCR thresholds."""
        self._capturing = True                  # the watchdog must not fire while we are hidden
        self._stop_watchdog()
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.withdraw()
        self.update_idletasks()
        self.after(GRAB_DELAY_MS, lambda: self._grab(capture_bbox))

    def _grab(self, capture_bbox: Tuple[int, int, int, int]):
        image, error = None, None
        try:
            image = ImageGrab.grab(bbox=capture_bbox, all_screens=True)
            log.info("captured bbox=%s size=%s", capture_bbox, image.size)
        except (OSError, ValueError) as exc:
            log.exception("screen capture failed")
            error = exc
        self._finish(image)                     # tear down before any modal dialog
        if error is not None:
            messagebox.showerror(f"{APP_NAME} - Capture Error", str(error))

    def _on_cancel(self, event=None):
        self._finish(None)

    def _on_destroy(self, event):
        if event.widget is self:
            self._finish(None)           # last resort, so the caller is never left hanging

    def _finish(self, image: Optional[Image.Image]):
        if self._finished:
            return
        self._finished = True
        self._stop_watchdog()
        callback, self._on_complete = self._on_complete, None
        try:
            self.attributes('-topmost', False)
        except tk.TclError:
            pass
        try:
            self.grab_release()
        except tk.TclError:
            pass
        try:
            if self.winfo_exists():
                self.destroy()
        except tk.TclError:
            pass
        if callback:
            callback(image)


class ResultPopup(tk.Toplevel):
    def __init__(self, parent, result: OCRResult, timeout: int = 10):
        super().__init__(parent)
        self.overrideredirect(True)
        self.configure(bg=THEME["bg_main"])
        self.attributes('-topmost', True)

        scale = dpi_scale(self)
        w, h = int(400 * scale), int(260 * scale)
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

        unreliable = result.confidence < CONF_WARN
        frame = tk.Frame(self, bg=THEME["bg_main"],
                         highlightbackground=THEME["warning"] if unreliable else THEME["accent"],
                         highlightthickness=1)
        frame.pack(fill=tk.BOTH, expand=True)

        if unreliable:
            headline, colour = "! COPIED - LOW CONFIDENCE", THEME["warning"]
            subtitle = (f"Tesseract is only {result.confidence:.0f}% sure. "
                        "Check the text before you use it.")
        else:
            headline, colour = "✓ COPIED TO CLIPBOARD", THEME["success"]
            subtitle = "Press Ctrl+V to paste"

        tk.Label(frame, text=headline, font=("Segoe UI", 11, "bold"),
                 bg=THEME["bg_main"], fg=colour).pack(pady=(15, 2))
        tk.Label(frame, text=subtitle, font=("Segoe UI", 9), wraplength=int(360 * scale),
                 bg=THEME["bg_main"], fg=THEME["instruction"]).pack(pady=(0, 10))

        preview = result.text.replace('\n', ' ')
        if len(preview) > 150:
            preview = preview[:150] + "..."

        tk.Label(
            frame, text=preview, font=("Consolas", 9),
            bg=THEME["bg_header"], fg=THEME["fg_text"],
            wraplength=int(360 * scale), justify="left", padx=10, pady=10
        ).pack(fill=tk.X, padx=20)

        self.lbl_timer = tk.Label(
            frame, text=f"Auto-closing in {timeout}s",
            font=("Segoe UI", 8), bg=THEME["bg_main"], fg="#666"
        )
        self.lbl_timer.pack(pady=(10, 5))

        tk.Button(
            frame, text="OK", command=self.destroy,
            bg=THEME["accent"], fg="white",
            activebackground=THEME["accent_hov"], activeforeground="white",
            bd=0, padx=25, pady=4, cursor="hand2"
        ).pack(pady=10)

        self._start_timer(timeout)

    def _start_timer(self, remaining: int):
        if not self.winfo_exists():
            return
        if remaining > 0:
            self.lbl_timer.config(text=f"Closing in {remaining}s")
            self.after(1000, lambda: self._start_timer(remaining - 1))
        else:
            self.destroy()


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.configure(bg=THEME["bg_main"])
        self.root.report_callback_exception = self._on_tk_error

        try:
            icon_path = get_resource_path("aa.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except tk.TclError:
            log.exception("could not load the window icon")

        self.config = ConfigManager.load()
        saved_lang = self.config.get("lang")
        self.ocr = OCREngine(saved_lang if isinstance(saved_lang, str) else DEFAULT_LANG)
        self.results: "queue.Queue[Tuple[str, Any]]" = queue.Queue()
        self.busy = False

        self._setup_window_geometry()
        self._build_custom_titlebar()
        self._build_main_ui()

        self.root.after(10, lambda: force_taskbar_visibility(self.root))

    def _on_tk_error(self, exc_type, exc_value, exc_tb):
        # in a --noconsole build these would otherwise vanish silently
        log.error("unhandled callback error", exc_info=(exc_type, exc_value, exc_tb))
        messagebox.showerror(APP_NAME, f"Unexpected error:\n{exc_value}")

    def _setup_window_geometry(self):
        scale = dpi_scale(self.root)
        try:
            self.root.tk.call('tk', 'scaling', 96.0 * scale / 72.0)
        except tk.TclError:
            pass

        w, h = int(WINDOW_W * scale), int(WINDOW_H * scale)
        v_x, v_y, v_width, v_height = virtual_screen()
        x, y = self.config.get("x"), self.config.get("y")
        if not isinstance(x, (int, float)) or not isinstance(y, (int, float)):
            x, y = v_x + (v_width - w) // 2, v_y + (v_height - h) // 2
        # a position saved on a monitor that is now unplugged would be unreachable
        x = max(v_x, min(int(x), v_x + v_width - w))
        y = max(v_y, min(int(y), v_y + v_height - h))
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _build_custom_titlebar(self):
        self.title_bar = tk.Frame(self.root, bg=THEME["bg_header"], height=30)
        self.title_bar.pack(fill=tk.X, side=tk.TOP)
        self.title_bar.pack_propagate(False)

        self.title_bar.bind("<Button-1>", self._start_move)
        self.title_bar.bind("<B1-Motion>", self._do_move)

        tk.Label(
            self.title_bar, text=f"{APP_NAME}",
            bg=THEME["bg_header"], fg=THEME["fg_text"], font=("Segoe UI", 9, "bold")
        ).pack(side=tk.LEFT, padx=10)

        self._create_titlebar_btn("✕", self._on_close, THEME["close_hov"])
        self._create_titlebar_btn("—", self._minimize, THEME["min_hov"])

    def _create_titlebar_btn(self, text, command, hover_color):
        btn = tk.Button(
            self.title_bar, text=text, command=command,
            bg=THEME["bg_header"], fg=THEME["fg_text"],
            bd=0, font=("Arial", 9, "bold"), width=4,
            activebackground=hover_color, activeforeground="white"
        )
        btn.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_main_ui(self):
        main_frame = tk.Frame(self.root, bg=THEME["bg_main"])
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.bind("<Button-1>", self._start_move)
        main_frame.bind("<B1-Motion>", self._do_move)

        self.capture_btn = tk.Button(
            main_frame, text="CAPTURE ZONE", command=self._start_snip,
            font=("Segoe UI", 10, "bold"), bg=THEME["accent"], fg="white",
            activebackground=THEME["accent_hov"], activeforeground="white",
            bd=0, cursor="hand2", padx=20, pady=10
        )
        self.capture_btn.place(relx=0.5, rely=0.42, anchor=tk.CENTER)

        self._build_language_picker(main_frame)

    def _build_language_picker(self, parent):
        choices = self.ocr.language_choices()
        row = tk.Frame(parent, bg=THEME["bg_main"])
        row.place(relx=0.5, rely=0.82, anchor=tk.CENTER)

        tk.Label(row, text="Language", font=("Segoe UI", 8), bg=THEME["bg_main"],
                 fg=THEME["instruction"]).pack(side=tk.LEFT, padx=(0, 6))

        if not choices:
            tk.Label(row, text="no language data found", font=("Segoe UI", 8, "bold"),
                     bg=THEME["bg_main"], fg=THEME["warning"]).pack(side=tk.LEFT)
            return

        self.lang_var = tk.StringVar(value=self.ocr.lang)
        picker = tk.OptionMenu(row, self.lang_var, *choices, command=self._on_language_change)
        picker.config(font=("Segoe UI", 8), bg=THEME["bg_header"], fg=THEME["fg_text"],
                      activebackground=THEME["min_hov"], activeforeground="white",
                      bd=0, highlightthickness=0, cursor="hand2", padx=8, pady=2)
        picker["menu"].config(bg=THEME["bg_header"], fg=THEME["fg_text"],
                              activebackground=THEME["accent"], bd=0)
        picker.pack(side=tk.LEFT)

    def _on_language_change(self, value: str):
        self.ocr.lang = value
        self.config["lang"] = value
        ConfigManager.save(self.config)
        log.info("language set to %s", value)

    def _start_move(self, event):
        self.last_x = event.x
        self.last_y = event.y

    def _do_move(self, event):
        deltax = event.x - self.last_x
        deltay = event.y - self.last_y
        x = self.root.winfo_x() + deltax
        y = self.root.winfo_y() + deltay
        self.root.geometry(f"+{x}+{y}")

    def _minimize(self):
        self.root.overrideredirect(False)
        self.root.iconify()
        self.root.bind("<Map>", self._restore_window)

    def _restore_window(self, event):
        if event.widget is not self.root:
            return                       # child widgets map too; only the window matters
        if self.root.state() == 'normal':
            self.root.overrideredirect(True)
            self.root.unbind("<Map>")
            force_taskbar_visibility(self.root)

    def _on_close(self):
        self.config["x"] = self.root.winfo_x()
        self.config["y"] = self.root.winfo_y()
        self.config["lang"] = self.ocr.lang
        self.config["version"] = VERSION
        ConfigManager.save(self.config)
        log.info("--- shutting down ---")
        self.root.destroy()

    def _set_busy(self, busy: bool):
        self.busy = busy
        self.capture_btn.config(
            text="READING..." if busy else "CAPTURE ZONE",
            state=tk.DISABLED if busy else tk.NORMAL,
            cursor="watch" if busy else "hand2"
        )

    def _start_snip(self):
        if self.busy:
            return
        self.root.withdraw()
        self.root.after(150, self._open_overlay)

    def _open_overlay(self):
        try:
            SnippingOverlay(self.root, self._process_snip)
        except tk.TclError:
            log.exception("could not open the capture overlay")
            self._restore_main_window()
            messagebox.showerror(APP_NAME, "Could not open the capture overlay.")

    def _restore_main_window(self):
        self.root.deiconify()
        self.root.overrideredirect(True)
        force_taskbar_visibility(self.root)

    def _process_snip(self, img: Optional[Image.Image]):
        self._restore_main_window()
        if img is None:
            return
        self._set_busy(True)
        threading.Thread(target=self._ocr_worker, args=(img,), daemon=True).start()
        self.root.after(100, self._poll_ocr)

    def _ocr_worker(self, img: Image.Image):
        # OCR takes seconds on a large capture; on the UI thread that freezes the app
        try:
            self.results.put(('ok', self.ocr.extract_text(img)))
        except Exception as exc:                     # noqa: BLE001 - surfaced to the user
            log.exception("OCR failed")
            self.results.put(('error', exc))

    def _poll_ocr(self):
        try:
            kind, payload = self.results.get_nowait()
        except queue.Empty:
            self.root.after(100, self._poll_ocr)
            return

        self._set_busy(False)
        if kind == 'error':
            messagebox.showerror(APP_NAME, f"Could not read the image:\n{payload}")
            return
        self._deliver(payload)

    def _deliver(self, result: OCRResult):
        if not result.text.strip():
            log.info("no text found")
            messagebox.showwarning(APP_NAME, "No text found.")
            return
        error = self._copy_to_clipboard(result.text)
        if error:
            messagebox.showerror(APP_NAME, f"Could not copy to the clipboard:\n{error}")
            return
        ResultPopup(self.root, result)

    @staticmethod
    def _copy_to_clipboard(text: str) -> Optional[str]:
        """Copy, then read back to confirm. Another app can own or wipe the clipboard, and
        claiming success without checking is how text goes missing after a capture."""
        last_error = "the clipboard did not accept the text"
        for attempt in range(3):
            try:
                pyperclip.copy(text)
                if pyperclip.paste() == text:
                    return None
                log.warning("clipboard read-back mismatch on attempt %d", attempt + 1)
            except Exception as exc:                 # noqa: BLE001 - clipboard can be locked
                last_error = str(exc) or exc.__class__.__name__
                log.warning("clipboard attempt %d failed: %s", attempt + 1, last_error)
            time.sleep(0.12)
        log.error("giving up on the clipboard: %s", last_error)
        return last_error

    def run(self):
        self.root.mainloop()


def main():
    setup_logging()
    _configure_win32()
    enable_high_dpi_awareness()
    try:
        App().run()
    except Exception:                                # noqa: BLE001 - last chance to log
        log.exception("fatal error")
        raise


if __name__ == "__main__":
    main()
