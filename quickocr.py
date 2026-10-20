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
import tkinter.font as tkfont
from ctypes import wintypes
from logging.handlers import RotatingFileHandler
from tkinter import messagebox
from typing import Any, Callable, Dict, List, NamedTuple, Optional, Tuple

import pytesseract
import pyperclip
from PIL import ImageDraw, ImageGrab, Image, ImageOps, ImageTk
from pytesseract import Output

APP_NAME = "QuickOCR"
VERSION = "1.2.0"
ICON_FILENAME = "quickocr.ico"
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
WINDOW_W, WINDOW_H = 340, 236
POPUP_W, POPUP_H = 420, 252
POPUP_SECONDS = 10

THEME = {
    "bg":           "#18181b",
    "surface":      "#232329",
    "surface_hov":  "#2d2d35",
    "border":       "#34343d",
    "fg":           "#ececf1",
    "muted":        "#9d9da8",
    "subtle":       "#6c6c77",
    "accent":       "#3b82f6",
    "accent_hov":   "#5593ff",
    "accent_press": "#2a6ce0",
    "accent_off":   "#2a3f63",
    "close_hov":    "#e81123",
    "success":      "#34d399",
    "warning":      "#f5a524",
}

FONT_UI = "Segoe UI"
FONT_SEMIBOLD = "Segoe UI Semibold"
FONT_MONO = "Consolas"
FONT_GLYPH = "Segoe MDL2 Assets"     # ships with Windows 10 and 11
GLYPH_CLOSE, GLYPH_MINIMIZE, GLYPH_CHEVRON = "", "", ""
GLYPH_CHECK, GLYPH_WARNING, GLYPH_CAPTURE = "", "", ""

GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080
SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN = 76, 77
SM_CXVIRTUALSCREEN, SM_CYVIRTUALSCREEN = 78, 79
DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
MONITOR_DEFAULTTONEAREST = 2
DWMWA_WINDOW_CORNER_PREFERENCE, DWMWCP_ROUND = 33, 2
DWMWA_BORDER_COLOR = 34

log = logging.getLogger(APP_NAME)


class MONITORINFO(ctypes.Structure):
    _fields_ = [('cbSize', wintypes.DWORD),
                ('rcMonitor', wintypes.RECT),
                ('rcWork', wintypes.RECT),
                ('dwFlags', wintypes.DWORD)]


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
        user32.GetCursorPos.argtypes = [ctypes.POINTER(wintypes.POINT)]
        user32.MonitorFromPoint.argtypes = [wintypes.POINT, wintypes.DWORD]
        user32.MonitorFromPoint.restype = wintypes.HMONITOR
        user32.GetMonitorInfoW.argtypes = [wintypes.HMONITOR, ctypes.POINTER(MONITORINFO)]
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


def monitor_at_cursor(work_area: bool = False) -> Tuple[int, int, int, int]:
    """(left, top, right, bottom) of the monitor under the mouse, in physical pixels.

    The virtual desktop is the bounding box of every monitor, so its centre can fall on a
    bezel and its top edge can sit above a shorter monitor - nowhere anyone can see.
    """
    if os.name == 'nt':
        try:
            user32 = ctypes.windll.user32
            point = wintypes.POINT()
            if user32.GetCursorPos(ctypes.byref(point)):
                monitor = user32.MonitorFromPoint(point, MONITOR_DEFAULTTONEAREST)
                info = MONITORINFO(cbSize=ctypes.sizeof(MONITORINFO))
                if monitor and user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
                    rect = info.rcWork if work_area else info.rcMonitor
                    return (rect.left, rect.top, rect.right, rect.bottom)
        except (AttributeError, OSError, ctypes.ArgumentError):
            log.exception("could not find the monitor under the cursor")
    x, y, width, height = virtual_screen()
    return (x, y, x + width, y + height)


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


def apply_window_chrome(window: tk.Misc) -> bool:
    """Rounded corners and a hairline border on Windows 11, like a native window.

    A borderless window otherwise has square corners and no edge, which reads as unfinished.
    Returns False where Windows does not support it (Windows 10), so the caller can keep
    drawing its own 1px border instead.
    """
    if os.name != 'nt':
        return False
    try:
        window.update_idletasks()
        hwnd = wintypes.HWND(toplevel_hwnd(window))
        dwm = ctypes.windll.dwmapi
        corner = ctypes.c_int(DWMWCP_ROUND)
        rounded = dwm.DwmSetWindowAttribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE,
                                            ctypes.byref(corner), ctypes.sizeof(corner))
        red, green, blue = (int(THEME["border"][i:i + 2], 16) for i in (1, 3, 5))
        colour = wintypes.DWORD(red | green << 8 | blue << 16)       # COLORREF is 0x00BBGGRR
        bordered = dwm.DwmSetWindowAttribute(hwnd, DWMWA_BORDER_COLOR,
                                             ctypes.byref(colour), ctypes.sizeof(colour))
        return rounded == 0 and bordered == 0                         # S_OK
    except (AttributeError, OSError, tk.TclError):
        log.info("window chrome not available on this system")
        return False


def rounded_image(master: tk.Misc, width: int, height: int, radius: float, fill: str,
                  background: str) -> ImageTk.PhotoImage:
    """An anti-aliased rounded rectangle. Tk's canvas draws curves with jagged edges."""
    k = 4
    image = Image.new('RGB', (width * k, height * k), background)
    ImageDraw.Draw(image).rounded_rectangle(
        (0, 0, width * k - 1, height * k - 1), radius=radius * k, fill=fill)
    return ImageTk.PhotoImage(image.resize((width, height), Image.Resampling.LANCZOS),
                              master=master)


def icon_image(master: tk.Misc, size: int) -> Optional[ImageTk.PhotoImage]:
    try:
        with Image.open(get_resource_path(ICON_FILENAME)) as icon:
            icon.size = max(icon.info.get('sizes', {icon.size}))
            icon.load()
            frame = icon.convert('RGBA').resize((size, size), Image.Resampling.LANCZOS)
        return ImageTk.PhotoImage(frame, master=master)
    except (OSError, ValueError):
        log.exception("could not load the icon image")
        return None


class RoundedButton(tk.Canvas):
    """A flat button with rounded corners and hover, press and disabled states."""

    def __init__(self, parent: tk.Misc, text: str, command: Callable[[], None],
                 width: int, height: int, background: str, glyph: str = "",
                 primary: bool = True, font_size: int = 10, glyph_after: bool = False):
        super().__init__(parent, width=width, height=height, bg=background,
                         highlightthickness=0, bd=0, cursor="hand2")
        self._command = command
        self._glyph_after = glyph_after
        self._enabled = True
        self._hover = self._pressed = False
        self._width, self._height = width, height
        if primary:
            palette = (THEME["accent"], THEME["accent_hov"], THEME["accent_press"],
                       THEME["accent_off"])
            self._fg, self._fg_off = "#ffffff", "#9fb3d6"
        else:
            palette = (THEME["surface"], THEME["surface_hov"], THEME["border"],
                       THEME["surface"])
            self._fg, self._fg_off = THEME["fg"], THEME["subtle"]
        radius = min(height / 2, 8 * height / 36)
        self._images = {state: rounded_image(self, width, height, radius, colour, background)
                        for state, colour in zip(('normal', 'hover', 'pressed', 'disabled'),
                                                 palette)}
        self._shape = self.create_image(0, 0, anchor='nw', image=self._images['normal'])

        self._font = tkfont.Font(self, family=FONT_SEMIBOLD, size=font_size)
        self._glyph_font = tkfont.Font(self, family=FONT_GLYPH,
                                       size=font_size - 2 if glyph_after else font_size + 1)
        self._glyph = self.create_text(0, height / 2, text=glyph, anchor='w',
                                       font=self._glyph_font, fill=self._fg)
        self._label = self.create_text(0, height / 2, text=text, anchor='w',
                                       font=self._font, fill=self._fg)
        self._layout()

        self.bind("<Enter>", lambda _e: self._set(hover=True))
        self.bind("<Leave>", lambda _e: self._set(hover=False))
        self.bind("<ButtonPress-1>", lambda _e: self._set(pressed=True))
        self.bind("<ButtonRelease-1>", self._on_release)

    def _layout(self):
        glyph = self.itemcget(self._glyph, 'text')
        text_w = self._font.measure(self.itemcget(self._label, 'text'))
        glyph_w = self._glyph_font.measure(glyph) if glyph else 0
        gap = int(self._font.measure(' ') * 1.6) if glyph else 0
        x = (self._width - glyph_w - gap - text_w) / 2
        if self._glyph_after:
            self.coords(self._label, x, self._height / 2)
            self.coords(self._glyph, x + text_w + gap, self._height / 2 + 1)
        else:
            self.coords(self._glyph, x, self._height / 2 + 1)
            self.coords(self._label, x + glyph_w + gap, self._height / 2)

    def _set(self, hover=None, pressed=None):
        if hover is not None:
            self._hover = hover
        if pressed is not None:
            self._pressed = pressed
        if not self._enabled:
            state = 'disabled'
        elif self._pressed and self._hover:
            state = 'pressed'
        else:
            state = 'hover' if self._hover else 'normal'
        self.itemconfigure(self._shape, image=self._images[state])
        colour = self._fg if self._enabled else self._fg_off
        self.itemconfigure(self._label, fill=colour)
        self.itemconfigure(self._glyph, fill=colour)

    def _on_release(self, _event):
        fire = self._pressed and self._hover and self._enabled
        self._set(pressed=False)
        if fire:
            self._command()

    def set_state(self, text: Optional[str] = None, enabled: bool = True,
                  glyph: Optional[str] = None):
        if text is not None:
            self.itemconfigure(self._label, text=text)
        if glyph is not None:
            self.itemconfigure(self._glyph, text=glyph)
        self._enabled = enabled
        self.configure(cursor="hand2" if enabled else "watch")
        self._layout()
        self._set()


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
        """Write a temp file and swap it in, so a failed or interrupted save keeps the old
        config instead of leaving it truncated."""
        tmp_path = None
        try:
            path = ConfigManager._get_path()
            tmp_path = path + '.tmp'
            with open(tmp_path, 'w', encoding='utf-8') as f:
                json.dump(data, f)
            os.replace(tmp_path, path)
        except (OSError, ValueError, TypeError) as exc:
            log.warning("could not write config: %s", exc)
            if tmp_path:
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass


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
        # Tesseract reads its language data from TESSDATA_PREFIX before its own folder, so a
        # separate Tesseract install would otherwise hijack the bundled data: eng+fra quietly
        # loses French, and any language missing from that other folder fails outright.
        stray = os.environ.get('TESSDATA_PREFIX')
        if stray and os.path.normcase(os.path.abspath(stray)) != os.path.normcase(self.tessdata_dir):
            log.info("ignoring TESSDATA_PREFIX=%s in favour of the bundled data", stray)
        os.environ['TESSDATA_PREFIX'] = self.tessdata_dir
        self.available = self.available_languages()
        if self.supports(DEFAULT_LANG):
            fallback = DEFAULT_LANG
        else:
            fallback = self.available[0] if self.available else 'eng'
        self.lang = lang if self.supports(lang) else fallback
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
        """Raises the last error if every variant failed - a missing tesseract.exe or
        language file must not be reported to the user as "No text found"."""
        best = OCRResult('', 0.0, 'none')
        started = time.monotonic()
        last_took = 0.0
        last_error: Optional[Exception] = None
        succeeded = False
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
            except (pytesseract.TesseractError, RuntimeError, OSError, ValueError) as exc:
                log.exception("variant %s failed", variant)
                last_error = exc
                continue
            finally:
                last_took = time.monotonic() - attempt
            succeeded = True
            log.info("variant=%s conf=%.1f chars=%d in %.2fs",
                     variant, result.confidence, len(result.text), last_took)
            # on equal confidence, some text beats none
            if ((result.confidence, bool(result.text.strip()))
                    > (best.confidence, bool(best.text.strip()))):
                best = result
            if best.confidence >= CONF_ACCEPT:
                break
        if not succeeded and last_error is not None:
            raise last_error
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
        # on the monitor the user is looking at; canvas coordinates start at the virtual origin
        left, top, right, _ = monitor_at_cursor()
        self.canvas.create_text(
            (left + right) // 2 - self.v_x, top - self.v_y + 40,
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


LANGUAGE_NAMES = {
    'eng': "English", 'fra': "French", 'deu': "German", 'spa': "Spanish", 'ita': "Italian",
    'por': "Portuguese", 'nld': "Dutch", 'pol': "Polish", 'rus': "Russian", 'ukr': "Ukrainian",
    'tur': "Turkish", 'swe': "Swedish", 'dan': "Danish", 'nor': "Norwegian", 'fin': "Finnish",
    'ces': "Czech", 'ell': "Greek", 'ara': "Arabic", 'heb': "Hebrew", 'hin': "Hindi",
    'jpn': "Japanese", 'kor': "Korean", 'chi_sim': "Chinese (Simplified)",
    'chi_tra': "Chinese (Traditional)", 'vie': "Vietnamese",
}


def language_label(lang: str) -> str:
    """'eng+fra' -> 'English + French'; unknown codes are shown as they are."""
    return " + ".join(LANGUAGE_NAMES.get(part, part) for part in lang.split('+'))


def blend(colour: str, background: str, amount: float) -> str:
    """colour laid over background at the given opacity."""
    fg = [int(colour[i:i + 2], 16) for i in (1, 3, 5)]
    bg = [int(background[i:i + 2], 16) for i in (1, 3, 5)]
    return '#%02x%02x%02x' % tuple(round(b + (f - b) * amount) for f, b in zip(fg, bg))


class Dropdown(tk.Toplevel):
    """A themed list of choices under a button. Windows draws tk.Menu natively - white
    border, system colours - which clashes with the rest of the window."""

    def __init__(self, anchor: tk.Misc, choices: List[Tuple[str, str]], current: str,
                 on_pick: Callable[[str], None], scale: float):
        super().__init__(anchor)
        self.withdraw()
        self.overrideredirect(True)
        self.attributes('-topmost', True)
        self.configure(bg=THEME["border"])
        self._on_pick = on_pick
        surface, padding = THEME["surface"], int(round(4 * scale))

        body = tk.Frame(self, bg=surface)
        body.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        inner = tk.Frame(body, bg=surface)
        inner.pack(fill=tk.BOTH, expand=True, padx=padding, pady=padding)
        for value, label in choices:
            self._add_row(inner, value, label, value == current, scale)

        self.update_idletasks()
        width = max(self.winfo_reqwidth(), anchor.winfo_width())
        self.geometry(f"{width}x{self.winfo_reqheight()}"
                      f"+{anchor.winfo_rootx() + anchor.winfo_width() - width}"
                      f"+{anchor.winfo_rooty() + anchor.winfo_height() + padding}")
        self.deiconify()
        if apply_window_chrome(self):
            body.pack_configure(padx=0, pady=0)
        self.bind("<Escape>", lambda _e: self.destroy())
        self.bind("<FocusOut>", self._on_focus_out)
        self.focus_force()

    def _add_row(self, parent: tk.Misc, value: str, label: str, selected: bool, scale: float):
        surface = THEME["surface"]
        row = tk.Frame(parent, bg=surface, cursor="hand2")
        row.pack(fill=tk.X)
        check = tk.Label(row, text=GLYPH_CHECK if selected else "", font=(FONT_GLYPH, 8),
                         width=2, bg=surface, fg=THEME["accent"])
        check.pack(side=tk.LEFT, padx=(int(6 * scale), 0))
        text = tk.Label(row, text=label, font=(FONT_UI, 9), anchor='w', bg=surface,
                        fg=THEME["fg"], padx=int(6 * scale), pady=int(5 * scale))
        text.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, int(18 * scale)))
        parts = (row, check, text)

        def paint(colour):
            for part in parts:
                part.configure(bg=colour)
        for part in parts:
            part.bind("<Enter>", lambda _e: paint(THEME["surface_hov"]))
            part.bind("<Leave>", lambda _e: paint(surface))
            part.bind("<ButtonRelease-1>", lambda _e: self._pick(value))

    def _pick(self, value: str):
        self.destroy()
        self._on_pick(value)

    def _on_focus_out(self, _event):
        # focus moving between our own rows also fires FocusOut; close only once it has
        # gone somewhere else - the main window, another app, the desktop
        self.after(1, self._close_unless_focused)

    def _close_unless_focused(self):
        try:
            if not self.winfo_exists():
                return
            focused = self.focus_get()
            if focused is None or focused.winfo_toplevel() is not self:
                self.destroy()
        except (tk.TclError, KeyError):
            self.destroy()


class ResultPopup(tk.Toplevel):
    """Confirms the copy and previews the text, then gets out of the way by itself.

    The countdown pauses while the pointer is over the popup, so it never closes on
    someone who is still reading it.
    """

    PREVIEW_LINES = 4
    TICK_MS = 50

    def __init__(self, parent, result: OCRResult, timeout: int = POPUP_SECONDS):
        super().__init__(parent)
        self.withdraw()                      # build first, then show at the right size
        self.overrideredirect(True)
        self.attributes('-topmost', True)
        self.configure(bg=THEME["border"])
        self._timeout = float(timeout)
        self._elapsed = 0.0

        scale = dpi_scale(self)

        def px(value: float) -> int:
            return int(round(value * scale))

        bg = THEME["bg"]
        width = px(POPUP_W)
        self.body = tk.Frame(self, bg=bg)
        self.body.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        if result.confidence < CONF_WARN:
            glyph, colour = GLYPH_WARNING, THEME["warning"]
            headline = "Copied, but check the text"
            subtitle = f"Tesseract is only {result.confidence:.0f}% sure of this read."
        else:
            glyph, colour = GLYPH_CHECK, THEME["success"]
            headline = "Copied to clipboard"
            subtitle = "Press Ctrl+V to paste it anywhere."

        # countdown bar along the bottom edge, packed first so it stays lowest
        self.bar = tk.Canvas(self.body, height=px(3), bg=bg, highlightthickness=0, bd=0)
        self.bar.pack(side=tk.BOTTOM, fill=tk.X)
        self._bar_fill = self.bar.create_rectangle(0, 0, width, px(3), width=0,
                                                   fill=THEME["accent"])

        footer = tk.Frame(self.body, bg=bg)
        footer.pack(side=tk.BOTTOM, fill=tk.X, padx=px(20), pady=(px(14), px(14)))
        RoundedButton(footer, "Done", self.destroy, width=px(88), height=px(32),
                      background=bg, font_size=9).pack(side=tk.RIGHT)
        words = len(result.text.split())
        tk.Label(footer, text=f"{words} word{'s' if words != 1 else ''}  ·  "
                              f"{result.confidence:.0f}% confidence",
                 font=(FONT_UI, 8), bg=bg, fg=THEME["subtle"]).pack(side=tk.LEFT)

        header = tk.Frame(self.body, bg=bg)
        header.pack(fill=tk.X, padx=px(20), pady=(px(18), px(12)))
        badge = tk.Canvas(header, width=px(36), height=px(36), bg=bg,
                          highlightthickness=0, bd=0)
        self._badge = rounded_image(badge, px(36), px(36), px(18), blend(colour, bg, 0.16), bg)
        badge.create_image(0, 0, anchor='nw', image=self._badge)
        badge.create_text(px(18), px(18) + 1, text=glyph, font=(FONT_GLYPH, 12), fill=colour)
        badge.pack(side=tk.LEFT)
        titles = tk.Frame(header, bg=bg)
        titles.pack(side=tk.LEFT, fill=tk.X, padx=(px(12), 0))
        tk.Label(titles, text=headline, font=(FONT_SEMIBOLD, 12), bg=bg, fg=THEME["fg"],
                 anchor='w').pack(fill=tk.X)
        tk.Label(titles, text=subtitle, font=(FONT_UI, 9), bg=bg, fg=THEME["muted"],
                 anchor='w').pack(fill=tk.X)

        panel = tk.Frame(self.body, bg=THEME["surface"])
        panel.pack(fill=tk.X, padx=px(20))
        mono = tkfont.Font(self, family=FONT_MONO, size=9)
        columns = max(10, (width - px(64)) // max(1, mono.measure('0')))
        tk.Label(panel, text=self._preview(result.text, columns), font=mono,
                 bg=THEME["surface"], fg=THEME["fg"], justify=tk.LEFT,
                 anchor='nw').pack(fill=tk.X, padx=px(12), pady=px(10))

        self.update_idletasks()
        height = self.winfo_reqheight()
        # centre on the monitor where the snip was taken, not always on the primary one
        left, top, right, bottom = monitor_at_cursor(work_area=True)
        x = left + max(0, (right - left - width) // 2)
        y = top + max(0, (bottom - top - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.deiconify()
        if apply_window_chrome(self):
            self.body.pack_configure(padx=0, pady=0)
        self.after(self.TICK_MS, self._tick)

    @classmethod
    def _preview(cls, text: str, columns: int) -> str:
        lines = [line.strip() for line in text.strip().splitlines() if line.strip()]
        shown = [line if len(line) <= columns else line[:columns - 1] + "…"
                 for line in lines[:cls.PREVIEW_LINES]]
        hidden = len(lines) - len(shown)
        if hidden > 0:
            shown[-1] = f"… and {hidden + 1} more lines"
        return "\n".join(shown)

    def _pointer_inside(self) -> bool:
        x, y = self.winfo_pointerxy()
        return (self.winfo_rootx() <= x < self.winfo_rootx() + self.winfo_width()
                and self.winfo_rooty() <= y < self.winfo_rooty() + self.winfo_height())

    def _tick(self):
        try:
            if not self.winfo_exists():
                return
            if not self._pointer_inside():
                self._elapsed += self.TICK_MS / 1000
            if self._elapsed >= self._timeout:
                self.destroy()
                return
            remaining = 1 - self._elapsed / self._timeout
            self.bar.coords(self._bar_fill, 0, 0, self.bar.winfo_width() * remaining,
                            self.bar.winfo_height())
            self.after(self.TICK_MS, self._tick)
        except tk.TclError:
            pass                                 # destroyed between ticks


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()                 # no half-built flash while the UI is assembled
        self.root.overrideredirect(True)
        self.root.configure(bg=THEME["border"])
        self.root.title(APP_NAME)
        self.root.report_callback_exception = self._on_tk_error
        # closing from the taskbar while minimized must still save the config
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        try:
            icon_path = get_resource_path(ICON_FILENAME)
            if os.path.exists(icon_path):
                self.root.iconbitmap(default=icon_path)   # default: dialogs inherit it too
        except tk.TclError:
            log.exception("could not load the window icon")

        self.config = ConfigManager.load()
        saved_lang = self.config.get("lang")
        self.ocr = OCREngine(saved_lang if isinstance(saved_lang, str) else DEFAULT_LANG)
        self.results: "queue.Queue[Tuple[str, Any]]" = queue.Queue()
        self.busy = False
        self.popup: Optional[ResultPopup] = None
        self._drag_offset: Optional[Tuple[int, int]] = None

        self._setup_window_geometry()
        self.frame = tk.Frame(self.root, bg=THEME["bg"])
        self.frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        self._build_custom_titlebar()
        self._build_main_ui()

        self.root.deiconify()
        self.root.after(10, self._apply_chrome)

    def _px(self, value: float) -> int:
        return int(round(value * getattr(self, 'scale', 1.0)))

    def _apply_chrome(self):
        force_taskbar_visibility(self.root)
        if apply_window_chrome(self.root):
            self.frame.pack_configure(padx=0, pady=0)     # Windows draws the border itself

    def _on_tk_error(self, exc_type, exc_value, exc_tb):
        # in a --noconsole build these would otherwise vanish silently
        log.error("unhandled callback error", exc_info=(exc_type, exc_value, exc_tb))
        messagebox.showerror(APP_NAME, f"Unexpected error:\n{exc_value}")

    def _setup_window_geometry(self):
        scale = self.scale = dpi_scale(self.root)
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

    def _drag_handle(self, *widgets: tk.Misc):
        for widget in widgets:
            widget.bind("<Button-1>", self._start_move)
            widget.bind("<B1-Motion>", self._do_move)

    def _build_custom_titlebar(self):
        px, bg = self._px, THEME["bg"]
        self.title_bar = tk.Frame(self.frame, bg=bg, height=px(40))
        self.title_bar.pack(fill=tk.X, side=tk.TOP)
        self.title_bar.pack_propagate(False)
        handles: List[tk.Misc] = [self.title_bar]

        self._title_icon = icon_image(self.root, px(18))
        if self._title_icon is not None:
            icon = tk.Label(self.title_bar, image=self._title_icon, bg=bg)
            icon.pack(side=tk.LEFT, padx=(px(14), px(8)))
            handles.append(icon)
        title = tk.Label(self.title_bar, text=APP_NAME, font=(FONT_SEMIBOLD, 10),
                         bg=bg, fg=THEME["fg"])
        title.pack(side=tk.LEFT, padx=(0 if self._title_icon else px(14), px(6)))
        version = tk.Label(self.title_bar, text=f"v{VERSION}", font=(FONT_UI, 8),
                           bg=bg, fg=THEME["subtle"])
        version.pack(side=tk.LEFT, pady=(px(2), 0))
        handles += [title, version]
        self._drag_handle(*handles)

        self._titlebar_button(GLYPH_CLOSE, self._on_close, THEME["close_hov"])
        self._titlebar_button(GLYPH_MINIMIZE, self._minimize, THEME["surface_hov"])

    def _titlebar_button(self, glyph: str, command: Callable[[], None], hover: str):
        bg = THEME["bg"]
        button = tk.Label(self.title_bar, text=glyph, font=(FONT_GLYPH, 8), bg=bg,
                          fg=THEME["muted"], padx=self._px(16))
        button.pack(side=tk.RIGHT, fill=tk.Y)
        button.bind("<Enter>", lambda _e: button.config(bg=hover, fg="#ffffff"))
        button.bind("<Leave>", lambda _e: button.config(bg=bg, fg=THEME["muted"]))

        def release(event):
            if 0 <= event.x < button.winfo_width() and 0 <= event.y < button.winfo_height():
                command()
        button.bind("<ButtonRelease-1>", release)

    def _build_main_ui(self):
        px, bg = self._px, THEME["bg"]
        main = tk.Frame(self.frame, bg=bg)
        main.pack(fill=tk.BOTH, expand=True, padx=px(20))

        footer = tk.Frame(main, bg=bg)
        footer.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, px(14)))
        tk.Frame(main, bg=THEME["border"], height=1).pack(side=tk.BOTTOM, fill=tk.X,
                                                          pady=(0, px(12)))

        content = tk.Frame(main, bg=bg)
        content.pack(expand=True)
        self.capture_btn = RoundedButton(
            content, "Capture text", self._start_snip, width=px(WINDOW_W - 40),
            height=px(46), background=bg, glyph=GLYPH_CAPTURE, font_size=11)
        self.capture_btn.pack()
        hint = tk.Label(content, text="Drag a box around any text  ·  Esc cancels",
                        font=(FONT_UI, 9), bg=bg, fg=THEME["muted"])
        hint.pack(pady=(px(10), 0))
        self._drag_handle(main, content, hint, footer)

        self._build_language_picker(footer)

    def _build_language_picker(self, parent):
        px, bg = self._px, THEME["bg"]
        choices = self.ocr.language_choices()
        label = tk.Label(parent, text="Language", font=(FONT_UI, 9), bg=bg, fg=THEME["muted"])
        label.pack(side=tk.LEFT)
        self._drag_handle(label)

        if not choices:
            tk.Label(parent, text="no language data found", font=(FONT_SEMIBOLD, 9),
                     bg=bg, fg=THEME["warning"]).pack(side=tk.RIGHT)
            return

        self.lang_choices = [(code, language_label(code)) for code in choices]
        self.lang_dropdown: Optional[Dropdown] = None
        self._dropdown_closed_at = 0.0

        measure = tkfont.Font(self.root, family=FONT_SEMIBOLD, size=9).measure
        widest = max(measure(language_label(code)) for code in choices)
        self.lang_button = RoundedButton(
            parent, language_label(self.ocr.lang), self._show_language_menu,
            width=widest + px(48), height=px(28), background=bg, glyph=GLYPH_CHEVRON,
            primary=False, font_size=9, glyph_after=True)
        self.lang_button.pack(side=tk.RIGHT)

    def _show_language_menu(self):
        # clicking the button while the list is open closes it via focus-out first;
        # don't reopen it straight away, so the button toggles
        if time.monotonic() - self._dropdown_closed_at < 0.25:
            return
        self.lang_dropdown = Dropdown(self.lang_button, self.lang_choices, self.ocr.lang,
                                      self._on_language_change, self.scale)
        self.lang_dropdown.bind("<Destroy>", self._on_dropdown_closed, add='+')

    def _on_dropdown_closed(self, event):
        if event.widget is self.lang_dropdown:
            self.lang_dropdown = None
            self._dropdown_closed_at = time.monotonic()

    def _on_language_change(self, value: str):
        self.ocr.lang = value
        self.config["lang"] = value
        ConfigManager.save(self.config)
        if hasattr(self, 'lang_button'):
            self.lang_button.set_state(text=language_label(value))
        log.info("language set to %s", value)

    def _start_move(self, event):
        self._drag_offset = (event.x_root - self.root.winfo_x(),
                             event.y_root - self.root.winfo_y())

    def _do_move(self, event):
        if self._drag_offset is None:
            return
        dx, dy = self._drag_offset
        self.root.geometry(f"+{event.x_root - dx}+{event.y_root - dy}")

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
            self._apply_chrome()

    def _on_close(self):
        # a minimized window sits at (-32000, -32000); keep the last real position instead
        if self.root.state() == 'normal':
            self.config["x"] = self.root.winfo_x()
            self.config["y"] = self.root.winfo_y()
        self.config["lang"] = self.ocr.lang
        self.config["version"] = VERSION
        ConfigManager.save(self.config)
        log.info("--- shutting down ---")
        self.root.destroy()

    def _set_busy(self, busy: bool):
        self.busy = busy
        self.capture_btn.set_state(text="Reading text…" if busy else "Capture text",
                                   enabled=not busy)

    def _start_snip(self):
        if self.busy:
            return
        self._close_popup()             # it is topmost, so it would end up inside the snip
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
        self._close_popup()
        self.popup = ResultPopup(self.root, result)

    def _close_popup(self):
        popup, self.popup = self.popup, None
        try:
            if popup is not None and popup.winfo_exists():
                popup.destroy()
        except tk.TclError:
            pass

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
