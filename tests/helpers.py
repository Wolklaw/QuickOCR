"""Shared helpers for the QuickOCR tests."""

import os
import random
import sys
import time
import unittest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

SENTENCE = "Balance 1,024.50 USD - Invoice #7781"


def load_quickocr():
    """Import the app, skipping the whole module if its dependencies are absent."""
    try:
        import quickocr
    except ImportError as exc:                       # pragma: no cover - environment issue
        raise unittest.SkipTest(f"quickocr could not be imported: {exc}")
    return quickocr


def require_windows():
    if os.name != 'nt':
        raise unittest.SkipTest("QuickOCR is a Windows application")


def require_tesseract():
    exe = os.path.join(REPO_ROOT, 'Tesseract-OCR', 'tesseract.exe')
    if not os.path.exists(exe):
        raise unittest.SkipTest(f"bundled tesseract not found at {exe}")
    return exe


def font(size=15):
    from PIL import ImageFont
    for name in ("segoeui.ttf", "arial.ttf", "DejaVuSans.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    raise unittest.SkipTest("no scalable font available to render test images")


def render_text(text=SENTENCE, fg=(0, 0, 0), bg=(255, 255, 255), noise=0, size=15, width=460):
    """A strip of text, optionally speckled to imitate noisy game or card artwork."""
    from PIL import Image, ImageDraw
    image = Image.new("RGB", (width, 44), bg)
    ImageDraw.Draw(image).text((8, 11), text, fill=fg, font=font(size))
    if noise:
        random.seed(7)
        pixels = image.load()
        for y in range(image.height):
            for x in range(image.width):
                r, g, b = pixels[x, y]
                shift = random.randint(-noise, noise)
                pixels[x, y] = (max(0, min(255, r + shift)),
                                max(0, min(255, g + shift)),
                                max(0, min(255, b + shift)))
    return image


def pump(widget, seconds):
    """Run the Tk event loop for a while without blocking on mainloop()."""
    deadline = time.time() + seconds
    while time.time() < deadline:
        widget.update()
        time.sleep(0.02)


def pump_until(widget, predicate, timeout=5.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        widget.update()
        if predicate():
            return True
        time.sleep(0.02)
    return False


def normalise(text):
    return ''.join(text.split())
