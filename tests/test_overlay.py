"""Capture overlay tests.

These open real windows and briefly take the foreground - that is the point. The overlay
covers the whole desktop while the main window is hidden, so an overlay that cannot be
dismissed, or that stays on top of whatever the user brings up next, looks exactly like a
frozen machine. Every exit path is checked here.
"""

import ctypes
import os
import subprocess
import sys
import time
import tkinter as tk
import unittest

from helpers import (REPO_ROOT, load_quickocr, pump, pump_until, require_windows)

quickocr = load_quickocr()


class OverlayTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        require_windows()
        quickocr._configure_win32()
        quickocr.enable_high_dpi_awareness()

    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.calls = []

    def tearDown(self):
        try:
            self.root.destroy()
        except tk.TclError:
            pass
        time.sleep(0.4)          # let Windows settle before the next test grabs the foreground

    def open_overlay(self):
        overlay = quickocr.SnippingOverlay(self.root, self.calls.append)
        pump(self.root, 0.3)
        return overlay

    def settle(self, overlay):
        pump_until(self.root, lambda: bool(self.calls) and not overlay.winfo_exists())
        pump(self.root, 0.2)

    def assertCancelled(self, overlay):
        self.settle(overlay)
        self.assertEqual(len(self.calls), 1, "callback must fire exactly once")
        self.assertIsNone(self.calls[0])
        self.assertFalse(overlay.winfo_exists(), "overlay must be destroyed")


class TestOverlayAlwaysCloses(OverlayTestCase):
    def test_release_without_press(self):
        """v1.0.1 returned early here, leaving the overlay covering the desktop forever."""
        overlay = self.open_overlay()
        overlay.canvas.event_generate('<ButtonRelease-1>', x=500, y=500)
        self.assertCancelled(overlay)

    def test_escape_binding(self):
        overlay = self.open_overlay()
        overlay.event_generate('<Escape>')
        self.assertCancelled(overlay)

    def test_right_click(self):
        overlay = self.open_overlay()
        overlay.event_generate('<ButtonPress-3>', x=500, y=500)
        self.assertCancelled(overlay)

    def test_single_click_is_too_small_a_selection(self):
        overlay = self.open_overlay()
        overlay.canvas.event_generate('<ButtonPress-1>', x=400, y=400)
        overlay.canvas.event_generate('<ButtonRelease-1>', x=402, y=401)
        self.assertCancelled(overlay)

    def test_finish_is_idempotent(self):
        overlay = self.open_overlay()
        overlay._finish(None)
        overlay._finish(None)
        pump(self.root, 0.2)
        self.assertEqual(len(self.calls), 1)


class TestOverlayTakesFocus(OverlayTestCase):
    def test_overlay_becomes_the_foreground_window(self):
        """Tk's winfo_id() is an inner child window and Windows blocks foreground changes
        from a background process; if either is mishandled, Escape can never fire."""
        overlay = self.open_overlay()
        try:
            self.assertEqual(ctypes.windll.user32.GetForegroundWindow(),
                             quickocr.toplevel_hwnd(overlay))
        finally:
            overlay._finish(None)

    def test_a_real_escape_keystroke_dismisses_it(self):
        overlay = self.open_overlay()
        ctypes.windll.user32.keybd_event(0x1B, 0, 0, 0)          # VK_ESCAPE down
        ctypes.windll.user32.keybd_event(0x1B, 0, 2, 0)          # and up
        self.assertCancelled(overlay)


class TestCapture(OverlayTestCase):
    def test_drag_returns_the_requested_region_undimmed(self):
        from PIL import ImageGrab, ImageStat
        box = (300, 300, 700, 560)
        before = ImageStat.Stat(
            ImageGrab.grab(bbox=box, all_screens=True).convert('L')).mean[0]

        overlay = self.open_overlay()
        x0, y0 = box[0] - overlay.v_x, box[1] - overlay.v_y
        x1, y1 = box[2] - overlay.v_x, box[3] - overlay.v_y
        overlay.canvas.event_generate('<ButtonPress-1>', x=x0, y=y0)
        overlay.canvas.event_generate('<B1-Motion>', x=x1, y=y1)
        overlay.canvas.event_generate('<ButtonRelease-1>', x=x1, y=y1)
        self.settle(overlay)

        self.assertEqual(len(self.calls), 1)
        image = self.calls[0]
        self.assertIsNotNone(image, "a real drag must produce an image")
        self.assertEqual(image.size, (box[2] - box[0], box[3] - box[1]))

        during = ImageStat.Stat(image.convert('L')).mean[0]
        self.assertAlmostEqual(
            during, before, delta=3.0,
            msg="the dimming overlay was captured into the snip; it must be hidden first")


class TestWatchdog(OverlayTestCase):
    def test_yields_the_screen_when_another_app_takes_over(self):
        """The Task Manager case: a topmost full-desktop window must not hide whatever the
        user brings up next, or the desktop looks frozen with no way out."""
        overlay = self.open_overlay()
        self.assertTrue(overlay.attributes('-topmost'))

        rival = subprocess.Popen(
            [sys.executable, os.path.join(REPO_ROOT, 'tests', 'rival_window.py'), '15'],
            stdout=subprocess.PIPE, text=True, cwd=os.path.join(REPO_ROOT, 'tests'))
        try:
            rival_hwnd = int(rival.stdout.readline().strip())
            closed = pump_until(self.root, lambda: overlay._finished, timeout=6.0)
            self.assertTrue(closed, "overlay should stand down once it loses the foreground")
            self.assertEqual(self.calls, [None])
            self.assertEqual(ctypes.windll.user32.GetForegroundWindow(), rival_hwnd,
                             "the other window must actually end up in front")
        finally:
            rival.kill()
            rival.stdout.close()
            rival.wait(timeout=5)

    def test_hard_lifetime_ceiling(self):
        original = quickocr.OVERLAY_MAX_SECONDS
        quickocr.OVERLAY_MAX_SECONDS = 1
        try:
            overlay = self.open_overlay()
            self.assertTrue(pump_until(self.root, lambda: overlay._finished, timeout=5.0),
                            "overlay must close itself once it hits the ceiling")
            self.assertEqual(self.calls, [None])
        finally:
            quickocr.OVERLAY_MAX_SECONDS = original

    def test_watchdog_does_not_disturb_a_normal_capture(self):
        overlay = self.open_overlay()
        pump(self.root, 1.2)                     # several watchdog ticks
        self.assertFalse(overlay._finished, "overlay must survive while it holds the screen")
        x0, y0 = 300 - overlay.v_x, 300 - overlay.v_y
        overlay.canvas.event_generate('<ButtonPress-1>', x=x0, y=y0)
        overlay.canvas.event_generate('<B1-Motion>', x=x0 + 400, y=y0 + 260)
        overlay.canvas.event_generate('<ButtonRelease-1>', x=x0 + 400, y=y0 + 260)
        self.settle(overlay)
        self.assertIsNotNone(self.calls[0])


if __name__ == '__main__':
    unittest.main()
