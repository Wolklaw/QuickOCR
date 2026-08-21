"""Config, window placement and clipboard tests.

The config lives in %APPDATA%, so every test here redirects that to a temporary directory -
running the suite must never disturb a real installation.
"""

import json
import os
import tempfile
import tkinter as tk
import unittest

from helpers import load_quickocr, require_windows

quickocr = load_quickocr()


class RedirectedDataDir(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._original = quickocr.user_data_dir
        quickocr.user_data_dir = lambda: self._tmp.name

    def tearDown(self):
        quickocr.user_data_dir = self._original
        self._tmp.cleanup()

    def write_config(self, raw):
        with open(os.path.join(self._tmp.name, quickocr.CONFIG_FILENAME),
                  'w', encoding='utf-8') as handle:
            handle.write(raw)


class TestConfig(RedirectedDataDir):
    def test_missing_config_is_empty(self):
        self.assertEqual(quickocr.ConfigManager.load(), {})

    def test_corrupt_json_does_not_raise(self):
        self.write_config("{not json at all")
        self.assertEqual(quickocr.ConfigManager.load(), {})

    def test_non_object_json_is_rejected(self):
        for raw in ("[1, 2, 3]", '"a string"', "42", "null"):
            with self.subTest(raw=raw):
                self.write_config(raw)
                self.assertEqual(quickocr.ConfigManager.load(), {})

    def test_round_trip(self):
        quickocr.ConfigManager.save({"x": 10, "y": 20, "lang": "eng"})
        self.assertEqual(quickocr.ConfigManager.load(), {"x": 10, "y": 20, "lang": "eng"})

    def test_unicode_values_round_trip(self):
        quickocr.ConfigManager.save({"note": "café ünïcode 日本"})
        self.assertEqual(quickocr.ConfigManager.load()["note"], "café ünïcode 日本")

    def test_save_of_unserialisable_data_does_not_raise(self):
        quickocr.ConfigManager.save({"bad": object()})       # logged, not raised
        self.assertEqual(quickocr.ConfigManager.load(), {})


class TestUserDataDir(unittest.TestCase):
    def test_survives_a_missing_appdata(self):
        require_windows()
        saved = os.environ.pop('APPDATA', None)
        try:
            self.assertTrue(os.path.isdir(quickocr.user_data_dir()))
        finally:
            if saved is not None:
                os.environ['APPDATA'] = saved


class TestWindowPlacement(RedirectedDataDir):
    """A position saved on a monitor that is later unplugged must not hide the window,
    because the borderless window has no titlebar to drag it back with."""

    def setUp(self):
        super().setUp()
        require_windows()
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.app = quickocr.App.__new__(quickocr.App)
        self.app.root = self.root

    def tearDown(self):
        try:
            self.root.destroy()
        except tk.TclError:
            pass
        super().tearDown()

    def assertOnScreen(self, config):
        self.app.config = config
        self.app._setup_window_geometry()
        self.root.update_idletasks()
        v_x, v_y, v_width, v_height = quickocr.virtual_screen()
        x, y = self.root.winfo_x(), self.root.winfo_y()
        self.assertTrue(v_x <= x <= v_x + v_width, f"x={x} is off the desktop")
        self.assertTrue(v_y <= y <= v_y + v_height, f"y={y} is off the desktop")

    def test_no_saved_position(self):
        self.assertOnScreen({})

    def test_position_far_off_screen(self):
        self.assertOnScreen({"x": 99999, "y": 99999})
        self.assertOnScreen({"x": -99999, "y": -99999})

    def test_non_numeric_position(self):
        self.assertOnScreen({"x": "left", "y": None})


class TestClipboard(unittest.TestCase):
    """v1.0.1 announced 'COPIED TO CLIPBOARD' whether or not the copy worked."""

    def setUp(self):
        try:
            import pyperclip
        except ImportError as exc:
            self.skipTest(f"pyperclip unavailable: {exc}")
        # pyperclip.copy and .paste start as lazy stubs that pick a backend on first use and
        # reassign both module attributes. Trigger that now, or it happens mid-test and
        # silently undoes the patch below.
        pyperclip.paste()
        self.pyperclip = pyperclip
        self._real_copy = pyperclip.copy

    def tearDown(self):
        self.pyperclip.copy = self._real_copy

    def test_successful_copy_is_confirmed(self):
        text = "clipboard check € é"
        self.assertIsNone(quickocr.App._copy_to_clipboard(text))
        self.assertEqual(self.pyperclip.paste(), text)

    def test_silent_no_op_copy_is_detected(self):
        self.pyperclip.copy = lambda text: None
        self.assertIsNotNone(quickocr.App._copy_to_clipboard("never actually copied"))

    def test_failing_copy_is_reported(self):
        def locked(text):
            raise RuntimeError("clipboard is locked by another app")
        self.pyperclip.copy = locked
        error = quickocr.App._copy_to_clipboard("nope")
        self.assertIsNotNone(error)
        self.assertIn("locked", error)


if __name__ == '__main__':
    unittest.main()
