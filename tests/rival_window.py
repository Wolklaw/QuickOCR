"""A separate process that opens a window and takes the foreground.

Stands in for Task Manager in the overlay watchdog test. It has to be a real second process:
a window inside the test process is blocked by the overlay's Tk grab, so it would never take
the foreground and the test would pass or fail for the wrong reason.

Prints its window handle on stdout, then pumps events until told to quit.
"""

import sys
import time
import tkinter as tk

from helpers import load_quickocr

quickocr = load_quickocr()


def main(seconds=15.0):
    quickocr._configure_win32()
    root = tk.Tk()
    root.title("QuickOCR test rival")
    root.geometry("320x140+60+60")
    tk.Label(root, text="stand-in for Task Manager").pack(expand=True)
    root.update()

    quickocr.force_foreground(quickocr.toplevel_hwnd(root))
    root.update()
    print(quickocr.toplevel_hwnd(root), flush=True)

    deadline = time.time() + seconds
    while time.time() < deadline:
        root.update()
        time.sleep(0.02)


if __name__ == '__main__':
    main(float(sys.argv[1]) if len(sys.argv) > 1 else 15.0)
