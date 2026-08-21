# QuickOCR tests

```bat
run_tests.bat
```

or directly:

```bat
py -m unittest discover -s tests -t tests -v
```

Plain `unittest` from the standard library - no extra test dependency. Needs the runtime
packages from `requirements.txt` and the bundled `Tesseract-OCR/tesseract.exe`; anything
missing makes the affected tests skip rather than fail.

**These tests open real windows and briefly take the foreground.** That is deliberate: the
capture overlay covers the entire desktop, and the bugs worth guarding against only appear
against the real Windows focus and compositing behaviour. Expect the screen to flicker for
about twenty seconds. Don't type during the run.

## What each file covers

| File | Covers |
|---|---|
| `test_ocr.py` | Text that v1.0.1 garbled or erased (grey and low-contrast), variant selection, language discovery, non-ASCII passthrough |
| `test_overlay.py` | Every way the overlay must close, that it takes keyboard focus, that captures exclude the overlay's own dimming, and the watchdog |
| `test_app.py` | Config corruption, window placement, clipboard read-back verification |

`rival_window.py` is a helper, not a test: it is a second process that steals the foreground,
standing in for Task Manager. It has to be a separate process - a window inside the test
process is blocked by the overlay's Tk grab and would never take the foreground.

## Regressions these lock down

* A `ButtonRelease` with no `ButtonPress` used to leave the overlay covering the desktop forever.
* Escape could never fire, because `winfo_id()` is an inner child window and Windows blocks
  foreground changes from a background process.
* A topmost full-desktop overlay hid anything the user brought up next, Task Manager included.
* A fixed threshold erased any text lighter than about `#5a5a5a`, silently.
* The result popup claimed "COPIED TO CLIPBOARD" whether or not the copy succeeded.
