# Changelog

## 1.2.0 - Final release (2026-10-19)

> **QuickOCR has reached end of life.** The Windows Snipping Tool can now copy text straight
> out of a screenshot (its *Text actions* button, on Windows 11), which covers most of what
> QuickOCR does. I finished this release to finish the project. It still works and stays
> available to download, but no further updates or support are planned.

### New look
* New icon: viewfinder brackets around lines of text, drawn sharp at every size from 16 to
  256 px. It replaces the old raster image.
* Redesigned main window: rounded corners and a hairline border on Windows 11, a proper title
  bar with the icon and version, a rounded **Capture text** button with hover and busy states,
  and a themed language dropdown that shows names ("English + French") instead of codes.
* Redesigned result popup: a clear success or warning badge, a multi-line preview of what
  was copied, word count and confidence, and a countdown bar that pauses while your
  pointer is over the popup.
* The exe now carries version information (name, version, publisher) in
  *Properties > Details* and Task Manager.

## 1.1.1 (2026-10-19)
* A `TESSDATA_PREFIX` left behind by a separate Tesseract install no longer replaces the
  bundled language data. That bug made `eng+fra` quietly lose French, and French on its own
  read nothing.
* A missing `tesseract.exe` or language file is now reported as an error, not as
  "No text found".
* The "press Esc to cancel" hint is drawn on the monitor you are using. On some
  multi-monitor layouts it used to land off-screen.
* The result popup opens on the monitor you are using, not always the primary one.
* Starting a new capture closes the previous result popup, so it can't end up inside the snip.
* Closing the app from the taskbar while minimized now saves your settings.
* Settings are written safely, so a failed save no longer wipes the previous ones.
* A saved language that is no longer installed falls back to English + French.

## 1.1.0 (2026-08-20)
* The capture overlay can always be dismissed (Esc, right-click or a single click), and it
  gets out of the way if another window such as Task Manager comes to the front.
* Much better reading of grey, low-contrast and coloured text: several preprocessing
  variants are tried and the most confident read wins.
* A warning appears when the read is unreliable, and the clipboard copy is verified.
* Languages are discovered from the installed language data and can be picked in the app.
* Correct capture regions on mixed-DPI multi-monitor setups.
* A log is written to `%APPDATA%\QuickOCR\quickocr.log`.
