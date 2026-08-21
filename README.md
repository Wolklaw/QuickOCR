[🇫🇷 Version Française](README_FR.md)
# QuickOCR
A portable, offline OCR utility for Windows. Instantly capture and extract text from images, videos, games, and uncopiable UI elements.

[![VirusTotal](https://img.shields.io/badge/VirusTotal-Scan_Result-blue?logo=virustotal)](https://www.virustotal.com/gui/file/b0f5e4cbd0048ef9d6329f7d71e0a05cceda1d62596f63d01b0e3d1489617298/detection)
[![Softpedia](https://img.shields.io/badge/Softpedia-Reviewed_4.5%2F5-brightgreen)](https://www.softpedia.com/get/Office-tools/Text-editors/QuickOCR.shtml)

## Features
* **Visual Snipping:** Draw a box on your screen to capture text.
* **Bilingual:** Ships with English & French, read together or one at a time.
* **Add your own language:** Drop any Tesseract `.traineddata` file in and pick it in the app.
* **Anti-Holographic:** Special filters to read text on noisy/colored backgrounds (Game menus, Trading cards).
* **Portable:** Single .exe file. No installation. No admin rights.

## How to Use
1. Download `QuickOCR.exe` from the link below.
2. Run `QuickOCR.exe` (It is a single file, no installation needed).
3. Click **CAPTURE ZONE**.
4. Draw a box around any text.
5. The text is automatically copied to your clipboard.

To cancel a capture, click once without dragging, right-click, or press **Esc**.

## Troubleshooting
* **The screen is dimmed and nothing happens.** That is the capture overlay. Press **Esc**,
  right-click, or click once to dismiss it.
* **The text came out wrong.** The result popup shows a warning when Tesseract is unsure of
  the read, so check it before pasting.
* **Text in another language comes out garbled** (`Größe` read as `GroBe`, `niño` as `nino`).
  QuickOCR can only recognise the languages it ships with. Add the language, see below.
* **Reporting a bug.** QuickOCR writes a log to `%APPDATA%\QuickOCR\quickocr.log`.
  Attaching it to an issue makes the problem far easier to track down.

## Adding a language
1. Download the `.traineddata` file for your language from
   [tessdata_fast](https://github.com/tesseract-ocr/tessdata_fast) (for example `deu.traineddata`).
2. Put it in `Tesseract-OCR/tessdata/`.
3. Pick it from the **Language** dropdown in the app.

Building from source picks up every language present automatically. Note that Tesseract is
most accurate with one language selected, so prefer a single language over a long combination.

## Requirements
* Windows 10/11
* No other dependencies (Tesseract is bundled).

## Download
[DOWNLOAD LATEST VERSION](https://github.com/Wolklaw/QuickOCR/releases/latest)

## Building from source
```bat
build.bat
```
Installs the dependencies, stages the minimal Tesseract bundle and produces `dist/QuickOCR.exe`.
Run `run_tests.bat` to execute the test suite (see [tests/README.md](tests/README.md)).

## Reviews
QuickOCR was reviewed by **Softpedia**, which rated it **4.5/5** and awarded its
**Certified 100% Clean** badge.

> "Capture text from anywhere on your screen, even on tricky or colorful backgrounds,
> and instantly convert it into editable English or French text"
> — [Softpedia review by Alexandra Sava](https://www.softpedia.com/get/Office-tools/Text-editors/QuickOCR.shtml)

## Legal & License

**License**
This project is licensed under the **GNU GPLv3 License** - see the [LICENSE](LICENSE) file for details.

**What this means:**
* You **cannot** close the source code and sell this application.
* Any modifications you distribute must also be open source under GPLv3.
* This software is free and open source. **Do not pay for this software.**
