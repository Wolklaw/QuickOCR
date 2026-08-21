"""Stage the smallest Tesseract-OCR tree QuickOCR actually needs, into build/bundle.

The upstream Tesseract install directory is ~91 MB and holds 18 executables (including an
NSIS uninstaller), Java jars for the training GUI and a stack of pango/cairo/ICU DLLs that
only the training tools use. PyInstaller --onefile re-extracts everything to %TEMP% on every
launch, which is slow and makes antivirus heuristics unhappy, so only ship the closure of
tesseract.exe's own imports plus the language data.

Run directly to see what would be staged:  py make_bundle.py
"""

import os
import shutil
import struct
import sys

SOURCE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Tesseract-OCR')
STAGE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         'build', 'bundle', 'Tesseract-OCR')
# osd is orientation detection, which QuickOCR never asks for (it always passes --psm 6)
SKIP_LANGUAGES = ('osd.traineddata',)


def imported_dlls(pe_path):
    """Names in a PE file's import directory. Empty list if it cannot be parsed."""
    try:
        with open(pe_path, 'rb') as handle:
            data = handle.read()
        pe = struct.unpack_from('<I', data, 0x3c)[0]
        if data[pe:pe + 4] != b'PE\0\0':
            return []
        section_count = struct.unpack_from('<H', data, pe + 6)[0]
        opt_size = struct.unpack_from('<H', data, pe + 20)[0]
        opt = pe + 24
        is_pe32_plus = struct.unpack_from('<H', data, opt)[0] == 0x20b
        import_rva = struct.unpack_from('<I', data, opt + (112 if is_pe32_plus else 96) + 8)[0]
        if not import_rva:
            return []

        sections = []
        for i in range(section_count):
            base = opt + opt_size + i * 40
            virtual_size = struct.unpack_from('<I', data, base + 8)[0]
            virtual_addr = struct.unpack_from('<I', data, base + 12)[0]
            raw_size = struct.unpack_from('<I', data, base + 16)[0]
            raw_ptr = struct.unpack_from('<I', data, base + 20)[0]
            sections.append((virtual_addr, max(virtual_size, raw_size), raw_ptr))

        def to_offset(rva):
            for virtual_addr, size, raw_ptr in sections:
                if virtual_addr <= rva < virtual_addr + size:
                    return raw_ptr + (rva - virtual_addr)
            return None

        names, offset = [], to_offset(import_rva)
        if offset is None:
            return []
        while True:
            entry = data[offset:offset + 20]
            if len(entry) < 20 or entry == b'\0' * 20:
                break
            name_offset = to_offset(struct.unpack_from('<I', entry, 12)[0])
            if name_offset:
                end = data.index(b'\0', name_offset)
                names.append(data[name_offset:end].decode('ascii', 'replace'))
            offset += 20
        return names
    except (OSError, ValueError, struct.error, IndexError) as exc:
        print(f"  ! could not read imports from {os.path.basename(pe_path)}: {exc}")
        return []


def required_files():
    available = {name.lower(): name for name in os.listdir(SOURCE_DIR)
                 if name.lower().endswith('.dll')}
    needed, pending, seen = set(), ['tesseract.exe'], set()
    while pending:
        current = pending.pop()
        if current.lower() in seen:
            continue
        seen.add(current.lower())
        path = os.path.join(SOURCE_DIR, current)
        if not os.path.exists(path):
            continue
        for name in imported_dlls(path):
            key = name.lower()
            if key in available and key not in needed:
                needed.add(key)
                pending.append(available[key])
    return ['tesseract.exe'] + sorted(available[key] for key in needed)


def main():
    if not os.path.isdir(SOURCE_DIR):
        sys.exit(f"missing {SOURCE_DIR}")

    files = required_files()
    if os.path.isdir(STAGE_DIR):
        shutil.rmtree(STAGE_DIR)
    os.makedirs(os.path.join(STAGE_DIR, 'tessdata'), exist_ok=True)

    total = 0
    for name in files:
        source = os.path.join(SOURCE_DIR, name)
        if not os.path.exists(source):
            sys.exit(f"missing required file: {name}")
        shutil.copy2(source, os.path.join(STAGE_DIR, name))
        total += os.path.getsize(source)

    # every language present is shipped, so adding one is a matter of dropping the
    # .traineddata into Tesseract-OCR/tessdata - no code or build change needed
    languages = sorted(
        name for name in os.listdir(os.path.join(SOURCE_DIR, 'tessdata'))
        if name.endswith('.traineddata') and name not in SKIP_LANGUAGES
    )
    if not languages:
        sys.exit("no .traineddata found in Tesseract-OCR/tessdata")
    for name in languages:
        source = os.path.join(SOURCE_DIR, 'tessdata', name)
        shutil.copy2(source, os.path.join(STAGE_DIR, 'tessdata', name))
        total += os.path.getsize(source)

    configs = os.path.join(SOURCE_DIR, 'tessdata', 'configs')
    if os.path.isdir(configs):
        shutil.copytree(configs, os.path.join(STAGE_DIR, 'tessdata', 'configs'))

    original = sum(
        os.path.getsize(os.path.join(root, name))
        for root, _, names in os.walk(SOURCE_DIR) for name in names
    )
    print(f"staged {len(files) + len(languages)} files into {STAGE_DIR}")
    print(f"  languages: {', '.join(n[:-len('.traineddata')] for n in languages)}")
    print(f"  {original / 1e6:.1f} MB  ->  {total / 1e6:.1f} MB "
          f"({100 - total * 100 / original:.0f}% smaller)")


if __name__ == "__main__":
    main()
