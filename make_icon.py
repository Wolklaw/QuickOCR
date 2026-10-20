"""Draw QuickOCR's icon: viewfinder brackets around lines of text on a blue tile.

Writes quickocr.ico (16-256 px) and icon.png (256 px, for the README). Each size is drawn on
its own rather than shrunk from one big image, so strokes stay crisp at 16 px.

Run:  py make_icon.py
"""

import os

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
SIZES = (16, 20, 24, 32, 40, 48, 64, 128, 256)
SUPERSAMPLE = 8
TOP_LEFT = (84, 148, 255)          # gradient start
BOTTOM_RIGHT = (30, 64, 206)       # gradient end
INK = (255, 255, 255, 255)


def gradient(size):
    """Diagonal gradient, top-left to bottom-right."""
    small = Image.new('RGB', (64, 64))
    pixels = small.load()
    for y in range(64):
        for x in range(64):
            t = (x + y) / 126
            pixels[x, y] = tuple(round(a + (b - a) * t) for a, b in zip(TOP_LEFT, BOTTOM_RIGHT))
    return small.resize((size, size), Image.Resampling.BICUBIC)


def draw_icon(size):
    s = size * SUPERSAMPLE
    px = s / size                                   # one final pixel, in supersampled units

    def stroke(fraction, minimum):
        return max(fraction * s, minimum * px)

    image = Image.new('RGBA', (s, s), (0, 0, 0, 0))
    margin = round(s * (0.02 if size >= 48 else 0))
    mask = Image.new('L', (s, s), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        (margin, margin, s - 1 - margin, s - 1 - margin), radius=s * 0.23, fill=255)
    image.paste(gradient(s), (0, 0), mask)

    draw = ImageDraw.Draw(image)

    # viewfinder brackets
    width = stroke(0.065, 1.6)
    inset = s * (0.2 if size >= 24 else 0.17)
    arm = s * 0.19
    half = width / 2
    for cx, cy, dx, dy in ((inset, inset, 1, 1), (s - inset, inset, -1, 1),
                           (inset, s - inset, 1, -1), (s - inset, s - inset, -1, -1)):
        points = [(cx, cy + dy * arm), (cx, cy), (cx + dx * arm, cy)]
        draw.line(points, fill=INK, width=round(width), joint='curve')
        for x, y in (points[0], points[2]):          # round the open ends
            draw.ellipse((x - half, y - half, x + half, y + half), fill=INK)

    # lines of text
    bar = stroke(0.075, 1.6)
    left = s * 0.33
    if size >= 24:
        rows = ((0.385, 0.34), (0.5, 0.34), (0.615, 0.2))
    else:
        rows = ((0.42, 0.34), (0.58, 0.2))
    for centre, length in rows:
        y = s * centre
        draw.rounded_rectangle((left, y - bar / 2, left + s * length, y + bar / 2),
                               radius=bar / 2, fill=INK)

    return image.resize((size, size), Image.Resampling.LANCZOS)


def main():
    frames = [draw_icon(size) for size in SIZES]
    largest = frames[-1]
    largest.save(os.path.join(HERE, 'quickocr.ico'),
                 sizes=[(size, size) for size in SIZES], append_images=frames[:-1])
    largest.save(os.path.join(HERE, 'icon.png'))
    print(f"wrote quickocr.ico ({', '.join(map(str, SIZES))} px) and icon.png")


if __name__ == "__main__":
    main()
