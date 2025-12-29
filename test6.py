# tools/generate_app_icon.py
# Generates a Windows .ico using only code (no external image files).

import os
from PIL import Image, ImageDraw, ImageFont

NAVY = (11, 47, 91, 255)       # #0B2F5B
NAVY_2 = (16, 58, 107, 255)    # #103A6B
WHITE = (255, 255, 255, 255)
ACCENT = (234, 242, 255, 255)  # light accent

def _rounded_rect(draw, xy, r, fill):
    x0, y0, x1, y1 = xy
    draw.rounded_rectangle([x0, y0, x1, y1], radius=r, fill=fill)

def _make_icon(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    pad = max(2, size // 16)
    r = max(6, size // 6)

    # base tile
    _rounded_rect(d, (pad, pad, size - pad, size - pad), r, NAVY)

    # subtle top band
    band_h = max(8, size // 4)
    _rounded_rect(d, (pad, pad, size - pad, pad + band_h), r, NAVY_2)

    # simple "grid" motif (like power grid nodes/lines)
    # positions scaled
    nodes = [
        (size * 0.30, size * 0.55),
        (size * 0.55, size * 0.45),
        (size * 0.70, size * 0.62),
        (size * 0.42, size * 0.72),
    ]
    line_w = max(2, size // 20)
    node_r = max(3, size // 18)

    # lines
    d.line([nodes[0], nodes[1]], fill=ACCENT, width=line_w)
    d.line([nodes[1], nodes[2]], fill=ACCENT, width=line_w)
    d.line([nodes[0], nodes[3]], fill=ACCENT, width=line_w)
    d.line([nodes[3], nodes[2]], fill=ACCENT, width=line_w)

    # nodes
    for (x, y) in nodes:
        d.ellipse(
            (x - node_r, y - node_r, x + node_r, y + node_r),
            fill=WHITE,
            outline=ACCENT,
            width=max(1, line_w // 2),
        )

    # "CC" letters (Contingency Comparison) – simple, readable
    # Use default PIL font (no external font files needed)
    try:
        # bigger sizes look nicer, but default font is limited; still fine
        font = ImageFont.load_default()
    except Exception:
        font = None

    text = "CC"
    # crude centering for default font
    tw, th = d.textsize(text, font=font)
    tx = int(size * 0.12)
    ty = int(size * 0.12)
    d.text((tx, ty), text, fill=WHITE, font=font)

    return img

def main():
    out_dir = os.path.join(os.path.dirname(__file__), "..", "assets")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "app.ico")

    # multi-size ICO is important for Windows
    sizes = [16, 24, 32, 48, 64, 128, 256]
    images = [_make_icon(s) for s in sizes]

    # Save ICO (Pillow supports this)
    images[0].save(out_path, format="ICO", sizes=[(s, s) for s in sizes])
    print(f"Saved: {out_path}")

if __name__ == "__main__":
    main()