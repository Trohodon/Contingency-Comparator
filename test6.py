# tools/icon_builder.py
# One-time icon generator for the Contingency Comparison Tool.
# Outputs:
#   - assets/app.ico    (Windows icon, multi-size)
#   - assets/app_256.png (optional preview)
#
# Requirements:
#   pip install pillow

import os
from PIL import Image, ImageDraw, ImageFont

NAVY = (11, 47, 91, 255)       # #0B2F5B
NAVY_2 = (16, 58, 107, 255)    # #103A6B
WHITE = (255, 255, 255, 255)
ACCENT = (234, 242, 255, 255)  # light accent
SHADOW = (0, 0, 0, 60)

SIZES = [16, 24, 32, 48, 64, 128, 256]


def rounded_rect(draw: ImageDraw.ImageDraw, xy, r: int, fill):
    draw.rounded_rectangle(xy, radius=r, fill=fill)


def measure_text(draw: ImageDraw.ImageDraw, text: str, font):
    """
    Pillow-safe text measurement:
    - Prefer textbbox (new)
    - Fall back to font.getsize (old)
    """
    try:
        # Pillow >= 8-ish
        bbox = draw.textbbox((0, 0), text, font=font)
        w = bbox[2] - bbox[0]
        h = bbox[3] - bbox[1]
        return w, h
    except Exception:
        try:
            return font.getsize(text)
        except Exception:
            # last resort
            return (len(text) * 8, 12)


def make_icon(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    pad = max(2, size // 16)
    r = max(5, size // 6)

    # shadow
    rounded_rect(d, (pad + 1, pad + 2, size - pad + 1, size - pad + 2), r, SHADOW)

    # base tile
    rounded_rect(d, (pad, pad, size - pad, size - pad), r, NAVY)

    # top band
    band_h = max(8, size // 4)
    rounded_rect(d, (pad, pad, size - pad, pad + band_h), r, NAVY_2)

    # grid motif
    nodes = [
        (size * 0.30, size * 0.58),
        (size * 0.55, size * 0.46),
        (size * 0.72, size * 0.62),
        (size * 0.42, size * 0.76),
    ]
    line_w = max(2, size // 22)
    node_r = max(3, size // 20)

    d.line([nodes[0], nodes[1]], fill=ACCENT, width=line_w)
    d.line([nodes[1], nodes[2]], fill=ACCENT, width=line_w)
    d.line([nodes[0], nodes[3]], fill=ACCENT, width=line_w)
    d.line([nodes[3], nodes[2]], fill=ACCENT, width=line_w)

    for (x, y) in nodes:
        d.ellipse(
            (x - node_r, y - node_r, x + node_r, y + node_r),
            fill=WHITE,
            outline=ACCENT,
            width=max(1, line_w // 2),
        )

    # Text mark: "CC"
    # Use a bold-ish default font if possible; otherwise default.
    try:
        font = ImageFont.truetype("segoeuib.ttf", max(10, size // 3))  # Segoe UI Bold
    except Exception:
        try:
            font = ImageFont.truetype("arialbd.ttf", max(10, size // 3))  # Arial Bold
        except Exception:
            font = ImageFont.load_default()

    text = "CC"
    tw, th = measure_text(d, text, font)

    # place in top-left area with good padding
    tx = int(size * 0.14)
    ty = int(size * 0.10)

    # slight shadow for readability
    d.text((tx + 1, ty + 1), text, fill=(0, 0, 0, 90), font=font)
    d.text((tx, ty), text, fill=WHITE, font=font)

    return img


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(here, ".."))
    out_dir = os.path.join(project_root, "assets")
    os.makedirs(out_dir, exist_ok=True)

    ico_path = os.path.join(out_dir, "app.ico")
    png_path = os.path.join(out_dir, "app_256.png")

    images = [make_icon(s) for s in SIZES]

    # save preview png
    images[-1].save(png_path, format="PNG")

    # save multi-size ico
    images[0].save(ico_path, format="ICO", sizes=[(s, s) for s in SIZES])

    print("Icon generated successfully:")
    print(" -", ico_path)
    print(" -", png_path)


if __name__ == "__main__":
    main()