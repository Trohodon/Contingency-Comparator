# tools/icon_builder.py
# Retro-futuristic icon generator for: Dominion Contingency Comparator (DCC)
# Outputs:
#   - assets/app.ico     (Windows icon, multi-size)
#   - assets/app_256.png (preview)
#
# Requirements:
#   py -m pip install pillow --user

import os
import math
from PIL import Image, ImageDraw, ImageFont, ImageFilter

SIZES = [16, 24, 32, 48, 64, 128, 256]


# ---------- Color palette (synthwave) ----------
DEEP = (10, 12, 28, 255)            # near-black navy
PURPLE = (130, 58, 240, 255)        # neon purple
PINK = (255, 64, 180, 255)          # hot pink
CYAN = (64, 220, 255, 255)          # neon cyan
TEAL = (40, 255, 200, 255)
WHITE = (245, 248, 255, 255)


def lerp(a, b, t: float):
    return int(a + (b - a) * t)


def lerp_rgba(c1, c2, t: float):
    return (
        lerp(c1[0], c2[0], t),
        lerp(c1[1], c2[1], t),
        lerp(c1[2], c2[2], t),
        lerp(c1[3], c2[3], t),
    )


def make_vertical_gradient(w: int, h: int, top, mid, bottom):
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    px = img.load()
    for y in range(h):
        t = y / max(1, (h - 1))
        if t < 0.55:
            tt = t / 0.55
            c = lerp_rgba(top, mid, tt)
        else:
            tt = (t - 0.55) / 0.45
            c = lerp_rgba(mid, bottom, tt)
        for x in range(w):
            px[x, y] = c
    return img


def rounded_rect(draw: ImageDraw.ImageDraw, xy, r: int, fill, outline=None, width=1):
    # Pillow supports rounded_rectangle in modern versions; this is standard
    draw.rounded_rectangle(xy, radius=r, fill=fill, outline=outline, width=width)


def add_glow(base: Image.Image, glow_color, blur_radius: int, intensity: float = 1.0):
    """
    Creates a glow by extracting alpha -> tint -> blur -> composite.
    """
    alpha = base.split()[-1]
    glow = Image.new("RGBA", base.size, glow_color)
    glow.putalpha(alpha)
    glow = glow.filter(ImageFilter.GaussianBlur(blur_radius))

    if intensity != 1.0:
        # scale glow alpha
        ga = glow.split()[-1]
        ga = ga.point(lambda p: int(p * intensity))
        glow.putalpha(ga)

    return glow


def draw_synth_grid(d: ImageDraw.ImageDraw, size: int, pad: int):
    """
    Perspective-ish grid lines near the bottom. Kept simple so it scales down well.
    """
    w = h = size
    horizon_y = int(h * 0.55)
    bottom_y = h - pad

    # horizontal lines
    num_h = 6
    for i in range(num_h):
        t = i / max(1, (num_h - 1))
        # ease spacing (denser near horizon)
        tt = t * t
        y = int(lerp(horizon_y, bottom_y, tt))
        col = lerp_rgba((255, 64, 180, 40), (64, 220, 255, 90), t)
        d.line([(pad, y), (w - pad, y)], fill=col, width=max(1, size // 64))

    # vertical converging lines
    center_x = w // 2
    num_v = 7
    spread = int(w * 0.55)
    for i in range(num_v):
        t = i / max(1, (num_v - 1))
        x = int(center_x - spread // 2 + t * spread)
        col = lerp_rgba((64, 220, 255, 50), (255, 64, 180, 80), abs(t - 0.5) * 2)
        d.line([(x, bottom_y), (center_x, horizon_y)], fill=col, width=max(1, size // 64))


def measure_text(draw: ImageDraw.ImageDraw, text: str, font):
    # Pillow-safe measurement
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        return (bbox[2] - bbox[0], bbox[3] - bbox[1])
    except Exception:
        try:
            return font.getsize(text)
        except Exception:
            return (len(text) * 8, 12)


def load_best_font(size: int):
    # Try modern Windows fonts first; fall back cleanly.
    candidates = [
        ("segoeuib.ttf", int(size * 0.34)),     # Segoe UI Bold
        ("seguisb.ttf", int(size * 0.34)),      # Segoe UI Semibold
        ("arialbd.ttf", int(size * 0.34)),      # Arial Bold
        ("arial.ttf", int(size * 0.34)),
    ]
    for name, pt in candidates:
        try:
            return ImageFont.truetype(name, max(10, pt))
        except Exception:
            pass
    return ImageFont.load_default()


def make_icon(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))

    pad = max(2, size // 18)
    r = max(6, size // 5)

    # ----- Background tile (gradient) -----
    bg = make_vertical_gradient(
        size, size,
        top=DEEP,
        mid=(40, 22, 70, 255),
        bottom=(12, 8, 30, 255)
    )

    # Add a subtle "sun" glow behind the horizon
    sun = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    sd = ImageDraw.Draw(sun)
    cx = size // 2
    cy = int(size * 0.50)
    sun_r = int(size * 0.36)
    sd.ellipse(
        (cx - sun_r, cy - sun_r, cx + sun_r, cy + sun_r),
        fill=(255, 64, 180, 110)
    )
    sun = sun.filter(ImageFilter.GaussianBlur(max(2, size // 18)))
    bg = Image.alpha_composite(bg, sun)

    # Clip to rounded tile
    tile = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    td = ImageDraw.Draw(tile)
    mask = Image.new("L", (size, size), 0)
    md = ImageDraw.Draw(mask)
    md.rounded_rectangle((pad, pad, size - pad, size - pad), radius=r, fill=255)
    tile.paste(bg, (0, 0), mask)

    # Border stroke
    td.rounded_rectangle(
        (pad, pad, size - pad, size - pad),
        radius=r,
        outline=(80, 120, 255, 80),
        width=max(1, size // 64),
        fill=None
    )

    img = Image.alpha_composite(img, tile)

    # ----- Grid overlay (bottom) -----
    grid = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    gd = ImageDraw.Draw(grid)
    draw_synth_grid(gd, size, pad)
    grid = grid.filter(ImageFilter.GaussianBlur(max(0, size // 180)))
    img = Image.alpha_composite(img, grid)

    # ----- Neon “signal” ring + node graph (hinting power grid) -----
    motif = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(motif)

    ring_r = int(size * 0.26)
    ring_w = max(2, size // 26)
    cx, cy = int(size * 0.72), int(size * 0.30)
    d.ellipse(
        (cx - ring_r, cy - ring_r, cx + ring_r, cy + ring_r),
        outline=(64, 220, 255, 220),
        width=ring_w
    )
    # inner ring
    d.ellipse(
        (cx - int(ring_r * 0.62), cy - int(ring_r * 0.62),
         cx + int(ring_r * 0.62), cy + int(ring_r * 0.62)),
        outline=(255, 64, 180, 180),
        width=max(1, ring_w - 1)
    )

    # node graph
    nodes = [
        (int(size * 0.24), int(size * 0.42)),
        (int(size * 0.42), int(size * 0.34)),
        (int(size * 0.54), int(size * 0.46)),
        (int(size * 0.38), int(size * 0.56)),
    ]
    lw = max(2, size // 48)
    d.line([nodes[0], nodes[1]], fill=(64, 220, 255, 170), width=lw)
    d.line([nodes[1], nodes[2]], fill=(255, 64, 180, 160), width=lw)
    d.line([nodes[0], nodes[3]], fill=(255, 64, 180, 140), width=lw)
    d.line([nodes[3], nodes[2]], fill=(64, 220, 255, 140), width=lw)

    nr = max(3, size // 26)
    for (x, y) in nodes:
        d.ellipse((x - nr, y - nr, x + nr, y + nr), fill=WHITE, outline=(64, 220, 255, 200), width=max(1, lw // 2))

    # glow pass
    glow = add_glow(motif, glow_color=(64, 220, 255, 255), blur_radius=max(2, size // 14), intensity=1.2)
    img = Image.alpha_composite(img, glow)
    img = Image.alpha_composite(img, motif)

    # ----- DCC mark (top-left) -----
    txt_layer = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    td = ImageDraw.Draw(txt_layer)
    font = load_best_font(size)

    text = "DCC"
    tw, th = measure_text(td, text, font)

    tx = int(size * 0.12)
    ty = int(size * 0.10)

    # neon underline bar behind text
    bar_h = max(6, int(th * 0.65))
    bar_w = max(24, int(tw * 1.10))
    bar = Image.new("RGBA", (bar_w, bar_h), (0, 0, 0, 0))
    bd = ImageDraw.Draw(bar)
    rounded_rect(bd, (0, 0, bar_w, bar_h), r=max(4, bar_h // 2),
                fill=(255, 64, 180, 140), outline=(64, 220, 255, 180), width=max(1, size // 128))
    bar = bar.filter(ImageFilter.GaussianBlur(max(0, size // 220)))
    txt_layer.alpha_composite(bar, (tx - int(size * 0.01), ty + int(th * 0.55)))

    # text glow
    td.text((tx + 1, ty + 1), text, font=font, fill=(0, 0, 0, 120))
    td.text((tx, ty), text, font=font, fill=WHITE)

    # glow around text (by blurring alpha)
    txt_glow = add_glow(txt_layer, glow_color=(255, 64, 180, 255), blur_radius=max(2, size // 18), intensity=1.1)
    img = Image.alpha_composite(img, txt_glow)
    img = Image.alpha_composite(img, txt_layer)

    # ----- Final: slight sharpen-like pop for larger sizes to keep crisp -----
    if size >= 64:
        img = img.filter(ImageFilter.UnsharpMask(radius=1, percent=120, threshold=2))

    return img


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(here, ".."))
    out_dir = os.path.join(project_root, "assets")
    os.makedirs(out_dir, exist_ok=True)

    ico_path = os.path.join(out_dir, "app.ico")
    png_path = os.path.join(out_dir, "app_256.png")

    images = [make_icon(s) for s in SIZES]

    # preview png
    images[-1].save(png_path, format="PNG")

    # multi-size .ico
    images[0].save(ico_path, format="ICO", sizes=[(s, s) for s in SIZES])

    print("Retro-futuristic icon generated successfully:")
    print(" -", ico_path)
    print(" -", png_path)


if __name__ == "__main__":
    main()