import os
import struct
import sys

try:
    from PIL import Image, ImageDraw, ImageFilter
except Exception:  # pragma: no cover
    Image = None
    ImageDraw = None
    ImageFilter = None


def _render_pixels(size: int):
    # Render a simple envelope icon at given size.
    w = h = int(size)
    bg = (0x2D, 0x7D, 0xD2, 0xFF)  # blue
    fg = (0xFF, 0xFF, 0xFF, 0xFF)  # white
    fg2 = (0xD9, 0xEE, 0xFF, 0xFF)  # light tint

    pixels = [[bg for _ in range(w)] for _ in range(h)]

    def sx(v: int) -> int:
        return max(0, min(w - 1, int(round(v * (w / 16.0)))))

    def sy(v: int) -> int:
        return max(0, min(h - 1, int(round(v * (h / 16.0)))))

    x0, y0, x1, y1 = sx(3), sy(5), sx(12), sy(12)
    for y in range(y0, y1 + 1):
        for x in range(x0, x1 + 1):
            pixels[y][x] = fg2

    for x in range(x0, x1 + 1):
        pixels[y0][x] = fg
        pixels[y1][x] = fg
    for y in range(y0, y1 + 1):
        pixels[y][x0] = fg
        pixels[y][x1] = fg

    steps = max(5, int(round(5 * (w / 16.0))))
    for i in range(0, steps):
        yy = y0 + i
        xl = x0 + i
        xr = x1 - i
        if 0 <= yy < h:
            if 0 <= xl < w:
                pixels[yy][xl] = fg
            if 0 <= xr < w:
                pixels[yy][xr] = fg

    cx = (x0 + x1) // 2
    for i in range(0, steps):
        x_l = x0 + i
        x_r = x1 - i
        yy = y0 + int(round(3 * (h / 16.0))) + i
        if yy <= y1 and 0 <= yy < h:
            if x_l <= cx and 0 <= x_l < w:
                pixels[yy][x_l] = fg
            if x_r >= cx and 0 <= x_r < w:
                pixels[yy][x_r] = fg

    bar_y = sy(3)
    for x in range(sx(4), sx(12) + 1):
        pixels[bar_y][x] = fg

    return pixels


def _bmp_from_pixels(pixels):
    h = len(pixels)
    w = len(pixels[0]) if h else 0

    xor = bytearray()
    for y in range(h - 1, -1, -1):
        for x in range(w):
            r, g, b, a = pixels[y][x]
            xor += bytes([b, g, r, a])

    row_bytes = ((w + 31) // 32) * 4
    and_mask = bytearray(row_bytes * h)

    return struct.pack(
        "<IIIHHIIIIII",
        40,
        w,
        h * 2,
        1,
        32,
        0,
        len(xor) + len(and_mask),
        0,
        0,
        0,
        0,
    ) + xor + and_mask


def make_icon(path: str, *, sizes=(16, 32, 48)) -> None:
    if Image is not None:
        make_icon_pillow(path, sizes=tuple(int(s) for s in (16, 32, 48, 256)))
        return

    images = []
    for s in sizes:
        px = _render_pixels(int(s))
        bmp = _bmp_from_pixels(px)
        images.append((int(s), bmp))

    icondir = struct.pack("<HHH", 0, 1, len(images))

    # Build entries first to calculate offsets.
    entries = bytearray()
    offset = 6 + (16 * len(images))
    for s, bmp in images:
        w = s if s < 256 else 0
        h = s if s < 256 else 0
        entries += struct.pack("<BBBBHHII", w, h, 0, 0, 1, 32, len(bmp), offset)
        offset += len(bmp)

    payload = bytearray()
    for _s, bmp in images:
        payload += bmp

    ico = icondir + entries + payload
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as f:
        f.write(ico)


def _lerp(a: int, b: int, t: float) -> int:
    return int(round(a + (b - a) * t))


def make_icon_pillow(path: str, *, sizes=(16, 32, 48, 256)) -> None:
    # Create a more "3D" icon at 256px, then let Pillow downscale for other sizes.
    base = 256
    img = Image.new("RGBA", (base, base), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Shadow
    shadow = Image.new("RGBA", (base, base), (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    pad = 26
    radius = 46
    shadow_rect = (pad + 8, pad + 14, base - pad + 8, base - pad + 14)
    shadow_draw.rounded_rectangle(shadow_rect, radius=radius, fill=(0, 0, 0, 140))
    shadow = shadow.filter(ImageFilter.GaussianBlur(14))
    img.alpha_composite(shadow)

    # Background rounded square with vertical gradient
    pad = 26
    rect = (pad, pad, base - pad, base - pad)
    bg = Image.new("RGBA", (base, base), (0, 0, 0, 0))
    bg_draw = ImageDraw.Draw(bg)
    bg_mask = Image.new("L", (base, base), 0)
    ImageDraw.Draw(bg_mask).rounded_rectangle(rect, radius=radius, fill=255)

    top = (0x35, 0x9B, 0xFF, 255)
    bottom = (0x1F, 0x5F, 0xC9, 255)
    for y in range(base):
        t = y / (base - 1)
        color = (
            _lerp(top[0], bottom[0], t),
            _lerp(top[1], bottom[1], t),
            _lerp(top[2], bottom[2], t),
            255,
        )
        bg_draw.line((0, y, base, y), fill=color)
    bg.putalpha(bg_mask)
    img.alpha_composite(bg)

    # Gloss highlight
    gloss = Image.new("RGBA", (base, base), (0, 0, 0, 0))
    gloss_draw = ImageDraw.Draw(gloss)
    gloss_rect = (pad + 10, pad + 10, base - pad - 10, pad + (base - 2 * pad) * 0.55)
    gloss_draw.rounded_rectangle(gloss_rect, radius=radius - 12, fill=(255, 255, 255, 70))
    gloss = gloss.filter(ImageFilter.GaussianBlur(2))
    img.alpha_composite(gloss)

    # Envelope (slight bevel)
    env = Image.new("RGBA", (base, base), (0, 0, 0, 0))
    env_draw = ImageDraw.Draw(env)
    ex0, ey0, ex1, ey1 = 62, 92, 194, 176
    env_draw.rounded_rectangle((ex0, ey0, ex1, ey1), radius=14, fill=(240, 248, 255, 255))
    env_draw.rounded_rectangle((ex0, ey0, ex1, ey1), radius=14, outline=(255, 255, 255, 180), width=3)
    # Flap
    env_draw.polygon([(ex0, ey0), ((ex0 + ex1) // 2, ey0 + 52), (ex1, ey0)], fill=(220, 238, 255, 255))
    env_draw.line([(ex0, ey0), ((ex0 + ex1) // 2, ey0 + 52)], fill=(255, 255, 255, 200), width=3)
    env_draw.line([((ex0 + ex1) // 2, ey0 + 52), (ex1, ey0)], fill=(255, 255, 255, 200), width=3)
    # Inner lines
    env_draw.line([(ex0 + 12, ey1 - 22), (ex1 - 12, ey1 - 22)], fill=(190, 215, 245, 255), width=4)
    env_draw.line([(ex0 + 12, ey1 - 40), (ex1 - 48, ey1 - 40)], fill=(190, 215, 245, 255), width=4)

    env = env.filter(ImageFilter.GaussianBlur(0.3))
    img.alpha_composite(env)

    os.makedirs(os.path.dirname(path), exist_ok=True)
    img.save(path, format="ICO", sizes=[(int(s), int(s)) for s in sizes])


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join("assets", "smtp_tool.ico")
    make_icon(out)
    print(f"Wrote {out}")
