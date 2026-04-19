"""Generate a multi-resolution disk-cylinder icon (icon.ico) for the app.
Run once: `python make_icon.py` — produces icon.ico in the same folder."""

from PIL import Image, ImageDraw


TOP_LIGHT = (91, 174, 226, 255)   # light blue (top cap highlight)
TOP_MAIN  = (59, 142, 208, 255)   # main top face (#3B8ED0)
BODY      = (31,  83, 141, 255)   # cylinder body (#1F538D)
SHADOW    = (17,  51,  92, 255)   # bottom shadow
EDGE      = (255, 255, 255, 90)   # subtle white rim
RING      = (255, 255, 255, 140)  # disk-platter divider


def draw_disk(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img, "RGBA")

    m = max(2, size // 10)                  # outer margin
    w = size - 2 * m                        # cylinder width
    top_h = max(4, size // 5)               # top/bottom ellipse height
    body_top = m + top_h // 2
    body_bot = size - m - top_h // 2

    # Shadow under cylinder
    sh = max(2, size // 32)
    d.ellipse(
        [m - sh // 2, size - m - top_h // 2, size - m + sh // 2, size - m + sh],
        fill=(0, 0, 0, 60),
    )

    # Front body
    d.rectangle([m, body_top, size - m, body_bot], fill=BODY)

    # Body side highlight (vertical gradient suggestion via thin strips)
    for i in range(max(1, size // 48)):
        d.line(
            [(m + i, body_top), (m + i, body_bot)],
            fill=(255, 255, 255, 40),
        )

    # Bottom curved edge (front half of ellipse, darker)
    d.chord(
        [m, body_bot - top_h // 2, size - m, body_bot + top_h // 2],
        start=0, end=180, fill=SHADOW,
    )

    # Middle platter rings
    rings = 2
    for i in range(1, rings + 1):
        ry = body_top + (body_bot - body_top) * i // (rings + 1)
        d.ellipse(
            [m, ry - top_h // 6, size - m, ry + top_h // 6],
            outline=RING, width=max(1, size // 96),
        )

    # Top face (main disk surface)
    d.ellipse([m, m, size - m, m + top_h], fill=TOP_MAIN)

    # Top face inner highlight (glossy look)
    pad = max(2, size // 16)
    d.ellipse(
        [m + pad, m + pad // 2, size - m - pad, m + top_h - pad // 2],
        fill=TOP_LIGHT,
    )

    # Top rim
    d.ellipse(
        [m, m, size - m, m + top_h],
        outline=EDGE, width=max(1, size // 96),
    )

    return img


def main():
    sizes = [16, 32, 48, 64, 128, 256]
    # Render at 256 and let .ico save downsample — sharper than drawing small sizes
    base = draw_disk(256)
    base.save("icon.ico", format="ICO", sizes=[(s, s) for s in sizes])
    print("Wrote icon.ico with sizes:", sizes)


if __name__ == "__main__":
    main()
