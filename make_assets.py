"""Generate the application icon and sidebar logo from one vector-like drawing."""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)


def _font(size: int) -> ImageFont.FreeTypeFont:
    for candidate in (r"C:\Windows\Fonts\arialbd.ttf", r"C:\Windows\Fonts\segoeuib.ttf"):
        try:
            return ImageFont.truetype(candidate, size)
        except OSError:
            continue
    return ImageFont.load_default()


def _gradient(size: int) -> Image.Image:
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    pixels = image.load()
    top = (30, 64, 175)
    bottom = (37, 99, 235)
    for y in range(size):
        t = y / max(1, size - 1)
        color = tuple(round(top[i] + (bottom[i] - top[i]) * t) for i in range(3))
        for x in range(size):
            pixels[x, y] = (*color, 255)
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, size - 1, size - 1), radius=int(size * 0.225), fill=255)
    image.putalpha(mask)
    glow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    ImageDraw.Draw(glow).ellipse(
        (-size * 0.15, -size * 0.42, size * 0.82, size * 0.45),
        fill=(96, 165, 250, 60),
    )
    glow.putalpha(Image.composite(glow.getchannel("A"), Image.new("L", (size, size), 0), mask))
    image = Image.alpha_composite(image, glow)
    image.putalpha(mask)
    return image


def draw_icon(size: int = 512) -> Image.Image:
    scale = size / 512
    canvas = _gradient(size)
    draw = ImageDraw.Draw(canvas, "RGBA")

    # Soft shadow behind the document.
    shadow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_draw.rounded_rectangle(
        (int(136 * scale), int(104 * scale), int(392 * scale), int(444 * scale)),
        radius=int(28 * scale),
        fill=(3, 15, 42, 105),
    )
    shadow = shadow.filter(ImageFilter.GaussianBlur(max(2, int(17 * scale))))
    canvas = Image.alpha_composite(canvas, shadow)
    draw = ImageDraw.Draw(canvas, "RGBA")

    # Small offset sheet, suggesting a queue of documents.
    draw.rounded_rectangle(
        (int(148 * scale), int(64 * scale), int(376 * scale), int(380 * scale)),
        radius=int(24 * scale),
        fill=(191, 214, 248, 220),
    )

    # Main white page.
    page_box = (
        int(128 * scale),
        int(90 * scale),
        int(384 * scale),
        int(434 * scale),
    )
    draw.rounded_rectangle(page_box, radius=int(28 * scale), fill=(248, 251, 255, 255))
    fold = [
        (int(322 * scale), int(90 * scale)),
        (int(384 * scale), int(152 * scale)),
        (int(322 * scale), int(152 * scale)),
    ]
    draw.polygon(fold, fill=(213, 227, 247, 255))

    # Word mark.
    font_w = _font(max(10, int(128 * scale)))
    text = "W"
    box = draw.textbbox((0, 0), text, font=font_w)
    text_w = box[2] - box[0]
    draw.text(
        (int(256 * scale) - text_w / 2, int(132 * scale)),
        text,
        font=font_w,
        fill=(28, 72, 190, 255),
    )

    # Text lines.
    line_color = (190, 204, 224, 255)
    draw.rounded_rectangle(
        (int(166 * scale), int(260 * scale), int(346 * scale), int(270 * scale)),
        radius=max(2, int(5 * scale)),
        fill=line_color,
    )
    draw.rounded_rectangle(
        (int(166 * scale), int(284 * scale), int(292 * scale), int(294 * scale)),
        radius=max(2, int(5 * scale)),
        fill=line_color,
    )

    # PDF badge.
    badge = (
        int(158 * scale),
        int(320 * scale),
        int(354 * scale),
        int(390 * scale),
    )
    draw.rounded_rectangle(badge, radius=int(18 * scale), fill=(219, 68, 55, 255))
    font_pdf = _font(max(8, int(46 * scale)))
    label = "PDF"
    bbox = draw.textbbox((0, 0), label, font=font_pdf)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    draw.text(
        (int(256 * scale) - tw / 2, int(355 * scale) - th / 2 - bbox[1]),
        label,
        font=font_pdf,
        fill=(255, 255, 255, 255),
    )
    return canvas


def main() -> None:
    icon = draw_icon(512)
    icon.save(ASSETS / "logo_512.png")
    icon.resize((128, 128), Image.Resampling.LANCZOS).save(ASSETS / "logo_128.png")
    icon.resize((64, 64), Image.Resampling.LANCZOS).save(ASSETS / "logo_64.png")
    icon.save(
        ASSETS / "app.ico",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )
    print("Wrote", ASSETS / "app.ico")


if __name__ == "__main__":
    main()
