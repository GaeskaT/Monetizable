import json
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont, PngImagePlugin

ROOT = Path(__file__).resolve().parent.parent
ART = ROOT / "assets" / "cover.png"
OUT = ROOT / "assets" / "listing"
QA = ROOT / "build" / "qa"
FONTS = Path("C:/Windows/Fonts")

NAVY = (16, 27, 46)
NAVY_TOP = (22, 37, 62)
NAVY_BOTTOM = (10, 18, 31)
CREAM = (244, 236, 220)
CREAM_SOFT = (216, 205, 185)
GOLD = (200, 161, 94)
MUTED = (184, 174, 157)

KIT = {}
for _index, _arg in enumerate(sys.argv):
    if _arg == "--config" and _index + 1 < len(sys.argv):
        KIT = json.loads(Path(sys.argv[_index + 1]).read_text(encoding="utf-8"))

if KIT:
    KIT_NO = KIT["kit"]
    ART = ROOT / KIT["cover_art"]
    FIGURE = KIT["figure"]
    TITLE_LINES = KIT["cover_lines"]
    QUOTE_LINES = KIT["quote_lines"]
    QUOTE_ATTR = KIT.get("quote_attribution", f"{FIGURE} · {KIT['primary_text']}")
    PRIMARY = KIT["primary_text"].upper()
    OUT_PREFIX = f"Sermon-Kit-{KIT_NO}"
    KIT_LABEL = f"SERMON KIT {KIT_NO}"
    TITLE_LABEL = KIT["title"].upper()
    COVER_RUN = KIT.get("cover_run", f"A complete {KIT['start_here_minutes']}-minute sermon")
else:
    KIT_NO = 1
    FIGURE = "Mary of Bethany"
    TITLE_LINES = ["What the Room", "Called Waste"]
    QUOTE_LINES = ["The room called it waste.", "Jesus called it beautiful."]
    QUOTE_ATTR = "Mary of Bethany · John 12:1–8"
    PRIMARY = "JOHN 12:1–8"
    OUT_PREFIX = "Sermon-Kit-1"
    KIT_LABEL = "SERMON KIT 1"
    TITLE_LABEL = "WHAT THE ROOM CALLED WASTE"
    COVER_RUN = "A complete 57-minute sermon"


def font(name, size):
    return ImageFont.truetype(str(FONTS / name), size)


def text_width(draw, text, fnt, tracking=0):
    if not text:
        return 0
    return sum(draw.textlength(ch, font=fnt) for ch in text) + tracking * (len(text) - 1)


def tracked(draw, text, fnt, tracking, fill, xy, align="left"):
    total = text_width(draw, text, fnt, tracking)
    x, y = xy
    if align == "center":
        x -= total / 2
    elif align == "right":
        x -= total
    for ch in text:
        draw.text((x, y), ch, font=fnt, fill=fill, anchor="lt")
        x += draw.textlength(ch, font=fnt) + tracking


def fit(draw, lines, name, max_width, start, minimum, tracking=0):
    for size in range(start, minimum - 1, -2):
        fnt = font(name, size)
        if all(text_width(draw, line, fnt, tracking) <= max_width for line in lines):
            return size, fnt
    fnt = font(name, minimum)
    return minimum, fnt


def gradient(width, height, top, bottom):
    strip = Image.new("RGB", (1, height))
    for y in range(height):
        t = y / max(height - 1, 1)
        strip.putpixel(
            (0, y),
            tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3)),
        )
    return strip.resize((width, height), Image.BILINEAR)


def alpha_gradient(width, height, top_alpha, bottom_alpha, color):
    strip = Image.new("L", (1, height))
    for y in range(height):
        t = y / max(height - 1, 1)
        strip.putpixel((0, y), int(top_alpha + (bottom_alpha - top_alpha) * t))
    mask = strip.resize((width, height), Image.BILINEAR)
    layer = Image.new("RGBA", (width, height), (*color, 255))
    layer.putalpha(mask)
    return layer


def crop_art(width, height, focus_x=0.5, focus_y=0.5):
    art = Image.open(ART).convert("RGB")
    scale = max(width / art.width, height / art.height)
    scaled = art.resize(
        (max(width, int(round(art.width * scale))), max(height, int(round(art.height * scale)))),
        Image.LANCZOS,
    )
    left = int(round((scaled.width - width) * focus_x))
    top = int(round((scaled.height - height) * focus_y))
    return scaled.crop((left, top, left + width, top + height))


def badge(draw, x, y, w, h, label, size=12, tracking=2.5):
    draw.rectangle([x, y, x + w, y + h], fill=NAVY, outline=GOLD, width=2)
    fnt = font("candarab.ttf", size)
    tw = text_width(draw, label, fnt, tracking)
    tracked(draw, label, fnt, tracking, GOLD, (x + (w - tw) / 2, y + (h - size * 1.35) / 2))


def save(image, path, title, description):
    OUT.mkdir(parents=True, exist_ok=True)
    info = PngImagePlugin.PngInfo()
    info.add_text("Title", title)
    info.add_text("Author", "Silenced Saints")
    info.add_text("Description", description)
    info.add_text("Software", "Silenced Saints build pipeline")
    image.convert("RGB").save(path, "PNG", optimize=True, pnginfo=info)
    print(f"{path.name}: {image.width}x{image.height}, {path.stat().st_size / 1024:.0f} KB")


def build_cover():
    width, height = 1600, 900
    split = 860
    canvas = gradient(width, height, NAVY_TOP, NAVY_BOTTOM).convert("RGBA")
    canvas.paste(crop_art(width - split, height, focus_y=0.5), (split, 0))
    draw = ImageDraw.Draw(canvas)
    draw.line([(split, 0), (split, height)], fill=GOLD, width=2)

    tracked(draw, "SILENCED SAINTS", font("candarab.ttf", 17), 5, GOLD, (84, 54))
    badge(draw, 638, 46, 156, 40, KIT_LABEL, 12, 2.5)
    tracked(draw, "PREACH THE FORGOTTEN", font("candara.ttf", 11.5), 4, MUTED, (84, 88))
    tracked(draw, "A SERMON FOR THE PULPIT", font("candarab.ttf", 13.5), 5, GOLD, (84, 180))

    size, title_font = fit(draw, TITLE_LINES, "georgia.ttf", 720, 96, 72)
    line_h = int(round(size * 1.1))
    draw.text((84, 222), TITLE_LINES[0], font=title_font, fill=CREAM)
    draw.text((84, 222 + line_h), TITLE_LINES[1], font=title_font, fill=CREAM)

    subtitle_y = 222 + 2 * line_h + 18
    draw.text((84, subtitle_y), FIGURE, font=font("georgiai.ttf", 38), fill=GOLD)
    rule_y = subtitle_y + 70
    draw.rectangle([84, rule_y, 424, rule_y + 2], fill=GOLD)

    desc_y = rule_y + 30
    draw.text(
        (84, desc_y),
        COVER_RUN,
        font=font("candara.ttf", 21),
        fill=CREAM,
    )
    tracked(
        draw,
        "MANUSCRIPT · OUTLINE · 12 SLIDES",
        font("candara.ttf", 14),
        1.2,
        CREAM_SOFT,
        (84, desc_y + 46),
    )
    tracked(
        draw,
        "SMALL-GROUP GUIDE · PREACHING LICENSE",
        font("candara.ttf", 14),
        1.2,
        CREAM_SOFT,
        (84, desc_y + 72),
    )
    draw.text(
        (84, desc_y + 122),
        "The sermon, not the study.",
        font=font("georgiai.ttf", 17),
        fill=MUTED,
    )

    tracked(draw, PRIMARY, font("candarab.ttf", 13), 2, GOLD, (84, 810))
    tracked(
        draw,
        "FIRST EDITION · OCTOBER 2026",
        font("candara.ttf", 11),
        2,
        MUTED,
        (84, 838),
    )

    save(
        canvas,
        OUT / f"{OUT_PREFIX}-Cover.png",
        f"{KIT['title'] if KIT else 'What the Room Called Waste'} - Sermon Kit {KIT_NO} Cover",
        f"Landscape listing cover for Preach the Forgotten Sermon Kit {KIT_NO}: {FIGURE}.",
    )


def build_title_promo():
    width, height = 1280, 720
    split = 700
    canvas = gradient(width, height, NAVY_TOP, NAVY_BOTTOM).convert("RGBA")
    canvas.paste(crop_art(width - split, height, focus_y=0.46), (split, 0))
    draw = ImageDraw.Draw(canvas)
    draw.line([(split, 0), (split, height)], fill=GOLD, width=2)

    tracked(draw, "SILENCED SAINTS", font("candarab.ttf", 13.5), 4, GOLD, (64, 42))
    badge(draw, 506, 35, 130, 33, KIT_LABEL, 10.5, 2)
    tracked(draw, "PREACH THE FORGOTTEN", font("candara.ttf", 10.5), 3, MUTED, (64, 68))
    tracked(draw, "A SERMON FOR THE PULPIT", font("candarab.ttf", 11), 4, GOLD, (64, 138))

    size, title_font = fit(draw, TITLE_LINES, "georgia.ttf", 600, 76, 58)
    line_h = int(round(size * 1.1))
    draw.text((64, 174), TITLE_LINES[0], font=title_font, fill=CREAM)
    draw.text((64, 174 + line_h), TITLE_LINES[1], font=title_font, fill=CREAM)

    subtitle_y = 174 + 2 * line_h + 12
    draw.text((64, subtitle_y), FIGURE, font=font("georgiai.ttf", 29), fill=GOLD)
    rule_y = subtitle_y + 52
    draw.rectangle([64, rule_y, 334, rule_y + 2], fill=GOLD)

    desc_y = rule_y + 22
    draw.text(
        (64, desc_y),
        COVER_RUN,
        font=font("candara.ttf", 17),
        fill=CREAM,
    )
    tracked(
        draw,
        "MANUSCRIPT · OUTLINE · 12 SLIDES · GUIDE · LICENSE",
        font("candara.ttf", 11.5),
        0.8,
        CREAM_SOFT,
        (64, desc_y + 36),
    )
    draw.text(
        (64, desc_y + 70),
        "The sermon, not the study.",
        font=font("georgiai.ttf", 14),
        fill=MUTED,
    )

    tracked(draw, PRIMARY, font("candarab.ttf", 10.5), 1.8, GOLD, (64, 642))
    tracked(
        draw,
        "FIRST EDITION · OCTOBER 2026",
        font("candara.ttf", 9),
        1.8,
        MUTED,
        (64, 662),
    )

    save(
        canvas,
        OUT / f"{OUT_PREFIX}-Promo-Title.png",
        f"{KIT['title'] if KIT else 'What the Room Called Waste'} - Sermon Kit {KIT_NO} Promo",
        f"16:9 title promo for Preach the Forgotten Sermon Kit {KIT_NO}: {FIGURE}.",
    )


def build_square():
    width, height = 1080, 1080
    image_h = 620
    canvas = gradient(width, height, NAVY_TOP, NAVY_BOTTOM).convert("RGBA")
    art = crop_art(width, image_h, focus_y=0.62)
    canvas.paste(art, (0, 0))
    canvas.alpha_composite(alpha_gradient(width, 175, 205, 0, NAVY), (0, 0))
    draw = ImageDraw.Draw(canvas)

    tracked(draw, "SILENCED SAINTS", font("candarab.ttf", 14), 4, CREAM, (54, 40))
    badge(draw, 868, 32, 158, 38, KIT_LABEL, 10.5, 2.2)
    draw.line([(0, image_h), (width, image_h)], fill=GOLD, width=2)

    tracked(
        draw,
        "A SERMON FOR THE PULPIT",
        font("candarab.ttf", 12.5),
        4,
        GOLD,
        (width / 2, 666),
        "center",
    )

    size, title_font = fit(draw, TITLE_LINES, "georgia.ttf", 940, 66, 52)
    line_h = int(round(size * 1.1))
    draw.text((width / 2, 704), TITLE_LINES[0], font=title_font, fill=CREAM, anchor="ma")
    draw.text((width / 2, 704 + line_h), TITLE_LINES[1], font=title_font, fill=CREAM, anchor="ma")

    subtitle_y = 704 + 2 * line_h + 12
    draw.text((width / 2, subtitle_y), FIGURE, font=font("georgiai.ttf", 26), fill=GOLD, anchor="ma")
    rule_y = subtitle_y + 46
    draw.rectangle([width / 2 - 48, rule_y, width / 2 + 48, rule_y + 2], fill=GOLD)

    draw.text(
        (width / 2, rule_y + 22),
        COVER_RUN,
        font=font("candara.ttf", 13.5),
        fill=CREAM,
        anchor="ma",
    )
    tracked(
        draw,
        "MANUSCRIPT · OUTLINE · 12 SLIDES · SMALL-GROUP GUIDE · LICENSE",
        font("candara.ttf", 11.5),
        0.8,
        MUTED,
        (width / 2, rule_y + 50),
        "center",
    )
    tracked(
        draw,
        "SILENCED SAINTS · PREACH THE FORGOTTEN",
        font("candara.ttf", 9.5),
        2.5,
        GOLD,
        (width / 2, 1030),
        "center",
    )

    save(
        canvas,
        OUT / f"{OUT_PREFIX}-Promo-Square.png",
        f"{KIT['title'] if KIT else 'What the Room Called Waste'} - Sermon Kit {KIT_NO} Square",
        f"Square social promo for Preach the Forgotten Sermon Kit {KIT_NO}: {FIGURE}.",
    )


def build_quote():
    width, height = 1080, 1080
    canvas = gradient(width, height, (24, 40, 66), NAVY_BOTTOM).convert("RGBA")

    glow = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    gd = ImageDraw.Draw(glow)
    gd.ellipse([130, 80, 950, 900], fill=(*GOLD, 22))
    canvas.alpha_composite(glow.filter(ImageFilter.GaussianBlur(150)))

    draw = ImageDraw.Draw(canvas)
    draw.rectangle([34, 34, width - 34, height - 34], outline=GOLD, width=2)

    tracked(draw, "SILENCED SAINTS", font("candarab.ttf", 12), 4, GOLD, (84, 78))
    tracked(
        draw,
        KIT_LABEL,
        font("candara.ttf", 10.5),
        3,
        MUTED,
        (width - 84, 80),
        "right",
    )

    draw.rectangle([width / 2 - 34, 258, width / 2 + 34, 260], fill=GOLD)

    tracked(
        draw,
        "A LINE FROM THE SERMON",
        font("candarab.ttf", 12.5),
        4,
        GOLD,
        (width / 2, 292),
        "center",
    )

    size, quote_font = fit(draw, QUOTE_LINES, "georgia.ttf", 940, 88, 58)
    line_h = int(round(size * 1.2))
    quote_y = 414
    draw.text((width / 2, quote_y), QUOTE_LINES[0], font=quote_font, fill=CREAM, anchor="ma")
    draw.text((width / 2, quote_y + line_h), QUOTE_LINES[1], font=quote_font, fill=CREAM, anchor="ma")

    rule_y = quote_y + 2 * line_h + 34
    draw.rectangle([width / 2 - 52, rule_y, width / 2 + 52, rule_y + 2], fill=GOLD)

    tracked(
        draw,
        TITLE_LABEL,
        font("candarab.ttf", 12.5),
        3,
        GOLD,
        (width / 2, rule_y + 30),
        "center",
    )
    draw.text(
        (width / 2, rule_y + 58),
        QUOTE_ATTR,
        font=font("candara.ttf", 12),
        fill=MUTED,
        anchor="ma",
    )

    tracked(
        draw,
        "SILENCED SAINTS · PREACH THE FORGOTTEN",
        font("candara.ttf", 9),
        2.5,
        MUTED,
        (width / 2, 980),
        "center",
    )

    save(
        canvas,
        OUT / f"{OUT_PREFIX}-Promo-Quote.png",
        f"{KIT['title'] if KIT else 'What the Room Called Waste'} - Quote Card",
        f"Square quote card from Preach the Forgotten Sermon Kit {KIT_NO}: {FIGURE}.",
    )


def build_contact_sheet():
    cover = Image.open(OUT / f"{OUT_PREFIX}-Cover.png")
    title = Image.open(OUT / f"{OUT_PREFIX}-Promo-Title.png")
    square = Image.open(OUT / f"{OUT_PREFIX}-Promo-Square.png")
    quote = Image.open(OUT / f"{OUT_PREFIX}-Promo-Quote.png")

    pad = 40
    gap = 28
    cell_w = 720
    cover_h = int(cover.height * cell_w / cover.width)
    title_h = int(title.height * cell_w / title.width)
    square_h = 420
    quote_h = 420
    sheet_w = pad * 2 + cell_w * 2 + gap
    sheet_h = pad * 2 + cover_h + gap + max(square_h, quote_h)
    sheet = gradient(sheet_w, sheet_h, (13, 22, 38), (8, 14, 25)).convert("RGB")

    sheet.paste(cover.resize((cell_w, cover_h), Image.LANCZOS), (pad, pad))
    sheet.paste(title.resize((cell_w, title_h), Image.LANCZOS), (pad + cell_w + gap, pad))
    sheet.paste(
        square.resize((square_h, square_h), Image.LANCZOS),
        (pad + (cell_w - square_h) // 2, pad + cover_h + gap),
    )
    sheet.paste(
        quote.resize((quote_h, quote_h), Image.LANCZOS),
        (pad + cell_w + gap + (cell_w - quote_h) // 2, pad + cover_h + gap),
    )
    name = "sermon-promo-contact-sheet.png" if not KIT else f"kit-{KIT_NO:02d}-promo-contact-sheet.png"
    sheet.save(QA / name)
    print(f"{name}: {sheet.width}x{sheet.height}")


if __name__ == "__main__":
    QA.mkdir(parents=True, exist_ok=True)
    build_cover()
    build_title_promo()
    build_square()
    build_quote()
    build_contact_sheet()
