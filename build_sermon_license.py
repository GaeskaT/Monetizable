import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pymupdf

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "build" / "sermon-license.template.html"
CSS = ROOT / "build" / "church-license.css"
QA = ROOT / "build" / "qa"
CHROME = Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe")

KIT = {}
for _index, _arg in enumerate(sys.argv):
    if _arg == "--config" and _index + 1 < len(sys.argv):
        KIT = json.loads(Path(sys.argv[_index + 1]).read_text(encoding="utf-8"))

REQUIRED_COMMON = [
    "Preaching License",
    "Permission to preach, edit, and record",
    "Licensed to",
    "Church / campus",
    "Order number",
    "The grant",
    "Permitted",
    "Not permitted",
    "One license covers one congregation or campus",
    "Multi-campus, paid-event, and publishing use require written permission",
    "This license does not expire",
    "Silenced Saints. All rights reserved",
    "First edition",
    "non-exclusive, non-transferable, perpetual license",
    "edit the manuscript for local context and your own voice",
    "record and stream the sermon",
    "print the outline and small-group guide",
    "Your edits are yours",
    "underlying manuscript remains the property of Silenced Saints",
    "No attribution is required",
    "Sharing the files with preachers outside your congregation",
    "Posting the manuscript publicly or selling it",
    "Using one license across multiple campuses",
    "paid conference or event without written permission",
]

if KIT:
    STEM = f"{KIT['prefix']}-Preaching-License"
    OUT_HTML = ROOT / f"{STEM}.html"
    OUT_PDF = ROOT / f"{STEM}.pdf"
    DELIVERY_PDF = ROOT / "delivery" / f"{STEM}.pdf"
    REPORT = ROOT / "build" / f"kit-{KIT['kit']:02d}-license-report.json"
    PREVIEW = QA / f"kit-{KIT['kit']:02d}-license.png"
    DOC_TITLE = f"Preach the Forgotten - Sermon Kit {KIT['kit']} - Preaching License"
    SERIES = f"Preach the Forgotten · Sermon Kit {KIT['kit']}"
    INTRO_KIT = f"Sermon Kit {KIT['kit']}: {KIT['figure']}"
    SERMON_TITLE = KIT["license_title"]
    FOOTER_KIT = f"Sermon Kit {KIT['kit']}"
    REQUIRED_TEXT = REQUIRED_COMMON + [
        f"Sermon Kit {KIT['kit']}",
        KIT["license_title"],
        KIT["figure"],
    ]
    META_TITLE = f"Preach the Forgotten - Sermon Kit {KIT['kit']} - Preaching License"
    META_SUBJECT = f"One-page preaching license for Sermon Kit {KIT['kit']}: {KIT['figure']}"
    META_KEYWORDS = KIT["keywords"]
else:
    OUT_HTML = ROOT / "Preach-the-Forgotten-Sermon-Kit-1-Preaching-License.html"
    OUT_PDF = ROOT / "Preach-the-Forgotten-Sermon-Kit-1-Preaching-License.pdf"
    DELIVERY_PDF = ROOT / "delivery" / "Preach-the-Forgotten-Sermon-Kit-1-Preaching-License.pdf"
    REPORT = ROOT / "build" / "sermon-license-report.json"
    PREVIEW = QA / "sermon-license.png"
    DOC_TITLE = "Preach the Forgotten - Sermon Kit 1 - Preaching License"
    SERIES = "Preach the Forgotten · Sermon Kit 1"
    INTRO_KIT = "Sermon Kit 1: Mary of Bethany"
    SERMON_TITLE = "What the Room Called Waste"
    FOOTER_KIT = "Sermon Kit 1"
    REQUIRED_TEXT = REQUIRED_COMMON + ["Sermon Kit 1", "What the Room Called Waste"]
    META_TITLE = "Preach the Forgotten - Sermon Kit 1 - Preaching License"
    META_SUBJECT = "One-page preaching license for Sermon Kit 1: Mary of Bethany"
    META_KEYWORDS = "preaching license, sermon license, Mary of Bethany, What the Room Called Waste, Preach the Forgotten, Silenced Saints"


def fail(msg):
    print(f"BUILD FAILED: {msg}", file=sys.stderr)
    sys.exit(1)


def normalize(text):
    return re.sub(r"\s+", " ", text).strip()


def print_pdf(html_path, pdf_path):
    if pdf_path.exists():
        pdf_path.unlink()
    profile = ROOT / "build" / ".chrome-profile-sermon-license"
    profile.mkdir(exist_ok=True)
    cmd = [
        str(CHROME),
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-extensions",
        f"--user-data-dir={profile}",
        "--no-pdf-header-footer",
        "--run-all-compositor-stages-before-draw",
        "--virtual-time-budget=20000",
        f"--print-to-pdf={pdf_path}",
        html_path.as_uri(),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=240)
    if not pdf_path.exists():
        fail(f"Chrome did not produce {pdf_path.name}\n{proc.stderr[-2000:]}")


def verify(pdf_path):
    doc = pymupdf.open(pdf_path)
    if doc.page_count != 1:
        doc.close()
        fail(f"expected 1 page, got {doc.page_count}")

    page = doc[0]
    width, height = round(page.rect.width), round(page.rect.height)
    if (width, height) != (612, 792):
        doc.close()
        fail(f"expected Letter 612x792, got {width}x{height}")

    text = normalize(page.get_text("text"))
    compact_text = re.sub(r"\s+", "", text).casefold()
    missing = [
        needle
        for needle in REQUIRED_TEXT
        if re.sub(r"\s+", "", needle).casefold() not in compact_text
    ]
    if missing:
        doc.close()
        fail(f"missing text: {missing}")
    if "\ufffd" in text:
        doc.close()
        fail("replacement characters found in extracted text")

    blocks = page.get_text("blocks")
    max_y = max(block[3] for block in blocks)
    min_x = min(block[0] for block in blocks)
    max_x = max(block[2] for block in blocks)
    if max_y > 752:
        doc.close()
        fail(f"text reaches y={max_y:.1f}, too close to the frame")
    if min_x < 24 or max_x > 588:
        doc.close()
        fail(f"text outside horizontal safe area: {min_x:.1f}–{max_x:.1f}")

    fonts = [font[3] for font in page.get_fonts(full=True)]
    lowered = " ".join(fonts).lower()
    if "georgia" not in lowered:
        doc.close()
        fail(f"Georgia not embedded; fonts: {fonts}")
    if "candara" not in lowered:
        doc.close()
        fail(f"Candara not embedded; fonts: {fonts}")

    doc.close()
    return {
        "text": text,
        "words": len(text.split()),
        "max_text_y": round(max_y, 1),
        "min_text_x": round(min_x, 1),
        "max_text_x": round(max_x, 1),
        "fonts": sorted(set(fonts)),
    }


def set_metadata(pdf_path):
    doc = pymupdf.open(pdf_path)
    doc.set_metadata(
        {
            "title": META_TITLE,
            "author": "Silenced Saints",
            "subject": META_SUBJECT,
            "keywords": META_KEYWORDS,
            "creator": "Silenced Saints",
            "producer": "Silenced Saints build pipeline",
        }
    )
    tmp = pdf_path.with_suffix(".meta.pdf")
    doc.save(tmp, garbage=4, deflate=True)
    doc.close()
    shutil.move(tmp, pdf_path)


def render_preview(pdf_path):
    QA.mkdir(exist_ok=True)
    png = PREVIEW
    doc = pymupdf.open(pdf_path)
    pix = doc[0].get_pixmap(dpi=140)
    pix.save(png)
    doc.close()
    return png


def main():
    if not CHROME.exists():
        fail(f"Chrome not found at {CHROME}")
    if not TEMPLATE.exists():
        fail(f"template not found: {TEMPLATE}")
    if not CSS.exists():
        fail(f"CSS not found: {CSS}")

    css = CSS.read_text(encoding="utf-8")
    html = (
        TEMPLATE.read_text(encoding="utf-8")
        .replace("{{CSS}}", css)
        .replace("{{DOC_TITLE}}", DOC_TITLE)
        .replace("{{SERIES}}", SERIES)
        .replace("{{INTRO_KIT}}", INTRO_KIT)
        .replace("{{SERMON_TITLE}}", SERMON_TITLE)
        .replace("{{FOOTER_KIT}}", FOOTER_KIT)
    )
    OUT_HTML.write_text(html, encoding="utf-8")

    print_pdf(OUT_HTML, OUT_PDF)
    verify(OUT_PDF)
    set_metadata(OUT_PDF)
    final = verify(OUT_PDF)
    png = render_preview(OUT_PDF)
    DELIVERY_PDF.parent.mkdir(exist_ok=True)
    shutil.copy2(OUT_PDF, DELIVERY_PDF)

    report = {
        "file": str(OUT_PDF),
        "html": str(OUT_HTML),
        "delivery": str(DELIVERY_PDF),
        "pages": 1,
        "page_size": [612, 792],
        "words": final["words"],
        "fonts": final["fonts"],
        "max_text_y": final["max_text_y"],
        "text_bounds_x": [final["min_text_x"], final["max_text_x"]],
        "missing_text": [],
        "pdf_kb": round(OUT_PDF.stat().st_size / 1024, 1),
        "html_kb": round(OUT_HTML.stat().st_size / 1024, 1),
        "preview": str(png),
        "checks": {
            "single_page": True,
            "letter_size": True,
            "required_text": True,
            "fonts_embedded": True,
            "inside_frame": True,
            "no_replacement_chars": True,
        },
    }
    REPORT.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report, indent=2))


if __name__ == "__main__":
    main()
