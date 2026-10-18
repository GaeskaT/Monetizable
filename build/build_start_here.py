import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pymupdf

ROOT = Path(__file__).resolve().parent.parent
BUILD = ROOT / "build"
DELIVERY = ROOT / "delivery"
QA = BUILD / "qa"
CSS = BUILD / "sermon-buyer.css"
TEMPLATE = BUILD / "start-here.template.html"
OUT_HTML = ROOT / "Preach-the-Forgotten-Sermon-Kit-1-Start-Here.html"
OUT_PDF = ROOT / "Preach-the-Forgotten-Sermon-Kit-1-Start-Here.pdf"
DELIVERY_PDF = DELIVERY / "Preach-the-Forgotten-Sermon-Kit-1-Start-Here.pdf"
REPORT = BUILD / "sermon-start-here-report.json"
CHROME = Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe")

KIT = {}
for _index, _arg in enumerate(sys.argv):
    if _arg == "--config" and _index + 1 < len(sys.argv):
        KIT = json.loads(Path(sys.argv[_index + 1]).read_text(encoding="utf-8"))

COMMON_REQUIRED = [
    "Start here",
    "From the folder to the pulpit",
    "The short version",
    "What’s in the folder",
    "00 · Start Here",
    "Preaching Manuscript",
    "One-Page Outline",
    "Small-Group Guide",
    "Slides (PPTX)",
    "Slides (PNG)",
    "Preaching License",
    "Make it yours",
    "Work in the DOCX, not the PDF",
    "Keep the evidence",
    "Take it to the pulpit",
    "The license at a glance",
    "One license covers one congregation or campus",
    "does not expire",
    "You may",
    "You may not",
    "Tell me how it went",
    "Do you preach",
    "Which of the nine forgotten figures should I write next?",
    "Where did you hear about this kit?",
    "within 14 days for a refund",
    "Silenced Saints",
]

if KIT:
    STEM = KIT["start_here_stem"]
    OUT_HTML = ROOT / f"{STEM}.html"
    OUT_PDF = ROOT / f"{STEM}.pdf"
    DELIVERY_PDF = DELIVERY / f"{STEM}.pdf"
    REPORT = BUILD / f"kit-{KIT['kit']:02d}-start-here-report.json"
    PREVIEW = QA / f"kit-{KIT['kit']:02d}-start-here.png"
    DOC_TITLE = f"Preach the Forgotten - Sermon Kit {KIT['kit']} - Start Here"
    SERIES = f"Preach the Forgotten · Sermon Kit {KIT['kit']}"
    SUBHEAD = f"Sermon Kit {KIT['kit']}: {KIT['figure']} — “{KIT['title']}”"
    WORD_COUNT = KIT["start_here_words"]
    RUN_ROWS = "\n".join(
        f"              <tr><td>{minutes}</td><td>{section}</td></tr>"
        for minutes, section in KIT["start_here_rows"]
    )
    RUNTIME_NOTE = KIT["start_here_runtime"]
    CUT_NOTE = KIT["start_here_cut"]
    GUARDRAILS = "\n".join(
        f"            <li>{item}</li>" for item in KIT["start_here_guardrails"]
    )
    EVIDENCE_NOTE = KIT["start_here_evidence"]
    META_TITLE = f"Preach the Forgotten - Sermon Kit {KIT['kit']} - Start Here"
    META_SUBJECT = f"Buyer guide for Sermon Kit {KIT['kit']}: {KIT['figure']}"
    META_KEYWORDS = KIT["keywords"]
    REQUIRED_TEXT = COMMON_REQUIRED + [
        SUBHEAD,
        KIT["figure"],
        KIT["title"],
        f"{KIT['start_here_words']} words",
        f"{KIT['start_here_minutes']} minutes",
    ] + [section for _, section in KIT["start_here_rows"]] + KIT.get(
        "start_here_required", []
    )
else:
    PREVIEW = QA / "start-here-final.png"
    DOC_TITLE = "Preach the Forgotten - Sermon Kit 1 - Start Here"
    SERIES = "Preach the Forgotten · Sermon Kit 1"
    SUBHEAD = "Sermon Kit 1: Mary of Bethany — “What the Room Called Waste”"
    WORD_COUNT = "6,653"
    RUN_ROWS = """              <tr><td>5</td><td>Introduction: The smell in the room</td></tr>
              <tr><td>6</td><td>The room full of evidence</td></tr>
              <tr><td>7</td><td>The woman who had been listening</td></tr>
              <tr><td>8</td><td>The jar that could not be saved</td></tr>
              <tr><td>8</td><td>The objection that sounds spiritual</td></tr>
              <tr><td>8</td><td>The burial she understood</td></tr>
              <tr><td>7</td><td>Application: What do you do with the jar?</td></tr>
              <tr><td>3</td><td>The invitation</td></tr>
              <tr><td>5</td><td>Response and closing</td></tr>"""
    RUNTIME_NOTE = "<strong>57 minutes</strong> with the full response; 55–60 with normal variation. The manuscript is 6,653 words at about 125 words per minute."
    CUT_NOTE = "<strong>If you have 50 minutes:</strong> shorten the introduction to 4, Movement 1 to 4, Movement 4 to 6, and the application to 5. Keep the 300-vs-30 contrast and all four applications, tightened. Never cut the three scenes at Jesus’ feet, the jar broken, the burial interpretation, the invitation, or the silence."
    GUARDRAILS = """            <li>This is not a stewardship lecture. Keep the focus on devotion, not other people’s money.</li>
            <li>Hold both halves of Deuteronomy 15:11: care for the poor and costly worship. Do not let the room choose.</li>
            <li>Do not cartoon Judas. His question is real; John names the motive.</li>
            <li>Do not claim more than the text says about how much Mary understood.</li>
            <li>Protect Movements 2, 3, and 5, the invitation, and the response.</li>"""
    EVIDENCE_NOTE = "The 300 denarii, the 30 pieces of silver, the Deuteronomy objection, and the distinction between Mary of Bethany, Mary Magdalene, and Luke 7’s woman are doing theological work."
    META_TITLE = "Preach the Forgotten - Sermon Kit 1 - Start Here"
    META_SUBJECT = "Buyer guide for Sermon Kit 1: Mary of Bethany"
    META_KEYWORDS = (
        "Mary of Bethany, What the Room Called Waste, Preach the Forgotten, "
        "Sermon Kit 1, Silenced Saints"
    )
    REQUIRED_TEXT = COMMON_REQUIRED + [
        "Sermon Kit 1: Mary of Bethany",
        "What the Room Called Waste",
        "6,653 words",
        "Introduction: The smell in the room",
        "The room full of evidence",
        "The woman who had been listening",
        "The jar that could not be saved",
        "The objection that sounds spiritual",
        "The burial she understood",
        "Application: What do you do with the jar?",
        "The invitation",
        "Response and closing",
        "57 minutes",
        "If you have 50 minutes",
        "Never cut the three scenes at Jesus’ feet",
        "Guardrails",
        "Deuteronomy 15:11",
    ]


def fail(msg):
    print(f"BUILD FAILED: {msg}", file=sys.stderr)
    sys.exit(1)


def normalize(text):
    return re.sub(r"\s+", " ", text).strip()


def compact(text):
    return re.sub(r"\s+", "", text).casefold()


def print_pdf(html_path, pdf_path):
    if pdf_path.exists():
        pdf_path.unlink()
    profile = BUILD / ".chrome-profile-start-here"
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
        "--virtual-time-budget=15000",
        f"--print-to-pdf={pdf_path}",
        html_path.as_uri(),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=240)
    if not pdf_path.exists():
        fail(f"Chrome did not produce {pdf_path.name}\n{proc.stderr[-2000:]}")


def verify(pdf_path):
    doc = pymupdf.open(pdf_path)
    if doc.page_count != 1:
        pages = doc.page_count
        doc.close()
        fail(f"expected 1 page, got {pages}")
    page = doc[0]
    width, height = round(page.rect.width), round(page.rect.height)
    if (width, height) != (612, 792):
        doc.close()
        fail(f"expected Letter 612x792, got {width}x{height}")

    text = normalize(page.get_text("text"))
    squeezed = compact(text)
    missing = [n for n in REQUIRED_TEXT if compact(n) not in squeezed]
    if missing:
        doc.close()
        fail(f"missing text {missing}")
    if "\ufffd" in text:
        doc.close()
        fail("replacement characters found")

    blocks = page.get_text("blocks")
    max_y = max(b[3] for b in blocks)
    min_x = min(b[0] for b in blocks)
    max_x = max(b[2] for b in blocks)
    if max_y > 772:
        doc.close()
        fail(f"text reaches y={max_y:.1f}")
    if min_x < 30 or max_x > 584:
        doc.close()
        fail(f"text outside safe area {min_x:.1f}-{max_x:.1f}")

    fonts = sorted({font[3] for font in page.get_fonts(full=True)})
    lowered = " ".join(fonts).lower()
    for family in ("georgia", "candara"):
        if family not in lowered:
            doc.close()
            fail(f"{family} not embedded; fonts={fonts}")

    doc.close()
    return {
        "pages": 1,
        "page_size": [612, 792],
        "words": len(text.split()),
        "fonts": fonts,
        "max_text_y": round(max_y, 1),
        "text_bounds_x": [round(min_x, 1), round(max_x, 1)],
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
    for path in (CSS, TEMPLATE):
        if not path.exists():
            fail(f"missing build input: {path}")

    css = CSS.read_text(encoding="utf-8")
    html = (
        TEMPLATE.read_text(encoding="utf-8")
        .replace("{{CSS}}", css)
        .replace("{{DOC_TITLE}}", DOC_TITLE)
        .replace("{{SERIES}}", SERIES)
        .replace("{{SUBHEAD}}", SUBHEAD)
        .replace("{{WORD_COUNT}}", WORD_COUNT)
        .replace("{{EVIDENCE_NOTE}}", EVIDENCE_NOTE)
        .replace("{{GUARDRAILS}}", GUARDRAILS)
        .replace("{{RUN_ROWS}}", RUN_ROWS)
        .replace("{{RUNTIME_NOTE}}", RUNTIME_NOTE)
        .replace("{{CUT_NOTE}}", CUT_NOTE)
        .replace("{{FIGURE}}", KIT.get("figure", "Mary of Bethany"))
    )
    OUT_HTML.write_text(html, encoding="utf-8")

    print_pdf(OUT_HTML, OUT_PDF)
    verify(OUT_PDF)
    set_metadata(OUT_PDF)
    final = verify(OUT_PDF)
    preview = render_preview(OUT_PDF)
    DELIVERY.mkdir(exist_ok=True)
    shutil.copy2(OUT_PDF, DELIVERY_PDF)

    report = {
        "file": str(OUT_PDF),
        "html": str(OUT_HTML),
        "delivery": str(DELIVERY_PDF),
        **final,
        "preview": str(preview),
        "pdf_kb": round(OUT_PDF.stat().st_size / 1024, 1),
        "html_kb": round(OUT_HTML.stat().st_size / 1024, 1),
        "status": "passed",
    }
    REPORT.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report, indent=2))


if __name__ == "__main__":
    main()
