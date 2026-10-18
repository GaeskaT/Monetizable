import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

import markdown
import pymupdf

from typography import smarten_markdown

ROOT = Path(__file__).resolve().parent.parent
BUILD = ROOT / "build"
DELIVERY = ROOT / "delivery"
QA = BUILD / "qa"
CSS = BUILD / "sermon-buyer.css"
TEMPLATE = BUILD / "sermon-onepager.template.html"
CHROME = Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe")

JOBS = [
    {
        "key": "outline",
        "source": ROOT / "SERMON-OUTLINE-WHAT-THE-ROOM-CALLED-WASTE.md",
        "stem": "Preach-the-Forgotten-Sermon-Kit-1-One-Page-Outline",
        "kicker": "One-page preaching outline",
        "subtitle": "Mary of Bethany · John 12:1–8",
        "footer_right": "One-page outline · Mary of Bethany",
        "pdf_title": "What the Room Called Waste - One-Page Preaching Outline",
        "required": [
            "What the Room Called Waste",
            "One-page preaching outline",
            "Primary text:",
            "Movement map",
            "Introduction: The smell in the room",
            "The room full of evidence",
            "The woman who had been listening",
            "The jar that could not be saved",
            "The objection that sounds spiritual",
            "The burial she understood",
            "Application: What do you do with the jar?",
            "The invitation",
            "Response and closing",
            "Lines to land",
            "Guardrails",
            "If you must cut",
            "The room called it waste. Jesus called it beautiful.",
            "57 minutes",
        ],
        "colgroup": True,
        "preview": "sermon-outline-final.png",
    },
    {
        "key": "guide",
        "source": ROOT / "SERMON-SMALL-GROUP-GUIDE-WHAT-THE-ROOM-CALLED-WASTE.md",
        "stem": "Preach-the-Forgotten-Sermon-Kit-1-Small-Group-Guide",
        "kicker": "Small-group guide",
        "subtitle": "Mary of Bethany · John 12:1–8",
        "footer_right": "Small-group guide · Mary of Bethany",
        "pdf_title": "What the Room Called Waste - Small-Group Guide",
        "required": [
            "What the Room Called Waste",
            "Small-group guide",
            "John 12:1–8",
            "Before you meet",
            "Open",
            "Observe",
            "Interpret",
            "Apply",
            "Close",
            "Leader notes",
            "Not a money talk",
            "The poor",
            "Judas",
            "Who this Mary is",
            "Grief and safety",
            "The gospel",
            "come closer",
        ],
        "colgroup": False,
        "preview": "sermon-guide-final.png",
    },
]


KIT = {}
for _index, _arg in enumerate(sys.argv):
    if _arg == "--config" and _index + 1 < len(sys.argv):
        KIT = json.loads(Path(sys.argv[_index + 1]).read_text(encoding="utf-8"))

if KIT:
    JOBS = [
        {
            "key": "outline",
            "source": ROOT / KIT["source_outline"],
            "stem": f"{KIT['prefix']}-One-Page-Outline",
            "kicker": "One-page preaching outline",
            "subtitle": KIT["outline_subtitle"],
            "footer_right": KIT["outline_footer_right"],
            "pdf_title": KIT["outline_pdf_title"],
            "required": KIT["outline_required"],
            "colgroup": True,
            "preview": f"kit-{KIT['kit']:02d}-outline.png",
        },
        {
            "key": "guide",
            "source": ROOT / KIT["source_guide"],
            "stem": f"{KIT['prefix']}-Small-Group-Guide",
            "kicker": "Small-group guide",
            "subtitle": KIT["guide_subtitle"],
            "footer_right": KIT["guide_footer_right"],
            "pdf_title": KIT["guide_pdf_title"],
            "required": KIT["guide_required"],
            "colgroup": False,
            "preview": f"kit-{KIT['kit']:02d}-guide.png",
        },
    ]
    REPORT = BUILD / f"kit-{KIT['kit']:02d}-onepagers-report.json"
else:
    REPORT = BUILD / "sermon-onepagers-report.json"


def fail(msg):
    print(f"BUILD FAILED: {msg}", file=sys.stderr)
    sys.exit(1)


def normalize(text):
    return re.sub(r"\s+", " ", text).strip()


def compact(text):
    return re.sub(r"\s+", "", text).casefold()


def parse_source(path):
    text = path.read_text(encoding="utf-8")
    lines = text.splitlines()
    if not lines or not lines[0].startswith("# "):
        fail(f"{path.name}: missing H1")
    title = lines[0][2:].strip()
    h2_index = next((i for i, line in enumerate(lines) if line.startswith("## ")), None)
    if h2_index is None:
        fail(f"{path.name}: missing H2 sections")
    meta_md = "\n".join(
        line for line in lines[1:h2_index] if not line.startswith("### ")
    ).strip()
    body_md = "\n".join(lines[h2_index:]).strip()
    return title, meta_md, body_md


def render_html(job, css, template):
    title, meta_md, body_md = parse_source(job["source"])
    meta = markdown.markdown(smarten_markdown(meta_md), extensions=["tables", "sane_lists"])
    content = markdown.markdown(smarten_markdown(body_md), extensions=["tables", "sane_lists"])
    if job["colgroup"]:
        content = content.replace(
            "<table>",
            '<table><colgroup><col style="width:0.42in">'
            '<col style="width:1.38in"><col style="width:5.6in"></colgroup>',
            1,
        )
    html = (
        template.replace("{{CSS}}", css)
        .replace("{{MASTHEAD}}", KIT.get("masthead_right", "PREACH THE FORGOTTEN · SERMON KIT 1"))
        .replace("{{DOC_TITLE}}", f"{title} — {job['kicker']}")
        .replace("{{KICKER}}", job["kicker"])
        .replace("{{TITLE}}", title)
        .replace("{{SUBTITLE}}", job["subtitle"])
        .replace("{{META}}", meta.strip())
        .replace("{{CONTENT}}", content.strip())
        .replace("{{FOOTER_LEFT}}", KIT.get("footer_left_short", "Preach the Forgotten · Sermon Kit 1"))
        .replace("{{FOOTER_RIGHT}}", job["footer_right"])
    )
    out = ROOT / f"{job['stem']}.html"
    out.write_text(html, encoding="utf-8")
    return out


def print_pdf(html_path, pdf_path):
    if pdf_path.exists():
        pdf_path.unlink()
    profile = BUILD / ".chrome-profile-sermon-buyers"
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


def verify(job, pdf_path):
    doc = pymupdf.open(pdf_path)
    if doc.page_count != len(job.get("expected_pages", [1])):
        pages = doc.page_count
        doc.close()
        fail(f"{job['key']}: expected 1 page, got {pages}")
    page = doc[0]
    width, height = round(page.rect.width), round(page.rect.height)
    if (width, height) != (612, 792):
        doc.close()
        fail(f"{job['key']}: expected Letter 612x792, got {width}x{height}")

    text = normalize(page.get_text("text"))
    squeezed = compact(text)
    missing = [n for n in job["required"] if compact(n) not in squeezed]
    if missing:
        doc.close()
        fail(f"{job['key']}: missing text {missing}")
    if "\ufffd" in text:
        doc.close()
        fail(f"{job['key']}: replacement characters found")

    blocks = page.get_text("blocks")
    max_y = max(b[3] for b in blocks)
    min_x = min(b[0] for b in blocks)
    max_x = max(b[2] for b in blocks)
    if max_y > 772:
        doc.close()
        fail(f"{job['key']}: text reaches y={max_y:.1f}")
    if min_x < 30 or max_x > 584:
        doc.close()
        fail(f"{job['key']}: text outside safe area {min_x:.1f}-{max_x:.1f}")

    fonts = sorted({font[3] for font in page.get_fonts(full=True)})
    lowered = " ".join(fonts).lower()
    for family in ("georgia", "candara"):
        if family not in lowered:
            doc.close()
            fail(f"{job['key']}: {family} not embedded; fonts={fonts}")

    doc.close()
    return {
        "pages": 1,
        "page_size": [612, 792],
        "words": len(text.split()),
        "fonts": fonts,
        "max_text_y": round(max_y, 1),
        "text_bounds_x": [round(min_x, 1), round(max_x, 1)],
    }


def set_metadata(pdf_path, job):
    doc = pymupdf.open(pdf_path)
    doc.set_metadata(
        {
            "title": job["pdf_title"],
            "author": "Silenced Saints",
            "subject": job["pdf_title"],
            "keywords": KIT.get(
                "onepager_keywords",
                "Mary of Bethany, What the Room Called Waste, Preach the Forgotten, "
                "Sermon Kit 1, Silenced Saints",
            ),
            "creator": "Silenced Saints",
            "producer": "Silenced Saints build pipeline",
        }
    )
    tmp = pdf_path.with_suffix(".meta.pdf")
    doc.save(tmp, garbage=4, deflate=True)
    doc.close()
    shutil.move(tmp, pdf_path)


def render_preview(pdf_path, name):
    QA.mkdir(exist_ok=True)
    png = QA / name
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
    template = TEMPLATE.read_text(encoding="utf-8")
    DELIVERY.mkdir(exist_ok=True)
    report = {}

    for job in JOBS:
        html_path = render_html(job, css, template)
        pdf_path = ROOT / f"{job['stem']}.pdf"
        print_pdf(html_path, pdf_path)
        first = verify(job, pdf_path)
        set_metadata(pdf_path, job)
        final = verify(job, pdf_path)
        preview = render_preview(pdf_path, job["preview"])
        delivered = DELIVERY / pdf_path.name
        shutil.copy2(pdf_path, delivered)
        report[job["key"]] = {
            "file": str(pdf_path),
            "html": str(html_path),
            "delivery": str(delivered),
            **final,
            "preview": str(preview),
            "pdf_kb": round(pdf_path.stat().st_size / 1024, 1),
            "status": "passed",
        }
        print(f"{job['key']}: one page, {final['words']} words, {report[job['key']]['pdf_kb']} KB")

    out = REPORT
    out.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(f"report: {out}")


if __name__ == "__main__":
    main()
