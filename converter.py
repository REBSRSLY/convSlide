"""PDF slide deck -> Word document conversion logic."""
import io
import re

import fitz  # PyMuPDF
from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches

BULLET_PREFIXES = ("•", "‣", "▪", "◦", "●", "○", "-", "*")
NUMBERED_RE = re.compile(r"^(\d+[.)]|[a-zA-Z][.)])\s+(.*)")
MIN_IMAGE_DIM = 40  # px; skip tiny icons/logos
MAX_IMAGE_WIDTH_IN = 5.5


def _is_bullet_line(text):
    stripped = text.strip()
    if not stripped:
        return False
    if stripped[0] in BULLET_PREFIXES:
        return True
    return bool(NUMBERED_RE.match(stripped))


def _strip_bullet(text):
    stripped = text.strip()
    if stripped and stripped[0] in BULLET_PREFIXES:
        return stripped[1:].strip()
    m = NUMBERED_RE.match(stripped)
    if m:
        return m.group(2)
    return stripped


def _looks_like_page_number(lines):
    return len(lines) == 1 and lines[0].strip().isdigit() and len(lines[0].strip()) <= 3


def _extract_page_blocks(page):
    raw = page.get_text("dict")
    blocks = []
    for b in raw.get("blocks", []):
        if b.get("type") != 0:  # 0 = text block
            continue
        lines, sizes = [], []
        for line in b.get("lines", []):
            spans = line.get("spans", [])
            line_text = "".join(s.get("text", "") for s in spans).strip()
            if not line_text:
                continue
            sizes.extend(s.get("size", 0) for s in spans)
            lines.append(line_text)
        if not lines or _looks_like_page_number(lines):
            continue
        blocks.append({
            "bbox": b["bbox"],
            "lines": lines,
            "avg_size": sum(sizes) / len(sizes) if sizes else 0,
        })
    blocks.sort(key=lambda blk: (round(blk["bbox"][1], 0), blk["bbox"][0]))
    return blocks


def _paragraphs_from_block(block):
    """Merge wrapped lines into paragraphs, splitting on new bullet/number markers."""
    paragraphs = []
    current_text, current_is_bullet = None, False
    for line in block["lines"]:
        if _is_bullet_line(line):
            if current_text:
                paragraphs.append((current_text, current_is_bullet))
            current_text, current_is_bullet = _strip_bullet(line), True
        elif current_text is None:
            current_text, current_is_bullet = line, False
        else:
            current_text += " " + line
    if current_text:
        paragraphs.append((current_text, current_is_bullet))
    return paragraphs


def _to_png_bytes(image_bytes):
    try:
        im = Image.open(io.BytesIO(image_bytes))
        if im.mode in ("CMYK", "P", "LA"):
            im = im.convert("RGB")
        out = io.BytesIO()
        im.save(out, format="PNG")
        return out.getvalue()
    except Exception:
        return None


def _extract_page_images(doc, page):
    images = []
    for img in page.get_images(full=True):
        xref = img[0]
        try:
            base = doc.extract_image(xref)
        except Exception:
            continue
        if base.get("width", 0) < MIN_IMAGE_DIM or base.get("height", 0) < MIN_IMAGE_DIM:
            continue
        png_bytes = _to_png_bytes(base["image"])
        if png_bytes:
            images.append(png_bytes)
    return images


def pdf_to_docx(pdf_bytes, title="Presentazione"):
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    document = Document()
    document.add_heading(title, level=0)

    for page_index in range(len(src)):
        page = src[page_index]
        blocks = _extract_page_blocks(page)
        images = _extract_page_images(src, page)

        if not blocks and not images:
            continue

        slide_title, body_blocks = None, blocks
        if blocks:
            max_size = max(b["avg_size"] for b in blocks)
            page_height = page.rect.height
            for b in blocks:
                if b["avg_size"] == max_size and b["bbox"][1] < page_height * 0.35:
                    slide_title = " ".join(b["lines"])
                    body_blocks = [x for x in blocks if x is not b]
                    break

        document.add_heading(slide_title or f"Slide {page_index + 1}", level=1)

        for b in body_blocks:
            for text, is_bullet in _paragraphs_from_block(b):
                if is_bullet:
                    document.add_paragraph(text, style="List Bullet")
                else:
                    document.add_paragraph(text)

        for img_bytes in images:
            try:
                document.add_picture(io.BytesIO(img_bytes), width=Inches(MAX_IMAGE_WIDTH_IN))
                document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception:
                continue

        if page_index < len(src) - 1:
            document.add_page_break()

    out = io.BytesIO()
    document.save(out)
    return out.getvalue()
