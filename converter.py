"""PDF slide deck -> Word document conversion logic."""
import io
import re
import statistics
from collections import Counter

import fitz  # PyMuPDF
from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches

BULLET_PREFIXES = ("•", "‣", "▪", "◦", "●", "○", "-", "*")
NUMBERED_RE = re.compile(r"^(\d+[.)]|[a-zA-Z][.)])\s+(.*)")

MIN_IMAGE_DIM = 40  # px; skip tiny icons
EDGE_BAND_RATIO = 0.08  # top/bottom 8% of the slide = header/footer band
SMALL_FONT_FACTOR = 0.65  # relative to median body font size
LARGE_IMAGE_AREA_RATIO = 0.12  # image covers >=12% of the slide -> treated as a chart/graphic
LARGE_IMAGE_WIDTH_IN = 6.0
SMALL_IMAGE_WIDTH_IN = 2.0
TITLE_SHADE_COLOR = "B8CCE4"
HEADING_SHADE_COLOR = "D9E2F3"
NARROW_MARGIN_IN = 0.5

STYLE_FOR_KIND = {"bullet": "List Bullet", "number": "List Number"}


def _line_kind(text):
    stripped = text.strip()
    if not stripped:
        return None
    if stripped[0] in BULLET_PREFIXES:
        return "bullet"
    if NUMBERED_RE.match(stripped):
        return "number"
    return None


def _strip_marker(text):
    stripped = text.strip()
    if stripped and stripped[0] in BULLET_PREFIXES:
        return stripped[1:].strip()
    m = NUMBERED_RE.match(stripped)
    if m:
        return m.group(2)
    return stripped


def _looks_like_page_number(lines):
    return len(lines) == 1 and lines[0].strip().isdigit() and len(lines[0].strip()) <= 3


def _normalize(text):
    return re.sub(r"\s+", " ", text.strip().lower())


def _block_key(block):
    return _normalize(" ".join(block["lines"]))


def _shade_paragraph(paragraph, hex_color):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    pPr.append(shd)


def _set_narrow_margins(document):
    for section in document.sections:
        section.top_margin = Inches(NARROW_MARGIN_IN)
        section.bottom_margin = Inches(NARROW_MARGIN_IN)
        section.left_margin = Inches(NARROW_MARGIN_IN)
        section.right_margin = Inches(NARROW_MARGIN_IN)


def _add_centered_page_number_footer(document):
    for section in document.sections:
        footer = section.footer
        footer.is_linked_to_previous = False
        paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        paragraph.text = ""
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = paragraph.add_run()
        fld_begin = OxmlElement("w:fldChar")
        fld_begin.set(qn("w:fldCharType"), "begin")
        instr_text = OxmlElement("w:instrText")
        instr_text.set(qn("xml:space"), "preserve")
        instr_text.text = "PAGE"
        fld_end = OxmlElement("w:fldChar")
        fld_end.set(qn("w:fldCharType"), "end")
        run._r.append(fld_begin)
        run._r.append(instr_text)
        run._r.append(fld_end)


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
    current_text, current_kind = None, None
    for line in block["lines"]:
        kind = _line_kind(line)
        if kind:
            if current_text:
                paragraphs.append((current_text, current_kind))
            current_text, current_kind = _strip_marker(line), kind
        elif current_text is None:
            current_text, current_kind = line, None
        else:
            current_text += " " + line
    if current_text:
        paragraphs.append((current_text, current_kind))
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


def _collect_document_data(src):
    """First pass: gather every page's text blocks and image placements so
    recurring headers/footers/logos can be recognized across the whole deck."""
    pages_blocks, pages_images = [], []
    text_freq = Counter()
    xref_freq = Counter()
    digest_freq = Counter()
    all_sizes = []

    for page in src:
        blocks = _extract_page_blocks(page)
        pages_blocks.append(blocks)
        for b in blocks:
            text_freq[_block_key(b)] += 1
            all_sizes.append(b["avg_size"])

        images = page.get_image_info(hashes=True, xrefs=True)
        pages_images.append(images)
        for info in images:
            if info.get("xref"):
                xref_freq[info["xref"]] += 1
            if info.get("digest"):
                digest_freq[info["digest"]] += 1

    median_size = statistics.median(all_sizes) if all_sizes else 12
    return pages_blocks, pages_images, text_freq, xref_freq, digest_freq, median_size * SMALL_FONT_FACTOR


def _filter_text_blocks(blocks, page_height, small_font_threshold, text_freq):
    filtered = []
    for b in blocks:
        in_edge_band = (
            b["bbox"][1] < page_height * EDGE_BAND_RATIO
            or b["bbox"][3] > page_height * (1 - EDGE_BAND_RATIO)
        )
        is_small = b["avg_size"] <= small_font_threshold
        if in_edge_band and is_small:
            continue
        if is_small and text_freq[_block_key(b)] >= 2:
            continue
        filtered.append(b)
    return filtered


def _filter_images(images, page_height, page_area, xref_freq, digest_freq):
    filtered = []
    for info in images:
        if info.get("width", 0) < MIN_IMAGE_DIM or info.get("height", 0) < MIN_IMAGE_DIM:
            continue
        bbox = info["bbox"]
        in_edge_band = (
            bbox[1] < page_height * EDGE_BAND_RATIO
            or bbox[3] > page_height * (1 - EDGE_BAND_RATIO)
        )
        if in_edge_band:
            continue
        if info.get("xref") and xref_freq[info["xref"]] >= 2:
            continue
        if info.get("digest") and digest_freq[info["digest"]] >= 2:
            continue
        filtered.append(info)
    return filtered


def _detect_title(blocks, page_height):
    if not blocks:
        return None, blocks
    max_size = max(b["avg_size"] for b in blocks)
    for b in blocks:
        if b["avg_size"] == max_size and b["bbox"][1] < page_height * 0.35:
            title = " ".join(b["lines"])
            body_blocks = [x for x in blocks if x is not b]
            return title, body_blocks
    return None, blocks


def pdf_to_docx(pdf_bytes, title="Presentazione"):
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_blocks, pages_images, text_freq, xref_freq, digest_freq, small_font_threshold = (
        _collect_document_data(src)
    )

    document = Document()
    _set_narrow_margins(document)
    _add_centered_page_number_footer(document)

    title_paragraph = document.add_heading(title, level=0)
    _shade_paragraph(title_paragraph, TITLE_SHADE_COLOR)

    prev_title_norm = None
    seen_in_section = set()

    for page_index, page in enumerate(src):
        page_height = page.rect.height
        page_area = page.rect.width * page_height

        blocks = _filter_text_blocks(pages_blocks[page_index], page_height, small_font_threshold, text_freq)
        images = _filter_images(pages_images[page_index], page_height, page_area, xref_freq, digest_freq)

        slide_title, body_blocks = _detect_title(blocks, page_height)
        title_norm = _normalize(slide_title) if slide_title else None

        if title_norm and title_norm != prev_title_norm:
            heading_paragraph = document.add_heading(slide_title, level=1)
            _shade_paragraph(heading_paragraph, HEADING_SHADE_COLOR)
            prev_title_norm = title_norm
            seen_in_section = set()

        for b in body_blocks:
            for text, kind in _paragraphs_from_block(b):
                norm = _normalize(text)
                if norm in seen_in_section:
                    continue
                seen_in_section.add(norm)
                style = STYLE_FOR_KIND.get(kind)
                document.add_paragraph(text, style=style)

        for info in images:
            try:
                base = src.extract_image(info["xref"])
            except Exception:
                continue
            png_bytes = _to_png_bytes(base["image"])
            if not png_bytes:
                continue
            bbox = info["bbox"]
            area_ratio = ((bbox[2] - bbox[0]) * (bbox[3] - bbox[1])) / page_area
            target_width = LARGE_IMAGE_WIDTH_IN if area_ratio >= LARGE_IMAGE_AREA_RATIO else SMALL_IMAGE_WIDTH_IN
            try:
                document.add_picture(io.BytesIO(png_bytes), width=Inches(target_width))
                document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception:
                continue

    out = io.BytesIO()
    document.save(out)
    return out.getvalue()
