"""PDF slide deck -> Word document conversion logic."""
import io
import re
import statistics
import unicodedata
from collections import Counter

import fitz  # PyMuPDF
from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, RGBColor

BULLET_PREFIXES = ("•", "‣", "▪", "◦", "●", "○", "-", "*", "­", "‐")
# The trailing text is optional: a numbered marker sometimes sits alone on
# its own PDF text line, with the item's actual text on the following line(s).
NUMBERED_RE = re.compile(r"^(\d+[.)]|[a-zA-Z][.)])(?:\s+(.*))?$")
CONTROL_CHAR_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")

MAX_LABEL_IMAGE_HEIGHT_PX = 200  # candidate height for a "text rendered as an image" label
MAX_LABEL_IMAGE_COLORS = 300  # real photos/diagrams have far more distinct colors than a flat text badge

BLOCKLIST_KEYWORDS = ("politecnico", "polimi", "motor system rehabilitation", "semester", "semestre")
ACADEMIC_YEAR_RE = re.compile(r"\b\d{4}\s*/\s*\d{4}\b")
DATE_RE = re.compile(r"\b\d{1,2}[./]\d{1,2}[./]\d{2,4}\b")
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
SHORT_BLOCK_CHAR_LIMIT = 40  # only auto-drop date-only blocks, not real content that mentions a date

MIN_IMAGE_DIM = 40  # px; skip tiny icons
EDGE_BAND_RATIO = 0.08  # top/bottom 8% of the slide = header/footer band
SMALL_FONT_FACTOR = 0.65  # relative to median body font size, for footer detection
H2_FONT_FACTOR = 1.35  # relative to median body font size, for sub-heading detection
H3_FONT_FACTOR = 1.15

LARGE_IMAGE_AREA_RATIO = 0.12  # image covers >=12% of the slide -> treated as a chart/graphic
LARGE_IMAGE_WIDTH_IN = 6.0
LARGE_IMAGE_MAX_HEIGHT_IN = 3.0
SMALL_IMAGE_WIDTH_IN = 2.0
SMALL_IMAGE_MAX_HEIGHT_IN = 1.5

DOCUMENT_FONT = "Aptos"
NARROW_MARGIN_IN = 0.5

# Approximation of Word's built-in "Shaded" design (Design tab > Style Set):
# solid color blocks for the top-level headings, fading to plain colored text
# for the deeper levels.
TITLE_SHADE_FILL = "1F4E79"
TITLE_FONT_COLOR = "FFFFFF"
H1_SHADE_FILL = "2E74B5"
H1_FONT_COLOR = "FFFFFF"
H2_SHADE_FILL = "DEEAF6"
H2_FONT_COLOR = "1F4E79"
H3_FONT_COLOR = "2E74B5"

STYLE_FOR_KIND = {"bullet": "List Bullet", "number": "List Number"}


RIGHT_ARROW_TEXT = "➞"
LEFT_ARROW_TEXT = "<--"


def _arrow_replacement(ch):
    """If ch is an arrow glyph (a named Unicode arrow, or an unmapped
    symbol-font Private Use Area glyph such as the Wingdings arrow
    PowerPoint uses for "leads to" bullets), return its text replacement;
    otherwise return None."""
    cp = ord(ch)
    try:
        name = unicodedata.name(ch)
    except ValueError:
        name = None
    if name and "ARROW" in name:
        is_left = "LEFT" in name and "RIGHT" not in name
        return LEFT_ARROW_TEXT if is_left else RIGHT_ARROW_TEXT
    if 0xE000 <= cp <= 0xF8FF:
        return RIGHT_ARROW_TEXT
    return None


def _convert_arrows(text):
    """Replace any Unicode arrow glyph (any direction, any arrow block) with
    an ASCII/plain-text arrow so it survives into Word text."""
    return "".join(_arrow_replacement(ch) or ch for ch in text)


def _sanitize_text(text):
    """Strip control characters that some fonts/PDF producers leave behind
    for glyphs they can't map (e.g. a missing symbol -> NUL byte); those
    would otherwise crash Word XML generation outright."""
    return CONTROL_CHAR_RE.sub("", text)


def _classify_line_marker(raw_text):
    """Look at a RAW (not yet arrow-converted) line and split off a leading
    list marker. Returns (kind, remainder) where remainder still needs
    _convert_arrows applied. A leading arrow glyph counts as a bullet too --
    some slide templates use an arrow instead of a dot as the bullet mark --
    so it's consumed as the marker instead of becoming inline "➞" text."""
    stripped = raw_text.strip()
    if not stripped:
        return None, stripped
    if stripped[0] in BULLET_PREFIXES:
        return "bullet", stripped[1:].strip()
    m = NUMBERED_RE.match(stripped)
    if m:
        return "number", (m.group(2) or "").strip()
    if _arrow_replacement(stripped[0]) is not None:
        return "bullet", stripped[1:].strip()
    return None, stripped


def _looks_like_page_number(lines):
    return len(lines) == 1 and lines[0].strip().isdigit() and len(lines[0].strip()) <= 3


def _looks_like_institutional_noise(lines):
    joined = " ".join(lines)
    lowered = joined.lower()
    if any(keyword in lowered for keyword in BLOCKLIST_KEYWORDS):
        return True
    if EMAIL_RE.search(joined):
        return True
    if len(joined) <= SHORT_BLOCK_CHAR_LIMIT and (ACADEMIC_YEAR_RE.search(joined) or DATE_RE.search(joined)):
        return True
    return False


def _normalize(text):
    return re.sub(r"\s+", " ", text.strip().lower())


def _block_key(block):
    return _normalize(" ".join(text for text, _kind in block["lines"]))


def _set_run_font(run, name=DOCUMENT_FONT):
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), name)


def _set_style_font(document, style_name, name=DOCUMENT_FONT):
    try:
        style = document.styles[style_name]
    except KeyError:
        return
    style.font.name = name
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), name)


def _apply_document_font(document):
    for style_name in ("Normal", "Title", "Heading 1", "Heading 2", "Heading 3", "List Bullet", "List Number"):
        _set_style_font(document, style_name)


def _shade_paragraph(paragraph, hex_color):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    pPr.append(shd)


def _style_heading(paragraph, fill_hex=None, font_hex=None):
    if fill_hex:
        _shade_paragraph(paragraph, fill_hex)
    if font_hex:
        for run in paragraph.runs:
            run.font.color.rgb = RGBColor.from_string(font_hex)


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
        _set_run_font(run)
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
        lines, sizes, fonts = [], [], []
        for line in b.get("lines", []):
            spans = line.get("spans", [])
            raw_text = _sanitize_text("".join(s.get("text", "") for s in spans).strip())
            if not raw_text:
                continue
            sizes.extend(s.get("size", 0) for s in spans)
            fonts.extend(s.get("font", "") for s in spans)
            kind, remainder = _classify_line_marker(raw_text)
            lines.append((_convert_arrows(remainder), kind))
        plain_lines = [text for text, _kind in lines]
        if not lines or _looks_like_page_number(plain_lines) or _looks_like_institutional_noise(plain_lines):
            continue
        blocks.append({
            "bbox": b["bbox"],
            "lines": lines,
            "avg_size": sum(sizes) / len(sizes) if sizes else 0,
            "all_bold": bool(fonts) and all("bold" in f.lower() for f in fonts),
        })
    blocks.sort(key=lambda blk: (round(blk["bbox"][1], 0), blk["bbox"][0]))
    return blocks


def _merge_lines_into_paragraphs(lines):
    """Merge a flat stream of (text, kind) lines into paragraphs, splitting on
    new bullet/number markers. Takes a plain line list rather than a single
    block's lines because PDF exporters sometimes split one bulleted
    sentence's wrapped continuation into a separate text block -- feeding in
    lines flattened across those block boundaries keeps such sentences whole."""
    paragraphs = []
    current_text, current_kind = None, None
    for text, kind in lines:
        if kind:
            if current_text:
                paragraphs.append((current_text, current_kind))
            current_text, current_kind = text, kind
        elif not current_text:
            # Nothing accumulated yet (or just an empty marker, e.g. a bullet
            # glyph on its own line) -- adopt this text but keep any kind
            # already established by that marker.
            current_text = text
        else:
            current_text += " " + text
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


def _looks_like_text_label_image(png_bytes):
    """Detect a short, near-flat-color image that's really a stylized text
    caption exported as a picture (e.g. WordArt-style headings/percentages
    rendered as a bitmap) rather than a real photo/diagram. Real diagrams
    have far more distinct colors even when small."""
    try:
        with Image.open(io.BytesIO(png_bytes)) as im:
            if im.height > MAX_LABEL_IMAGE_HEIGHT_PX:
                return False
            colors = im.convert("RGB").getcolors(maxcolors=MAX_LABEL_IMAGE_COLORS)
        return colors is not None
    except Exception:
        return False


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
    return pages_blocks, pages_images, text_freq, xref_freq, digest_freq, median_size


def _filter_text_blocks(blocks, page_height, small_font_threshold, text_freq, page_max_font):
    filtered = []
    for b in blocks:
        in_edge_band = (
            b["bbox"][1] < page_height * EDGE_BAND_RATIO
            or b["bbox"][3] > page_height * (1 - EDGE_BAND_RATIO)
        )
        is_small = b["avg_size"] <= small_font_threshold
        if in_edge_band and is_small:
            continue
        # Anything that repeats verbatim across pages and isn't this page's
        # own (largest-font) title is boilerplate: course name, professor,
        # lesson/date labels, etc. Repetition is checked regardless of font
        # size so it also catches boilerplate set in a normal body-text size.
        is_title_candidate = b["avg_size"] >= page_max_font
        if not is_title_candidate and text_freq[_block_key(b)] >= 2:
            continue
        filtered.append(b)
    return filtered


def _filter_images(images, page_height, page_area, xref_freq, digest_freq):
    filtered = []
    for info in images:
        if info.get("width", 0) < MIN_IMAGE_DIM or info.get("height", 0) < MIN_IMAGE_DIM:
            continue
        bbox = info["bbox"]
        # Only treat as a corner logo/watermark if it sits ENTIRELY within the
        # top/bottom margin band; a large chart merely extending past that
        # line (but mostly in the body) must survive.
        fully_in_top_band = bbox[3] <= page_height * EDGE_BAND_RATIO
        fully_in_bottom_band = bbox[1] >= page_height * (1 - EDGE_BAND_RATIO)
        if fully_in_top_band or fully_in_bottom_band:
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
            title = " ".join(text for text, _kind in b["lines"])
            body_blocks = [x for x in blocks if x is not b]
            return title, body_blocks
    return None, blocks


def _looks_like_cover_slide_label(block):
    """Short, non-list fragments like 'Lesson', 'of', a lecture topic, or a
    professor's name, typically laid out as separate boxes on a title
    slide rather than real sentences."""
    if block["lines"][0][1]:  # starts with a bullet/number marker
        return False
    word_count = sum(len(text.split()) for text, _kind in block["lines"])
    return len(block["lines"]) <= 2 and word_count <= 4


def _classify_subheading(block, median_size):
    """A short, single-line block that is noticeably larger than normal body
    text (but isn't the slide title) is treated as a Heading 2/3. A short,
    entirely bold line at ordinary body size (e.g. "Pros"/"Cons",
    "Advantages") is also a sub-header even without a size difference."""
    if len(block["lines"]) != 1 or block["lines"][0][1]:
        return None
    text = block["lines"][0][0]
    if block["avg_size"] >= median_size * H2_FONT_FACTOR:
        return "heading2"
    if block["avg_size"] >= median_size * H3_FONT_FACTOR:
        return "heading3"
    if block.get("all_bold") and len(text.split()) <= 4:
        return "heading3"
    return None


def _add_sized_picture(document, png_bytes, target_width_in, max_height_in):
    width_kwargs = {"width": Inches(target_width_in)}
    try:
        with Image.open(io.BytesIO(png_bytes)) as im:
            px_w, px_h = im.size
        if px_w and px_h:
            projected_height_in = target_width_in * (px_h / px_w)
            if projected_height_in > max_height_in:
                width_kwargs = {"height": Inches(max_height_in)}
    except Exception:
        pass
    document.add_picture(io.BytesIO(png_bytes), **width_kwargs)
    document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER


def extract_content(pdf_bytes, title="Presentazione"):
    """Parse the PDF into an ordered list of content items (title, headings,
    paragraphs, images) independent of how the Word document gets built, so
    the same extraction can feed both the original and a translated build."""
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_blocks, pages_images, text_freq, xref_freq, digest_freq, median_size = (
        _collect_document_data(src)
    )
    small_font_threshold = median_size * SMALL_FONT_FACTOR

    items = [{"kind": "title", "text": title}]
    prev_title_norm = None
    seen_in_section = set()

    for page_index, page in enumerate(src):
        page_height = page.rect.height
        page_area = page.rect.width * page_height
        raw_blocks = pages_blocks[page_index]
        page_max_font = max((b["avg_size"] for b in raw_blocks), default=0)

        blocks = _filter_text_blocks(raw_blocks, page_height, small_font_threshold, text_freq, page_max_font)
        images = _filter_images(pages_images[page_index], page_height, page_area, xref_freq, digest_freq)

        slide_title, body_blocks = _detect_title(blocks, page_height)

        is_first_page = page_index == 0
        if is_first_page:
            # A cover slide's stray labels (lesson number, topic, professor
            # name) aren't real content. Only drop the page's own "title"
            # too when NOTHING real is left once those are removed -- a
            # short but genuine slide title (e.g. "Analisi della stagione")
            # must survive when it's followed by real body content.
            filtered_body = [b for b in body_blocks if not _looks_like_cover_slide_label(b)]
            if not filtered_body and slide_title and _looks_like_cover_slide_label({"lines": [(slide_title, None)]}):
                slide_title = None
            body_blocks = filtered_body

        title_norm = _normalize(slide_title) if slide_title else None

        if title_norm and title_norm != prev_title_norm:
            items.append({"kind": "heading1", "text": slide_title})
            prev_title_norm = title_norm
            seen_in_section = set()

        pending_text, pending_kind = None, None

        def _flush_pending():
            nonlocal pending_text, pending_kind
            if pending_text:
                norm = _normalize(pending_text)
                if norm not in seen_in_section:
                    seen_in_section.add(norm)
                    items.append({"kind": "paragraph", "text": pending_text, "list_style": pending_kind})
            pending_text, pending_kind = None, None

        for b in body_blocks:
            subheading_kind = _classify_subheading(b, median_size)
            if subheading_kind:
                _flush_pending()
                subheading_text = b["lines"][0][0]
                text_norm = _normalize(subheading_text)
                if text_norm not in seen_in_section:
                    seen_in_section.add(text_norm)
                    items.append({"kind": subheading_kind, "text": subheading_text})
                continue

            block_paragraphs = _merge_lines_into_paragraphs(b["lines"])
            if not block_paragraphs:
                continue

            first_text, first_kind = block_paragraphs[0]
            rest = block_paragraphs[1:]
            # A bulleted/numbered sentence's wrapped continuation sometimes
            # lands in a separate PDF text block from its own marker (no
            # marker of its own). Only bridge that gap when the still-open
            # item actually had a marker -- two adjacent PLAIN blocks (e.g.
            # independent diagram captions/labels) must stay separate.
            if pending_text is not None and pending_kind and not first_kind:
                pending_text += " " + first_text
            else:
                _flush_pending()
                pending_text, pending_kind = first_text, first_kind

            for text, kind in rest:
                _flush_pending()
                pending_text, pending_kind = text, kind

        _flush_pending()

        for info in images:
            try:
                base = src.extract_image(info["xref"])
            except Exception:
                continue
            png_bytes = _to_png_bytes(base["image"])
            if not png_bytes:
                continue
            if _looks_like_text_label_image(png_bytes):
                continue
            bbox = info["bbox"]
            area_ratio = ((bbox[2] - bbox[0]) * (bbox[3] - bbox[1])) / page_area
            is_large = area_ratio >= LARGE_IMAGE_AREA_RATIO
            if is_first_page and not is_large:
                # Cover slides carry logos/decorative badges, not content;
                # only a genuinely large graphic there is worth keeping.
                continue
            if is_large:
                target_width, max_height = LARGE_IMAGE_WIDTH_IN, LARGE_IMAGE_MAX_HEIGHT_IN
            else:
                target_width, max_height = SMALL_IMAGE_WIDTH_IN, SMALL_IMAGE_MAX_HEIGHT_IN
            items.append({
                "kind": "image",
                "png_bytes": png_bytes,
                "target_width": target_width,
                "max_height": max_height,
            })

    return items


_HEADING_STYLE = {
    "title": (0, TITLE_SHADE_FILL, TITLE_FONT_COLOR),
    "heading1": (1, H1_SHADE_FILL, H1_FONT_COLOR),
    "heading2": (2, H2_SHADE_FILL, H2_FONT_COLOR),
    "heading3": (3, None, H3_FONT_COLOR),
}


def build_document(items):
    """Render a list of content items (see extract_content) into a styled
    Word document and return its bytes."""
    document = Document()
    _apply_document_font(document)
    _set_narrow_margins(document)
    _add_centered_page_number_footer(document)

    for item in items:
        kind = item["kind"]
        if kind in _HEADING_STYLE:
            level, fill_hex, font_hex = _HEADING_STYLE[kind]
            paragraph = document.add_heading(item["text"], level=level)
            _style_heading(paragraph, fill_hex, font_hex)
        elif kind == "paragraph":
            style = STYLE_FOR_KIND.get(item.get("list_style"))
            paragraph = document.add_paragraph(item["text"], style=style)
            for run in paragraph.runs:
                _set_run_font(run)
        elif kind == "image":
            try:
                _add_sized_picture(document, item["png_bytes"], item["target_width"], item["max_height"])
            except Exception:
                continue

    out = io.BytesIO()
    document.save(out)
    return out.getvalue()


def translate_items(items, target_lang="it"):
    """Return a copy of items with every text field translated; falls back
    to the original text for anything that can't be translated (e.g. no
    network access), so a translation hiccup never breaks the conversion."""
    from deep_translator import GoogleTranslator

    translator = GoogleTranslator(source="auto", target=target_lang)
    text_indices = [i for i, item in enumerate(items) if item["kind"] != "image" and item["text"].strip()]
    texts = [items[i]["text"] for i in text_indices]

    new_items = [dict(item) for item in items]
    if not texts:
        return new_items

    translated = None
    try:
        translated = translator.translate_batch(texts)
    except Exception:
        translated = None

    if translated and len(translated) == len(texts):
        for idx, translated_text in zip(text_indices, translated):
            if translated_text:
                new_items[idx]["text"] = translated_text
    else:
        for idx in text_indices:
            try:
                result = translator.translate(new_items[idx]["text"])
                if result:
                    new_items[idx]["text"] = result
            except Exception:
                continue

    return new_items


def pdf_to_docx(pdf_bytes, title="Presentazione"):
    return build_document(extract_content(pdf_bytes, title))


def pdf_to_docx_italian(pdf_bytes, title="Presentazione"):
    items = extract_content(pdf_bytes, title)
    return build_document(translate_items(items, target_lang="it"))
