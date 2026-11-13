import os
import re
import tempfile
from pathlib import Path
from xml.sax.saxutils import escape

from docx import Document
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, PageBreak, Image as ReportImage
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def _as_float(value, default=None):
    try:
        if value in (None, "", "Not specified", "Not detected"):
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def _set_run_font(run, font_name, font_size):
    run.font.name = font_name
    run.font.size = Pt(font_size)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)
    rfonts.set(qn("w:eastAsia"), font_name)


def _iter_document_blocks(document):
    """Yield top-level paragraphs and tables in their original document order."""
    body = document._element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield DocxParagraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield DocxTable(child, document)



def _remove_previous_manual_page_break(blocks, block_index):
    """Avoid a blank page when a manual break already precedes a heading."""
    if block_index <= 0:
        return
    previous = blocks[block_index - 1]
    if not isinstance(previous, DocxParagraph) or previous.text.strip():
        return
    for page_break in previous._p.xpath('.//w:br[@w:type="page"]'):
        page_break.getparent().remove(page_break)
    previous.paragraph_format.page_break_before = None


def _format_paragraph(paragraph, font_name, font_size, line_spacing):
    paragraph.paragraph_format.line_spacing = line_spacing
    paragraph.paragraph_format.first_line_indent = Pt(6)
    for run in paragraph.runs:
        _set_run_font(run, font_name, font_size)


def format_docx_report(input_path, output_path, guidelines):
    """Apply supported college rules to an editable DOCX report."""
    document = Document(input_path)

    font_name = guidelines.get("font_family")
    font_size = _as_float(guidelines.get("font_size"))
    line_spacing = _as_float(guidelines.get("line_spacing"))
    left_margin = _as_float(guidelines.get("left_margin"))
    other_margin = _as_float(guidelines.get("other_margins"))

    if not font_name or font_name == "Not specified":
        font_name = "Times New Roman"
    if font_size is None:
        font_size = 12.0
    if line_spacing is None:
        line_spacing = 1.5
    if left_margin is None:
        left_margin = 1.5
    if other_margin is None:
        other_margin = 1.0

    for section in document.sections:
        section.left_margin = Inches(left_margin)
        section.right_margin = Inches(other_margin)
        section.top_margin = Inches(other_margin)
        section.bottom_margin = Inches(other_margin)

    normal_style = document.styles["Normal"]
    normal_style.font.name = font_name
    normal_style.font.size = Pt(font_size)
    normal_style.paragraph_format.line_spacing = line_spacing

    blocks = list(_iter_document_blocks(document))
    formatting_started = False
    in_table_of_contents = False
    in_references = False
    current_main_heading = ""

    from format_analyzer import detect_heading

    for block_index, block in enumerate(blocks):
        if isinstance(block, DocxParagraph):
            paragraph = block
            text = paragraph.text.strip()
            if not text:
                continue

            normalized = re.sub(
                r"^\s*\d+(?:\.\d+)*[.)]?\s*", "", text
            ).strip().lower().rstrip(":")
            style_name = paragraph.style.name if paragraph.style else ""
            is_word_heading = bool(re.match(r"Heading\s*[1-9]", style_name, re.I))
            is_uppercase = bool(re.search(r"[A-Za-z]", text)) and text == text.upper()

            if normalized == "abstract":
                formatting_started = True

            # Preserve the supplied title and certificate pages as they are.
            if not formatting_started:
                continue

            _format_paragraph(paragraph, font_name, font_size, line_spacing)

            # The contents heading is formatted, but its listed chapter names
            # are entries, not chapter headings. Keep them together on the TOC page.
            if normalized in {"table of contents", "contents"}:
                in_table_of_contents = True
                paragraph.text = text.upper()
                _remove_previous_manual_page_break(blocks, block_index)
                paragraph.paragraph_format.page_break_before = True
                paragraph.paragraph_format.keep_with_next = True
                for run in paragraph.runs:
                    run.bold = True
                    _set_run_font(run, font_name, font_size)
                current_main_heading = normalized
                continue

            # Leave TOC mode only when a real styled or uppercase heading appears.
            if in_table_of_contents:
                if is_word_heading or is_uppercase:
                    in_table_of_contents = False
                else:
                    # TOC entries keep normal weight and do not trigger page breaks.
                    for run in paragraph.runs:
                        run.bold = False
                    continue

            # References are ordinary bibliography entries, even when numbered.
            if normalized in {"references", "bibliography"}:
                in_references = True
                paragraph.text = text.upper()
                _remove_previous_manual_page_break(blocks, block_index)
                paragraph.paragraph_format.page_break_before = True
                paragraph.paragraph_format.keep_with_next = True
                for run in paragraph.runs:
                    run.bold = True
                    _set_run_font(run, font_name, font_size)
                current_main_heading = normalized
                continue

            if in_references:
                if not (is_word_heading or is_uppercase):
                    for run in paragraph.runs:
                        run.bold = False
                    continue
                in_references = False

            # Numbered list entries should not be promoted to headings. This
            # prevents advantages and bibliography items from becoming bold.
            has_word_numbering = paragraph._p.pPr is not None and paragraph._p.pPr.numPr is not None
            numbered_text = bool(re.match(r"^\s*\d+(?:\.\d+)*[.)]?\s+", text))
            numbered_list_entry = numbered_text and not is_uppercase and not is_word_heading

            detected = None
            if not numbered_list_entry:
                detected = detect_heading(
                    text,
                    bold=any(run.bold for run in paragraph.runs),
                    style_name=style_name
                )
            if has_word_numbering and not is_word_heading:
                detected = None

            if detected:
                level = detected.get("level", 1)
                if level == 1:
                    _remove_previous_manual_page_break(blocks, block_index)
                    paragraph.paragraph_format.page_break_before = True
                    paragraph.text = text.upper()
                    current_main_heading = normalized
                else:
                    paragraph.text = text.title()

                paragraph.paragraph_format.keep_with_next = True
                for run in paragraph.runs:
                    run.bold = True
                    _set_run_font(run, font_name, font_size)
            else:
                # Body paragraphs and subpoints must remain normal weight.
                if current_main_heading == "advantages":
                    for run in paragraph.runs:
                        run.bold = False

        elif isinstance(block, DocxTable) and formatting_started:
            # Format report tables but do not change their structure or borders.
            for row in block.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if paragraph.text.strip():
                            _format_paragraph(
                                paragraph, font_name, font_size, line_spacing
                            )

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    document.save(output_path)
    return output_path


def _register_pdf_fonts(requested_font="Times New Roman"):
    """Use an installed matching font where possible, with standard PDF fallbacks."""
    requested = (requested_font or "Times New Roman").lower()

    candidates = {
        "times": (
            [
                r"C:\Windows\Fonts\times.ttf",
                "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman.ttf",
                "/usr/share/fonts/truetype/liberation2/LiberationSerif-Regular.ttf"
            ],
            [
                r"C:\Windows\Fonts\timesbd.ttf",
                "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman_Bold.ttf",
                "/usr/share/fonts/truetype/liberation2/LiberationSerif-Bold.ttf"
            ],
            "Times-Roman", "Times-Bold"
        ),
        "arial": (
            [
                r"C:\Windows\Fonts\arial.ttf",
                "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf",
                "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"
            ],
            [
                r"C:\Windows\Fonts\arialbd.ttf",
                "/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf",
                "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf"
            ],
            "Helvetica", "Helvetica-Bold"
        ),
        "calibri": (
            [
                r"C:\Windows\Fonts\calibri.ttf",
                "/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf"
            ],
            [
                r"C:\Windows\Fonts\calibrib.ttf",
                "/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf"
            ],
            "Helvetica", "Helvetica-Bold"
        )
    }

    if "arial" in requested:
        regular_paths, bold_paths, fallback_regular, fallback_bold = candidates["arial"]
    elif "calibri" in requested:
        regular_paths, bold_paths, fallback_regular, fallback_bold = candidates["calibri"]
    else:
        regular_paths, bold_paths, fallback_regular, fallback_bold = candidates["times"]

    regular_name, bold_name = fallback_regular, fallback_bold
    for candidate in regular_paths:
        if os.path.exists(candidate):
            try:
                pdfmetrics.registerFont(TTFont("UniDraftRegular", candidate))
                regular_name = "UniDraftRegular"
                break
            except Exception:
                pass
    for candidate in bold_paths:
        if os.path.exists(candidate):
            try:
                pdfmetrics.registerFont(TTFont("UniDraftBold", candidate))
                bold_name = "UniDraftBold"
                break
            except Exception:
                pass
    return regular_name, bold_name


def _extract_cover_images(input_path, temp_dir, max_images=2):
    """Extract likely cover-page images so the rebuilt PDF can retain a logo."""
    extension = Path(input_path).suffix.lower()
    saved_paths = []

    if extension == ".docx":
        source_doc = Document(input_path)
        seen_relationships = set()
        # Cover-page logos are commonly placed in the first few paragraphs.
        # Also inspect the first table, where some college cover templates put logos.
        candidate_paragraphs = list(source_doc.paragraphs[:12])
        for table in source_doc.tables[:1]:
            for row in table.rows:
                for cell in row.cells:
                    candidate_paragraphs.extend(cell.paragraphs)
        for paragraph in candidate_paragraphs:
            for blip in paragraph._p.xpath('.//a:blip'):
                relationship_id = blip.get(qn('r:embed'))
                if not relationship_id or relationship_id in seen_relationships:
                    continue
                seen_relationships.add(relationship_id)
                image_part = source_doc.part.related_parts.get(relationship_id)
                if image_part is None:
                    continue
                suffix = Path(str(image_part.partname)).suffix or ".png"
                image_path = os.path.join(temp_dir, f"cover_image_{len(saved_paths) + 1}{suffix}")
                with open(image_path, "wb") as image_file:
                    image_file.write(image_part.blob)
                saved_paths.append(image_path)
                if len(saved_paths) >= max_images:
                    return saved_paths

    elif extension == ".pdf":
        import fitz
        with fitz.open(input_path) as source_pdf:
            if len(source_pdf):
                seen_xrefs = set()
                for image_info in source_pdf[0].get_images(full=True):
                    xref = image_info[0]
                    if xref in seen_xrefs:
                        continue
                    seen_xrefs.add(xref)
                    extracted = source_pdf.extract_image(xref)
                    image_bytes = extracted.get("image")
                    if not image_bytes:
                        continue
                    suffix = "." + extracted.get("ext", "png")
                    image_path = os.path.join(temp_dir, f"cover_image_{len(saved_paths) + 1}{suffix}")
                    with open(image_path, "wb") as image_file:
                        image_file.write(image_bytes)
                    saved_paths.append(image_path)
                    if len(saved_paths) >= max_images:
                        break

    return saved_paths


def create_formatted_pdf(input_path, output_path, guidelines):
    """
    Rebuild a clean, consistently formatted PDF from extracted text.
    This reflows text; it cannot preserve the exact original PDF layout/images.
    """
    from content_extractor import extract_document_structure

    structure = extract_document_structure(input_path)
    title = (structure.get("title") or "").strip()
    sections = structure.get("sections") or []
    tables = structure.get("tables") or []
    # PDF text extraction already includes table cell text in page order. Adding
    # the extracted grid again would duplicate that content, so keep PDF-source
    # tables as extracted text. DOCX-source tables are rendered as real tables.
    if Path(input_path).suffix.lower() == ".pdf":
        tables = []

    if not title and not sections:
        raise ValueError(
            "No selectable text was found in this PDF. Scanned PDFs require OCR before formatting."
        )

    font_name = guidelines.get("font_family")
    font_size = _as_float(guidelines.get("font_size"), 12.0) or 12.0
    line_spacing = _as_float(guidelines.get("line_spacing"), 1.5) or 1.5
    left_margin = _as_float(guidelines.get("left_margin"), 1.5)
    other_margin = _as_float(guidelines.get("other_margins"), 1.0)
    left_margin = left_margin if left_margin is not None else 1.5
    other_margin = other_margin if other_margin is not None else 1.0

    regular_font, bold_font = _register_pdf_fonts(font_name)
    leading = font_size * line_spacing

    styles = getSampleStyleSheet()
    body_style = ParagraphStyle(
        "UniDraftBody",
        parent=styles["BodyText"],
        fontName=regular_font,
        fontSize=font_size,
        leading=leading,
        alignment=TA_LEFT,
        spaceAfter=6,
        allowWidows=0,
        allowOrphans=0
    )
    heading_style = ParagraphStyle(
        "UniDraftHeading",
        parent=body_style,
        fontName=bold_font,
        fontSize=font_size,
        leading=leading,
        spaceBefore=12,
        spaceAfter=8,
        keepWithNext=True
    )
    title_style = ParagraphStyle(
        "UniDraftTitle",
        parent=heading_style,
        fontSize=font_size,
        leading=font_size * line_spacing,
        alignment=TA_CENTER,
        spaceAfter=24
    )

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=left_margin * inch,
        rightMargin=other_margin * inch,
        topMargin=other_margin * inch,
        bottomMargin=other_margin * inch,
        title=title or "Formatted Report",
        author="UniDraft AI"
    )

    story = []
    # Keep likely college-logo images from the cover page in the generated PDF.
    # The editable DOCX is formatted in place and retains all of its original images.
    image_temp_dir = tempfile.TemporaryDirectory(prefix="unidraft_images_")
    try:
        cover_images = _extract_cover_images(input_path, image_temp_dir.name)
        for image_path in cover_images:
            image_flowable = ReportImage(image_path, width=1.8 * inch, height=1.35 * inch, kind="proportional")
            image_flowable.hAlign = "CENTER"
            story.append(image_flowable)
            story.append(Spacer(1, 6))

        if title:
            story.append(Paragraph(escape(title), title_style))
            story.append(Spacer(1, 8))

        for section in sections:
            heading = (section.get("heading") or "").strip()
            content = (section.get("content") or "").strip()

            if heading and heading.lower() not in {"general content", "general"}:
                # Each extracted main section begins on a new page in the rebuilt PDF.
                if story:
                    story.append(PageBreak())
                # The college guideline asks for bold, uppercase main article titles.
                story.append(Paragraph(escape(heading.upper()), heading_style))

            if content:
                for block in re.split(r"\n\s*\n|\n", content):
                    block = block.strip()
                    if block:
                        safe_text = escape(block).replace("\n", "<br/>")
                        story.append(Paragraph(safe_text, body_style))

        for table_data in tables:
            rows = table_data.get("rows") or []
            if rows:
                safe_rows = [
                    [Paragraph(escape(str(cell)), body_style) for cell in row]
                    for row in rows
                ]
                report_table = Table(safe_rows, repeatRows=1, hAlign="LEFT")
                report_table.setStyle(TableStyle([
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4)
                ]))
                story.append(Spacer(1, 8))
                story.append(report_table)

        if not story:
            raise ValueError("The report does not contain enough extractable text to format.")

        def add_page_number(canvas, document_obj):
            canvas.saveState()
            canvas.setFont(regular_font, 9)
            canvas.drawRightString(
                A4[0] - other_margin * inch,
                0.55 * inch,
                str(document_obj.page)
            )
            canvas.restoreState()

        doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
        return output_path
    finally:
        image_temp_dir.cleanup()
