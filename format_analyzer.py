
from pathlib import Path
from collections import Counter
import re
import fitz
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --------------------------------------------------

# COMMON HEADING DETECTION

# --------------------------------------------------
def detect_heading(text, font_size=None, average_size=None,
                   bold=False, style_name=""):
    """
    Detect headings using numbering, capitalization,
    title case, font size, bold formatting and Word styles.
    """

    text = " ".join(text.split())
    if not text or len(text) > 150:
        return None

    words = text.split()
    if len(words) > 15:
        return None

    # 1. Built-in Word heading styles
    if style_name and re.match(r"Heading\s*[1-9]", style_name, re.I):
        match = re.search(r"\d+", style_name)
        level = int(match.group()) if match else 1
        return {
            "text": text,
            "level": level,
            "type": "Word Heading Style"
        }

    # 2. Numbered headings: 1. Introduction, 1.1 Objectives
    numbered = re.match(
            r"^(\d+(?:\.\d+)*\.?)\s+(.+)$",
            text
    )

    if numbered:
        numbering = numbered.group(1).rstrip(".")
        level = numbering.count(".") + 1
        return {
            "text": text,
            "level": min(level, 6),
            "type": "Numbered Heading"
        }

    # 3. Roman numeral headings: CHAPTER I, II, III
    if re.match(
        r"^(?:CHAPTER\s+[IVXLCDM]+|[IVXLCDM]+[.):])\s+\S",
        text,
        re.I
    ):
        return {
            "text": text,
            "level": 1,
            "type": "Roman Numeral Heading"
        }


    # 4. All uppercase headings
    letters = [c for c in text if c.isalpha()]
    if letters and all(c.isupper() for c in letters):
        if len(words) <= 12:
            return {
                "text": text,
                "level": 1,
                "type": "Uppercase Heading"
            }

    # Avoid sentence-like text
    if text.endswith((".", "?", "!", ";")):
        return None

    # 5. Normal Title Case headings
    title_words = [
        word for word in words
        if any(char.isalpha() for char in word)
    ]

    small_words = {
        "a", "an", "the", "and", "or", "but",
        "in", "on", "at", "to", "of", "for",
        "with", "by", "from", "as", "is"
    }

    if title_words:
        capitalized = sum(
            1 for word in title_words
            if (
                word[0].isupper()
                or word.lower() in small_words
            )
        )

        title_ratio = capitalized / len(title_words)

        is_title_case = (
            len(words) <= 12
            and len(text) <= 100
            and title_ratio >= 0.85
            and any(word[0].isupper() for word in title_words)
        )


        # Short title-case phrases are accepted.
        # Longer phrases need additional visual evidence.
        visual_support = (
            bold
            or (
                font_size is not None
                and average_size is not None
                and font_size > average_size * 1.08
            )
        )

        if is_title_case and (
            len(words) <= 6 or visual_support
        ):
            return {
                "text": text,
                "level": 1,
                "type": "Title Case Heading"
            }

    # 6. Bold or larger-font heading
    visual_heading = (
        bold
        or (
            font_size is not None
            and average_size is not None
            and font_size > average_size * 1.18
        )
    )

    if visual_heading and len(words) <= 12:
        return {
            "text": text,
            "level": 1,
            "type": "Formatting-Based Heading"
        }
    return None

# --------------------------------------------------

# ALIGNMENT DETECTION

# --------------------------------------------------

def get_alignment(alignment):
    alignment_map = {
        WD_ALIGN_PARAGRAPH.LEFT: "Left",
        WD_ALIGN_PARAGRAPH.CENTER: "Center",
        WD_ALIGN_PARAGRAPH.RIGHT: "Right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "Justify"
    }
    return alignment_map.get(alignment, "Not specified")

# --------------------------------------------------

# DOCX ANALYSIS

# --------------------------------------------------


def analyze_docx(file_path, document_role="student"):
    """Analyze DOCX typography, paragraph layout, headings and page setup."""
    document = Document(file_path)

    font_names = []
    font_sizes = []
    heading_styles = []
    alignments = []
    paragraph_spacings = []
    first_line_indents = []
    left_indents = []
    detected_headings = []
    paragraph_details = []

    paragraphs = document.paragraphs

    def clean_heading_text(text):
        text = " ".join(text.split())
        text = re.sub(r"^\d+(?:\.\d+)*[.)]?\s*", "", text)
        return text.strip(" :").lower()

    def inherited_paragraph_value(paragraph, attribute):
        value = getattr(paragraph.paragraph_format, attribute, None)
        if value is not None:
            return value
        style = paragraph.style
        while style is not None:
            value = getattr(style.paragraph_format, attribute, None)
            if value is not None:
                return value
            style = style.base_style
        return None

    def inherited_run_value(run, attribute):
        value = getattr(run.font, attribute, None)
        if value is not None:
            return value
        style = run.style
        while style is not None:
            value = getattr(style.font, attribute, None)
            if value is not None:
                return value
            style = style.base_style
        value = getattr(paragraph_style_font, attribute, None)
        return value

    paragraph_style_font = document.styles["Normal"].font

    certificate_index = None
    body_start_index = 0
    if document_role == "student":
        for index, paragraph in enumerate(paragraphs):
            if clean_heading_text(paragraph.text) == "abstract":
                body_start_index = index
                break

        for index, paragraph in enumerate(paragraphs):
            if clean_heading_text(paragraph.text) in {"certificate", "certificate page"}:
                certificate_index = index
                break

        if certificate_index is not None:
            title_page_content = any(
                paragraph.text.strip()
                for paragraph in paragraphs[:certificate_index]
            )
            if title_page_content:
                detected_headings.append({
                    "text": "Title page",
                    "level": 1,
                    "type": "Detected Title Page"
                })

    in_table_of_contents = False
    in_references_section = False

    for index, paragraph in enumerate(paragraphs):
        text = paragraph.text.strip()
        if not text:
            continue

        normalized_text = clean_heading_text(text)
        style_name = paragraph.style.name if paragraph.style else ""
        is_word_heading_style = style_name.lower().startswith("heading")

        if document_role == "student":
            if in_references_section:
                continue

            metadata_labels = (
                "submitted by", "student name", "under the guidance of",
                "project guide", "head of department", "department of",
                "academic year", "date"
            )
            if normalized_text.startswith(metadata_labels):
                continue

            if certificate_index is not None and index < certificate_index:
                continue

            if normalized_text in {"table of contents", "contents"}:
                detected_headings.append({
                    "text": text, "level": 1,
                    "type": "Table of Contents Heading"
                })
                in_table_of_contents = True
                continue

            if in_table_of_contents:
                is_actual_acknowledgment = (
                    normalized_text in {"acknowledgment", "acknowledgement"}
                    and (text.isupper() or is_word_heading_style)
                )
                is_next_real_section_heading = (
                    is_word_heading_style
                    and normalized_text not in {"table of contents", "contents"}
                )
                if is_actual_acknowledgment or is_next_real_section_heading:
                    in_table_of_contents = False
                else:
                    continue

        alignments.append(get_alignment(paragraph.alignment))
        if style_name.startswith("Heading"):
            heading_styles.append(style_name)

        spacing = inherited_paragraph_value(paragraph, "line_spacing")
        if spacing is not None:
            try:
                spacing_value = round(float(spacing), 2)
                paragraph_spacings.append(str(spacing_value))
            except (TypeError, ValueError):
                pass

        first_indent = inherited_paragraph_value(paragraph, "first_line_indent")
        left_indent = inherited_paragraph_value(paragraph, "left_indent")
        first_indent_pt = round(first_indent.pt, 1) if first_indent is not None else 0.0
        left_indent_pt = round(left_indent.pt, 1) if left_indent is not None else 0.0
        # The college sample cover/certificate pages have their own fixed
        # layout; check paragraph indentation from the main report onward.
        if document_role != "student" or index >= body_start_index:
            first_line_indents.append(first_indent_pt)
            left_indents.append(left_indent_pt)

        paragraph_font_sizes = []
        paragraph_fonts = []
        bold_runs = 0
        total_runs = 0

        for run in paragraph.runs:
            if not run.text.strip():
                continue
            total_runs += 1

            run_bold = inherited_run_value(run, "bold")
            if run_bold is True:
                bold_runs += 1

            run_font_name = inherited_run_value(run, "name")
            run_font_size = inherited_run_value(run, "size")

            if run_font_name:
                font_names.append(run_font_name)
                paragraph_fonts.append(run_font_name)

            if run_font_size:
                size = round(run_font_size.pt, 1)
                font_sizes.append(size)
                paragraph_font_sizes.append(size)

        paragraph_size = max(paragraph_font_sizes) if paragraph_font_sizes else None
        is_bold = total_runs > 0 and bold_runs / total_runs >= 0.5

        heading = detect_heading(
            text,
            font_size=paragraph_size,
            average_size=paragraph_size,
            bold=is_bold,
            style_name=style_name
        )

        if heading:
            detected_headings.append(heading)

        paragraph_details.append({
            "text": text,
            "is_heading": bool(heading),
            "heading_level": heading.get("level", 0) if heading else 0,
            "bold": is_bold,
            "all_caps": bool(re.search(r"[A-Za-z]", text)) and text == text.upper(),
            "font_size": paragraph_size,
            "font_family": paragraph_fonts[0] if paragraph_fonts else None,
            "first_line_indent_pt": first_indent_pt,
            "left_indent_pt": left_indent_pt,
            "line_spacing": str(spacing) if spacing is not None else None
        })

        if document_role == "student" and normalized_text in {
            "references", "bibliography", "works cited"
        }:
            if not heading:
                detected_headings.append({
                    "text": text, "level": 1, "type": "References Heading"
                })
            in_references_section = True

    sections = []
    for section in document.sections:
        sections.append({
            "page_width": round(section.page_width.inches, 2),
            "page_height": round(section.page_height.inches, 2),
            "top_margin": round(section.top_margin.inches, 2),
            "bottom_margin": round(section.bottom_margin.inches, 2),
            "left_margin": round(section.left_margin.inches, 2),
            "right_margin": round(section.right_margin.inches, 2)
        })

    return {
        "file_type": "DOCX",
        "font_names": Counter(font_names).most_common(5),
        "font_sizes": Counter(font_sizes).most_common(5),
        "heading_styles": sorted(set(heading_styles)),
        "detected_headings": detected_headings,
        "heading_count": len(detected_headings),
        "alignments": Counter(alignments).most_common(),
        "line_spacing": Counter(paragraph_spacings).most_common(5),
        "first_line_indents": Counter(first_line_indents).most_common(5),
        "left_indents": Counter(left_indents).most_common(5),
        "paragraph_details": paragraph_details,
        "page_layout": sections,
        "analysis_note": "DOCX formatting and heading analysis completed."
    }


def analyze_pdf(file_path, document_role="student"):
    """Analyze PDF typography and estimate page layout from visible text."""
    font_names = []
    font_sizes = []
    alignment_counts = Counter()
    detected_headings = []
    paragraph_details = []
    line_spacing_estimates = []
    page_layout = []
    total_pages = 0
    in_table_of_contents = False
    in_references_section = False
    seen_certificate = False
    title_page_has_content = False

    def clean_pdf_heading(value):
        value = " ".join(value.split())
        value = re.sub(r"^\d+(?:\.\d+)*[.)]?\s*", "", value)
        return value.strip(" :").lower()

    with fitz.open(file_path) as pdf:
        total_pages = len(pdf)

        if document_role == "student":
            has_certificate_heading = False
            for scan_page in pdf:
                if any(
                    clean_pdf_heading(line) in {"certificate", "certificate page"}
                    for line in scan_page.get_text("text").splitlines()
                    if line.strip()
                ):
                    has_certificate_heading = True
                    break
            # If no certificate page is present, do not suppress the entire report.
            seen_certificate = not has_certificate_heading

        for page_number, page in enumerate(pdf, start=1):
            page_data = page.get_text("dict")
            page_lines = []
            text_boxes = []
            page_line_gaps = []

            for block in page_data.get("blocks", []):
                if "lines" not in block:
                    continue

                block_lines = []
                for line in block["lines"]:
                    spans = [span for span in line.get("spans", []) if span["text"].strip()]
                    if not spans:
                        continue

                    text = "".join(span["text"] for span in spans).strip()
                    if not text:
                        continue

                    sizes = [span["size"] for span in spans]
                    average_line_size = sum(sizes) / len(sizes)
                    bold = any(
                        "bold" in span["font"].lower()
                        or bool(span.get("flags", 0) & 16)
                        for span in spans
                    )

                    font_names.extend(span["font"] for span in spans)
                    font_sizes.extend(round(span["size"], 1) for span in spans)

                    bbox = line["bbox"]
                    # Exclude likely running headers/footers (including page numbers)
                    # from page-margin estimates.
                    if bbox[1] > page.rect.height * 0.07 and bbox[3] < page.rect.height * 0.92:
                        text_boxes.append(bbox)
                    item = {"text": text, "size": average_line_size, "bold": bold, "bbox": bbox}
                    page_lines.append(item)
                    block_lines.append(item)

                for previous, current in zip(block_lines, block_lines[1:]):
                    gap = current["bbox"][1] - previous["bbox"][1]
                    base_size = (previous["size"] + current["size"]) / 2
                    if base_size > 0 and 0.8 <= gap / base_size <= 2.5:
                        page_line_gaps.append(round(gap / base_size, 1))

            typical_size = (
                sorted(line["size"] for line in page_lines)[len(page_lines) // 2]
                if page_lines else None
            )

            for line in page_lines:
                text = line["text"]
                normalized = clean_pdf_heading(text)
                heading = None
                should_record = True

                if document_role == "student":
                    if not seen_certificate:
                        if normalized in {"certificate", "certificate page"}:
                            if title_page_has_content:
                                detected_headings.append({
                                    "text": "Title page",
                                    "level": 1,
                                    "type": "Detected Title Page"
                                })
                            seen_certificate = True
                            heading = {
                                "text": text,
                                "level": 1,
                                "type": "Certificate Heading",
                                "page": page_number
                            }
                            detected_headings.append(heading)
                        else:
                            if text.strip():
                                title_page_has_content = True
                            should_record = False

                    else:
                        metadata_labels = (
                            "submitted by", "student name",
                            "under the guidance of", "project guide",
                            "head of department", "department of",
                            "academic year", "date"
                        )
                        if normalized.startswith(metadata_labels):
                            should_record = False
                        elif normalized in {"table of contents", "contents"}:
                            heading = {
                                "text": text,
                                "level": 1,
                                "type": "Table of Contents Heading",
                                "page": page_number
                            }
                            detected_headings.append(heading)
                            in_table_of_contents = True
                        elif in_table_of_contents:
                            if (
                                normalized in {"acknowledgment", "acknowledgement"}
                                and text.isupper()
                            ):
                                in_table_of_contents = False
                            else:
                                should_record = False
                        elif in_references_section:
                            should_record = False

                if should_record and heading is None:
                    heading = detect_heading(
                        text,
                        font_size=line["size"],
                        average_size=typical_size,
                        bold=line["bold"]
                    )
                    if heading:
                        heading["page"] = page_number
                        detected_headings.append(heading)

                if should_record:
                    paragraph_details.append({
                        "text": text,
                        "is_heading": bool(heading),
                        "heading_level": heading.get("level", 0) if heading else 0,
                        "bold": line["bold"],
                        "all_caps": bool(re.search(r"[A-Za-z]", text)) and text == text.upper(),
                        "font_size": round(line["size"], 1),
                        "font_family": None,
                        "first_line_indent_pt": None,
                        "left_indent_pt": None,
                        "line_spacing": None,
                        "page": page_number
                    })

                if document_role == "student" and normalized in {
                    "references", "bibliography", "works cited"
                }:
                    in_references_section = True

                x0, y0, x1, y1 = line["bbox"]
                center = (x0 + x1) / 2
                page_width = page.rect.width
                if abs(center - page_width / 2) < 20:
                    alignment_counts["Center"] += 1
                elif x0 > page_width * 0.55:
                    alignment_counts["Right"] += 1
                else:
                    alignment_counts["Left"] += 1

            line_spacing_estimates.extend(page_line_gaps)

            if text_boxes:
                min_x = min(box[0] for box in text_boxes)
                min_y = min(box[1] for box in text_boxes)
                max_x = max(box[2] for box in text_boxes)
                max_y = max(box[3] for box in text_boxes)
                page_layout.append({
                    "width": round(page.rect.width / 72, 2),
                    "height": round(page.rect.height / 72, 2),
                    "left_margin": round(min_x / 72, 2),
                    "top_margin": round(min_y / 72, 2),
                    "right_margin": (
                        round((page.rect.width - max_x) / 72, 2)
                        if max_x >= page.rect.width * 0.80 else None
                    ),
                    "bottom_margin": (
                        round((page.rect.height - max_y) / 72, 2)
                        if max_y >= page.rect.height * 0.75 else None
                    ),
                    "margin_estimate": True
                })
            else:
                page_layout.append({
                    "width": round(page.rect.width / 72, 2),
                    "height": round(page.rect.height / 72, 2),
                    "left_margin": None, "top_margin": None,
                    "right_margin": None, "bottom_margin": None,
                    "margin_estimate": True
                })

    return {
        "file_type": "PDF",
        "font_names": Counter(font_names).most_common(5),
        "font_sizes": Counter(font_sizes).most_common(5),
        "heading_styles": [],
        "detected_headings": detected_headings,
        "heading_count": len(detected_headings),
        "alignments": alignment_counts.most_common(),
        "line_spacing": Counter(line_spacing_estimates).most_common(5),
        "first_line_indents": [],
        "left_indents": [],
        "paragraph_details": paragraph_details,
        "page_layout": page_layout,
        "total_pages": total_pages,
        "analysis_note": (
            "PDF formatting and heading analysis completed. Page margins and "
            "line spacing are estimated from extracted text positions; paragraph "
            "indentation and Word heading styles cannot be verified reliably in PDF."
        )
    }


def build_format_blueprint(profile):
    """
    Organize extracted formatting information
    into a structured college format blueprint.
    """
    font_names = profile.get("font_names", [])
    font_sizes = profile.get("font_sizes", [])
    alignments = profile.get("alignments", [])
    line_spacing = profile.get("line_spacing", [])
    page_layout = profile.get("page_layout", [])
    headings = profile.get("detected_headings", [])

    # Most frequently used font
    preferred_font = (
        font_names[0][0] if font_names else "Not detected"
    )


    # Most frequently used font size
    body_font_size = (
        font_sizes[0][0] if font_sizes else "Not detected"
    )

    # Most frequently used alignment
    preferred_alignment = (
        alignments[0][0] if alignments else "Not detected"
    )

    # Most frequently detected line spacing
    preferred_line_spacing = (
        line_spacing[0][0] if line_spacing else "Not detected"
    )

    # Page layout information
    page_information = {}

    if page_layout:
        page_information = page_layout[0]

    # Organize heading hierarchy
    main_headings = []
    subheadings = []

    for heading in headings:
        if heading.get("level", 1) == 1:
            main_headings.append(heading)
        else:
            subheadings.append(heading)

    blueprint = {
        "font_family": preferred_font,
        "body_font_size": body_font_size,
        "alignment": preferred_alignment,
        "line_spacing": preferred_line_spacing,
        "page_layout": page_information,
        "total_headings": len(headings),
        "main_headings": main_headings,
        "subheadings": subheadings,
        "all_headings": headings
    }
    return blueprint


def analyze_reference(file_path, document_role="student"):

    extension = Path(file_path).suffix.lower()

    if extension == ".docx":
        profile = analyze_docx(
            file_path,
            document_role=document_role
        )

    elif extension == ".pdf":
        profile = analyze_pdf(
            file_path,
            document_role=document_role
        )

    else:
        raise ValueError(
            "Unsupported reference document format."
        )

    profile["format_blueprint"] = build_format_blueprint(profile)

    return profile