
import re
import pymupdf

from pathlib import Path
from urllib.parse import urlparse, unquote
from docx import Document


# --------------------------------------------------
# 1. FILE PATH HANDLING
# --------------------------------------------------

def normalize_file_path(file_path):
    """Convert Windows paths and local file URLs into valid paths."""

    file_path = str(file_path).strip().strip('"').strip("'")

    if file_path.lower().startswith("file:"):
        parsed = urlparse(file_path)

        file_path = unquote(parsed.path)

        if re.match(r"^/[A-Za-z]:", file_path):
            file_path = file_path[1:]

        if not file_path and parsed.netloc:
            file_path = parsed.netloc + parsed.path

    return Path(file_path)


# --------------------------------------------------
# 2. TEXT CLEANING
# --------------------------------------------------

def clean_text(text):
    """Remove unnecessary spaces and empty lines."""

    if not text:
        return ""

    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)

    return text.strip()


# --------------------------------------------------
# 3. HEADING DETECTION
# --------------------------------------------------

def is_heading(text, style_name=""):
    """Identify headings using Word styles and patterns."""

    text = clean_text(text)

    if not text:
        return False

    style_name = (style_name or "").lower()

    if "heading" in style_name or "title" in style_name:
        return True

    # Numbered chapter headings are usually written as "1. INTRODUCTION".
    # A single number followed by ordinary sentence/title text is more likely
    # a list item or a table-of-contents entry, so do not promote it to a heading.
    numbered = re.match(r"^(\d+(?:\.\d+)*\.?)\s+(.+)$", text)
    if numbered:
        numbering, heading_text = numbered.groups()
        letters = [char for char in heading_text if char.isalpha()]
        is_uppercase_heading = bool(letters) and all(char.isupper() for char in letters)
        is_numbered_subheading = numbering.count(".") >= 1 and len(heading_text.split()) <= 12 and not heading_text.endswith((".", "?", "!"))
        if is_uppercase_heading or is_numbered_subheading:
            return True

    if re.match(r"^\bchapter\s+\d+\s+\S", text, re.IGNORECASE):
        return True

    # Roman numeral headings
    if re.match(
        r"^[IVXLCDM]+\.\s+\S",
        text,
        re.IGNORECASE
    ):
        return True

    common_headings = {
        "abstract",
        "acknowledgment",
        "acknowledgement",
        "acknowledgments",
        "acknowledgements",
        "introduction",
        "problem statement",
        "objectives",
        "scope",
        "literature review",
        "methodology",
        "proposed system",
        "system architecture",
        "implementation",
        "results",
        "discussion",
        "conclusion",
        "future scope",
        "references",
        "bibliography"
    }

    if text.lower().rstrip(":") in common_headings:
        return True

    return False


# --------------------------------------------------
# 4. AUTOMATIC GUIDELINE DETECTION
# --------------------------------------------------

def extract_required_sections(text):
    """
    Extract the report sequence from the arrangement
    of report section.

    Supports different section names instead of relying
    only on predefined heading names.
    """

    lines = text.splitlines()

    arrangement_found = False
    collected_sections = []

    for line in lines:

        line = clean_text(line)

        if not line:
            continue

        upper_line = line.upper()

        if "ARRANGEMENT OF REPORT" in upper_line:
            arrangement_found = True
            continue

        if arrangement_found:

            if (
                "INSTRUCTIONS TO BE FOLLOWED" in upper_line
                or "INSTRUCTIONS FOR PREPARING" in upper_line
            ):
                break

            # Ignore standalone page numbers that PDF extraction may return
            # as separate lines between the required report sections.
            if re.fullmatch(r"(?:page\s*)?\d+", line, re.IGNORECASE):
                continue

            # Remove list numbering and bullet symbols.
            line = re.sub(
                r"^\s*(?:\d+\s*[.)]|[•➢▪-])\s*",
                "",
                line
            ).strip()

            line = line.rstrip(":").strip()

            # Ignore explanatory sentences
            if len(line) > 100:
                continue

            if line.lower() in {
                "chapters",
                "chapter"
            }:
                continue

            # Avoid adding introductory instructions
            if any(
                phrase in line.lower()
                for phrase in [
                    "sequence in which",
                    "material should be",
                    "as follows",
                    "arranged and bound"
                ]
            ):
                continue

            if line and line not in collected_sections:
                collected_sections.append(line)

    return collected_sections


def extract_formatting_rules(text):
    """
    Identify font, size, spacing, margins,
    indentation and heading formatting rules.
    """

    rules = {
        "font_name": None,
        "font_size_pt": None,
        "line_spacing": None,
        "margins_inches": {
            "left": None,
            "right": None,
            "top": None,
            "bottom": None
        },
        "chapter_titles_bold": None,
        "chapter_titles_uppercase": None,
        "paragraph_indent_pt": None,
        "subtitle_case": None,
        "figure_table_graph_numbering": None
    }

    lines = text.splitlines()

    for line in lines:

        line = clean_text(line)
        lower_line = line.lower()

        # Font name
        match = re.search(
            r"^\s*font\s*:\s*([A-Za-z][A-Za-z ]*)",
            line,
            re.IGNORECASE
        )

        if match and "font size" not in lower_line:
            rules["font_name"] = match.group(1).strip()

        # Font size
        match = re.search(
            r"font\s+size\s*:\s*(\d+(?:\.\d+)?)",
            line,
            re.IGNORECASE
        )

        if match:
            rules["font_size_pt"] = float(match.group(1))

        # Line spacing
        match = re.search(
            r"line\s+spacing\s*:\s*(\d+(?:\.\d+)?)",
            line,
            re.IGNORECASE
        )

        if match:
            rules["line_spacing"] = float(match.group(1))

        # Left margin
        match = re.search(
            r"left\s*:\s*(\d+(?:\.\d+)?)",
            line,
            re.IGNORECASE
        )

        if "margin" in lower_line and match:
            rules["margins_inches"]["left"] = float(
                match.group(1)
            )

        # Other margins
        match = re.search(
            r"all\s+other\s+margins?\s*:\s*(\d+(?:\.\d+)?)",
            line,
            re.IGNORECASE
        )

        if match:
            value = float(match.group(1))

            rules["margins_inches"]["right"] = value
            rules["margins_inches"]["top"] = value
            rules["margins_inches"]["bottom"] = value

        # Chapter title formatting
        if (
            "chapter titles" in lower_line
            or "main article titles" in lower_line
        ):

            if "bold" in lower_line:
                rules["chapter_titles_bold"] = True

            if "all capitals" in lower_line or "uppercase" in lower_line:
                rules["chapter_titles_uppercase"] = True


        # Paragraph indentation (supports words and numbers)
        match = re.search(
            r"\b(one|two|three|four|five|six|seven|eight|nine|ten|\d+(?:\.\d+)?)\s*[- ]?point\s+indent",
            line,
            re.IGNORECASE
        )

        if match:
            values = {
                "one": 1, "two": 2, "three": 3,
                "four": 4, "five": 5, "six": 6,
                "seven": 7, "eight": 8, "nine": 9,
                "ten": 10
            }

            value = match.group(1).lower()

            rules["paragraph_indent_pt"] = (
                values[value] if value in values else float(value)
            )

        # Subtitle formatting
        if "subtitle" in lower_line:

            if (
                "first alphabet" in lower_line
                or "first letter" in lower_line
            ):
                rules["subtitle_case"] = "title_case"

        # Figure, table and graph numbering
        if (
            "figure" in lower_line
            and "table" in lower_line
            and "graph" in lower_line
        ):

            if "chapter wise" in lower_line or "chapter-wise" in lower_line:
                rules["figure_table_graph_numbering"] = "chapter_wise"

    return rules


def extract_special_instructions(text):
    """
    Identify additional instructions related to
    page count, abstract, contents, title pages and references.
    """

    instructions = {
        "minimum_pages": None,
        "maximum_pages": None,
        "includes_certificates": None,
        "abstract_length": None,
        "table_of_contents_page_numbers": None,
        "preserve_title_certificate_design": None,
        "reference_author_year_order": None
    }

    # Page count
    match = re.search(
        r"(\d+)\s*(?:to|-)\s*(\d+)\s*pages?",
        text,
        re.IGNORECASE
    )

    if match:
        instructions["minimum_pages"] = int(match.group(1))
        instructions["maximum_pages"] = int(match.group(2))

    if re.search(
        r"pages?\s*\(including certificates\)",
        text,
        re.IGNORECASE
    ):
        instructions["includes_certificates"] = True

    # Abstract length
    if re.search(
        r"abstract.{0,100}one page",
        text,
        re.IGNORECASE | re.DOTALL
    ):
        instructions["abstract_length"] = "one_page"

    # Table of contents
    if re.search(
        r"table of contents.{0,200}page numbers",
        text,
        re.IGNORECASE | re.DOTALL
    ):
        instructions["table_of_contents_page_numbers"] = True

    # Title and certificate design
    if (
        re.search(r"title page", text, re.IGNORECASE)
        and re.search(r"certificate page", text, re.IGNORECASE)
        and re.search(
            r"do not change.{0,100}(font|size|color|border)",
            text,
            re.IGNORECASE | re.DOTALL
        )
    ):
        instructions["preserve_title_certificate_design"] = True

    # Reference format
    if re.search(
        r"author.{0,100}followed by year",
        text,
        re.IGNORECASE | re.DOTALL
    ):
        instructions["reference_author_year_order"] = True

    return instructions


def analyze_guidelines(text):
    """
    Main guideline analysis function.

    Separates formatting rules, required sections
    and special instructions.
    """

    text = clean_text(text)

    formatting_rules = extract_formatting_rules(text)

    required_sections = extract_required_sections(text)

    special_instructions = extract_special_instructions(text)

    detected_rules = sum(
        1
        for value in formatting_rules.values()
        if value is not None
    )

    detected_rules += sum(
        1
        for value in special_instructions.values()
        if value is not None
    )

    return {
        "document_type": (
            "college_report_guidelines"
            if required_sections or detected_rules >= 2
            else "unknown"
        ),

        "formatting_rules": formatting_rules,

        "required_sections": required_sections,

        "special_instructions": special_instructions,

        "detection_summary": {
            "formatting_rules_detected": detected_rules,
            "sections_detected": len(required_sections)
        }
    }



# --------------------------------------------------
# CERTIFICATE PAGE DETECTION
# --------------------------------------------------

def _certificate_match(text):
    """Return whether page/document text contains strong certificate evidence."""
    normalized = clean_text(text).lower()
    if not normalized:
        return False, []

    evidence = []
    lines = [clean_text(line).lower().strip(" .:-") for line in text.splitlines()]

    # A standalone certificate heading is strong evidence.
    if any(re.fullmatch(r"(?:certificate|certificate page|certification|bonafide certificate|certificate of completion)", line) for line in lines):
        evidence.append("Certificate heading detected")

    # Common wording used on academic project certificates.
    phrases = [
        (r"this is to certify", "'This is to certify' wording detected"),
        (r"certified that", "'Certified that' wording detected"),
        (r"bonafide student", "Bonafide student wording detected"),
        (r"has been submitted in partial fulfilment", "Academic submission certificate wording detected"),
        (r"has been submitted in partial fulfillment", "Academic submission certificate wording detected"),
    ]
    for pattern, label in phrases:
        if re.search(pattern, normalized, re.IGNORECASE):
            evidence.append(label)

    # Avoid treating a passing mention in references/body text as a certificate page.
    strong_heading = "Certificate heading detected" in evidence
    academic_phrase = any("wording detected" in item for item in evidence)
    return bool(strong_heading or (academic_phrase and len(normalized) > 80)), evidence


def detect_certificate_page(file_path):
    """Detect an academic certificate page in a DOCX or PDF document.

    PDF page numbers are exact. DOCX pagination depends on Word layout, so the
    detector reports document-level presence without guessing a page number.
    """
    path = normalize_file_path(file_path)
    extension = path.suffix.lower()

    if extension == ".pdf":
        detected_pages = []
        with pymupdf.open(str(path)) as pdf:
            for page_number, page in enumerate(pdf, start=1):
                page_text = page.get_text("text") or ""
                found, evidence = _certificate_match(page_text)
                if found:
                    detected_pages.append({
                        "page": page_number,
                        "evidence": evidence
                    })

        return {
            "status": "Detected" if detected_pages else "Not detected",
            "found": bool(detected_pages),
            "pages": detected_pages,
            "page_number_available": True,
            "message": (
                "Certificate page detected on PDF page(s): " +
                ", ".join(str(item["page"]) for item in detected_pages)
                if detected_pages else
                "No certificate page was identified in the PDF text. If this is a scanned PDF, OCR may be required."
            )
        }

    if extension == ".docx":
        document = Document(str(path))
        parts = [paragraph.text for paragraph in document.paragraphs]
        for table in document.tables:
            for row in table.rows:
                parts.append(" ".join(cell.text for cell in row.cells))
        found, evidence = _certificate_match("\n".join(parts))
        return {
            "status": "Detected" if found else "Not detected",
            "found": found,
            "pages": [],
            "page_number_available": False,
            "evidence": evidence,
            "message": (
                "Certificate content detected in the DOCX. Exact page number is not reported because Word pagination can change with fonts, margins and printer settings."
                if found else
                "No clear certificate heading or standard certificate wording was identified in the DOCX."
            )
        }

    return {
        "status": "Unsupported format",
        "found": False,
        "pages": [],
        "page_number_available": False,
        "message": "Certificate detection supports DOCX and PDF files."
    }

# --------------------------------------------------
# VISUAL ASSET DETECTION
# --------------------------------------------------

def _detect_docx_visual_assets(document):
    """Count embedded pictures, charts and tables in a DOCX document."""
    pictures = 0
    charts = 0
    for shape in document.inline_shapes:
        # python-docx uses 3 for picture and 12 for chart in its inline-shape enum.
        shape_type = int(getattr(shape.type, "value", shape.type))
        if shape_type == 3:
            pictures += 1
        elif shape_type == 12:
            charts += 1

    # Captions such as "Figure 2.1: ..." or "Chart 3.1: ..." can indicate
    # a chart that was inserted as a picture rather than as a native Word chart.
    caption_count = sum(
        1 for paragraph in document.paragraphs
        if re.search(r"\b(?:chart|graph)\s*\d|\bfigure\s*\d", paragraph.text, re.I)
    )
    charts = max(charts, caption_count)
    table_count = len(document.tables)
    return {
        "picture_count": pictures,
        "chart_count": charts,
        "table_count": table_count,
        "logo_detected": pictures > 0,
        "logo_note": (
            "One or more embedded pictures were found. A picture on the first page may be the college logo; please visually confirm it."
            if pictures else
            "No embedded picture was found in the DOCX."
        ),
        "chart_note": (
            f"Detected {charts} embedded chart(s)."
            if charts else
            "No native embedded charts were detected. Charts inserted as ordinary pictures are counted as pictures."
        )
    }


def _detect_pdf_visual_assets(file_path):
    """Detect PDF images, vector drawing regions and tables where supported."""
    picture_count = 0
    chart_count = 0
    table_count = 0
    with pymupdf.open(str(file_path)) as pdf:
        for page in pdf:
            images = page.get_images(full=True)
            picture_count += len(images)
            drawings = page.get_drawings()
            page_text = page.get_text("text") or ""
            has_chart_caption = bool(re.search(
                r"\b(?:chart|graph)\s*\d|\bfigure\s*\d", page_text, re.I
            ))
            # Vector charts commonly contain multiple drawing paths. Captions
            # plus image/vector content improve detection, but remain heuristic.
            if len(drawings) >= 12 or (has_chart_caption and (images or drawings)):
                chart_count += 1
            try:
                found = page.find_tables()
                table_count += len(found.tables)
            except Exception:
                # Table detection depends on PDF layout and PyMuPDF capabilities.
                pass

    return {
        "picture_count": picture_count,
        "chart_count": chart_count,
        "table_count": table_count,
        "logo_detected": picture_count > 0,
        "logo_note": (
            "One or more embedded images were found. An image on the first page may be the college logo; please visually confirm it."
            if picture_count else
            "No embedded images were found in the PDF."
        ),
        "chart_note": (
            f"Possible vector chart/diagram content was detected on {chart_count} page(s). This is a heuristic and needs visual confirmation."
            if chart_count else
            "No likely vector chart pages were detected. Raster charts may be included in the picture count."
        )
    }


# --------------------------------------------------
# 5. DOCX EXTRACTION
# --------------------------------------------------

def extract_docx(file_path):

    document = Document(str(file_path))

    result = {
        "title": "",
        "sections": [],
        "tables": [],
        "headers": [],
        "footers": [],
        "total_pages": None,
        "certificate_detection": detect_certificate_page(file_path),
        "visual_assets": _detect_docx_visual_assets(document)
    }

    paragraphs = document.paragraphs

    raw_text = []

    for paragraph in paragraphs:
        if paragraph.text.strip():
            raw_text.append(paragraph.text)

    # Find document title
    title_paragraph = None

    for paragraph in paragraphs:

        if (
            paragraph.text.strip()
            and paragraph.style
            and paragraph.style.name.lower() == "title"
        ):
            title_paragraph = paragraph
            break

    if title_paragraph is None:

        for paragraph in paragraphs:

            if paragraph.text.strip():
                title_paragraph = paragraph
                break

    if title_paragraph is not None:
        result["title"] = clean_text(title_paragraph.text)

    current_section = None
    in_table_of_contents = False
    in_references = False

    for paragraph in paragraphs:
        text = clean_text(paragraph.text)
        if not text or paragraph is title_paragraph:
            continue

        style_name = paragraph.style.name if paragraph.style else ""
        normalized = re.sub(r"^\s*\d+(?:\.\d+)*[.)]?\s*", "", text).strip().lower().rstrip(":")
        has_heading_style = bool(re.match(r"Heading\s*[1-9]", style_name, re.I))
        uppercase_heading = bool(re.search(r"[A-Za-z]", text)) and text == text.upper()

        if normalized in {"table of contents", "contents"}:
            is_section_heading = True
            in_table_of_contents = True
        elif in_table_of_contents and not (has_heading_style or uppercase_heading):
            is_section_heading = False
        else:
            if in_table_of_contents:
                in_table_of_contents = False
            if normalized in {"references", "bibliography"}:
                in_references = True
                is_section_heading = True
            elif in_references and not (has_heading_style or uppercase_heading):
                is_section_heading = False
            else:
                if in_references:
                    in_references = False
                is_section_heading = is_heading(text, style_name)

        if is_section_heading:
            current_section = {"heading": text, "content": ""}
            result["sections"].append(current_section)
        else:
            if current_section is None:
                current_section = {"heading": "General Content", "content": ""}
                result["sections"].append(current_section)
            if current_section["content"]:
                current_section["content"] += "\n"
            current_section["content"] += text

    # Extract tables
    for table_index, table in enumerate(
        document.tables,
        start=1
    ):

        table_data = []

        for row in table.rows:

            row_data = []

            for cell in row.cells:
                row_data.append(clean_text(cell.text))

            table_data.append(row_data)

        result["tables"].append({
            "table_number": table_index,
            "rows": table_data
        })

        # Include table content in guideline analysis
        for row in table_data:
            raw_text.append(" ".join(row))

    # Extract headers and footers
    for section in document.sections:

        for paragraph in section.header.paragraphs:

            text = clean_text(paragraph.text)

            if text and text not in result["headers"]:
                result["headers"].append(text)

        for paragraph in section.footer.paragraphs:

            text = clean_text(paragraph.text)

            if text and text not in result["footers"]:
                result["footers"].append(text)

    # NEW: Automatic guideline analysis
    result["guidelines"] = analyze_guidelines(
        "\n".join(raw_text)
    )

    return result


# --------------------------------------------------
# 6. PDF EXTRACTION
# --------------------------------------------------

def extract_pdf(file_path):
    """Extract structured information from PDF."""

    result = {
        "title": "",
        "sections": [],
        "tables": [],
        "headers": [],
        "footers": [],
        "total_pages": 0,
        "certificate_detection": detect_certificate_page(file_path),
        "visual_assets": _detect_pdf_visual_assets(file_path)
    }

    current_section = None
    raw_text = []

    with pymupdf.open(str(file_path)) as pdf:

        result["total_pages"] = len(pdf)

        for page in pdf:

            page_height = page.rect.height

            blocks = page.get_text("blocks")

            for block in blocks:

                x0, y0, x1, y1, text = block[:5]

                text = clean_text(text)

                if not text:
                    continue

                if y1 < page_height * 0.08:

                    if text not in result["headers"]:
                        result["headers"].append(text)

                elif y0 > page_height * 0.92:

                    if text not in result["footers"]:
                        result["footers"].append(text)

            page_text = page.get_text("text")

            raw_text.append(page_text)

            lines = page_text.splitlines()

            for line in lines:

                text = clean_text(line)

                if not text:
                    continue

                if re.fullmatch(
                    r"(?:page\s*)?\d+",
                    text,
                    re.IGNORECASE
                ):
                    continue

                if not result["title"]:
                    result["title"] = text
                    continue

                if is_heading(text):

                    current_section = {
                        "heading": text,
                        "content": ""
                    }

                    result["sections"].append(current_section)

                else:

                    if current_section is None:

                        current_section = {
                            "heading": "General Content",
                            "content": ""
                        }

                        result["sections"].append(current_section)

                    if current_section["content"]:
                        current_section["content"] += "\n"

                    current_section["content"] += text

    # Detect and extract PDF tables when PyMuPDF can identify their cell layout.
    with pymupdf.open(str(file_path)) as pdf:
        table_number = 0
        for page_number, page in enumerate(pdf, start=1):
            try:
                found_tables = page.find_tables()
                for found_table in found_tables.tables:
                    extracted_rows = found_table.extract()
                    if not extracted_rows:
                        continue
                    table_number += 1
                    result["tables"].append({
                        "table_number": table_number,
                        "page": page_number,
                        "rows": [
                            [clean_text(cell or "") for cell in row]
                            for row in extracted_rows
                        ]
                    })
            except Exception:
                # Keep text extraction available even if a PDF table is complex.
                continue

    # NEW: Automatic guideline analysis
    result["guidelines"] = analyze_guidelines(
        "\n".join(raw_text)
    )

    return result


# --------------------------------------------------
# 7. MAIN EXTRACTION FUNCTION
# --------------------------------------------------

def extract_document_structure(file_path):
    """
    Main function to extract structured information.

    Supports DOCX and PDF files.
    """

    file_path = normalize_file_path(file_path)

    if not file_path.exists():

        raise FileNotFoundError(
            f"Document not found: {file_path}"
        )

    if not file_path.is_file():

        raise ValueError(
            "The provided path is not a file."
        )

    extension = file_path.suffix.lower()

    if extension == ".docx":
        return extract_docx(file_path)

    elif extension == ".pdf":
        return extract_pdf(file_path)

    else:

        raise ValueError(
            "Unsupported file format. "
            "Please upload a DOCX or PDF file."
        )
