import re
import os
import uuid
from pathlib import Path
from collections import Counter

import fitz
from docx import Document
from flask import Flask, render_template, request, send_from_directory, abort
from werkzeug.utils import secure_filename

from format_analyzer import analyze_reference
from content_extractor import extract_document_structure
from format_compliance import build_compliance_report, update_generated_pdf_page_count
from report_formatter import format_docx_report, create_formatted_pdf


app = Flask(__name__)

# Application configuration
UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"docx", "pdf"}
MAX_FILE_SIZE = 32 * 1024 * 1024
MAX_INDIVIDUAL_FILE_SIZE = 15 * 1024 * 1024
OUTPUT_FOLDER = "outputs"

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_SIZE

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# --------------------------------------------------
# FILE VALIDATION AND TEXT EXTRACTION
# --------------------------------------------------

def allowed_file(filename):
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS
    )


def read_docx(file_path):
    document = Document(file_path)
    extracted_text = []

    for paragraph in document.paragraphs:
        if paragraph.text.strip():
            extracted_text.append(paragraph.text)

    # Extract text from tables
    for table in document.tables:
        for row in table.rows:
            row_text = [cell.text.strip() for cell in row.cells]
            extracted_text.append(" | ".join(row_text))

    return "\n".join(extracted_text)


def read_pdf(file_path):
    extracted_text = []

    with fitz.open(file_path) as pdf:
        for page in pdf:
            extracted_text.append(page.get_text())

    return "\n".join(extracted_text)


def extract_text(file_path):
    extension = Path(file_path).suffix.lower()

    if extension == ".docx":
        return read_docx(file_path)

    if extension == ".pdf":
        return read_pdf(file_path)

    raise ValueError("Unsupported document format.")


# --------------------------------------------------
# HEADING COMPARISON
# --------------------------------------------------

def normalize_heading(text):
    """Remove numbering, punctuation and extra spaces."""

    text = text.lower().strip()
    text = re.sub(r"^\s*\d+(?:\.\d+)*[.)]?\s*", "", text)
    text = re.sub(r"[^\w\s]", "", text)
    return " ".join(text.split())


HEADING_ALIASES = {
    "background": "introduction",
    "background of study": "introduction",
    "background of the study": "introduction",
    "overview": "introduction",

    "related work": "literature review",
    "related studies": "literature review",
    "review of literature": "literature review",
    "existing work": "literature review",

    "technologies used": "technology",
    "technology used": "technology",
    "tools and technologies": "technology",
    "tools and technology": "technology",

    "benefits": "advantages",
    "merits": "advantages",

    "limitations": "disadvantages",
    "drawbacks": "disadvantages",
    "demerits": "disadvantages",

    "applications": "application",
    "use cases": "application",
    "practical applications": "application",

    "future work": "future scope",
    "future enhancements": "future scope",
    "future improvements": "future scope",
    "scope for future work": "future scope",

    "summary": "conclusion",
    "concluding remarks": "conclusion",

    "acknowledgement": "acknowledgment",
    "acknowledgements": "acknowledgment",
    "acknowledgments": "acknowledgment",

    "contents": "table of contents",
    "index": "table of contents",

    "certificate": "certificate page",

    "bibliography": "references",
    "works cited": "references"
}


def canonical_heading(text):
    normalized = normalize_heading(text)
    return HEADING_ALIASES.get(normalized, normalized)


def compare_headings(reference_profile, student_profile, required_sections=None):
    """
    Compare student headings against college requirements.
    Uses extracted required sections whenever available.
    """

    if required_sections:
        reference_headings = [
            {"text": section, "level": 1}
            for section in required_sections
        ]
    else:
        reference_headings = reference_profile.get("detected_headings", [])

    student_headings = student_profile.get("detected_headings", [])

    reference_lookup = {}

    for index, heading in enumerate(reference_headings):
        key = canonical_heading(heading["text"])
        reference_lookup.setdefault(key, []).append(index)

    matched = []
    missing = []
    extra = []
    level_differences = []
    matched_reference_positions = []

    exact_matches = 0
    semantic_matches = 0

    for student_heading in student_headings:
        student_text = student_heading["text"]
        key = canonical_heading(student_text)

        if reference_lookup.get(key):
            reference_index = reference_lookup[key].pop(0)
            reference_heading = reference_headings[reference_index]
            matched_reference_positions.append(reference_index)

            exact_match = (
                normalize_heading(reference_heading["text"])
                == normalize_heading(student_text)
            )

            if exact_match:
                match_type = "Exact Match"
                exact_matches += 1
            else:
                match_type = "Semantic Match"
                semantic_matches += 1

            matched.append({
                "reference": reference_heading["text"],
                "student": student_text,
                "match_type": match_type
            })

            reference_level = reference_heading.get("level", 1)
            student_level = student_heading.get("level", 1)

            if reference_level != student_level:
                level_differences.append({
                    "heading": student_text,
                    "reference_level": reference_level,
                    "student_level": student_level
                })

        else:
            extra.append(student_text)

    for positions in reference_lookup.values():
        for position in positions:
            missing.append(reference_headings[position]["text"])

    order_correct = (
        matched_reference_positions
        == sorted(matched_reference_positions)
    )

    return {
        "total_reference_headings": len(reference_headings),
        "total_student_headings": len(student_headings),
        "exact_matches": exact_matches,
        "semantic_matches": semantic_matches,
        "matched": matched,
        "missing": missing,
        "extra": extra,
        "level_differences": level_differences,
        "order_correct": order_correct,
        "heading_order_status": "Correct" if order_correct else "Different"
    }


# --------------------------------------------------
# READ WRITTEN COLLEGE REQUIREMENTS
# --------------------------------------------------

def extract_written_guidelines(reference_text):
    """
    Read the requirements stated in the guideline document.
    These values are intentionally kept separate from the
    formatting physically detected in the uploaded DOCX/PDF.
    """

    text = " ".join(reference_text.split())

    def find_number(pattern):
        match = re.search(pattern, text, re.IGNORECASE)
        return float(match.group(1)) if match else "Not specified"

    font_match = re.search(
        r"\bFont\s*:\s*([A-Za-z ]+?)(?=\s+Font Size|\s+Font\s+Size|$)",
        text,
        re.IGNORECASE
    )
    font_family = font_match.group(1).strip() if font_match else "Not specified"

    font_size = find_number(
        r"\bFont\s*Size\s*:\s*(\d+(?:\.\d+)?)"
    )

    line_spacing = find_number(
        r"\bLine\s*Spacing\s*:\s*(\d+(?:\.\d+)?)"
    )

    left_margin = find_number(
        r"\bLeft\s*:\s*(\d+(?:\.\d+)?)\s*[\"”]?"
    )

    other_margin_match = re.search(
        r"all\s+other\s+margins\s*:\s*(\d+(?:\.\d+)?)",
        text,
        re.IGNORECASE
    )
    other_margin = (
        float(other_margin_match.group(1))
        if other_margin_match else "Not specified"
    )

    page_count_match = re.search(
        r"(\d+)\s*to\s*(\d+)\s*pages",
        text,
        re.IGNORECASE
    )

    page_count = (
        f'{page_count_match.group(1)}–{page_count_match.group(2)} pages'
        if page_count_match else "Not specified"
    )

    rules = []

    if re.search(r"bold\s+and\s+all\s+capitals", text, re.IGNORECASE):
        rules.append("Chapter titles and main article titles must be bold and in ALL CAPITALS.")

    if re.search(r"six\s+point\s+indent", text, re.IGNORECASE):
        rules.append("Use a 6-point indent before every paragraph.")

    if re.search(r"first\s+alphabet\s+of\s+each\s+word", text, re.IGNORECASE):
        rules.append("Subtitles must use Title Case (first letter of each word capitalized).")

    if re.search(r"figure.*table.*graph", text, re.IGNORECASE):
        rules.append("Number figures, tables and graphs chapter-wise (for example, 2.1, 2.2).")

    if re.search(r"one\s+page\s+synopsis", text, re.IGNORECASE):
        rules.append("The abstract should be a one-page synopsis.")

    if re.search(r"page\s+numbers", text, re.IGNORECASE):
        rules.append("The table of contents should include page numbers.")

    return {
        "font_family": font_family,
        "font_size": font_size,
        "line_spacing": line_spacing,
        "left_margin": left_margin,
        "other_margins": other_margin,
        "page_count": page_count,
        "rules": rules
    }


def apply_guidelines_to_blueprint(format_profile, required_sections, written):
    """
    Make the main blueprint show the written college requirements.
    Actual DOCX/PDF measurements remain in the detected-formatting
    sections below the blueprint.
    """

    blueprint = format_profile.get("format_blueprint", {})

    blueprint["font_family"] = written["font_family"]
    blueprint["body_font_size"] = written["font_size"]
    blueprint["alignment"] = "Not specified in written guidelines"
    blueprint["line_spacing"] = written["line_spacing"]

    # Keep physical page dimensions detected from the file,
    # but show the required margins from the written guideline.
    detected_page = (
        format_profile.get("page_layout", [{}])[0]
        if format_profile.get("page_layout") else {}
    )

    blueprint["page_layout"] = {
        "page_width": detected_page.get("page_width", detected_page.get("width", "Not detected")),
        "page_height": detected_page.get("page_height", detected_page.get("height", "Not detected")),
        "top_margin": written["other_margins"],
        "bottom_margin": written["other_margins"],
        "left_margin": written["left_margin"],
        "right_margin": written["other_margins"]
    }

    # The reference's required sections are more reliable than
    # heuristic heading detection in a guideline/instruction document.
    required_heading_items = [
        {"text": section, "level": 1, "type": "Extracted College Requirement"}
        for section in required_sections
    ]

    blueprint["total_headings"] = len(required_heading_items)
    blueprint["main_headings"] = required_heading_items
    blueprint["subheadings"] = []
    blueprint["all_headings"] = required_heading_items

    # Make the heading list shown in the reference analysis consistent.
    format_profile["detected_headings"] = required_heading_items
    format_profile["heading_count"] = len(required_heading_items)

    blueprint["written_guidelines"] = written
    format_profile["format_blueprint"] = blueprint
    format_profile["analysis_note"] = (
        "College requirements were extracted from the written guideline. "
        "Actual formatting detected in the uploaded reference file is "
        "shown separately below."
    )


# --------------------------------------------------
# LANDING PAGE
# --------------------------------------------------

@app.route("/")
def home():
    return render_template("index.html")


# --------------------------------------------------
# UPLOAD AND ANALYZE DOCUMENTS
# --------------------------------------------------

@app.route("/upload", methods=["GET", "POST"])
def upload():

    results = []
    format_profile = None
    student_profile = None
    comparison = None
    error = None
    required_sections = None
    written_guidelines = {}
    compliance_report = None
    downloads = []
    formatting_warning = None
    certificate_detection = None
    visual_detection = None

    if request.method == "POST":

        reference = request.files.get("reference")
        student = request.files.get("student")

        if not reference or not student:
            error = "Please upload both documents before proceeding."

        elif not reference.filename or not student.filename:
            error = "Please select both documents before proceeding."

        elif (
            not allowed_file(reference.filename)
            or not allowed_file(student.filename)
        ):
            error = "Only PDF and DOCX files are supported."

        else:
            try:
                documents = [
                    ("College Reference Report", reference),
                    ("Student Original Report", student)
                ]

                for label, uploaded_file in documents:

                    filename = secure_filename(uploaded_file.filename)
                    unique_name = f"{uuid.uuid4().hex}_{filename}"

                    file_path = os.path.join(
                        app.config["UPLOAD_FOLDER"],
                        unique_name
                    )

                    uploaded_file.save(file_path)

                    try:
                        if os.path.getsize(file_path) > MAX_INDIVIDUAL_FILE_SIZE:
                            raise ValueError("Each uploaded file must be 15 MB or smaller.")

                        extracted_text = extract_text(file_path)

                        # Analyze the two documents using different roles.
                        if label == "College Reference Report":
                            document_profile = analyze_reference(
                                file_path,
                                document_role="reference"
                            )

                            format_profile = document_profile

                            # Extract written requirements from the reference.
                            reference_structure = extract_document_structure(
                                file_path
                            )

                            guidelines = reference_structure.get(
                                "guidelines", {}
                            )

                            required_sections = guidelines.get(
                                "required_sections", []
                            )

                            written_guidelines = extract_written_guidelines(
                                extracted_text
                            )

                            apply_guidelines_to_blueprint(
                                format_profile,
                                required_sections,
                                written_guidelines
                            )

                            print(
                                "\nCOLLEGE REQUIRED SECTIONS:",
                                required_sections
                            )

                            print(
                                "\nWRITTEN COLLEGE GUIDELINES:",
                                written_guidelines
                            )

                        else:
                            document_profile = analyze_reference(
                                file_path,
                                document_role="student"
                            )

                            student_profile = document_profile

                            # Detect whether the student report contains an
                            # academic certificate page/content.
                            student_structure = extract_document_structure(file_path)
                            certificate_detection = student_structure.get(
                                "certificate_detection"
                            )
                            visual_detection = student_structure.get("visual_assets")

                            # Compare the student's actual formatting with the
                            # written college requirements.
                            compliance_report = build_compliance_report(
                                written_guidelines,
                                student_profile,
                                extracted_text
                            )

                            # Generate corrected files while the uploaded
                            # student report is still available.
                            output_id = uuid.uuid4().hex[:10]
                            original_stem = Path(filename).stem
                            extension = Path(filename).suffix.lower()

                            try:
                                if extension == ".docx":
                                    formatted_docx = f"{original_stem}_formatted_{output_id}.docx"
                                    formatted_docx_path = os.path.join(
                                        OUTPUT_FOLDER, formatted_docx
                                    )
                                    format_docx_report(
                                        file_path,
                                        formatted_docx_path,
                                        written_guidelines
                                    )
                                    downloads.append({
                                        "label": "Download formatted DOCX",
                                        "filename": formatted_docx
                                    })

                                    # Also provide a PDF copy of the formatted
                                    # DOCX using a clean, reflowed layout.
                                    formatted_pdf = f"{original_stem}_formatted_{output_id}.pdf"
                                    formatted_pdf_path = os.path.join(
                                        OUTPUT_FOLDER, formatted_pdf
                                    )
                                    create_formatted_pdf(
                                        formatted_docx_path,
                                        formatted_pdf_path,
                                        written_guidelines
                                    )
                                    downloads.append({
                                        "label": "Download formatted PDF",
                                        "filename": formatted_pdf
                                    })
                                    with fitz.open(formatted_pdf_path) as generated_pdf_doc:
                                        compliance_report = update_generated_pdf_page_count(
                                            compliance_report,
                                            written_guidelines,
                                            generated_pdf_doc.page_count
                                        )
                                else:
                                    formatted_pdf = f"{original_stem}_formatted_{output_id}.pdf"
                                    formatted_pdf_path = os.path.join(
                                        OUTPUT_FOLDER, formatted_pdf
                                    )
                                    create_formatted_pdf(
                                        file_path,
                                        formatted_pdf_path,
                                        written_guidelines
                                    )
                                    downloads.append({
                                        "label": "Download formatted PDF",
                                        "filename": formatted_pdf
                                    })
                                    with fitz.open(formatted_pdf_path) as generated_pdf_doc:
                                        compliance_report = update_generated_pdf_page_count(
                                            compliance_report,
                                            written_guidelines,
                                            generated_pdf_doc.page_count
                                        )
                            except Exception as format_error:
                                app.logger.exception("Automatic formatting failed")
                                formatting_warning = str(format_error)

                            print(
                                "\nSTUDENT HEADINGS:",
                                student_profile["detected_headings"]
                            )

                        results.append({
                            "name": label,
                            "filename": filename,
                            "text": extracted_text[:10000],
                            "characters": len(extracted_text)
                        })

                    finally:
                        if os.path.exists(file_path):
                            os.remove(file_path)

                # Compare after both documents have been analyzed.
                if format_profile and student_profile:
                    comparison = compare_headings(
                        format_profile,
                        student_profile,
                        required_sections
                    )

                    print("\nHEADING COMPARISON:", comparison)

            except ValueError as processing_error:
                app.logger.warning("Document processing could not continue: %s", processing_error)
                error = str(processing_error)
                results = []
                format_profile = None
                student_profile = None
                comparison = None
                compliance_report = None
                downloads = []
                certificate_detection = None
                visual_detection = None

            except Exception:
                app.logger.exception("Document processing failed")

                error = (
                    "Document processing failed. "
                    "Please check your files and try again."
                )

                results = []
                format_profile = None
                student_profile = None
                comparison = None
                compliance_report = None
                downloads = []
                certificate_detection = None
                visual_detection = None

    return render_template(
        "upload.html",
        results=results,
        error=error,
        format_profile=format_profile,
        student_profile=student_profile,
        comparison=comparison,
        compliance_report=compliance_report,
        downloads=downloads,
        formatting_warning=formatting_warning,
        certificate_detection=certificate_detection,
        visual_detection=visual_detection
    )


# --------------------------------------------------
# DOWNLOAD GENERATED REPORTS
# --------------------------------------------------

@app.route("/download/<filename>")
def download_report(filename):
    safe_name = secure_filename(filename)
    if not safe_name or safe_name != filename:
        abort(404)
    file_path = os.path.join(OUTPUT_FOLDER, safe_name)
    if not os.path.isfile(file_path):
        abort(404)
    return send_from_directory(OUTPUT_FOLDER, safe_name, as_attachment=True)


# --------------------------------------------------
# HANDLE FILES LARGER THAN 15 MB EACH
# --------------------------------------------------

@app.errorhandler(413)
def file_too_large(error):
    return (
        "Upload too large. Maximum request size is 32 MB; each file must be 15 MB or smaller.",
        413
    )


if __name__ == "__main__":
    app.run(debug=os.environ.get("FLASK_DEBUG", "0") == "1")
