import re
from statistics import median


def _top_value(items):
    return items[0][0] if items else None


def _number(value):
    try:
        if value in (None, "", "Not specified", "Not detected"):
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def _normal_font(value):
    value = (value or "").lower()
    value = re.sub(r"^[a-z]{6}\+", "", value)
    value = re.sub(r"[^a-z]", "", value)
    if "timesnewroman" in value or value in {"timesroman", "times"}:
        return "times new roman"
    return value


def _add(checks, name, expected, actual, status, note=""):
    checks.append({
        "name": name,
        "expected": expected,
        "actual": actual if actual not in (None, "") else "Not detected",
        "status": status,
        "note": note
    })


def build_compliance_report(guidelines, student_profile, student_text=""):
    """Compare student document measurements with extracted written rules."""
    guidelines = guidelines or {}
    profile = student_profile or {}
    checks = []

    expected_font = guidelines.get("font_family")
    actual_font = _top_value(profile.get("font_names", []))
    if expected_font and actual_font:
        passed = _normal_font(expected_font) == _normal_font(actual_font)
        _add(checks, "Font family", expected_font, actual_font,
             "Passed" if passed else "Needs attention")
    else:
        _add(checks, "Font family", expected_font or "Not specified",
             actual_font, "Not verifiable")

    expected_size = _number(guidelines.get("font_size"))
    actual_size = _number(_top_value(profile.get("font_sizes", [])))
    if expected_size is not None and actual_size is not None:
        passed = abs(expected_size - actual_size) <= 0.5
        _add(checks, "Font size", f"{expected_size:g} pt", f"{actual_size:g} pt",
             "Passed" if passed else "Needs attention",
             "Compared with the most frequently detected font size.")
    else:
        _add(checks, "Font size", guidelines.get("font_size", "Not specified"),
             actual_size, "Not verifiable")

    expected_spacing = _number(guidelines.get("line_spacing"))
    spacing_items = profile.get("line_spacing", [])
    actual_spacing = _number(_top_value(spacing_items))
    if expected_spacing is not None and actual_spacing is not None:
        passed = abs(expected_spacing - actual_spacing) <= 0.15
        _add(checks, "Line spacing", f"{expected_spacing:g}",
             f"{actual_spacing:g}", "Passed" if passed else "Needs attention",
             "PDF line spacing is estimated from text-line positions.")
    else:
        _add(checks, "Line spacing", guidelines.get("line_spacing", "Not specified"),
             actual_spacing, "Not verifiable",
             "This value could not be reliably detected.")

    page_layouts = profile.get("page_layout", [])
    if page_layouts:
        layout = page_layouts[0]
        expected_left = _number(guidelines.get("left_margin"))
        expected_other = _number(guidelines.get("other_margins"))
        for key, label, expected in (
            ("left_margin", "Left margin", expected_left),
            ("top_margin", "Top margin", expected_other),
            ("bottom_margin", "Bottom margin", expected_other),
            ("right_margin", "Right margin", expected_other)
        ):
            actual = _number(layout.get(key))
            if expected is None:
                _add(checks, label, "Not specified", actual, "Not verifiable")
            elif actual is None:
                _add(checks, label, f"{expected:g} in", None, "Not verifiable")
            else:
                tolerance = 0.15 if profile.get("file_type") == "PDF" else 0.05
                passed = abs(expected - actual) <= tolerance
                note = "Estimated from PDF text placement." if profile.get("file_type") == "PDF" else ""
                _add(checks, label, f"{expected:g} in", f"{actual:g} in",
                     "Passed" if passed else "Needs attention", note)
    else:
        for label in ("Left margin", "Top margin", "Bottom margin", "Right margin"):
            _add(checks, label, "College guideline", None, "Not verifiable")

    expected_indent = 6.0
    indent_items = profile.get("first_line_indents", [])
    if profile.get("file_type") == "DOCX" and indent_items:
        total = sum(count for _, count in indent_items) or 1
        matching = sum(
            count for value, count in indent_items
            if _number(value) is not None and abs(float(value) - expected_indent) <= 1.0
        )
        ratio = matching / total
        _add(checks, "Paragraph first-line indent", "6 pt",
             f"{ratio * 100:.0f}% of paragraphs have 6 pt",
             "Passed" if ratio >= 0.8 else "Needs attention",
             "Checked paragraph indentation in the DOCX file.")
    else:
        _add(checks, "Paragraph first-line indent", "6 pt", None,
             "Not verifiable",
             "PDF does not reliably expose Word paragraph indentation.")

    paragraph_details = profile.get("paragraph_details", [])
    main_headings = [
        item for item in paragraph_details
        if item.get("is_heading") and item.get("heading_level", 0) == 1
    ]
    if main_headings:
        good = sum(1 for item in main_headings if item.get("bold") and item.get("all_caps"))
        ratio = good / len(main_headings)
        _add(checks, "Main headings: bold and ALL CAPITALS",
             "All main headings bold and uppercase",
             f"{good}/{len(main_headings)} compliant",
             "Passed" if ratio == 1 else "Needs attention")
    else:
        _add(checks, "Main headings: bold and ALL CAPITALS",
             "Bold and uppercase", None, "Not verifiable",
             "No main headings were confidently identified.")

    subheadings = [
        item for item in paragraph_details
        if item.get("is_heading") and item.get("heading_level", 0) > 1
    ]
    if subheadings:
        def is_title_case(text):
            words = re.findall(r"[A-Za-z][A-Za-z'-]*", re.sub(r"^\s*\d+(?:\.\d+)*[.)]?\s*", "", text))
            return bool(words) and all(word[0].isupper() for word in words)
        good = sum(1 for item in subheadings if is_title_case(item.get("text", "")))
        _add(checks, "Subheading capitalization", "First letter of each word capitalized",
             f"{good}/{len(subheadings)} compliant",
             "Passed" if good == len(subheadings) else "Needs attention")
    else:
        _add(checks, "Subheading capitalization", "Title Case", None,
             "Not applicable", "No subheadings were identified.")

    caption_lines = [
        line.strip() for line in (student_text or "").splitlines()
        if re.match(r"^\s*(?:figure|fig\.?|table|graph)\s+", line, re.IGNORECASE)
    ]
    if caption_lines:
        numbered = sum(
            1 for line in caption_lines
            if re.match(
                r"^\s*(?:figure|fig\.?|table|graph)\s+\d+\.\d+\b",
                line,
                re.IGNORECASE
            )
        )
        _add(checks, "Figure, table and graph numbering",
             "Chapter-wise numbering (e.g. 2.1, 2.2)",
             f"{numbered}/{len(caption_lines)} captions follow chapter-wise numbering",
             "Passed" if numbered == len(caption_lines) else "Needs attention")
    else:
        _add(checks, "Figure, table and graph numbering",
             "Chapter-wise numbering", None, "Not applicable",
             "No figure, table or graph captions were detected.")

    page_range = guidelines.get("page_count", "Not specified")
    range_match = re.search(r"(\d+)\s*[–-]\s*(\d+)", str(page_range))
    total_pages = profile.get("total_pages")
    if range_match and total_pages is not None:
        low, high = int(range_match.group(1)), int(range_match.group(2))
        passed = low <= total_pages <= high
        _add(checks, "Report page count", f"{low}–{high} pages",
             f"{total_pages} pages", "Passed" if passed else "Needs attention")
    elif range_match:
        _add(checks, "Report page count", page_range, None,
             "Not verifiable", "Page count is not available for DOCX without pagination.")

    applicable = [item for item in checks if item["status"] in {"Passed", "Needs attention"}]
    passed_count = sum(1 for item in applicable if item["status"] == "Passed")
    score = round((passed_count / len(applicable)) * 100) if applicable else 0

    return {
        "checks": checks,
        "passed": passed_count,
        "failed": sum(1 for item in applicable if item["status"] == "Needs attention"),
        "not_verifiable": sum(1 for item in checks if item["status"] == "Not verifiable"),
        "not_applicable": sum(1 for item in checks if item["status"] == "Not applicable"),
        "score": score,
        "total_checked": len(applicable)
    }


def update_generated_pdf_page_count(compliance_report, guidelines, page_count):
    """Add/update a page-count warning using the PDF generated by UniDraft AI."""
    if not compliance_report or page_count is None:
        return compliance_report

    page_range = str((guidelines or {}).get("page_count", "Not specified"))
    match = re.search(r"(\d+)\s*[–-]\s*(\d+)", page_range)
    if not match:
        return compliance_report

    minimum, maximum = int(match.group(1)), int(match.group(2))
    passed = minimum <= int(page_count) <= maximum
    status = "Passed" if passed else "Needs attention"
    actual = f"{int(page_count)} pages (generated PDF)"
    note = (
        "Page count is within the college range. No report content was removed."
        if passed else
        "The generated PDF is outside the college page range. This is a warning only; "
        "UniDraft AI has not deleted any content. DOCX pagination may differ slightly "
        "when opened in Microsoft Word."
    )

    check = next(
        (item for item in compliance_report.get("checks", [])
         if item.get("name") == "Report page count"),
        None
    )
    if check is None:
        check = {"name": "Report page count"}
        compliance_report.setdefault("checks", []).append(check)

    check.update({
        "expected": f"{minimum}–{maximum} pages",
        "actual": actual,
        "status": status,
        "note": note
    })

    applicable = [
        item for item in compliance_report["checks"]
        if item.get("status") in {"Passed", "Needs attention"}
    ]
    passed_count = sum(1 for item in applicable if item.get("status") == "Passed")
    compliance_report["passed"] = passed_count
    compliance_report["failed"] = sum(
        1 for item in applicable if item.get("status") == "Needs attention"
    )
    compliance_report["not_verifiable"] = sum(
        1 for item in compliance_report["checks"]
        if item.get("status") == "Not verifiable"
    )
    compliance_report["not_applicable"] = sum(
        1 for item in compliance_report["checks"]
        if item.get("status") == "Not applicable"
    )
    compliance_report["total_checked"] = len(applicable)
    compliance_report["score"] = round(
        passed_count / len(applicable) * 100
    ) if applicable else 0
    return compliance_report
