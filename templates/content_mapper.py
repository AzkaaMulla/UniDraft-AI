
import re
from difflib import SequenceMatcher


# Common headings that may have different names
# but represent similar academic sections.

HEADING_GROUPS = {
    "introduction": [
        "introduction",
        "overview",
        "background"
    ],

    "literature review": [
        "literature review",
        "related work",
        "existing work",
        "review of literature"
    ],

    "methodology": [
        "methodology",
        "research methodology",
        "system methodology",
        "system method",
        "proposed methodology",
        "proposed system"
    ],

    "results and discussion": [
        "results and discussion",
        "results",
        "experimental results",
        "analysis and results"
    ],

    "conclusion": [
        "conclusion",
        "conclusions",
        "final remarks",
        "summary and conclusion"
    ],

    "references": [
        "references",
        "bibliography",
        "reference list"
    ]
}


def normalize_heading(heading):
    """
    Remove numbering and unnecessary symbols
    from a heading.
    """

    heading = heading.lower().strip()

    # Remove numbering such as 1., 2.1, Chapter 1
    heading = re.sub(
        r"^(chapter\s+)?\d+(\.\d+)*[\s.):\-]+",
        "",
        heading
    )

    # Remove extra symbols
    heading = re.sub(r"[^a-z0-9\s]", "", heading)

    # Remove extra spaces
    heading = " ".join(heading.split())

    return heading


def get_heading_group(heading):
    """
    Find the common category of a heading.
    """

    normalized = normalize_heading(heading)

    for group, variations in HEADING_GROUPS.items():

        if normalized in variations:
            return group

    return normalized


def is_heading(line):
    """
    Identify possible report headings.
    """

    line = line.strip()

    if not line:
        return False

    if len(line) > 100:
        return False

    # Numbered headings: 1. INTRODUCTION, 2.1 OBJECTIVES
    numbered_heading = re.match(
        r"^(chapter\s+)?\d+(\.\d+)*[\s.):\-]+\S",
        line,
        re.IGNORECASE
    )

    if numbered_heading:
        return True

    # Short headings written in uppercase
    if line.isupper() and any(char.isalpha() for char in line):
        return True

    return False


def extract_sections(document_text):
    """
    Extract headings and their related content
    from a document.
    """

    sections = []
    current_heading = None
    current_content = []

    lines = document_text.splitlines()

    for line in lines:

        line = line.strip()

        if not line:
            continue

        if is_heading(line):

            # Save the previous section
            if current_heading is not None:

                sections.append({
                    "heading": current_heading,
                    "content": "\n".join(current_content).strip()
                })

            # Start a new section
            current_heading = line
            current_content = []

        else:

            if current_heading is not None:
                current_content.append(line)

    # Save the final section
    if current_heading is not None:

        sections.append({
            "heading": current_heading,
            "content": "\n".join(current_content).strip()
        })

    return sections


def calculate_similarity(reference_heading, student_heading):
    """
    Calculate similarity between two headings.
    """

    reference_group = get_heading_group(reference_heading)
    student_group = get_heading_group(student_heading)

    # Exact category match
    if reference_group == student_group:
        return 1.0

    # Compare heading names
    similarity = SequenceMatcher(
        None,
        reference_group,
        student_group
    ).ratio()

    return round(similarity, 2)


def map_report_sections(reference_text, student_text):
    """
    Match college reference headings with
    student report headings.
    """

    reference_sections = extract_sections(reference_text)
    student_sections = extract_sections(student_text)

    mappings = []
    used_student_sections = set()

    for reference in reference_sections:

        best_match = None
        best_score = 0

        for index, student in enumerate(student_sections):

            if index in used_student_sections:
                continue

            score = calculate_similarity(
                reference["heading"],
                student["heading"]
            )

            if score > best_score:
                best_score = score
                best_match = (index, student)

        if best_match and best_score >= 0.55:

            index, student = best_match

            used_student_sections.add(index)

            mappings.append({
                "reference_heading": reference["heading"],
                "student_heading": student["heading"],
                "similarity": round(best_score * 100, 1),
                "status": "Matched",
                "student_content": student["content"]
            })

        else:

            mappings.append({
                "reference_heading": reference["heading"],
                "student_heading": "No matching section found",
                "similarity": 0,
                "status": "Not Matched",
                "student_content": ""
            })

    # Find student headings that were not matched
    unmatched_sections = []

    for index, student in enumerate(student_sections):

        if index not in used_student_sections:

            unmatched_sections.append({
                "heading": student["heading"],
                "content": student["content"]
            })

    return {
        "total_reference_sections": len(reference_sections),
        "total_student_sections": len(student_sections),
        "matched_sections": sum(
            1 for item in mappings
            if item["status"] == "Matched"
        ),
        "unmatched_sections": unmatched_sections,
        "mappings": mappings
    }