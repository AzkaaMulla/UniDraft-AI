
import json
from content_extractor import extract_document_structure


file_path = input(
    "Enter the complete path of your report: "
)

try:

    result = extract_document_structure(file_path)

    print("\n========== EXTRACTED REPORT ==========\n")

    print(
        json.dumps(
            result,
            indent=4,
            ensure_ascii=False
        )
    )

    print("\n========== EXTRACTION SUMMARY ==========")

    print("Title:", result["title"])
    print("Sections:", len(result["sections"]))
    print("Tables:", len(result["tables"]))
    print("Headers:", len(result["headers"]))
    print("Footers:", len(result["footers"]))
    print("Total Pages:", result["total_pages"])
    print("Visual Assets:", result.get("visual_assets", {}))

except Exception as error:

    print("\nExtraction Error:", error)
