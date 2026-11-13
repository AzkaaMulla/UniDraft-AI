# UniDraft AI

**AI-Powered College Report Analysis and Formatting System**

UniDraft AI helps students compare an academic report with college guidelines, identify formatting issues, and generate a consistently formatted copy.

## Features

- Upload a college reference document and a student report.
- Supports DOCX and text-based PDF input.
- Extracts written college requirements such as font, font size, line spacing, margins, and required sections.
- Detects report headings and compares the student's heading sequence with the college's required sections.
- Checks font family, font size, line spacing, page margins, paragraph indentation, heading capitalization, and figure/table/graph numbering where the file format exposes the information.
- Generates an editable formatted DOCX and a reflowed PDF for DOCX student reports.
- Generates a reflowed PDF for PDF student reports.
- Starts each detected main section on a new page; subheadings remain within their section.
- Provides a downloadable compliance report through the results page.

## Technology Stack

- Python
- Flask
- python-docx
- PyMuPDF
- ReportLab
- HTML, CSS and JavaScript

## Project Structure

```text
UniDraft-AI/
├── app.py
├── content_extractor.py
├── format_analyzer.py
├── format_compliance.py
├── report_formatter.py
├── requirements.txt
├── templates/
│   ├── index.html
│   └── upload.html
├── static/
│   └── style.css
├── uploads/
│   └── .gitkeep
└── outputs/
    └── .gitkeep
```

## Setup Instructions

### 1. Clone the repository

```bash
git clone YOUR_GITHUB_REPOSITORY_URL
cd UniDraft-AI
```

### 2. Create a virtual environment

Windows:

```bash
python -m venv venv
venv\Scripts\activate
```

macOS / Linux:

```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Run the application

```bash
python app.py
```

Open the local address shown in the terminal (normally `http://127.0.0.1:5000`).

## How to Use

1. Open UniDraft AI in your browser.
2. Upload the college guideline/reference document.
3. Upload the original student report.
4. Select **Analyze My Documents**.
5. Review the college format blueprint, heading comparison, and formatting compliance report.
6. Download the generated formatted document.

## PDF Support and Limitations

- Text-based PDFs can be analysed and reconstructed into a clean formatted PDF.
- Scanned/image-only PDFs require OCR before their text can be analysed or reformatted.
- Reconstructed PDFs use extracted text and tables where available. They may not preserve the exact original page breaks, images, diagrams, charts, or complex layouts.
- PDF margin and line-spacing checks are estimates based on text positions.
- DOCX files provide more reliable measurements for paragraph indentation and editable formatting.
- Automatic formatting does not rewrite the student's academic content.

## Notes

- Each uploaded file can be up to 15 MB.
- Uploaded source documents are temporarily processed and removed after analysis.
- Generated files are stored in the local `outputs/` folder. That folder is excluded from Git tracking.
- Keep the original report as a backup before using a formatted copy.

## Future Improvements

- OCR for scanned PDFs.
- Better preservation of images, diagrams, and complex tables.
- More advanced heading hierarchy and caption detection.
- User accounts and saved report history.


### Certificate page detection
- Detects common academic certificate headings and wording in DOCX and PDF uploads.
- Reports exact page numbers for PDF files. DOCX page numbers are not reported because pagination depends on Word layout settings.
- Scanned/image-only PDFs may need OCR before certificate text can be detected.

### Formatting safeguards

- Main report headings start on a new page; subheadings and their content remain with the chapter.
- Table of Contents entries are kept as normal-weight entries on the contents page and are not mistaken for chapter headings.
- Numbered advantages and bibliography entries remain normal weight rather than being promoted to bold headings.
- The page-count rule is reported as a warning against the generated PDF copy. UniDraft AI does not delete report content to force a page limit; DOCX pagination can vary slightly between Word installations.


## Visual asset detection and preservation

- DOCX analysis counts embedded pictures, native charts, and tables. Figure/chart captions can also help identify charts inserted as pictures.
- PDF analysis detects embedded images, attempts table extraction with PyMuPDF, and uses captions/vector drawing patterns as a heuristic for possible charts or diagrams.
- The editable formatted DOCX retains its original embedded pictures, charts, and tables because formatting is applied to the existing document rather than rebuilding it from plain text.
- When a PDF is rebuilt, likely cover-page images (including a college logo) are carried into the new PDF. Exact placement of inline images and charts elsewhere in the report is not guaranteed.
- PDF table detection is best-effort. For PDF uploads, table cell text is retained in the reflowed text rather than duplicated as a second table; DOCX tables are rendered as tables in the generated PDF.
- Logo and chart identification is a detection aid, not a visual classifier. Users should confirm whether a detected picture is actually the college logo or a chart.
