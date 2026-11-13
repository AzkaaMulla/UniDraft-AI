::: {align="center"}
`<img src="https://capsule-render.vercel.app/api?type=rect&color=gradient&customColorList=12,18,20&height=190&section=header&text=UniDraft%20AI&fontSize=58&fontColor=ffffff&animation=fadeIn&desc=Smart%20Academic%20Report%20Formatting&descAlignY=75&descSize=18" width="100%" alt="UniDraft AI banner"/>`{=html}
✨ Turn your academic content into a neatly structured report.
```{=html}
<p>
```
`<img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"/>`{=html}
`<img src="https://img.shields.io/badge/Flask-222222?style=for-the-badge&logo=flask&logoColor=white" alt="Flask"/>`{=html}
`<img src="https://img.shields.io/badge/DOCX-2B579A?style=for-the-badge&logo=microsoftword&logoColor=white" alt="DOCX"/>`{=html}
`<img src="https://img.shields.io/badge/PDF-Basic%20Support-BD5757?style=for-the-badge&logo=adobeacrobatreader&logoColor=white" alt="PDF"/>`{=html}
```{=html}
</p>
```
Developed by Saad
:::
---
🌷 About the Project
UniDraft AI is a Python-based web application designed to make
academic report formatting easier and more organized.
Students often spend a lot of time adjusting report headings, fonts,
spacing, and document structure to match their college's prescribed
format. UniDraft AI aims to simplify this process by analyzing a college
reference document and using the extracted formatting information as a
guide for the student's report.
The application works with uploaded documents, extracts their content
and formatting-related information, and presents the analysis through a
web interface.
> **Project vision:** Spend less time fixing document layouts and more
> time improving the content of your report.
💡 Why UniDraft AI?
---
Common challenge                    UniDraft AI approach
---
Manually checking college           Analyze a reference document
formatting rules
Finding report headings             Extract and compare heading
information
Understanding document structure    Organize extracted content into
sections
Repeating formatting checks         Create a reusable college format
blueprint
⚙️ How UniDraft AI Works
The following flow illustrates the intended end-to-end workflow.
``` mermaid
%%{init: {"theme":"base","themeVariables":{"background":"#F7F4FF","primaryColor":"#E8E0FF","primaryTextColor":"#30264A","primaryBorderColor":"#A895D5","lineColor":"#9A87C8","secondaryColor":"#E1F0FF","tertiaryColor":"#FCE9F1","fontFamily":"Arial"}}}%%
flowchart TD
    A(["🌸 Start"]) --> B["Open UniDraft AI"]
    B --> C["Upload college reference document"]
    C --> D["Analyze reference formatting"]
    D --> E["Generate College Format Blueprint"]
    E --> F["Upload student report"]
    F --> G["Extract report content and structure"]
    G --> H["Compare headings and formatting"]
    H --> I["Identify differences"]
    I --> J["Display analysis / results"]
    J --> K(["✨ End"])
    classDef start fill:#E8DFFF,stroke:#967BCB,color:#30264A,stroke-width:2px
    classDef process fill:#E4F1FF,stroke:#78A6D6,color:#203A56,stroke-width:1.5px
    classDef analysis fill:#FCE8F0,stroke:#D58DAA,color:#563044,stroke-width:1.5px
    classDef finish fill:#E2F4E9,stroke:#76B58C,color:#244A32,stroke-width:2px
    class A,K start
    class B,C,F process
    class D,E,G,H,I,J analysis
```
🔍 Detailed workflow
Reference upload: The user uploads a college-prescribed
reference document.
Reference analysis: The application analyzes available
formatting details, such as font family, font size, line spacing and
heading hierarchy.
Blueprint generation: The extracted preferences are organized
into a College Format Blueprint.
Student report upload: The user uploads the report that needs to
be checked.
Content extraction: The application extracts the report title,
headings, sections and other supported document information.
Comparison: The extracted report structure is compared with the
reference information.
Results: The application displays the analysis so the user can
identify possible formatting and structural differences.
🧩 System Architecture
``` mermaid
%%{init: {"theme":"base","themeVariables":{"background":"#F5F8FF","primaryColor":"#E5EDFF","primaryTextColor":"#253451","primaryBorderColor":"#829BCB","lineColor":"#8B9FC5","secondaryColor":"#E8F6F2","tertiaryColor":"#FFF0E8","fontFamily":"Arial"}}}%%
flowchart TB
    U["👩‍🎓 User"] --> UI["🌐 Web Interface<br/>HTML Templates"]
    UI --> APP["🐍 Flask Application<br/>app.py"]
    APP --> FA["📐 Format Analyzer<br/>format_analyzer.py"]
    APP --> CE["📄 Content Extractor<br/>content_extractor.py"]
    FA --> BP["College Format Blueprint"]
    CE --> DATA["Extracted Report Data"]
    BP --> CMP["🔎 Comparison / Analysis"]
    DATA --> CMP
    CMP --> OUT["📊 Results displayed to user"]
    classDef user fill:#FCE8F0,stroke:#D58DAA,color:#563044,stroke-width:2px
    classDef ui fill:#E5EDFF,stroke:#829BCB,color:#253451
    classDef backend fill:#E8F6F2,stroke:#79B5A1,color:#24483D
    classDef result fill:#FFF0E8,stroke:#D6A17C,color:#573B27
    class U user
    class UI ui
    class APP,FA,CE,BP,DATA,CMP backend
    class OUT result
```
🛠️ Technologies Used
Technology / Library   Purpose
---
Python                 Main programming language
Flask                  Web application and request handling
PyMuPDF (`fitz`)       PDF document reading and text extraction
python-docx            Reading and processing DOCX documents
HTML                   Web page structure
CSS                    User interface styling
JavaScript             Client-side interactions
Werkzeug               Secure handling of uploaded filenames
The table describes the technologies used in the current project setup;
additional libraries may be added as development continues.
📁 Project Structure
``` text
UniDraft-AI/
│
├── app.py
├── format_analyzer.py
├── content_extractor.py
│
├── templates/
│   ├── index.html
│   ├── upload.html
│   └── ...
│
├── static/
│   ├── css/
│   ├── js/
│   └── images/
│
├── uploads/
│
├── requirements.txt
└── README.md
```
Your actual folder and template names may differ. Update this tree to
match the files in your project.
🚀 Getting Started
Follow these steps to run UniDraft AI locally.
1. Clone the repository
``` bash
git clone https://github.com/YOUR-USERNAME/UniDraft-AI.git
```
2. Open the project folder
``` bash
cd UniDraft-AI
```
3. Create a virtual environment (recommended)
``` bash
python -m venv venv
```
Activate it on Windows:
``` bash
venv\Scripts\activate
```
4. Install dependencies
If your project has a `requirements.txt` file:
``` bash
pip install -r requirements.txt
```
Otherwise, install the libraries used by the current application:
``` bash
pip install Flask PyMuPDF python-docx Werkzeug
```
5. Run the application
``` bash
python app.py
```
Open the local URL shown in your terminal (commonly
`http://127.0.0.1:5000/`).
📌 Current Project Modules
---
Module                              Responsibility
---
Reference Format Analyzer           Reads the reference document and
detects available formatting
preferences
College Format Blueprint            Organizes the detected preferences
into a readable summary
Content Extractor                   Extracts report content and
document structure
Heading Comparison                  Compares detected report headings
with the reference structure
Web Interface                       Allows users to upload documents
and view results
🌟 Key Highlights
Reference-based formatting analysis.
Support for PDF and DOCX document processing, depending on the
implemented extraction path.
Extraction of report titles, headings and section content.
College Format Blueprint generation.
Heading comparison and structured analysis.
Browser-based interface powered by Flask.
🔮 Future Enhancements
Automatic formatting correction for uploaded reports.
More flexible heading detection for different college templates.
Improved handling of tables, images, headers and footers.
Downloadable corrected reports in DOCX and PDF formats.
Detailed formatting suggestions and validation reports.
More advanced document comparison and accuracy checks.
👩‍💻 Developer
::: {align="center"}
Saad
Developer of UniDraft AI
Building practical solutions with Python and web technologies.
:::
🤝 Contributions
Suggestions, feedback and contributions are welcome. If you find an
issue or have an idea to improve UniDraft AI, feel free to open an issue
or submit a pull request.
📄 License
A license has not yet been specified for this project. Add a `LICENSE`
file before distributing the project under an open-source license.
---
::: {align="center"}
💜 UniDraft AI
Smarter formatting. Better presentation. More time for ideas.
`<sub>`{=html}Made with care by Saad ✨`</sub>`{=html}
:::