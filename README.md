# AI Resume Parser & Standardizer
This application converts unstructured resumes (PDF/DOCX) into structured, company-standardized resumes using Azure OpenAI, Streamlit, and custom DOCX templates.

The system extracts candidate information, generates standardized resumes, produces structured tables for quick HR screening, and also creates a clean replica of the original resume without formatting decorations.

## Key features

### AI Resume Parsing
- Extracts structured information from resumes using Azure OpenAI
- Identifies candidate details such as name, role, experience, skills, certifications, education, and projects

### Resume Standardization
- Converts resumes into a company-approved format
- Generates two versions:
  - Company Version – internal format
  - Non-Company Version – external/client format
  - Clean Version - retains exact resume format and content, but wothout decorators like borders, page numbers, underlines, hyperlinks, etc

### Clean Resume Generation
Generates a clean version of the original resume that preserves the content and layout while removing visual decorations such as:
- Page numbers
- Borders
- Hyperlinks
- Underlines
- Header/footer styling

### Candidate Screening Table
Automatically generates a structured table containing:
- Contact information
- Visa status
- Relocation availability
- Interview availability
- Certification availability
- Education details

### Skill Matrix Generation
Builds a dynamic skill matrix based on tools and technologies used in project experience.

### Database Logging
Each processed resume is recorded in Azure SQL to maintain an audit trail of processed documents.

### LLM Observability & Tracing
Includes local LLM tracing using **Arize Phoenix (Docker)** which captures:
- LLM inputs
- LLM outputs
- token usage
- cost tracking
- model metadata
- request traces

### Structured Outputs
For each resume, the system generates:
- Parsed resume JSON
- Standardized DOCX resumes
- Clean replica resume (decoration-free)
- Candidate summary Excel file
- LLM usage logs and traces

> Create an empty **output** folder in the same directory as the code files

## Tech Stack
- Frontend: Streamlit
- AI: Azure OpenAI
- Observability: Arize Phoenix, LangSmith (optional)
- Document Processing: pdfplumber, python-docx, docxtpl
- Database: Azure SQL
- Containerization: Docker

## Running the Application
Add the required credentials in the **.env** file.

### Local Run
```
pip install -r requirements.txt
streamlit run app_docx_output.py
```
Then open:
```
http://localhost:8501
```
### Docker run:
```
docker compose up --build
```
The app will be available at:
```
http://localhost:8501
```
For accessing the Phoenix Observability Dashboard, open:
```
127.0.0.1:6006/projects
```
