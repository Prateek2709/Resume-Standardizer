# AI Resume Parser & Standardizer
An AI-powered application that converts unstructured resumes (PDF/DOCX) into structured, company-standardized resumes using Azure OpenAI, Streamlit, and custom DOCX templates.

The system extracts candidate information, generates standardized resumes, and produces structured tables for quick HR screening.

## Key features
### AI Resume Parsing
- Extracts structured information from resumes using Azure OpenAI
- Identifies candidate details such as name, role, experience, skills, certifications, education, and projects

### Resume Standardization
- Converts resumes into a company-approved format
- Generates two versions:
  - Company Version – internal format
  - Non-Company Version – external/client format

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

### Structured Outputs
For each resume, the system generates:
- Parsed resume JSON
- Standardized DOCX resumes
- Candidate summary Excel file
- LLM usage logs (both in the terminal and on a LangSmith dashboard, if connected)

## Tech Stack
- Frontend: Streamlit
- AI: Azure OpenAI
- Observability: LangSmith
- Document Processing: pdfplumber, python-docx, docxtpl
- Database: Azure SQL
- Containerization: Docker

## Running the Application
Add the required credentials in the **.env** file.
```
pip install -r requirements.txt
streamlit run app_docx_output.py
```
Then open:
```
http://localhost:8501
```
The Docker setup exposes the Streamlit port and persists generated outputs.
