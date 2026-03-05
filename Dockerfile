# Streamlit Dockerfile (DOCX-only firm resume output)
FROM python:3.13-slim

# Prevent Python from writing .pyc files and enable unbuffered logs
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# Minimal build deps for common Python wheels
# Minimal build deps for common Python wheels + ODBC Driver 18 prereqs
# --- ODBC Driver 18 + unixODBC (Debian slim, no apt-key) ---
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    ca-certificates \
    gnupg \
    apt-transport-https \
    unixodbc \
    unixodbc-dev \
 && mkdir -p /etc/apt/keyrings \
 && curl -fsSL https://packages.microsoft.com/keys/microsoft.asc \
    | gpg --dearmor -o /etc/apt/keyrings/microsoft.gpg \
 && chmod go+r /etc/apt/keyrings/microsoft.gpg \
 && echo "deb [arch=amd64 signed-by=/etc/apt/keyrings/microsoft.gpg] https://packages.microsoft.com/debian/12/prod bookworm main" \
    > /etc/apt/sources.list.d/mssql-release.list \
 && apt-get update \
 && ACCEPT_EULA=Y apt-get install -y --no-install-recommends msodbcsql18 \
 && rm -rf /var/lib/apt/lists/*
# ----------------------------------------------------------

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

RUN pip install --no-cache-dir python-docx docxtpl openpyxl pdfplumber pyodbc

# Copy application code (includes templates/, Company_Template.docx, logo, etc.)
COPY . .

# Streamlit port
EXPOSE 8501

# Run Streamlit app
CMD ["streamlit", "run", "app_docx_output.py", "--server.port=8501", "--server.address=0.0.0.0"]