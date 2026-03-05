import pyodbc
from datetime import datetime, timezone
from dotenv import load_dotenv
import os

load_dotenv()

AZURE_SQL_SERVER = os.getenv("AZURE_SQL_SERVER")
AZURE_SQL_DATABASE = os.getenv("AZURE_SQL_DATABASE")
AZURE_SQL_USERNAME = os.getenv("AZURE_SQL_USERNAME")
AZURE_SQL_PASSWORD = os.getenv("AZURE_SQL_PASSWORD")

def get_sql_conn():
    return pyodbc.connect(
        "DRIVER={ODBC Driver 18 for SQL Server};"
        f"SERVER={AZURE_SQL_SERVER};"
        f"DATABASE={AZURE_SQL_DATABASE};"
        f"UID={AZURE_SQL_USERNAME};"
        f"PWD={AZURE_SQL_PASSWORD};"
        "Encrypt=yes;"
        "TrustServerCertificate=no;"   # better for Azure SQL in production
        "Connection Timeout=30;"
    )

def insert_resume_upload(resume_name: str):
    sql = """
        INSERT INTO dbo.resume_tracker (resume_name, created_date)
        VALUES (?, ?)
    """
    now_utc = datetime.now(timezone.utc)

    conn = get_sql_conn()
    try:
        cur = conn.cursor()
        cur.execute(sql, (resume_name, now_utc))
        conn.commit()
    finally:
        try:
            cur.close()
        except Exception:
            pass
        conn.close()