from dotenv import load_dotenv
import os
from sqlalchemy import create_engine
import urllib

load_dotenv()

driver = os.getenv("DB_DRIVER").strip()
server = os.getenv("DB_SERVER").strip()
database = os.getenv("DB_NAME").strip()
username = os.getenv("DB_USER").strip()
password = os.getenv("DB_PASS").strip()

conn_str = (
    f"DRIVER={{{driver}}};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"UID={username};"
    f"PWD={password};"
    f"TrustServerCertificate=yes;"
)

engine = create_engine(
    f"mssql+pyodbc:///?odbc_connect={urllib.parse.quote_plus(conn_str)}",
    fast_executemany=True
)

