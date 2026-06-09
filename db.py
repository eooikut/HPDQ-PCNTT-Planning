from dotenv import load_dotenv
import os
from sqlalchemy import create_engine
import urllib

load_dotenv()

# ==========================================
# 1. CẤU HÌNH KẾT NỐI MS SQL SERVER (Cũ)
# ==========================================
driver = os.getenv("DB_DRIVER", "").strip()
server = os.getenv("DB_SERVER", "").strip()
database = os.getenv("DB_NAME", "").strip()
username = os.getenv("DB_USER", "").strip()
password = os.getenv("DB_PASS", "").strip()

conn_str = (
    f"DRIVER={{{driver}}};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"UID={username};"
    f"PWD={password};"
    f"TrustServerCertificate=yes;"
)
SQLALCHEMY_URI = f"mssql+pyodbc:///?odbc_connect={urllib.parse.quote_plus(conn_str)}"
engine = create_engine(
    SQLALCHEMY_URI,
    fast_executemany=True
)

# ==========================================
# 2. CẤU HÌNH KẾT NỐI MYSQL (Mới)
# ==========================================
mysql_server = os.getenv("MYSQL_SERVER", "localhost").strip()
mysql_database = os.getenv("MYSQL_DB", "bkmis_hpsdq").strip()
mysql_username = os.getenv("MYSQL_USER", "root").strip()
mysql_password = os.getenv("MYSQL_PASS", "").strip()
mysql_port = os.getenv("MYSQL_PORT", "3306").strip()

# Dùng pymysql làm driver kết nối chuẩn cho MySQL
MYSQL_URI = f"mysql+pymysql://{mysql_username}:{urllib.parse.quote_plus(mysql_password)}@{mysql_server}:{mysql_port}/{mysql_database}?charset=utf8mb4"

engine_mysql = create_engine(
    MYSQL_URI,
    pool_recycle=3600 # Tự động kết nối lại để tránh lỗi "MySQL server has gone away"
)