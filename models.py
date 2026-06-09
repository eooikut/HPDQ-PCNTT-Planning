# models.py
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

# Khởi tạo đối tượng db
db = SQLAlchemy()

# 1. Bảng Cấu Hình
class ConfigMacThep(db.Model):
    __tablename__ = 'Config_MacThep'
    
    ID = db.Column(db.Integer, primary_key=True, autoincrement=True)
    Loai_MTC = db.Column(db.String(50), nullable=False)
    Mac_Thep = db.Column(db.String(100), nullable=False)
    Tieu_Chuan = db.Column(db.String(255))
    License_No = db.Column(db.String(100))

# 2. Bảng Dữ liệu Đơn hàng
class MTCOrder(db.Model):
    __tablename__ = 'MTC_Orders'
    
    ID = db.Column(db.Integer, primary_key=True, autoincrement=True)
    Loai_MTC = db.Column(db.String(50), nullable=False)
    SO_Number = db.Column(db.String(100), nullable=False)
    Mac_Thep = db.Column(db.String(100), nullable=False)
    Noi_Dung_Chinh = db.Column(db.Text) # Tương đương NVARCHAR(MAX)
    Trang_Thai = db.Column(db.String(50), default='Process')
    Ngay_Tao = db.Column(db.DateTime, default=datetime.now)
    linkPng = db.Column(db.String(255), nullable=True)
    nhan_mau = db.Column(db.Unicode(255), nullable=True)
    pic = db.Column(db.Unicode(255), nullable=True)
    ten_tau = db.Column(db.Unicode(255), nullable=True)
    ghi_chu = db.Column(db.UnicodeText, nullable=True)