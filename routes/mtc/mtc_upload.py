# routes/upload.py
from flask import Blueprint, request, jsonify, render_template
import pandas as pd
from models import db, MTCOrder

mtc_upload_bp = Blueprint('mtc_upload_bp', __name__)

# Giao diện tải file lên
@mtc_upload_bp.route('/mtc_upload', methods=['GET'])
def upload_page():
    return render_template('mtc/mtc_upload.html')

# API xử lý file Excel
@mtc_upload_bp.route('/api/mtc/upload-excel', methods=['POST'])
def process_excel():
    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "Không tìm thấy file gửi lên!"}), 400
        
    file = request.files['file']
    if file.filename == '':
        return jsonify({"status": "error", "message": "Bạn chưa chọn file nào!"}), 400

    try:
        # Đọc Excel bằng Pandas
        df = pd.read_excel(file)
        
        # Làm sạch dữ liệu: Xóa các khoảng trắng thừa ở tên cột (đề phòng file Excel lỗi)
        df.columns = df.columns.str.strip()
        
        # Kiểm tra xem các cột bắt buộc có tồn tại không
        required_cols = ['Loai_MTC', 'SO_Number']
        for col in required_cols:
            if col not in df.columns:
                return jsonify({"status": "error", "message": f"File Excel thiếu cột bắt buộc: {col}"}), 400

        # Xử lý cột Mác Thép: Nếu trống (NaN) hoặc không có, gán thành 'DEFAULT'
        if 'Mac_Thep' in df.columns:
            df['Mac_Thep'] = df['Mac_Thep'].fillna('DEFAULT')
        else:
            df['Mac_Thep'] = 'DEFAULT'

        # Duyệt qua từng dòng và insert vào Database
        for index, row in df.iterrows():
            new_order = MTCOrder(
                Loai_MTC = str(row['Loai_MTC']).strip(),
                SO_Number = str(row['SO_Number']).strip(),
                Mac_Thep = str(row['Mac_Thep']).strip(),
                Noi_Dung_Chinh = str(row.get('Noi_Dung_Chinh', '')).strip() if pd.notna(row.get('Noi_Dung_Chinh')) else ''
            )
            db.session.add(new_order)
        
        # Lưu thay đổi xuống Database
        db.session.commit()
        
        return jsonify({"status": "success", "message": f"Đã lưu thành công {len(df)} dòng dữ liệu vào hệ thống!"}), 200

    except Exception as e:
        db.session.rollback()
        return jsonify({"status": "error", "message": f"Lỗi trong quá trình xử lý: {str(e)}"}), 500