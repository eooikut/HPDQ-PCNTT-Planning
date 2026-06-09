# routes/config.py
from flask import Blueprint, render_template, request, jsonify
from models import db, ConfigMacThep

mtc_config_bp = Blueprint('mtc_config_bp', __name__)

# 1. Trang giao diện quản lý cấu hình
@mtc_config_bp.route('/mtc_config')
def config_page():
    # Lấy toàn bộ dữ liệu cấu hình hiện có
    configs = ConfigMacThep.query.order_by(ConfigMacThep.Loai_MTC, ConfigMacThep.Mac_Thep).all()
    return render_template('mtc/mtc_config.html', configs=configs)

# 2. API Lưu cấu hình (Dùng chung cho cả Thêm mới và Sửa)
@mtc_config_bp.route('/api/mtc/config/save', methods=['POST'])
def save_config():
    try:
        config_id = request.form.get('id')
        loai_mtc = request.form.get('loai_mtc').strip()
        mac_thep = request.form.get('mac_thep').strip()
        tieu_chuan = request.form.get('tieu_chuan').strip()
        license_no = request.form.get('license_no').strip()

        if config_id: 
            # Nếu có ID -> Cập nhật dữ liệu cũ
            config = ConfigMacThep.query.get(config_id)
            if config:
                config.Loai_MTC = loai_mtc
                config.Mac_Thep = mac_thep
                config.Tieu_Chuan = tieu_chuan
                config.License_No = license_no
        else: 
            # Nếu không có ID -> Tạo dữ liệu mới
            new_config = ConfigMacThep(
                Loai_MTC=loai_mtc,
                Mac_Thep=mac_thep,
                Tieu_Chuan=tieu_chuan,
                License_No=license_no
            )
            db.session.add(new_config)

        db.session.commit()
        return jsonify({"status": "success", "message": "Đã lưu cấu hình thành công!"})
    except Exception as e:
        db.session.rollback()
        return jsonify({"status": "error", "message": f"Lỗi: {str(e)}"}), 500

# 3. API Xóa cấu hình
@mtc_config_bp.route('/api/mtc/config/delete/<int:id>', methods=['POST'])
def delete_config(id):
    config = ConfigMacThep.query.get_or_404(id)
    db.session.delete(config)
    db.session.commit()
    return jsonify({"status": "success"})