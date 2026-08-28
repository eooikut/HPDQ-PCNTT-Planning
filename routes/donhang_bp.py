import os
import logging
from datetime import datetime
from flask import Blueprint, render_template, request, jsonify, session, flash, redirect, url_for
from werkzeug.utils import secure_filename
import pandas as pd
from db import engine
from sqlalchemy import text
from auth.decorator import login_required, permission_required

donhang_bp = Blueprint('donhang_bp', __name__, template_folder='templates')

# ==========================================
# MA TRẬN PHÂN QUYỀN (COLUMN-LEVEL RBAC)
# ==========================================
FIELD_PERMISSIONS = {
    'dung_sai': ['edit_pkd'], 'nm_san_xuat': ['edit_pkd'], 'nhip': ['edit_pkd'], 
    'tg_can_hang': ['edit_pkd'], 'xndn': ['edit_pkd'], 'so_mapping': ['edit_pkd'], 
    'cw': ['edit_pkd'], 'tdc_rieng': ['edit_pkd'], 'skinpass': ['edit_pkd'], 
    'yeu_cau_dac_biet': ['edit_pkd'], 'mac_thep': ['edit_pkd'], 'do_day': ['edit_pkd'], 
    'kho_rong': ['edit_pkd'], 'tong_luong_pkd': ['edit_pkd'], 'muc_dich_su_dung': ['edit_pkd'], 
    'tc_yeu_cau_kh': ['edit_pkd'],
    'tdc_code': ['edit_pkd', 'edit_pdv'], 'chieu_day_muc_tieu': ['edit_pkd', 'edit_pdv'],
    'tong_lsx': ['edit_pkh'], 'thang': ['edit_pkh'], 'kh': ['edit_pkh'],
    'mvt_hspm': ['edit_pkh'], 'mvt_mdd': ['edit_pkh'], 'po_duc': ['edit_pkh'],
    'tc_hien_tai_sap': ['edit_pkh'], 'tc_sap_hspm': ['edit_pkh'],
    'tinh_trang_sx': ['edit_pkh'], 'tinh_trang_mapping': ['edit_pkh'],
    'luong_ban': ['edit_pkh'], 'luong_map': ['edit_pkh'],
    'ton_kho_chua_map': ['edit_pkh'], 'chua_nhap_kho': ['edit_pkh'],
    'po_can_204': ['edit_pkh'], 'po_skin_206': ['edit_pkh'],
    'mac_phoi': ['edit_pcnlt'],
    'kich_thuoc_phoi': ['edit_nmsx'], 'chieu_dai_phoi': ['edit_nmsx'], 'khoi_luong_phoi': ['edit_nmsx']
}

# =========================================================================
# CÁC HÀM HỖ TRỢ CHUẨN HÓA DỮ LIỆU ĐƠN HÀNG (ĐỘ DÀY & NGÀY THÁNG)
# =========================================================================
def format_doday(val):
    """
    Chuẩn hóa độ dày: Luôn giữ ít nhất 1 số thập phân '.0' (VD: 2 -> '2.0', 5.9 -> '5.9', 2.35 -> '2.35')
    """
    try:
        if val == '' or val is None: return ''
        f = float(val)
        s = f"{f:.4f}".rstrip('0')
        if s.endswith('.'): return s + '0'  # 2. -> 2.0
        parts = s.split('.')
        if len(parts) == 2 and len(parts[1]) < 1: return f"{f:.1f}"
        return s
    except (ValueError, TypeError):
        return str(val).strip()

def format_khorong(val):
    """ Chuẩn hóa khổ rộng: Loại bỏ đuôi '.0' (VD: 1500.0 -> '1500') """
    try:
        if val == '' or val is None: return ''
        return f"{float(val):g}"
    except (ValueError, TypeError):
        return str(val).strip()

def clean_date_str(val):
    """
    Chuẩn hóa ngày tháng về dạng chuẩn 07/08/2026.
    Ép mọi kiểu dữ liệu từ Excel trả về đúng con số người dùng mong đợi.
    """
    if val == '' or val is None or pd.isna(val) or str(val).strip().lower() in ['nan', 'none', '']:
        return ''
    
    # 1. Nếu Excel đọc dạng Datetime (vd: 2026-07-08 00:00:00) 
    # -> Dùng %m/%d/%Y để bốc đúng số 07/08/2026 ra theo ý người dùng.
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime('%m/%d/%Y')
        
    val_str = str(val).strip()
    if ' 00:00:00' in val_str:
        val_str = val_str.replace(' 00:00:00', '')
        
    # 2. Nếu chuỗi dạng YYYY-MM-DD (vd: 2026-07-08) -> 07/08/2026
    if '-' in val_str:
        parts = val_str.split('-')
        if len(parts) == 3 and len(parts[0]) == 4:
            return f"{int(parts[1]):02d}/{int(parts[2]):02d}/{parts[0]}"
            
    # 3. Nếu chuỗi dạng D/M/YYYY (vd: 7/8/2026 hoặc 07/08/2026)
    if '/' in val_str:
        parts = val_str.split('/')
        if len(parts) == 3:
            if len(parts[2]) == 4:
                return f"{int(parts[0]):02d}/{int(parts[1]):02d}/{parts[2]}"
            elif len(parts[0]) == 4:
                return f"{int(parts[2]):02d}/{int(parts[1]):02d}/{parts[0]}"
                
    return val_str

# ==========================================
# ROUTE 1: RENDER GIAO DIỆN CHÍNH
# ==========================================
@donhang_bp.route('/don-hang', methods=['GET'])
@login_required
def index():
    user_perms = session.get('permissions', [])
    role = session.get('role')
    
    editable_fields = list(FIELD_PERMISSIONS.keys()) if role == 'admin' else [
        field for field, req_perms in FIELD_PERMISSIONS.items() 
        if any(p in user_perms for p in req_perms)
    ]
    return render_template('donhang_list.html', editable_fields=editable_fields)

# ==========================================
# ROUTE 2: LẤY DỮ LIỆU ĐỔ VÀO LƯỚI (GET)
# ==========================================
@donhang_bp.route('/api/don-hang/list', methods=['GET'])
@login_required
def get_list():
    try:
        with engine.connect() as conn:
            query = text("""
                SELECT m.*, 
                    (SELECT STRING_AGG(ma_po, ', ') FROM TB_PO_DETAIL p WHERE p.id_donhang = m.id_donhang AND p.loai_po = 'CAN_204') as po_can_204,
                    (SELECT STRING_AGG(ma_po, ', ') FROM TB_PO_DETAIL p WHERE p.id_donhang = m.id_donhang AND p.loai_po = 'SKIN_206') as po_skin_206
                FROM TB_DON_HANG_MASTER m ORDER BY m.id_donhang DESC
            """)
            result = conn.execute(query).mappings().fetchall()
            
            data = []
            for row in result:
                row_dict = dict(row)
                if row_dict.get('created_at'): 
                    row_dict['created_at'] = row_dict['created_at'].strftime("%d/%m/%Y")
                if row_dict.get('updated_at'): 
                    row_dict['updated_at'] = row_dict['updated_at'].strftime("%d/%m/%Y %H:%M:%S")
                
                data.append(row_dict)
            return jsonify({"status": "success", "data": data})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

# =========================================================================
# CẤU HÌNH HỆ THỐNG GHI LOG IMPORT EXCEL (FILE: logs/import_excel.log)
# =========================================================================
log_dir = "logs"
if not os.path.exists(log_dir):
    os.makedirs(log_dir, exist_ok=True)

log_file_path = os.path.join(log_dir, "import_excel.log")
logging.basicConfig(
    filename=log_file_path,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8"
)
logger = logging.getLogger("ImportExcelLogger")

# =========================================================================
# ROUTE 3: IMPORT EXCEL CÓ GHI LOG CHI TIẾT ĐỂ DIAGNOSTIC / DEBUG
# =========================================================================
@donhang_bp.route('/api/don-hang/upload', methods=['POST'])
@login_required
def upload_excel():
    logger.info("=" * 60)
    logger.info(f"BẮT ĐẦU PHIÊN IMPORT EXCEL - USER ID: {session.get('user_id')} - ROLE: {session.get('role')}")

    if session.get('role') != 'admin' and 'edit_pkd' not in session.get('permissions', []):
        return jsonify({"status": "error", "message": "Bạn không có quyền Import dữ liệu Đơn hàng!"}), 403

    file = request.files.get('file')
    if not file or not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({"status": "error", "message": "Vui lòng chọn file Excel hợp lệ (.xlsx, .xls)!"}), 400

    try:
        df = pd.read_excel(file, header=9)
        df = df.fillna('')
        
        df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]
        cleaned_cols = df.columns.tolist()

        expected_columns_map = {
            "Khóa XNĐN": ['Số XNĐN', 'XNĐH', 'XNĐN', 'XN', 'Số XN'],
            "Mác thép": ['Mác thép', 'Mác', 'Grade'],
            "Độ dày": ['Độ dày (mm)', 'Độ dày', 'Chiều dày', 'Dày'],
            "Khổ rộng": ['Khổ rộng (mm)', 'Khổ rộng', 'Rộng', 'Khổ'],
            "Khối lượng": ['Khối lượng (tấn)', 'Khối lượng', 'Tổng lượng PKD', 'Tổng lượng', 'KL']
        }

        missing_criticals = []
        for label, possible_names in expected_columns_map.items():
            if not any(name in cleaned_cols for name in possible_names):
                missing_criticals.append(label)

        if missing_criticals:
            msg = f"File Excel thiếu nhóm cột bắt buộc: {', '.join(missing_criticals)}"
            return jsonify({"status": "error", "message": msg, "details": [f"Các cột từ file: {cleaned_cols}"]}), 400

        def get_val(row, possible_names, default=''):
            for name in possible_names:
                if name in row and str(row[name]).strip() != '':
                    return row[name]
            return default

        def to_float(val, default=0.0):
            try:
                if val == '' or val is None: return default
                return float(str(val).replace(',', '').strip())
            except (ValueError, TypeError):
                return default

        success_count = 0
        skipped_count = 0
        ui_logs = []

        with engine.begin() as conn:
            for index, row in df.iterrows():
                excel_row_number = index + 11

                xndn = str(get_val(row, ['Số XNĐN', 'XNĐH', 'XNĐN', 'XN', 'Số XN'])).strip()
                
                if not xndn or xndn.lower() in ['nan', 'none', '', 'tổng', 'tổng cộng']:
                    skipped_count += 1
                    if skipped_count <= 5:
                        ui_logs.append(f"Dòng {excel_row_number}: Bỏ qua do XNĐN trống hoặc dòng tổng ('{xndn}')")
                    continue

                mvt_mdd = str(get_val(row, ['MVT MDD', 'MVT', 'Mã vật tư', 'Mã VT'])).strip()
                if mvt_mdd.lower() in ['nan', 'none']: mvt_mdd = ''

                try:
                    nm_san_xuat   = str(get_val(row, ['Nhà máy sản xuất', 'NM sản xuất', 'Nhà máy', 'NMSX'])).strip()
                    mac_thep      = str(get_val(row, ['Mác thép', 'Mác', 'Grade'])).strip()
                    tc_yeu_cau_kh = str(get_val(row, ['Tiêu chuẩn yêu cầu KHÁCH HÀNG', 'Tiêu chuẩn', 'TC KH', 'Tiêu chuẩn YC'])).strip()
                    
                    do_day        = to_float(get_val(row, ['Độ dày (mm)', 'Độ dày', 'Chiều dày', 'Dày']))
                    kho_rong      = to_float(get_val(row, ['Khổ rộng (mm)', 'Khổ rộng', 'Rộng', 'Khổ']))
                    tong_luong    = to_float(get_val(row, ['Khối lượng (tấn)', 'Khối lượng', 'Tổng lượng PKD', 'Tổng lượng', 'KL']))
                    
                    dung_sai      = str(get_val(row, ['Dung sai khối lượng (%)', 'Dung sai khối lượng', 'Dung sai'])).strip()
                    cw            = str(get_val(row, ['KL cuộn (tấn)', 'KL cuộn', 'CW', 'Coil Weight'])).strip()
                    muc_dich      = str(get_val(row, ['Mục đích sử dụng', 'Mục đích', 'MDSD'])).strip()
                    
                    # === SỬ DỤNG HÀM CLEAN_DATE ĐỂ GIẢI QUYẾT LỖI NGÀY THÁNG ===
                    raw_date      = get_val(row, ['Thời gian cần hàng', 'TG Cần hàng', 'Ngày cần hàng', 'TGCH'])
                    tg_can_hang   = clean_date_str(raw_date)
                    
                    tdc_rieng     = str(get_val(row, ['TDC riêng', 'TDC Riêng'])).strip()
                    
                    skinpass_raw  = str(get_val(row, ['Skinpass', 'Skin pass'], 'No')).strip().capitalize()
                    skinpass      = 'Yes' if skinpass_raw in ['Yes', 'Y', 'Có', '1', 'True'] else 'No'
                    
                    yeu_cau_db    = str(get_val(row, ['Y/C đặc biệt', 'Yêu cầu đặc biệt', 'YC đặc biệt', 'YCĐB'])).strip()
                    cd_muc_tieu   = str(get_val(row, ['Chiều dày mục tiêu', 'CD mục tiêu'])).strip()
                    so_mapping    = str(get_val(row, ['SO', 'SO Mapping', 'Số SO'])).strip()
                    tdc_code      = str(get_val(row, ['TDC Code', 'TDC code', 'Mã TDC'])).strip()
                    nhip          = str(get_val(row, ['Nhịp', 'Nhịp SX', 'Nhịp sản xuất'])).strip()

                    # === SỬ DỤNG FORMAT_DODAY ĐỂ ĐẢM BẢO 2.0x ===
                    dd_str = format_doday(do_day)
                    kr_str = format_khorong(kho_rong)
                    mat_desc      = f"Thép cuộn cán nóng {dd_str}x{kr_str} {mac_thep}".strip()
                    mat_desc_hspm = f"Thép HRC HSPM {dd_str}x{kr_str} {mac_thep}".strip() if skinpass == 'Yes' else ""

                    query = text("""
                        INSERT INTO TB_DON_HANG_MASTER (
                            xndn, mvt_mdd, nm_san_xuat, nhip, tg_can_hang, so_mapping, cw, tdc_rieng, 
                            skinpass, yeu_cau_dac_biet, mac_thep, do_day, kho_rong, tong_luong_pkd, 
                            dung_sai, muc_dich_su_dung, tc_yeu_cau_kh, chieu_day_muc_tieu, tdc_code, 
                            material_description, material_description_hspm, created_by
                        )
                        VALUES (
                            :xndn, :mvt_mdd, :nm_sx, :nhip, :tg_can, :so_map, :cw, :tdc_rieng, 
                            :skinpass, :yeu_cau_db, :mac_thep, :do_day, :kho_rong, :tong_luong, 
                            :dung_sai, :muc_dich, :tc_kh, :cd_mt, :tdc_code, 
                            :mat_desc, :mat_desc_hspm, :uid
                        )
                    """)
                    
                    conn.execute(query, {
                        "xndn": xndn, "mvt_mdd": mvt_mdd, "nm_sx": nm_san_xuat, "nhip": nhip,
                        "tg_can": tg_can_hang, "so_map": so_mapping, "cw": cw, "tdc_rieng": tdc_rieng,
                        "skinpass": skinpass, "yeu_cau_db": yeu_cau_db, "mac_thep": mac_thep,
                        "do_day": do_day if do_day != 0.0 else 0,
                        "kho_rong": kho_rong if kho_rong != 0.0 else 0,
                        "tong_luong": tong_luong if tong_luong != 0.0 else 0,
                        "dung_sai": dung_sai, "muc_dich": muc_dich, "tc_kh": tc_yeu_cau_kh,
                        "cd_mt": cd_muc_tieu, "tdc_code": tdc_code,
                        "mat_desc": mat_desc, "mat_desc_hspm": mat_desc_hspm,
                        "uid": session.get('user_id')
                    })
                    success_count += 1

                except Exception as row_error:
                    logger.error(f"Lỗi SQL/Xử lý dòng {excel_row_number}: {str(row_error)}")
                    raise row_error

        summary_msg = f"Import thành công {success_count} dòng Đơn hàng vào hệ thống! (Đã bỏ qua {skipped_count} dòng)"
        return jsonify({
            "status": "success", 
            "message": summary_msg,
            "details": ui_logs
        }), 200

    except Exception as e:
        error_msg = str(e)
        return jsonify({
            "status": "error", 
            "message": f"Lỗi xử lý file Excel: {error_msg}",
            "details": ["Đã xảy ra lỗi nghiêm trọng. Vui lòng mở file log tại server: logs/import_excel.log để xem Traceback."]
        }), 500

# ==========================================
# ROUTE 4: CẬP NHẬT TRỰC TIẾP Ô (INLINE EDIT)
# ==========================================
@donhang_bp.route('/api/don-hang/update-master', methods=['PUT'])
@login_required
def update_master():
    data = request.json
    field, value, id_donhang = data.get('field'), data.get('value'), data.get('id_donhang')
    role, user_perms = session.get('role'), session.get('permissions', [])

    if role != 'admin':
        if field not in FIELD_PERMISSIONS or not any(p in user_perms for p in FIELD_PERMISSIONS[field]):
            return jsonify({"status": "error", "message": "BẠN KHÔNG CÓ QUYỀN SỬA CỘT NÀY!"}), 403

    try:
        with engine.begin() as conn:
            if field == 'tg_can_hang':
                value = clean_date_str(value)

            query = text(f"UPDATE TB_DON_HANG_MASTER SET {field} = :val, updated_at = GETDATE(), updated_by = :uid WHERE id_donhang = :id")
            conn.execute(query, {"val": value, "uid": session.get('user_id'), "id": id_donhang})
            
            updated_data = {}
            if field in ['do_day', 'kho_rong', 'mac_thep', 'skinpass']:
                row = conn.execute(text("SELECT do_day, kho_rong, mac_thep, skinpass FROM TB_DON_HANG_MASTER WHERE id_donhang = :id"), {"id": id_donhang}).fetchone()
                if row:
                    dd_str = format_doday(row[0])
                    kr_str = format_khorong(row[1])
                    mac = str(row[2]).strip() if row[2] else ""
                    sp = str(row[3]).strip() if row[3] else "No"

                    m_desc = f"Thép cuộn cán nóng {dd_str}x{kr_str} {mac}".strip()
                    m_hspm = f"Thép HRC HSPM {dd_str}x{kr_str} {mac}".strip() if sp == 'Yes' else ""

                    conn.execute(text("UPDATE TB_DON_HANG_MASTER SET material_description = :md, material_description_hspm = :mh WHERE id_donhang = :id"),
                        {"md": m_desc, "mh": m_hspm, "id": id_donhang})
                    
                    updated_data = {"material_description": m_desc, "material_description_hspm": m_hspm}

        return jsonify({"status": "success", "message": "Đã lưu!", "updated_data": updated_data})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

# ==========================================
# ROUTE 5: CẬP NHẬT MÃ PO (CÓ RÀNG BUỘC SKINPASS)
# ==========================================
@donhang_bp.route('/api/don-hang/update-po', methods=['PUT'])
@login_required
def update_po():
    role, user_perms = session.get('role'), session.get('permissions', [])
    if role != 'admin' and 'edit_pkh' not in user_perms:
        return jsonify({"status": "error", "message": "CHỈ PHÒNG KẾ HOẠCH ĐƯỢC SỬA MÃ PO!"}), 403

    data = request.json
    id_donhang = data.get('id_donhang')
    po_can, po_skin = data.get('po_can_204', ''), data.get('po_skin_206', '')

    try:
        with engine.begin() as conn:
            check = conn.execute(text("SELECT skinpass FROM TB_DON_HANG_MASTER WHERE id_donhang = :id"), {"id": id_donhang}).fetchone()
            is_skinpass = check and (check[0] == 'Yes')

            conn.execute(text("DELETE FROM TB_PO_DETAIL WHERE id_donhang = :id"), {"id": id_donhang})

            def insert_po(po_str, loai):
                if not po_str: return
                for p in [x.strip() for x in po_str.split(',') if x.strip()]:
                    conn.execute(text("INSERT INTO TB_PO_DETAIL (id_donhang, loai_po, ma_po) VALUES (:id, :l, :m)"),
                                 {"id": id_donhang, "l": loai, "m": p})

            insert_po(po_can, 'CAN_204')
            if is_skinpass: insert_po(po_skin, 'SKIN_206')

        return jsonify({"status": "success", "message": "Đã cập nhật PO!"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500