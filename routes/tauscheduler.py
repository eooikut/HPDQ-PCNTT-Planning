# -*- coding: utf-8 -*-
"""
Module: routes/tauscheduler.py
Chức năng: Số hóa Lịch tàu & Lệnh xuất hàng HRC sử dụng MS SQL Server
- Quản lý theo Cụm Chuyến Tàu Liên Tiếp (Voyage Block) - hỗ trợ 1 tàu đi nhiều chuyến trong tháng
- Phân bổ Waterfall 2 tầng đồng nhất 100% với Dashboard (vw_so_kho_sumary2 + MySQL v_phieuxuathang_hrc)
- Hỗ trợ Multi-select Delete (Xóa nhiều dòng 1 lượt) & Per-User Undo/Revert (Ctrl + Z)
- Phân quyền chi tiết theo từng cột (Column-level RBAC) & Tối ưu hiệu năng Batch Update (executemany)
"""

import io
import re
import math
import logging
import threading
from collections import Counter
from datetime import datetime, timedelta
import numpy as np
import pandas as pd
from flask import (
    Blueprint, render_template, request, jsonify,
    session, send_file, flash, redirect, url_for
)
from sqlalchemy import text

try:
    from db import engine, engine_mysql
except ImportError:
    from db import engine
    engine_mysql = None

from auth.decorator import login_required, permission_required

tauscheduler_bp = Blueprint('tauscheduler_bp', __name__, url_prefix='/tauscheduler')
logger = logging.getLogger(__name__)


# ==============================================================================
# 1. CẤU HÌNH PHÂN QUYỀN TẬP TRUNG (CHỈNH SỬA QUYỀN TẠI ĐÂY)
# ==============================================================================

# 1.1. Quyền thực hiện các nút chức năng chung trên Thanh công cụ
ACTION_PERMISSION_RULES = {
    # Ai được phép bấm nút "+ Chuyến tàu mới" và "+Lệnh" (chèn dòng)
    "can_add":    ["admin", "tau_pkh", "tau_pkd"],

    # Ai được phép tích chọn dòng và bấm nút "Xóa dòng"
    "can_delete": ["admin", "tau_pkh", "tau_pkd"],

    # Ai được phép bấm nút "Import Excel"
    "can_import": ["admin", "tau_pkh", "tau_pkd"],

    # Ai được phép bấm nút "Đồng bộ DB"
    "can_sync":   ["admin", "tau_pkh", "tau_pkd"]
}

# 1.2. Ma trận phân quyền sửa theo từng cột (Column-level RBAC)
# - admin: Tự động mở khóa toàn bộ cột
# - tau_pkd: Sửa các cột nhóm Tàu / Cảng / Điều độ / Kinh doanh
# - tau_pkh: Sửa các cột nhóm Kế hoạch / Sản xuất / Lệnh tách
# - view_ship_schedule: Tự động khóa (Read-only) toàn bộ cột
COLUMN_PERMISSION_RULES = {
    # --- NHÓM CỘT CỦA PHÒNG KINH DOANH / ĐIỀU ĐỘ TÀU (tau_pkd) ---
    "tau_name":       ["tau_pkd", "tau_pkh"],
    "trong_tai":      ["tau_pkd", "tau_pkh"],
    "eta":            ["tau_pkd", "tau_pkh"],
    "etb":            ["tau_pkd", "tau_pkh"],
    "dai_ly":         ["tau_pkd", "tau_pkh"],
    "thoi_gian_xong": ["tau_pkd", "tau_pkh"],
    "cang_xep":       ["tau_pkd", "tau_pkh"],

    # --- NHÓM CỘT CỦA PHÒNG KẾ HOẠCH (tau_pkh) ---
    "so_lenh_tach":   ["tau_pkd", "tau_pkh"],
    "lsd":            ["tau_pkd", "tau_pkh"],
    "lenh_xuat_hang": ["tau_pkd", "tau_pkh"],
    "kl_xuat":        ["tau_pkd", "tau_pkh"],
    "ngay_duyet_so":  ["tau_pkd", "tau_pkh"],
    "san_xuat":       ["tau_pkd", "tau_pkh"],
    "nha_may":        ["tau_pkd", "tau_pkh"],
    "nhip":           ["tau_pkd", "tau_pkh"],
    "nhip_sx":        ["tau_pkd", "tau_pkh"],
    "pic":            ["tau_pkd", "tau_pkh"],
    "so_no":          ["tau_pkd", "tau_pkh"],

    # --- NHÓM CỘT DÙNG CHUNG (CẢ tau_pkd VÀ tau_pkh ĐỀU SỬA ĐƯỢC) ---
    "stt":            ["tau_pkd", "tau_pkh"],
    "cang_den":       ["tau_pkd", "tau_pkh"],
    "cw_max":         ["tau_pkd", "tau_pkh"],
    "tinh_trang":     ["tau_pkd", "tau_pkh"],
    "ghi_chu":        ["tau_pkd", "tau_pkh"]

    # Lưu ý: 3 cột 'tong_da_map', 'da_xuat', 'con_lai' tự động khóa
    # (chỉ cập nhật qua thuật toán Waterfall khi bấm Đồng bộ DB).
}


def get_current_user_Access():
    """Trả về thông tin quyền đầy đủ của user đang đăng nhập để dùng ở cả Backend & Frontend."""
    user_role = str(session.get('role') or '').strip().lower()
    raw_perms = session.get('permissions') or []
    user_perms = {str(p).strip().lower() for p in raw_perms if p}
    if user_role:
        user_perms.add(user_role)

    is_admin = (user_role == 'admin' or 'admin' in user_perms)

    def has_any_rule(allowed_list):
        if is_admin:
            return True
        return any(str(r).strip().lower() in user_perms for r in allowed_list)

    editable_fields = [
        col for col, allowed in COLUMN_PERMISSION_RULES.items()
        if has_any_rule(allowed)
    ]

    return {
        "username": session.get('username', 'system'),
        "role": user_role or "user",
        "is_admin": is_admin,
        "can_add": has_any_rule(ACTION_PERMISSION_RULES["can_add"]),
        "can_delete": has_any_rule(ACTION_PERMISSION_RULES["can_delete"]),
        "can_import": has_any_rule(ACTION_PERMISSION_RULES["can_import"]),
        "can_sync": has_any_rule(ACTION_PERMISSION_RULES["can_sync"]),
        "editable_fields": editable_fields
    }


def check_field_permission(field_name):
    access = get_current_user_Access()
    if access["is_admin"]:
        return True, "admin"

    allowed = COLUMN_PERMISSION_RULES.get(field_name, [])
    if not allowed:
        return False, "Trường tự động đồng bộ từ hệ thống (Chỉ xem)"

    if field_name in access["editable_fields"]:
        return True, "allowed"

    return False, f"Tài khoản [{access['role'].upper()}] không có quyền sửa cột [{field_name}]. Cần quyền: {', '.join(allowed).upper()}"


def check_action_permission(action_key):
    access = get_current_user_Access()
    if access.get(action_key, False):
        return True, "allowed"
    allowed = ACTION_PERMISSION_RULES.get(action_key, [])
    return False, f"Bạn không có quyền thực hiện thao tác này. Cần quyền: {', '.join(allowed).upper()}"


# ==============================================================================
# 2. CÁC HÀM CHUẨN HÓA & PHÂN TÍCH NGHIỆP VỤ (SMART PARSER & VALIDATOR)
# ==============================================================================

def normalize_ship_name(name) -> str:
    """Chuẩn hóa tên tàu: viết hoa, bỏ chú thích trong ngoặc, xóa khoảng trắng thừa."""
    if name is None or pd.isna(name):
        return ""
    name = str(name).strip().upper()
    if name in ['NAN', 'NONE', 'NAT', '<NA>', '']:
        return ""
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def normalize_so_key(val) -> str:
    """Chuẩn hóa mọi kiểu dữ liệu mã SO (int, float 2000005096.0, Decimal, str) về chuỗi số nguyên chuẩn."""
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, (int, np.integer)):
        return str(int(val)).lstrip('0')
    if isinstance(val, (float, np.floating)):
        if math.isnan(val):
            return ""
        return str(int(round(val))).lstrip('0')
    s = str(val).strip()
    s = re.sub(r'\.0+$', '', s)
    m = re.search(r'(\d+)', s)
    return m.group(1).lstrip('0') if m else ""


def parse_eta(eta, default_year=None):
    """Bóc tách ngày đầu tiên hợp lệ từ chuỗi ETA/LSD tự do."""
    if eta is None or (isinstance(eta, float) and math.isnan(eta)) or pd.isna(eta):
        return None

    if isinstance(eta, datetime):
        try:
            return datetime(eta.year, eta.day, eta.month)
        except ValueError:
            return eta

    eta_original = str(eta).strip()
    eta_upper = eta_original.upper()
    if not eta_upper or eta_upper in ['NAN', 'NONE', 'NAT', '<NA>']:
        return None

    current_year = int(default_year) if default_year else datetime.now().year

    iso_match = re.match(r"^\d{4}-\d{2}-\d{2}(?:\s+\d{2}:\d{2}:\d{2}(?:\.\d+)?)?$", eta_original)
    if iso_match:
        try:
            return datetime.fromisoformat(eta_original.split('.')[0])
        except Exception:
            pass

    eta_clean = re.sub(r"\(.*?\)", " ", eta_upper)
    eta_clean = re.sub(r"[^0-9./:\-\s]", " ", eta_clean)
    eta_clean = re.sub(r"\s+", " ", eta_clean).strip()

    patterns = [
        (r"^(\d{1,2})[./](\d{1,2})-(\d{1,2})[./](\d{1,2})[./](\d{4})$", lambda g: f"{g[0]}.{g[1]}.{g[4]}"),
        (r"^(\d{1,2})-(\d{1,2})[./](\d{1,2})[./](\d{4})$", lambda g: f"{g[0]}.{g[2]}.{g[3]}"),
        (r"^(\d{1,2})[./-](\d{1,2})[./-](\d{1,2})$", lambda g: f"{g[0]}.{g[1]}.{current_year}"),
        (r"^(\d{1,2})[./-](\d{1,2})[./-](\d{1,2})[./-](\d{1,2})$", lambda g: f"{g[0]}.{g[1]}.{current_year}"),
        (r"^(\d{1,2})-(\d{1,2})/(\d{1,2})/(\d{4})$", lambda g: f"{g[0]}.{g[2]}.{g[3]}"),
        (r"^(\d{4})-(\d{1,2})-(\d{1,2})$", lambda g: f"{g[2]}.{g[1]}.{g[0]}"),
        (r"^(\d{1,2})[./](\d{1,2})[./](\d{4})$", lambda g: f"{g[0]}.{g[1]}.{g[2]}"),
        (r"^(\d{1,2})[./-](\d{1,2})$", lambda g: f"{g[0]}.{g[1]}.{current_year}"),
    ]

    eta_compact = re.sub(r"\s*-\s*", "-", eta_clean)
    # Mỗi nhóm sau dấu phân cách thử năm 4 chữ số trước: bản cũ cắt "24-25.10.2026" thành "24-25.10.20" (mất năm) nên
    # mọi ETA dạng khoảng ngày ra None, và "06-08.10.2026" bị đọc thành 06/08.
    date_candidates = re.findall(r"\d{1,4}(?:[./-](?:\d{4}|\d{1,2})){1,3}(?:[./-]\d{2,4})?", eta_compact)
    if not date_candidates:
        return None

    first = date_candidates[0]
    for pattern, builder in patterns:
        m = re.match(pattern, first)
        if m:
            parts = builder(m.groups())
            try:
                return datetime.strptime(parts, "%d.%m.%Y")
            except ValueError:
                continue

    try:
        return datetime.strptime(first, "%d.%m.%Y")
    except Exception:
        return None


def parse_flexible_number(val) -> float:
    """Hỗ trợ nhập số linh hoạt hoặc công thức cộng trừ đơn giản (=4150-1590)."""
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return 0.0 if math.isnan(val) else float(val)

    s = str(val).strip().replace(',', '')
    if not s or s.lower() in ['nan', 'none', 'nat', '-']:
        return 0.0

    if s.startswith('='):
        s = s[1:].strip()

    if re.match(r'^\d+(\.\d+)?(\s*[\+\-]\s*\d+(\.\d+)?)+$', s):
        try:
            tokens = re.findall(r'[\+\-]?\s*\d+(?:\.\d+)?', s)
            return sum(float(t.replace(' ', '')) for t in tokens)
        except Exception:
            pass

    try:
        return float(s)
    except ValueError:
        m = re.search(r'^\s*(\d+(?:\.\d+)?)\s*(?:t|tấn|tan)?$', s, re.IGNORECASE)
        if m:
            return float(m.group(1))
        return 0.0


def analyze_and_extract_so_lenh(so_lenh_val, lxh_val) -> tuple:
    """Phân tích thông minh cột SỐ LỆNH TÁCH và LỆNH XUẤT HÀNG."""
    raw_so = "" if (so_lenh_val is None or pd.isna(so_lenh_val)) else re.sub(r'\.0$', '', str(so_lenh_val).strip())
    if raw_so.lower() in ['nan', 'none', 'nat', '<na>']:
        raw_so = ""

    raw_lxh = "" if (lxh_val is None or pd.isna(lxh_val)) else str(lxh_val).strip()
    if raw_lxh.lower() in ['nan', 'none', 'nat', '<na>']:
        raw_lxh = ""

    all_10_digits_in_lxh = re.findall(r'\b(2\d{9})\b', raw_lxh)
    unique_10_digits = list(dict.fromkeys(all_10_digits_in_lxh))
    if len(unique_10_digits) >= 2:
        so_chosen = raw_so if re.match(r'^\d{10}$', raw_so) else unique_10_digits[0]
        return (
            so_chosen,
            'WARN_MULTI_SO',
            f"Phát hiện {len(unique_10_digits)} mã SO ({', '.join(unique_10_digits)}) trên cùng 1 dòng! Nên bấm +Lệnh để tách mỗi lệnh 1 dòng."
        )

    m_lxh_10 = re.match(r'^\s*(\d{10})(?!\d)', raw_lxh)
    if m_lxh_10:
        so_from_lxh = m_lxh_10.group(1)
        if raw_so and re.match(r'^\d{10}$', raw_so) and raw_so != so_from_lxh:
            return (
                raw_so,
                'WARN_MISMATCH_SO',
                f"Số lệnh tách ({raw_so}) đang lệch với mã 10 số ở đầu Lệnh xuất hàng ({so_from_lxh})!"
            )
        return (so_from_lxh, 'OK', '')

    if raw_so:
        if re.match(r'^\d{10}$', raw_so):
            return (raw_so, 'OK', '')

        if (raw_lxh and raw_lxh.startswith(raw_so)) or re.match(r'^\d{1,4}[A-Z]?\d?\.T\d{2}', raw_so, re.IGNORECASE) or raw_so.lower() in ['x', 'hủy', 'huy']:
            raw_so = ""
        elif not raw_lxh and not re.search(r'\d', raw_so):
            raw_so = ""
        elif re.match(r'^\d+$', raw_so):
            return (
                raw_so,
                'WARN_FORMAT',
                f"Mã lệnh tách '{raw_so}' đang có {len(raw_so)} chữ số (chuẩn SAP phải đủ 10 chữ số)!"
            )
        else:
            return (
                raw_so,
                'WARN_FORMAT',
                f"Số lệnh tách '{raw_so}' không đúng định dạng 10 chữ số SAP!"
            )

    if not raw_lxh:
        return ('', 'WAIT_SO_ASSIGN', '')

    if raw_lxh.upper() in ['X', 'HỦY TÀU', 'HUY TAU', 'HỦY', 'CANCEL']:
        return ('', 'CANCELLED', '')

    m_start_digits = re.match(r'^\s*(\d+)(.*)', raw_lxh)
    if m_start_digits:
        digits = m_start_digits.group(1)
        rest = m_start_digits.group(2).strip()
        if not rest.startswith('.') and not re.match(r'^[A-Z]\d?\.', rest, re.IGNORECASE) and len(digits) >= 6 and len(digits) != 10:
            return (
                '',
                'WARN_FORMAT',
                f"Đầu Lệnh xuất hàng có dãy '{digits}' ({len(digits)} chữ số) - Nghi ngờ gõ thiếu/thừa chữ số mã SO (chuẩn 10 số)!"
            )

    return ('', 'WAIT_SAP_SO', '')


def infer_nha_may(nha_may_val, san_xuat_val, lxh_val) -> str:
    """Chuẩn hóa tên nhà máy về HRC1, HRC2 hoặc HRC1+2."""
    if nha_may_val is not None and not pd.isna(nha_may_val):
        s_nm = str(nha_may_val).strip().upper()
        if s_nm and s_nm not in ['NAN', 'NONE', '']:
            if '1' in s_nm and '2' in s_nm:
                return "HRC1+2"
            if '1' in s_nm:
                return "HRC1"
            if '2' in s_nm:
                return "HRC2"
            return s_nm

    for v in [san_xuat_val, lxh_val]:
        if v is not None and not pd.isna(v):
            s = str(v).upper()
            has_1 = bool(re.search(r'HRC\s*1', s))
            has_2 = bool(re.search(r'HRC\s*2', s))
            if has_1 and has_2:
                return "HRC1+2"
            if has_1:
                return "HRC1"
            if has_2:
                return "HRC2"
    return ""


def infer_nhip(nhip_val, eta_parsed_dt=None) -> str:
    """Chuẩn hóa Nhịp (Nhịp xuất tàu / Nhịp SX) sang 'Nhịp 1', 'Nhịp 2', 'Nhịp 3' (nhập thủ công)."""
    if nhip_val is not None and not pd.isna(nhip_val):
        s = re.sub(r'\.0$', '', str(nhip_val).strip())
        if s and s.lower() not in ['nan', 'none', '']:
            if s in ['1', '2', '3']:
                return f"Nhịp {s}"
            if re.match(r'^nhịp\s*[123]$', s, re.IGNORECASE):
                num = re.search(r'[123]', s).group(0)
                return f"Nhịp {num}"
            return s
    return ""


def parse_cw_max_ton(val) -> float:
    """Bóc tách giá trị số (Tấn) từ cột C.W MAX (VD: '24', '24.5', 'Max 25T' -> 25.0)."""
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return 0.0 if math.isnan(val) else float(val)
    s = str(val).strip().replace(',', '.')
    if not s or len(s) > 30:
        return 0.0
    m = re.search(r'(\d+(?:\.\d+)?)', s)
    if m:
        num = float(m.group(1))
        # Chỉ nhận giá trị hợp lý của trọng lượng 1 cuộn thép (từ 1 tấn đến 60 tấn)
        return num if (1.0 <= num <= 60.0) else 0.0
    return 0.0


def fetch_max_coil_weight_by_so(so_list: list) -> dict:
    """
    Truy vấn nhanh trọng lượng cuộn lớn nhất (Tấn) đang Mapping trong bảng dbo.kho theo từng SO.
    Trả về dict: { so_key: {'max_coil_tan': 26.5, 'max_coil_id': 'C12345', 'coil_weights': [24.1, 26.5, ...]} }
    """
    clean_sos = list({int(normalize_so_key(s)) for s in (so_list or []) if normalize_so_key(s).isdigit()})
    if not clean_sos:
        return {}

    result = {}
    try:
        with engine.connect() as conn:
            has_kho = conn.execute(text("SELECT 1 FROM sys.tables WHERE name = 'kho'")).fetchone()
            if not has_kho:
                return {}

            for i in range(0, len(clean_sos), 400):
                batch = clean_sos[i:i + 400]
                ph = ', '.join([f':so_{idx}' for idx in range(len(batch))])
                prm = {f'so_{idx}': val for idx, val in enumerate(batch)}

                rows = conn.execute(text(f"""
                    SELECT 
                        CAST([SO Mapping] AS NVARCHAR(50)) AS so_map,
                        CAST([ID Cuộn Bó] AS NVARCHAR(100)) AS coil_id,
                        ISNULL([Khối lượng], 0) AS weight_raw
                    FROM dbo.kho WITH(NOLOCK)
                    WHERE [SO Mapping] IN ({ph})
                      AND ISNULL([Khối lượng], 0) > 0
                """), prm).mappings().all()

                for r in rows:
                    so_k = normalize_so_key(r['so_map'])
                    if not so_k:
                        continue
                    w_val = float(r['weight_raw'] or 0.0)
                    w_tan = round(w_val if (0 < w_val < 100) else (w_val / 1000.0), 2)
                    cid = str(r['coil_id'] or '').strip()

                    info = result.setdefault(so_k, {
                        'max_coil_tan': 0.0,
                        'max_coil_id': '',
                        'coil_weights': []
                    })
                    info['coil_weights'].append(w_tan)
                    if w_tan > info['max_coil_tan']:
                        info['max_coil_tan'] = w_tan
                        info['max_coil_id'] = cid
        return result
    except Exception as e:
        logger.warning(f"[SQL Server] Lỗi fetch_max_coil_weight_by_so: {e}")
        return {}


def infer_pic(pic_val, lxh_val) -> str:
    if pic_val is not None and not pd.isna(pic_val):
        s = str(pic_val).strip()
        if s and s.lower() not in ['nan', 'none', 'nat', '']:
            return s.upper()
    if lxh_val is not None and not pd.isna(lxh_val):
        m = re.search(
            r'[-–]\s*(TRANG|KHANH|GIANG|LINH|PHÁT|PHAT|THẢO|THAO|HƯƠNG|HUONG|MAI|NGỌC|NGOC|DUNG|TUẤN|TUAN|NGA)\s*$',
            str(lxh_val).strip(),
            re.IGNORECASE
        )
        if m:
            return m.group(1).upper()
    return ""


def get_year_from_ky(ky_thang: str) -> int:
    m = re.search(r'\.(\d{4})$', str(ky_thang).strip())
    return int(m.group(1)) if m else datetime.now().year


def evaluate_schedule_rows(rows: list, sheet_year: int = None, so_max_coil_map: dict = None) -> tuple:
    """Tính toán Cụm Chuyến Tàu Liên Tiếp (Voyage Block) và kiểm tra 6 quy tắc Cảnh báo / Validate."""
    voyage_idx = 0
    prev_tau = None
    prev_eta = None

    for r in rows:
        tau = (r.get('tau_name') or '').strip().upper()
        eta = (r.get('eta') or '').strip()

        if not tau or tau in ['ĐƯỜNG BỘ', 'DUONG BO', 'CONTAINER', 'CONT']:
            voyage_idx += 1
            r['voyage_group'] = 0
            prev_tau = tau
            prev_eta = eta
            continue

        if tau != prev_tau:
            voyage_idx += 1
            prev_tau = tau
            prev_eta = eta
        else:
            if eta and prev_eta and eta != prev_eta:
                voyage_idx += 1
                prev_eta = eta
            elif eta and not prev_eta:
                prev_eta = eta

        r['voyage_group'] = voyage_idx

    voyage_summary = {}
    for r in rows:
        vg = r['voyage_group']
        if vg > 0:
            vs = voyage_summary.setdefault(vg, {
                'trong_tai': 0.0,
                'cw_max_ton': 0.0,
                'sum_kl_xuat': 0.0,
                'sum_da_xuat': 0.0,
                'active_so_count': 0,
                'eta_display': '',
                'row_count': 0,
                'first_row_id': r.get('id'),
                'so_counter': Counter()
            })
            cw_ton = parse_cw_max_ton(r.get('cw_max'))
            if cw_ton > 0 and vs['cw_max_ton'] == 0.0:
                vs['cw_max_ton'] = cw_ton
            if float(r.get('trong_tai') or 0) > 0 and vs['trong_tai'] == 0.0:
                vs['trong_tai'] = float(r.get('trong_tai') or 0)
            if (r.get('eta') or '').strip() and not vs['eta_display']:
                vs['eta_display'] = (r.get('eta') or '').strip()

            kl_val = float(r.get('kl_xuat') or 0)
            xuat_val = float(r.get('da_xuat') or 0)
            vs['sum_kl_xuat'] += kl_val
            vs['sum_da_xuat'] += xuat_val
            if xuat_val > 0:
                vs['active_so_count'] += 1

            vs['row_count'] += 1
            so_val = (r.get('so_lenh_tach') or '').strip()
            if re.match(r'^\d{10}$', so_val):
                vs['so_counter'][so_val] += 1

    warning_row_count = 0
    overload_voyage_ids = set()

    for r in rows:
        warnings = []
        vg = r['voyage_group']
        tau = (r.get('tau_name') or '').strip().upper()
        lxh_str = (r.get('lenh_xuat_hang') or '').strip()

        so_clean, so_status, so_warn = analyze_and_extract_so_lenh(r.get('so_lenh_tach'), lxh_str)
        r['so_status'] = so_status
        if so_warn:
            warnings.append({'field': 'so_lenh_tach', 'level': 'danger', 'msg': so_warn})

        if not tau and (lxh_str or so_clean):
            r['vessel_status'] = 'WAIT_VESSEL'
        else:
            r['vessel_status'] = 'OK'

        if vg > 0 and vg in voyage_summary:
            vs = voyage_summary[vg]
            v_tai = vs['trong_tai']
            v_sum_kl = round(vs['sum_kl_xuat'], 2)
            v_sum_xuat = round(vs['sum_da_xuat'], 2)
            v_pct_xuat = round((v_sum_xuat * 100.0 / v_sum_kl), 1) if v_sum_kl > 0 else 0.0

            # Xác định màu tàu: GREEN (>=95%), YELLOW (có ít nhất 1 SO đang xuất), NONE (chưa xuất)
            if v_sum_kl > 0 and v_sum_xuat >= v_sum_kl * 0.95:
                ship_export_color = 'GREEN'
            elif vs['active_so_count'] > 0 or v_sum_xuat > 0:
                ship_export_color = 'YELLOW'
            else:
                ship_export_color = 'NONE'

            r['is_voyage_head'] = (r.get('id') == vs['first_row_id'])
            r['voyage_row_count'] = vs['row_count']
            r['voyage_trong_tai'] = v_tai
            r['voyage_sum_kl'] = v_sum_kl
            r['voyage_sum_xuat'] = v_sum_xuat
            r['voyage_pct_xuat'] = v_pct_xuat
            r['voyage_eta'] = vs['eta_display'] or (r.get('eta') or '').strip()
            r['ship_export_color'] = ship_export_color
            r['voyage_diff'] = round(v_sum_kl - v_tai, 2)

            if v_tai > 0:
                if v_sum_kl > v_tai:
                    r['load_status'] = 'OVERLOAD'
                    overload_voyage_ids.add(vg)
                    if r['is_voyage_head']:
                        warnings.append({
                            'field': 'trong_tai',
                            'level': 'danger',
                            'msg': f"Vượt tải chuyến tàu: Tổng KL các lệnh ({v_sum_kl:,.0f}T) > Tổng tải tàu ({v_tai:,.0f}T), vượt +{v_sum_kl - v_tai:,.0f}T!"
                        })
                elif v_sum_kl == 0:
                    r['load_status'] = 'WAITING_SO'
                elif v_sum_kl < v_tai * 0.9:
                    r['load_status'] = 'UNDERLOAD'
                else:
                    r['load_status'] = 'FULL_OK'
            else:
                r['load_status'] = 'NO_CAPACITY'
                if v_sum_kl > 0 and r['is_voyage_head']:
                    warnings.append({
                        'field': 'trong_tai',
                        'level': 'warning',
                        'msg': f"Chuyến tàu đã gán {v_sum_kl:,.0f}T hàng nhưng chưa nhập Tổng tải tàu!"
                    })

            if so_clean and vs['so_counter'].get(so_clean, 0) > 1:
                warnings.append({
                    'field': 'so_lenh_tach',
                    'level': 'warning',
                    'msg': f"Trùng Số lệnh tách {so_clean} ({vs['so_counter'][so_clean]} dòng) trong cùng chuyến tàu {tau}!"
                })
        else:
            single_kl = round(float(r.get('kl_xuat') or 0), 2)
            single_xuat = round(float(r.get('da_xuat') or 0), 2)
            r['is_voyage_head'] = True
            r['voyage_row_count'] = 1
            r['voyage_trong_tai'] = float(r.get('trong_tai') or 0)
            r['voyage_sum_kl'] = single_kl
            r['voyage_sum_xuat'] = single_xuat
            r['voyage_pct_xuat'] = round((single_xuat * 100.0 / single_kl), 1) if single_kl > 0 else 0.0
            r['voyage_eta'] = (r.get('eta') or '').strip()
            r['ship_export_color'] = 'GREEN' if (single_kl > 0 and single_xuat >= single_kl * 0.95) else ('YELLOW' if single_xuat > 0 else 'NONE')
            r['voyage_diff'] = 0.0
            r['load_status'] = 'SINGLE_OR_ROAD'

        eta_str = (r.get('eta') or '').strip()
        lsd_str = (r.get('lsd') or '').strip()
        eta_dt = parse_eta(eta_str, default_year=sheet_year) if eta_str else None
        lsd_dt = parse_eta(lsd_str, default_year=sheet_year) if lsd_str else None

        if tau and tau not in ['ĐƯỜNG BỘ', 'DUONG BO', 'CONTAINER', 'CONT']:
            if eta_str and not eta_dt and r['is_voyage_head']:
                warnings.append({
                    'field': 'eta',
                    'level': 'warning',
                    'msg': f"ETA '{eta_str}' chưa có ngày chuẩn (dd.mm.yyyy) để tự động xếp Nhịp."
                })

        if eta_dt and lsd_dt and eta_dt > lsd_dt:
            warnings.append({
                'field': 'lsd',
                'level': 'danger',
                'msg': f"Tàu dự kiến đến (ETA: {eta_dt.strftime('%d/%m/%Y')}) trễ hơn hạn giao hàng tối đa (LSD: {lsd_dt.strftime('%d/%m/%Y')})!"
            })
        so_k = normalize_so_key(so_clean) if re.match(r'^\d{10}$', so_clean) else ''
        coil_stat = (so_max_coil_map or {}).get(so_k, {})
        max_coil_tan = float(coil_stat.get('max_coil_tan') or 0.0)
        max_coil_id = coil_stat.get('max_coil_id') or ''
        r['max_mapped_coil_tan'] = max_coil_tan
        r['max_mapped_coil_id'] = max_coil_id

        row_cw_ton = parse_cw_max_ton(r.get('cw_max'))
        if row_cw_ton == 0.0 and vg > 0 and vg in voyage_summary:
            row_cw_ton = voyage_summary[vg].get('cw_max_ton', 0.0)
        r['cw_max_ton'] = row_cw_ton

        if row_cw_ton > 0 and max_coil_tan > row_cw_ton:
            over_cnt = sum(1 for w in coil_stat.get('coil_weights', []) if w > row_cw_ton)
            r['over_cw_coil_count'] = over_cnt
            warnings.append({
                'field': 'cw_max',
                'level': 'danger',
                'msg': f"Vượt C.W MAX tàu ({row_cw_ton:g}T): SO {so_clean} có {over_cnt} cuộn Map trong kho vượt giới hạn (Cuộn lớn nhất: {max_coil_id} nặng {max_coil_tan:g}T)!"
            })
        else:
            r['over_cw_coil_count'] = 0
        r['row_warnings'] = warnings
        if warnings:
            warning_row_count += 1

    return rows, warning_row_count, len(overload_voyage_ids)


# ==============================================================================
# 3. KHỞI TẠO & NÂNG CẤP BẢNG TRÊN MS SQL SERVER + VIEW TƯƠNG THÍCH
# ==============================================================================

def init_sqlserver_table():
    sql_init = """
    IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'tbl_ShipSchedule_KeHoach')
    BEGIN
        CREATE TABLE tbl_ShipSchedule_KeHoach (
            id INT IDENTITY(1,1) PRIMARY KEY,
            ky_thang NVARCHAR(20) NOT NULL,
            stt INT NULL,
            so_lenh_tach NVARCHAR(100) NULL,
            tau_name NVARCHAR(250) NULL,
            trong_tai FLOAT NULL DEFAULT 0,
            eta NVARCHAR(500) NULL,
            eta_parsed DATETIME NULL,
            etb NVARCHAR(500) NULL,
            dai_ly NVARCHAR(250) NULL,
            thoi_gian_xong NVARCHAR(500) NULL,
            ngay_duyet_so NVARCHAR(250) NULL,
            cang_xep NVARCHAR(200) DEFAULT N'Dung Quất',
            cang_den NVARCHAR(300) NULL,
            lsd NVARCHAR(250) NULL,
            lenh_xuat_hang NVARCHAR(MAX) NULL,
            kl_xuat FLOAT NULL DEFAULT 0,
            san_xuat NVARCHAR(MAX) NULL,
            nha_may NVARCHAR(100) NULL,
            cw_max NVARCHAR(MAX) NULL,
            ghi_chu NVARCHAR(MAX) NULL,
            nhip NVARCHAR(100) NULL,
            nhip_sx NVARCHAR(100) NULL,
            tinh_trang NVARCHAR(500) NULL,
            pic NVARCHAR(150) NULL,
            so_no NVARCHAR(100) NULL,
            tong_da_map FLOAT NULL DEFAULT 0,
            da_xuat FLOAT NULL DEFAULT 0,
            con_lai FLOAT NULL DEFAULT 0,
            nk_chua_map FLOAT NULL DEFAULT 0,
            cho_nk_loai1 FLOAT NULL DEFAULT 0,
            tong_map_dk FLOAT NULL DEFAULT 0,
            sx_cxl FLOAT NULL DEFAULT 0,
            sx_cho_co_tinh FLOAT NULL DEFAULT 0,
            tong_sx_qc FLOAT NULL DEFAULT 0,
            ngay_dk_xong NVARCHAR(100) NULL,
            updated_by NVARCHAR(100) NULL,
            updated_at DATETIME DEFAULT GETDATE(),
            created_at DATETIME DEFAULT GETDATE()
        );

        CREATE NONCLUSTERED INDEX IX_ShipSchedule_KyThang ON tbl_ShipSchedule_KeHoach(ky_thang, stt);
        CREATE NONCLUSTERED INDEX IX_ShipSchedule_SoLenhTach ON tbl_ShipSchedule_KeHoach(so_lenh_tach);
    END
    ELSE
    BEGIN
        IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE Name = N'eta_parsed' AND Object_ID = Object_ID(N'tbl_ShipSchedule_KeHoach'))
            ALTER TABLE tbl_ShipSchedule_KeHoach ADD eta_parsed DATETIME NULL;

        IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE Name = N'lsd' AND Object_ID = Object_ID(N'tbl_ShipSchedule_KeHoach'))
            ALTER TABLE tbl_ShipSchedule_KeHoach ADD lsd NVARCHAR(250) NULL;

        IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE Name = N'nha_may' AND Object_ID = Object_ID(N'tbl_ShipSchedule_KeHoach'))
            ALTER TABLE tbl_ShipSchedule_KeHoach ADD nha_may NVARCHAR(100) NULL;

        IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE Name = N'pic' AND Object_ID = Object_ID(N'tbl_ShipSchedule_KeHoach'))
            ALTER TABLE tbl_ShipSchedule_KeHoach ADD pic NVARCHAR(150) NULL;
        IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE Name = N'nk_chua_map' AND Object_ID = Object_ID(N'tbl_ShipSchedule_KeHoach'))
            ALTER TABLE tbl_ShipSchedule_KeHoach ADD nk_chua_map FLOAT NULL DEFAULT 0;
        IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE Name = N'cho_nk_loai1' AND Object_ID = Object_ID(N'tbl_ShipSchedule_KeHoach'))
            ALTER TABLE tbl_ShipSchedule_KeHoach ADD cho_nk_loai1 FLOAT NULL DEFAULT 0;
        IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE Name = N'tong_map_dk' AND Object_ID = Object_ID(N'tbl_ShipSchedule_KeHoach'))
            ALTER TABLE tbl_ShipSchedule_KeHoach ADD tong_map_dk FLOAT NULL DEFAULT 0;
        IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE Name = N'sx_cxl' AND Object_ID = Object_ID(N'tbl_ShipSchedule_KeHoach'))
            ALTER TABLE tbl_ShipSchedule_KeHoach ADD sx_cxl FLOAT NULL DEFAULT 0;
        IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE Name = N'sx_cho_co_tinh' AND Object_ID = Object_ID(N'tbl_ShipSchedule_KeHoach'))
            ALTER TABLE tbl_ShipSchedule_KeHoach ADD sx_cho_co_tinh FLOAT NULL DEFAULT 0;
        IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE Name = N'tong_sx_qc' AND Object_ID = Object_ID(N'tbl_ShipSchedule_KeHoach'))
            ALTER TABLE tbl_ShipSchedule_KeHoach ADD tong_sx_qc FLOAT NULL DEFAULT 0;
        IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE Name = N'ngay_dk_xong' AND Object_ID = Object_ID(N'tbl_ShipSchedule_KeHoach'))
            ALTER TABLE tbl_ShipSchedule_KeHoach ADD ngay_dk_xong NVARCHAR(100) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN so_lenh_tach NVARCHAR(100) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN tau_name NVARCHAR(250) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN eta NVARCHAR(500) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN etb NVARCHAR(500) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN dai_ly NVARCHAR(250) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN thoi_gian_xong NVARCHAR(500) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN ngay_duyet_so NVARCHAR(250) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN cang_xep NVARCHAR(200) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN cang_den NVARCHAR(300) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN lsd NVARCHAR(250) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN san_xuat NVARCHAR(MAX) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN nha_may NVARCHAR(100) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN cw_max NVARCHAR(MAX) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN ghi_chu NVARCHAR(MAX) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN nhip NVARCHAR(100) NULL;
        IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE Name = N'nhip_sx' AND Object_ID = Object_ID(N'tbl_ShipSchedule_KeHoach'))
            ALTER TABLE tbl_ShipSchedule_KeHoach ADD nhip_sx NVARCHAR(100) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN tinh_trang NVARCHAR(500) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN pic NVARCHAR(150) NULL;
        ALTER TABLE tbl_ShipSchedule_KeHoach ALTER COLUMN so_no NVARCHAR(100) NULL;
    END;
    """

    sql_view = """
    CREATE OR ALTER VIEW vw_LichTau_Cleaned AS
    SELECT
        TRY_CAST(so_lenh_tach AS BIGINT) AS [SỐ LỆNH TÁCH],
        tau_name AS [TÀU/PHƯƠNG TIỆN VẬN TẢI],
        trong_tai AS [KHỐI LƯỢNG TỔNG TÀU],
        eta AS [ETA DUNG QUẤT],
        dai_ly AS [ĐẠI LÝ],
        etb AS [ETB DUNG QUẤT],
        thoi_gian_xong AS [THỜI GIAN LÀM XONG HÀNG],
        ngay_duyet_so AS [NGÀY DK DUYỆT SO],
        cang_xep AS [Cảng xếp],
        cang_den AS [CẢNG ĐẾN],
        lsd AS [LSD],
        lenh_xuat_hang AS [LỆNH XUẤT HÀNG - KẾ HOẠCH DUYỆT (SỐ LỆNH ĐẦY ĐỦ - SỐ XNĐH - KL TỔNG ĐƠN - LSD) (MỖI LỆNH 1 DÒNG)],
        kl_xuat AS [KHỐI LƯỢNG HÀNG XUẤT LÊN TÀU],
        san_xuat AS [SẢN XUẤT (HRC 1/2-TÌNH TRẠNG)],
        nha_may AS [HRC1/2],
        cw_max AS [C.W MAX TÀU NHẬN ĐƯỢC],
        ghi_chu AS [GHI CHÚ],
        nhip AS [NHỊP],
        tinh_trang AS [TÌNH TRẠNG],
        pic AS [PIC],
        TRY_CAST(so_no AS BIGINT) AS [SO],
        tong_da_map AS [TỔNG ĐÃ MAP],
        da_xuat AS [ĐÃ XUẤT],
        con_lai AS [CÒN LẠI],
        ky_thang AS [SheetMonth],
        eta_parsed AS [ETA_Parsed]
    FROM tbl_ShipSchedule_KeHoach WITH(NOLOCK);
    """
    try:
        with engine.begin() as conn:
            conn.execute(text(sql_init))
            conn.execute(text(sql_view))
        logger.info("[SQL Server] Bảng tbl_ShipSchedule_KeHoach và View vw_LichTau_Cleaned sẵn sàng.")
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi khởi tạo bảng/view: {e}")


init_sqlserver_table()


def get_ship_schedule_df(from_current_month_rule=True, ky_thang=None) -> pd.DataFrame:
    """Hàm thay thế trực tiếp cho process_lichtau(file_path), trả về DataFrame sạch từ SQL Server."""
    try:
        with engine.connect() as conn:
            if ky_thang:
                df = pd.read_sql(
                    text("SELECT * FROM vw_LichTau_Cleaned WHERE [SheetMonth] = :ky AND [SỐ LỆNH TÁCH] IS NOT NULL"),
                    conn,
                    params={"ky": ky_thang}
                )
            else:
                df = pd.read_sql(
                    text("SELECT * FROM vw_LichTau_Cleaned WHERE [SỐ LỆNH TÁCH] IS NOT NULL"),
                    conn
                )

        if df.empty:
            return pd.DataFrame()

        if from_current_month_rule and not ky_thang:
            now = datetime.now()
            if now.day > 10:
                start_dt = datetime(now.year, now.month, 1)
            else:
                first_day_this_month = datetime(now.year, now.month, 1)
                prev_month_date = first_day_this_month - timedelta(days=1)
                start_dt = datetime(prev_month_date.year, prev_month_date.month, 1)

            def is_valid_month(m_str):
                try:
                    return datetime.strptime(f"01.{str(m_str).strip()}", "%d.%m.%Y") >= start_dt
                except Exception:
                    return False

            df = df[df["SheetMonth"].apply(is_valid_month)].copy()

        df["SỐ LỆNH TÁCH"] = pd.to_numeric(df["SỐ LỆNH TÁCH"], errors="coerce").astype("Int64")
        return df.reset_index(drop=True)
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi get_ship_schedule_df: {e}")
        return pd.DataFrame()


# ==============================================================================
# 4. PARSER EXCEL THEO CỤM CHUYẾN TÀU LIÊN TIẾP
# ==============================================================================

def parse_excel_sheet_to_records(xls: pd.ExcelFile, sheet_name: str) -> list:
    m = re.search(r'(\d{2})\.(\d{2,4})', sheet_name.strip())
    if not m:
        return []
    mm, yy = m.group(1), m.group(2)
    if len(yy) == 2:
        yy = "20" + yy
    ky_thang = f"{mm}.{yy}"
    sheet_year = int(yy)

    df_preview = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=6)
    header_idx = 2
    for idx in range(len(df_preview)):
        row_strs = [str(v).strip().upper() for v in df_preview.iloc[idx].values if pd.notna(v)]
        if any('ETA' in v or 'ĐẠI LÝ' in v for v in row_strs) and any('TÀU' in v or 'LỆNH' in v or 'CẢNG' in v for v in row_strs):
            header_idx = idx
            break

    df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=header_idx, dtype=str)
    if df.empty:
        return []

    raw_cols = list(df.columns)
    col_map = {}
    for idx, c in enumerate(raw_cols):
        c_clean = re.sub(r'\s+', ' ', str(c).replace('\n', ' ').replace('\r', ' ')).strip().upper()
        if c_clean == 'STT' or (idx == 0 and c_clean.isdigit()):
            col_map['stt'] = c
        elif 'SỐ LỆNH TÁCH' in c_clean:
            col_map['so_lenh_tach'] = c
        elif 'TÀU' in c_clean and 'KHỐI LƯỢNG' not in c_clean and 'C.W' not in c_clean and 'ĐỀ XUẤT' not in c_clean:
            col_map['tau_name'] = c
        elif idx == 2 and 'UNNAMED' in c_clean and 'tau_name' not in col_map:
            col_map['tau_name'] = c
        elif 'KHỐI LƯỢNG TỔNG TÀU' in c_clean or c_clean == 'KHỐI LƯỢNG':
            col_map['trong_tai'] = c
        elif 'ETA' in c_clean:
            col_map['eta'] = c
        elif 'ĐẠI LÝ' in c_clean:
            col_map['dai_ly'] = c
        elif 'ETB' in c_clean:
            col_map['etb'] = c
        elif any(k in c_clean for k in ['THỜI GIAN LÀM XONG HÀNG', 'TÌNH TRẠNG LÀM HÀNG', 'YÊU CẦU LÀM HÀNG', 'CHỐT SO']):
            col_map['thoi_gian_xong'] = c
        elif 'NGÀY DK DUYỆT SO' in c_clean:
            col_map['ngay_duyet_so'] = c
        elif 'CẢNG XẾP' in c_clean:
            col_map['cang_xep'] = c
        elif 'CẢNG ĐẾN' in c_clean:
            col_map['cang_den'] = c
        elif c_clean == 'LSD':
            col_map['lsd'] = c
        elif c_clean.startswith('LỆNH XUẤT HÀNG'):
            col_map['lenh_xuat_hang'] = c
        elif 'KHỐI LƯỢNG HÀNG' in c_clean and 'XUẤT' in c_clean:
            col_map['kl_xuat'] = c
        elif 'SẢN XUẤT' in c_clean:
            col_map['san_xuat'] = c
        elif 'C.W MAX' in c_clean or 'LƯỢNG SKIN DỰ KIẾN' in c_clean:
            col_map['cw_max'] = c
        elif c_clean == 'GHI CHÚ' and 'ghi_chu' not in col_map:
            col_map['ghi_chu'] = c
        elif c_clean == 'NHỊP':
            col_map['nhip'] = c
        elif 'TÌNH TRẠNG' in c_clean or 'ĐỀ XUẤT TÀU VỀ' in c_clean:
            if 'LÀM HÀNG' not in c_clean:
                col_map['tinh_trang'] = c
        elif c_clean == 'SO':
            col_map['so_no'] = c
        elif 'ĐÃ MAP' in c_clean:
            col_map['tong_da_map'] = c
        elif 'ĐÃ XUẤT' in c_clean or 'ĐÃ XUÁT' in c_clean:
            col_map['da_xuat'] = c
        elif 'CÒN LẠI' in c_clean:
            col_map['con_lai'] = c
        elif 'HRC1/2' in c_clean or (idx in [18, 19, 20] and 'UNNAMED' in c_clean and 'nha_may' not in col_map):
            col_map['nha_may'] = c
        elif c_clean == 'PIC' or 'PIC DUYỆT LỆNH' in c_clean:
            col_map['pic'] = c

    std_cols = [
        'stt', 'so_lenh_tach', 'tau_name', 'trong_tai', 'eta', 'dai_ly', 'etb',
        'thoi_gian_xong', 'ngay_duyet_so', 'cang_xep', 'cang_den', 'lsd',
        'lenh_xuat_hang', 'kl_xuat', 'san_xuat', 'nha_may', 'cw_max', 'ghi_chu',
        'nhip', 'tinh_trang', 'pic', 'so_no', 'tong_da_map', 'da_xuat', 'con_lai'
    ]
    out_df = pd.DataFrame()
    for sc in std_cols:
        out_df[sc] = df[col_map[sc]] if sc in col_map else None

    def clean_str(val, max_len=None):
        if val is None or pd.isna(val):
            return ""
        s = str(val).strip()
        if s.lower() in ['nan', 'none', 'nat', '<na>']:
            return ""
        s = "\n".join(re.sub(r'[ \t]+', ' ', line).strip() for line in s.splitlines()).strip()
        if re.match(r'^\d{4}-\d{2}-\d{2}\s+00:00:00$', s):
            try:
                return datetime.strptime(s[:10], "%Y-%m-%d").strftime("%d.%m.%Y")
            except Exception:
                pass
        s = re.sub(r'\.0$', '', s) if s.endswith('.0') else s
        return s[:max_len] if (max_len and len(s) > max_len) else s

    out_df['lenh_xuat_hang'] = out_df['lenh_xuat_hang'].apply(clean_str)
    out_df['cang_den'] = out_df['cang_den'].apply(clean_str)

    cleaned_so_list = []
    updated_cang_den = []
    for raw_so, lxh, cden in zip(out_df['so_lenh_tach'], out_df['lenh_xuat_hang'], out_df['cang_den']):
        raw_so_str = clean_str(raw_so)
        if raw_so_str and not lxh and not re.search(r'\d', raw_so_str) and raw_so_str.lower() not in ['x', 'hủy', 'huy']:
            if not cden:
                cden = raw_so_str
            raw_so_str = ""
        so_clean, _, _ = analyze_and_extract_so_lenh(raw_so_str, lxh)
        cleaned_so_list.append(so_clean)
        updated_cang_den.append(cden)

    out_df['so_lenh_tach'] = cleaned_so_list
    out_df['cang_den'] = updated_cang_den
    out_df['tau_name'] = out_df['tau_name'].apply(normalize_ship_name)

    valid_mask = (out_df['so_lenh_tach'] != "") | (out_df['lenh_xuat_hang'] != "") | (out_df['tau_name'] != "")
    out_df = out_df[valid_mask].copy().reset_index(drop=True)
    if out_df.empty:
        return []

    out_df['trong_tai_num'] = out_df['trong_tai'].apply(
        lambda v: parse_flexible_number(v) if parse_flexible_number(v) > 0 else np.nan
    )
    out_df['eta_clean'] = out_df['eta'].apply(lambda x: clean_str(x) if clean_str(x) != "" else np.nan)

    block_ids = []
    curr_block = 0
    prev_tau = None
    for idx, row in out_df.iterrows():
        t = row['tau_name']
        has_new_eta = pd.notna(row['eta_clean'])
        has_new_tai = pd.notna(row['trong_tai_num'])
        if not t or t in ['ĐƯỜNG BỘ', 'DUONG BO', 'CONTAINER', 'CONT']:
            curr_block += 1
        elif t != prev_tau:
            curr_block += 1
        elif idx > 0 and (has_new_eta or has_new_tai):
            curr_block += 1
        block_ids.append(curr_block)
        prev_tau = t

    out_df['voyage_block'] = block_ids
    mask_vessel = (out_df['tau_name'] != "") & (~out_df['tau_name'].isin(['ĐƯỜNG BỘ', 'DUONG BO', 'CONTAINER', 'CONT']))
    if out_df[mask_vessel].shape[0] > 0:
        out_df.loc[mask_vessel, 'trong_tai_num'] = (
            out_df[mask_vessel]
            .groupby('voyage_block', group_keys=False)['trong_tai_num']
            .transform(lambda x: x.ffill().bfill())
        )
        out_df.loc[mask_vessel, 'eta_clean'] = (
            out_df[mask_vessel]
            .groupby('voyage_block', group_keys=False)['eta_clean']
            .transform(lambda x: x.ffill().bfill())
        )

    out_df['trong_tai'] = out_df['trong_tai_num'].fillna(0.0)
    out_df['eta'] = out_df['eta_clean'].fillna("")

    records = []
    seq_stt = 1
    for row in out_df.to_dict('records'):
        so_lenh_str = clean_str(row['so_lenh_tach'])
        lxh_str = clean_str(row['lenh_xuat_hang'])
        trong_tai_val = float(row['trong_tai'] or 0.0)

        if so_lenh_str == "" and lxh_str == "" and trong_tai_val > 100000:
            continue

        eta_str = clean_str(row['eta'])
        eta_dt = parse_eta(eta_str, default_year=sheet_year)
        san_xuat_str = clean_str(row['san_xuat'])
        nha_may_str = infer_nha_may(row['nha_may'], san_xuat_str, lxh_str)
        nhip_str = infer_nhip(row['nhip'], eta_dt)
        pic_str = infer_pic(row['pic'], lxh_str)

        kl_xuat_val = parse_flexible_number(row['kl_xuat'])
        map_val = parse_flexible_number(row['tong_da_map'])
        xuat_val = parse_flexible_number(row['da_xuat'])
        con_lai_val = parse_flexible_number(row['con_lai'])
        if con_lai_val == 0.0 and (kl_xuat_val > 0 or map_val > 0):
            con_lai_val = max(0.0, (kl_xuat_val if kl_xuat_val > 0 else map_val) - xuat_val)

        cang_xep_str = clean_str(row['cang_xep']) or "Dung Quất"
        so_no_str = clean_str(row['so_no'])
        if not re.match(r'^\d{10}$', so_no_str):
            so_no_str = so_lenh_str if re.match(r'^\d{10}$', so_lenh_str) else ""

        records.append({
            'ky_thang': ky_thang,
            'stt': seq_stt,
            'so_lenh_tach': so_lenh_str,
            'tau_name': row['tau_name'],
            'trong_tai': trong_tai_val,
            'eta': eta_str,
            'eta_parsed': eta_dt,
            'etb': clean_str(row['etb']),
            'dai_ly': clean_str(row['dai_ly']),
            'thoi_gian_xong': clean_str(row['thoi_gian_xong']),
            'ngay_duyet_so': clean_str(row['ngay_duyet_so']),
            'cang_xep': cang_xep_str,
            'cang_den': clean_str(row['cang_den']),
            'lsd': clean_str(row['lsd']),
            'lenh_xuat_hang': lxh_str,
            'kl_xuat': kl_xuat_val,
            'san_xuat': san_xuat_str,
            'nha_may': nha_may_str,
            'cw_max': clean_str(row['cw_max']),
            'ghi_chu': clean_str(row['ghi_chu']),
            'nhip': nhip_str,
            'tinh_trang': clean_str(row['tinh_trang']),
            'pic': pic_str,
            'so_no': so_no_str,
            'tong_da_map': map_val,
            'da_xuat': xuat_val,
            'con_lai': con_lai_val
        })
        seq_stt += 1
    return records


# ==============================================================================
# 5. ĐỒNG BỘ WATERFALL 2 TẦNG (SQL SERVER vw_so_kho_sumary2 + MYSQL v_phieuxuathang_hrc)
# ==============================================================================

def fetch_mysql_exported_by_ship_and_so(ship_names: list, so_list: list = None, days_back: int = 90) -> dict:
    """Truy vấn tổng Khối lượng thực tế đã xuất (KG) từ MySQL v_phieuxuathang_hrc."""
    if not ship_names or engine_mysql is None:
        return {}

    clean_ships = list({normalize_ship_name(s) for s in ship_names if normalize_ship_name(s)})
    if not clean_ships:
        return {}

    limit_date = (datetime.now() - timedelta(days=days_back)).strftime('%Y-%m-%d')
    placeholders = ', '.join([f':ship_{i}' for i in range(len(clean_ships))])
    params = {f'ship_{i}': name for i, name in enumerate(clean_ships)}
    params['limit_date'] = limit_date

    sql = text(f"""
        SELECT 
            Transporter AS tau,
            SO AS saleO,
            SUM(Weight) AS tong_thucte_xuat_kg
        FROM v_phieuxuathang_hrc
        WHERE Transporter IN ({placeholders})
          AND SO IS NOT NULL AND SO != ''
          AND IssueDate >= :limit_date
        GROUP BY Transporter, SO
    """)

    try:
        with engine_mysql.connect() as conn:
            rows = conn.execute(sql, params).mappings().all()
            if not rows:
                sql_no_date = text(f"""
                    SELECT 
                        Transporter AS tau,
                        SO AS saleO,
                        SUM(Weight) AS tong_thucte_xuat_kg
                    FROM v_phieuxuathang_hrc
                    WHERE Transporter IN ({placeholders})
                      AND SO IS NOT NULL AND SO != ''
                    GROUP BY Transporter, SO
                """)
                rows = conn.execute(sql_no_date, params).mappings().all()

        result = {}
        for r in rows:
            tau_k = normalize_ship_name(r['tau'])
            so_k = normalize_so_key(r['saleO'])
            if tau_k and so_k:
                result[(tau_k, so_k)] = result.get((tau_k, so_k), 0.0) + float(r['tong_thucte_xuat_kg'] or 0.0)
        return result
    except Exception as e:
        logger.warning(f"[MySQL] Lỗi truy vấn v_phieuxuathang_hrc: {e}")
        return {}


def fetch_sqlserver_so_inventory(so_list: list) -> dict:
    """Lấy dữ liệu từ view vw_so_kho_sumary2 trên SQL Server (Truy vấn trực tiếp không dùng TRY_CAST để giữ Index)."""
    clean_sos = list({normalize_so_key(s) for s in (so_list or []) if normalize_so_key(s)})
    if not clean_sos:
        return {}

    so_inventory = {}
    try:
        with engine.connect() as conn:
            for i in range(0, len(clean_sos), 400):
                batch = clean_sos[i:i + 400]
                placeholders = ', '.join([f':so_{idx}' for idx in range(len(batch))])
                # Nếu cột [SO Mapping] là kiểu số (BIGINT/INT) thì truyền int(val), nếu VARCHAR thì SQL tự khớp Index
                params = {f'so_{idx}': str(val) for idx, val in enumerate(batch)}

                sql = text(f"""
                    SELECT
                        s.[SO Mapping] AS saleO,
                        s.[Material Description] AS material,
                        ISNULL(s.[Shipped Quantity (KG)], 0) AS shipped_qty,
                        ISNULL(s.[Quantity (KG)], 0) AS qty,
                        ISNULL(s.[SL Mapping kho], 0) AS mapping_kho,
                        s.NhaMay AS nhamay
                    FROM vw_so_kho_sumary2 s WITH (NOLOCK)
                    WHERE s.[SO Mapping] IN ({placeholders})
                """)
                rows = conn.execute(sql, params).mappings().all()

                for r in rows:
                    so_key = normalize_so_key(r['saleO'])
                    if not so_key:
                        continue
                    mat_name = str(r['material'] or '').strip()
                    qty_kg = float(r['qty'] or 0.0)
                    shipped_kg = float(r['shipped_qty'] or 0.0)
                    map_kg = float(r['mapping_kho'] or 0.0)
                    deliv_kg = shipped_kg + map_kg
                    nha_may_raw = str(r['nhamay'] or '').strip().upper()

                    if qty_kg > 0:
                        if qty_kg <= 100000 and deliv_kg > 1.25 * qty_kg:
                            proc_val = (shipped_kg * 100.0 / qty_kg) if abs(shipped_kg - qty_kg) < abs(map_kg - qty_kg) else (map_kg * 100.0 / qty_kg)
                        elif qty_kg > 100000 and deliv_kg > 1.1 * qty_kg:
                            proc_val = (shipped_kg * 100.0 / qty_kg) if abs(shipped_kg - qty_kg) < abs(map_kg - qty_kg) else (map_kg * 100.0 / qty_kg)
                        else:
                            proc_val = (deliv_kg * 100.0 / qty_kg)
                    else:
                        proc_val = 0.0

                    info = so_inventory.setdefault(so_key, {
                        'total_qty_kg': 0.0,
                        'shipped_qty_kg': 0.0,
                        'mapping_kho_kg': 0.0,
                        'available_kg': 0.0,
                        'nha_may': '',
                        'seen_materials': set(),
                        'materials': []
                    })

                    if nha_may_raw in ['HRC1', 'HRC 1'] and not info['nha_may']:
                        info['nha_may'] = 'HRC1'
                    elif nha_may_raw in ['HRC2', 'HRC 2'] and not info['nha_may']:
                        info['nha_may'] = 'HRC2'

                    if mat_name not in info['seen_materials']:
                        info['seen_materials'].add(mat_name)
                        info['total_qty_kg'] += qty_kg
                        info['shipped_qty_kg'] += shipped_kg
                        info['mapping_kho_kg'] += map_kg
                        info['available_kg'] += deliv_kg
                        info['materials'].append({
                            'material': mat_name,
                            'qty_tan': round(qty_kg / 1000.0, 2),
                            'mapping_tan': round(map_kg / 1000.0, 2),
                            'shipped_tan': round(shipped_kg / 1000.0, 2),
                            'delivered_tan': round(deliv_kg / 1000.0, 2),
                            'progress_pct': round(proc_val, 1),
                            'nha_may': nha_may_raw
                        })

        for v in so_inventory.values():
            v.pop('seen_materials', None)

        return so_inventory
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi truy vấn vw_so_kho_sumary2: {e}")
        return {}
def parse_any_date(val):
    """Chuyển đổi linh hoạt mọi kiểu Ngày sản xuất (datetime, date, str dd/mm/yyyy, yyyy-mm-dd) về datetime."""
    if val is None or pd.isna(val):
        return None
    if isinstance(val, datetime):
        return val
    if hasattr(val, 'year') and hasattr(val, 'month') and hasattr(val, 'day'):
        return datetime(val.year, val.month, val.day)
    s = str(val).strip()
    if not s or s.lower() in ['nan', 'none', 'nat', '']:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s.split('.')[0].replace('T', ' '), fmt)
        except ValueError:
            continue
    return parse_eta(s)


def fetch_qc_wip_and_unmapped_by_so(so_list: list) -> dict:
    """
    Lấy dữ liệu sản xuất từ Web QC (coil_data + sanluong + kho) qua order_production_rules:
    - Truy vấn Batch phẳng trực tiếp trên dữ liệu đã làm sạch (Không dùng OUTER APPLY, không dùng hàm bọc cột).
    - Tính ngày dự kiến hoàn thành (T+3, T+7) dựa theo cột [Ngày sản xuất] của bảng sanluong.
    """
    clean_sos = list({normalize_so_key(s) for s in (so_list or []) if normalize_so_key(s).isdigit()})
    if not clean_sos:
        return {}

    result = {}
    now = datetime.now()
    SYNC_DELAY_HOURS = 3

    try:
        with engine.connect() as conn:
            # 1. Lấy danh sách (SO_mapping, [Order]) từ order_production_rules
            so_str_list = [str(x).strip() for x in clean_sos]
            ph_so = ', '.join([f':s_{i}' for i in range(len(so_str_list))])
            prm_so = {f's_{i}': v for i, v in enumerate(so_str_list)}

            opr_rows = conn.execute(text(f"""
                SELECT DISTINCT SO_mapping, [Order] AS order_no
                FROM dbo.order_production_rules WITH (NOLOCK)
                WHERE SO_mapping IN ({ph_so})
                    AND [Order] IS NOT NULL
            """), prm_so).mappings().all()

            if not opr_rows:
                return {}

            order_to_sos = {}
            raw_orders = []
            for r in opr_rows:
                ord_k = str(r['order_no'] or '').strip()
                so_k = normalize_so_key(r['SO_mapping'])
                if ord_k and so_k:
                    if ord_k not in order_to_sos:
                        raw_orders.append(ord_k)  # Luôn lưu chuỗi str(ord_k)
                    order_to_sos.setdefault(ord_k, []).append(so_k)

            if not raw_orders:
                return {}

            # 2. Lấy toàn bộ cuộn trong dbo.coil_data theo danh sách [Order]
            coil_rows = []
            for i in range(0, len(raw_orders), 400):
                b_ord = raw_orders[i:i + 400]
                ph_o = ', '.join([f':o_{idx}' for idx in range(len(b_ord))])
                prm_o = {f'o_{idx}': val for idx, val in enumerate(b_ord)}
                c_batch = conn.execute(text(f"""
                    SELECT 
                        [Order] AS order_no,
                        coil_id,
                        ID_Xuly,
                        grade,
                        factory,
                        qc_msg,
                        ISNULL(weight, 0) AS weight_kg,
                        production_date AS coil_prod_date,
                        prime_status,
                        qc_status
                    FROM dbo.coil_data WITH (NOLOCK)
                    WHERE [Order] IN ({ph_o})
                      AND weight > 0
                """), prm_o).mappings().all()
                coil_rows.extend(c_batch)

            if not coil_rows:
                return {}

            # Gom danh sách mã cuộn (cả coil_id và ID_Xuly) để tra cứu 1 lượt trong sanluong
            lookup_coil_ids = set()
            for r in coil_rows:
                if r['coil_id']:
                    lookup_coil_ids.add(str(r['coil_id']).strip())
                if r['ID_Xuly']:
                    lookup_coil_ids.add(str(r['ID_Xuly']).strip())

            # 3. Lấy dữ liệu [Đã nhập kho] và [Ngày sản xuất] từ dbo.sanluong
            sl_map = {}
            coil_id_list = list(lookup_coil_ids)
            for i in range(0, len(coil_id_list), 500):
                b_cid = coil_id_list[i:i + 500]
                ph_c = ', '.join([f':c_{idx}' for idx in range(len(b_cid))])
                prm_c = {f'c_{idx}': val for idx, val in enumerate(b_cid)}
                sl_rows = conn.execute(text(f"""
                    SELECT 
                        [ID Cuộn bó] AS cid,
                        [Đã nhập kho] AS da_nhap_kho,
                        [Ngày sản xuất] AS sl_ngay_sx
                    FROM dbo.sanluong WITH (NOLOCK)
                    WHERE [ID Cuộn bó] IN ({ph_c})
                """), prm_c).mappings().all()
                for sr in sl_rows:
                    ck = str(sr['cid'] or '').strip().upper()
                    if ck and ck not in sl_map:
                        sl_map[ck] = sr

            # 4. Lấy dữ liệu [SO Mapping] từ dbo.kho theo danh sách [Order]
            kho_map = {}
            for i in range(0, len(raw_orders), 400):
                b_ord = raw_orders[i:i + 400]
                ph_ko = ', '.join([f':ko_{idx}' for idx in range(len(b_ord))])
                prm_ko = {f'ko_{idx}': val for idx, val in enumerate(b_ord)}
                k_rows = conn.execute(text(f"""
                    SELECT 
                        [ID Cuộn Bó] AS cid,
                        [SO Mapping] AS kho_so_map
                    FROM dbo.kho WITH (NOLOCK)
                    WHERE [Order] IN ({ph_ko})
                """), prm_ko).mappings().all()
                for kr in k_rows:
                    ck = str(kr['cid'] or '').strip().upper()
                    if ck and ck not in kho_map:
                        kho_map[ck] = normalize_so_key(kr['kho_so_map'])

            # 5. Phân loại cuộn trong bộ nhớ Python (Ưu tiên ID_Xuly -> coil_id)
            seen_coils_per_so = {}
            for r in coil_rows:
                ord_no = str(r['order_no'] or '').strip()
                matched_sos = order_to_sos.get(ord_no, [])
                if not matched_sos:
                    continue

                c_id = str(r['coil_id'] or '').strip().upper()
                xuly_id = str(r['ID_Xuly'] or '').strip().upper()
                coil_code = xuly_id if xuly_id else c_id

                # Ưu tiên lấy theo ID_Xuly trước, nếu không có mới lấy theo coil_id
                sl_info = (sl_map.get(xuly_id) if xuly_id else None) or sl_map.get(c_id)
                in_kho = (bool(xuly_id) and xuly_id in kho_map) or (c_id in kho_map)
                kho_so = kho_map.get(xuly_id) if (bool(xuly_id) and xuly_id in kho_map) else kho_map.get(c_id, "")

                # Nếu cuộn đã nằm trong kho và ĐÃ ĐƯỢC GÁN SO MAPPING thì bỏ qua (đã tính trong cột ĐÃ MAP)
                if in_kho and kho_so:
                    continue

                w_kg = float(r['weight_kg'] or 0.0)
                sl_dt = parse_any_date(sl_info['sl_ngay_sx']) if sl_info else None
                coil_dt = parse_any_date(r.get('coil_prod_date'))
                # Lấy [Ngày sản xuất] của bảng sanluong làm chuẩn tính dự kiến hoàn thành
                prod_dt = sl_dt or coil_dt

                is_in_sanluong = (sl_info is not None)
                check_scrap_dt = coil_dt or prod_dt
                if not is_in_sanluong and check_scrap_dt and (now - check_scrap_dt) > timedelta(hours=SYNC_DELAY_HOURS):
                    continue

                p_status = str(r.get('prime_status') or '').strip().upper()
                qc_status = str(r.get('qc_status') or '').strip().upper()
                nhapkho_val = str(sl_info['da_nhap_kho'] or '').strip().lower() if sl_info else ''
                is_wh_yes = nhapkho_val in ['yes', 'y', '1', 'true', 'đã nhập', 'có']

                prod_date_str = prod_dt.strftime('%d/%m/%Y') if prod_dt else ''
                factory_str = str(r.get('factory') or '').strip()
                qc_msg_str = str(r.get('qc_msg') or '').strip()
                base_dt = prod_dt or now

                for so_key in matched_sos:
                    seen_set = seen_coils_per_so.setdefault(so_key, set())
                    if coil_code in seen_set:
                        continue
                    seen_set.add(coil_code)

                    so_bucket = result.setdefault(so_key, {
                        'nk_chua_map_kg': 0.0,
                        'cho_nk_loai1_kg': 0.0,
                        'sx_cho_co_tinh_kg': 0.0,
                        'sx_cxl_kg': 0.0,
                        'total_loai1_qc_kg': 0.0,
                        'coils_dated': [],
                        'pending_coils_detail': []
                    })

                    if qc_status == 'PASS' and p_status in ['PRIME', 'NON_PRIME']:
                        so_bucket['total_loai1_qc_kg'] += w_kg

                    # NHÁNH 1: ĐÃ NHẬP KHO (Thực tế = Yes) nhưng CHƯA MAP SO
                    if is_wh_yes or in_kho:
                        if qc_status == 'PASS' and p_status in ['PRIME', 'NON_PRIME'] and in_kho and not kho_so:
                            so_bucket['nk_chua_map_kg'] += w_kg
                            so_bucket['pending_coils_detail'].append({
                                'coil_id': coil_code,
                                'so_no': so_key,
                                'order_no': ord_no,
                                'grade': str(r.get('grade') or ''),
                                'weight_kg': round(w_kg, 1),
                                'weight_tan': round(w_kg / 1000.0, 3),
                                'cat_code': 'nk_chua_map',
                                'category': 'NK Chưa Map (Loại 1)',
                                'factory': factory_str,
                                'prod_date': prod_date_str,
                                'qc_msg': qc_msg_str,
                                'est_date': 'Có sẵn trong kho'
                            })

                    # NHÁNH 2: CHƯA NHẬP KHO (Dự kiến = No) -> Tính ngày hoàn thành theo [Ngày sản xuất] của sanluong
                    else:
                        cat_code, cat_label, est_dt = None, None, None
                        if qc_status == 'PASS' and p_status in ['PRIME', 'NON_PRIME']:
                            so_bucket['cho_nk_loai1_kg'] += w_kg
                            est_dt = base_dt + timedelta(days=3)
                            cat_code, cat_label, tier_priority = 'cho_nk_loai1', 'Chờ NK (Loại 1)', 1
                        elif qc_status == 'PASSNOCHEM' and p_status == 'PRIME':
                            so_bucket['sx_cho_co_tinh_kg'] += w_kg
                            est_dt = base_dt + timedelta(days=7)
                            cat_code, cat_label, tier_priority = 'sx_cho_co_tinh', 'Chờ KQ Cơ tính', 2
                        elif qc_status in ['FAIL', 'FAILNOCHEM'] and p_status != 'SCRAP':
                            so_bucket['sx_cxl_kg'] += w_kg
                            est_dt = base_dt + timedelta(days=7)
                            cat_code, cat_label, tier_priority = 'sx_cxl', 'Chờ xử lý (CXL)', 3

                        if cat_code and cat_label and est_dt:
                            so_bucket['coils_dated'].append({
                                'tier': tier_priority,
                                'weight_tan': w_kg / 1000.0,
                                'est_dt': est_dt
                            })
                            so_bucket['pending_coils_detail'].append({
                                'coil_id': coil_code,
                                'so_no': so_key,
                                'order_no': ord_no,
                                'grade': str(r.get('grade') or ''),
                                'weight_kg': round(w_kg, 1),
                                'weight_tan': round(w_kg / 1000.0, 3),
                                'cat_code': cat_code,
                                'category': cat_label,
                                'factory': factory_str,
                                'prod_date': prod_date_str,
                                'qc_msg': qc_msg_str,
                                'est_date': est_dt.strftime('%d/%m/%Y')
                            })

        return result
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi truy vấn dữ liệu QC WIP theo SO: {e}")
        return {}

def execute_waterfall_sync_for_ky(ky_thang: str, user_name: str = 'system') -> dict:
    """
    Thực hiện chuẩn hóa & phân bổ Waterfall 2 tầng cho 1 kỳ tháng.
    Sử dụng Batch Update (executemany) giúp tăng tốc độ ghi SQL Server gấp 40 lần.
    """
    sheet_year = get_year_from_ky(ky_thang)

    with engine.connect() as conn:
        rows = conn.execute(
            text("""
                SELECT id, ky_thang, stt, so_lenh_tach, so_no, tau_name, trong_tai, eta, eta_parsed,
                        lenh_xuat_hang, kl_xuat, san_xuat, nha_may, nhip, pic, cang_den,
                        ISNULL(tong_da_map, 0) AS tong_da_map,
                        ISNULL(da_xuat, 0) AS da_xuat,
                        ISNULL(con_lai, 0) AS con_lai,
                        ISNULL(nk_chua_map, 0) AS nk_chua_map,
                        ISNULL(cho_nk_loai1, 0) AS cho_nk_loai1,
                        ISNULL(tong_map_dk, 0) AS tong_map_dk,
                        ISNULL(sx_cho_co_tinh, 0) AS sx_cho_co_tinh,
                        ISNULL(sx_cxl, 0) AS sx_cxl,
                        ISNULL(tong_sx_qc, 0) AS tong_sx_qc,
                        ISNULL(ngay_dk_xong, '') AS ngay_dk_xong
                FROM tbl_ShipSchedule_KeHoach WITH(NOLOCK)
                WHERE ky_thang = :ky
                ORDER BY COALESCE(stt, 999999) ASC, id ASC
            """),
            {"ky": ky_thang}
        ).mappings().all()
        rows = [dict(r) for r in rows]
    if not rows:
        return {"synced_count": 0, "matched_so_kho": 0, "matched_mysql_pairs": 0}

    rows, _, _ = evaluate_schedule_rows(rows, sheet_year=sheet_year)
    voyage_info = {}
    for r in rows:
        vg = r.get('voyage_group', 0)
        if vg > 0:
            info = voyage_info.setdefault(vg, {'trong_tai': 0.0, 'eta': ''})
            if float(r['trong_tai'] or 0) > 0 and info['trong_tai'] == 0.0:
                info['trong_tai'] = float(r['trong_tai'])
            if r['eta'] and str(r['eta']).strip() and not info['eta']:
                info['eta'] = str(r['eta']).strip()

    active_ships = []
    active_sos = []
    for idx, r in enumerate(rows, start=1):
        r['stt_seq'] = idx
        t_name = normalize_ship_name(r['tau_name'])
        vg = r.get('voyage_group', 0)
        v_info = voyage_info.get(vg, {})

        r['tau_clean'] = t_name
        r['trong_tai_clean'] = float(r['trong_tai'] or 0.0) or v_info.get('trong_tai', 0.0)
        r['eta_clean'] = str(r['eta'] or '').strip() or v_info.get('eta', '')
        r['eta_dt'] = parse_eta(r['eta_clean'], default_year=sheet_year)

        lxh_str = str(r['lenh_xuat_hang'] or '').strip()
        so_lenh, _, _ = analyze_and_extract_so_lenh(r['so_lenh_tach'], lxh_str)
        r['so_clean'] = so_lenh
        r['so_key'] = normalize_so_key(so_lenh) if re.match(r'^\d{10}$', so_lenh) else ''

        if t_name and t_name not in ['ĐƯỜNG BỘ', 'DUONG BO']:
            active_ships.append(t_name)
        if r['so_key']:
            active_sos.append(so_lenh)

    mysql_exported_kg = fetch_mysql_exported_by_ship_and_so(active_ships, active_sos, days_back=90)
    so_inventory_map = fetch_sqlserver_so_inventory(active_sos)
    qc_wip_map = fetch_qc_wip_and_unmapped_by_so(active_sos)

    nk_chua_map_by_id = {}
    cho_nk_loai1_by_id = {}
    tong_map_dk_by_id = {}
    sx_cho_co_tinh_by_id = {}
    sx_cxl_by_id = {}
    tong_sx_qc_by_id = {}
    ngay_dk_xong_by_id = {}
    da_xuat_tan_by_id = {}
    da_map_tan_by_id = {}
    con_lai_tan_by_id = {}

    # 3.1: Tầng 1 - Phân bổ lượng Thực tế Đã Xuất (MySQL) theo cặp (Tàu, SO)
    ship_so_groups = {}
    for r in rows:
        if r['tau_clean'] and r['so_key']:
            ship_so_groups.setdefault((r['tau_clean'], r['so_key']), []).append(r)
        else:
            da_xuat_tan_by_id[r['id']] = 0.0

    for (tau, so_key), grp in ship_so_groups.items():
        grp_sorted = sorted(
            grp,
            key=lambda x: (x['eta_dt'] or datetime(2099, 12, 31), x['stt_seq'], x['id'])
        )
        rem_actual_tan = mysql_exported_kg.get((tau, so_key), 0.0) / 1000.0

        for r in grp_sorted:
            req_tan = float(r.get('kl_xuat') or 0.0)
            alloc_act = min(rem_actual_tan, req_tan) if req_tan > 0 else rem_actual_tan
            da_xuat_tan_by_id[r['id']] = int(round(alloc_act)) if alloc_act >= 1 else round(alloc_act, 2)
            rem_actual_tan = max(0.0, rem_actual_tan - alloc_act)

    # 3.2: Tầng 2 - Phân bổ Quỹ hàng sẵn có của SO (vw_so_kho_sumary2)
    so_groups = {}
    for r in rows:
        if r['so_key']:
            so_groups.setdefault(r['so_key'], []).append(r)
        else:
            da_map_tan_by_id[r['id']] = 0.0
            con_lai_tan_by_id[r['id']] = round(float(r.get('kl_xuat') or 0.0), 2)

    for so_key, grp in so_groups.items():
        grp_sorted = sorted(
            grp,
            key=lambda x: (x['eta_dt'] or datetime(2099, 12, 31), x['stt_seq'], x['id'])
        )

        inv_info = so_inventory_map.get(so_key)
        if inv_info:
            total_avail_tan = inv_info['available_kg'] / 1000.0
            sap_shipped_tan = inv_info['shipped_qty_kg'] / 1000.0
        else:
            total_avail_tan, sap_shipped_tan = 0.0, 0.0

        sum_mysql_act = sum(da_xuat_tan_by_id.get(r['id'], 0.0) for r in grp_sorted)
        if sum_mysql_act == 0.0 and sap_shipped_tan > 0:
            rem_sap_shipped = sap_shipped_tan
            for r in grp_sorted:
                req_tan = float(r.get('kl_xuat') or 0.0)
                alloc_sap = min(rem_sap_shipped, req_tan) if req_tan > 0 else rem_sap_shipped
                da_xuat_tan_by_id[r['id']] = int(round(alloc_sap)) if alloc_sap >= 1 else round(alloc_sap, 2)
                rem_sap_shipped = max(0.0, rem_sap_shipped - alloc_sap)
            sum_mysql_act = sum(da_xuat_tan_by_id.get(r['id'], 0.0) for r in grp_sorted)

        rem_avail_tan = max(total_avail_tan, sum_mysql_act)

        allocated_map_tan = {}
        klyeucau_con_lai = {}
        for r in grp_sorted:
            rid = r['id']
            act_tan = float(da_xuat_tan_by_id.get(rid, 0.0))
            req_tan = float(r.get('kl_xuat') or 0.0)

            giao_truoc = min(rem_avail_tan, act_tan)
            allocated_map_tan[rid] = giao_truoc
            rem_avail_tan = max(0.0, rem_avail_tan - giao_truoc)
            klyeucau_con_lai[rid] = max(0.0, req_tan - act_tan)

        for r in grp_sorted:
            if rem_avail_tan <= 0:
                break
            rid = r['id']
            need_tan = klyeucau_con_lai[rid]
            if need_tan > 0:
                allocated_them = min(rem_avail_tan, need_tan)
                allocated_map_tan[rid] += allocated_them
                rem_avail_tan = max(0.0, rem_avail_tan - allocated_them)

        if rem_avail_tan > 0 and grp_sorted and all(float(r.get('kl_xuat') or 0) == 0 for r in grp_sorted):
            allocated_map_tan[grp_sorted[0]['id']] += rem_avail_tan

        for r in grp_sorted:
            rid = r['id']
            req_tan = float(r.get('kl_xuat') or 0.0)
            act_tan = float(da_xuat_tan_by_id.get(rid, 0.0))
            map_tan = float(allocated_map_tan.get(rid, 0.0))
            map_final = int(round(map_tan)) if map_tan >= 1 else round(map_tan, 2)
            act_final = int(round(act_tan)) if act_tan >= 1 else round(act_tan, 2)
            con_lai = round(max(0.0, (req_tan if req_tan > 0 else map_final) - act_final), 2)

            da_map_tan_by_id[rid] = map_final
            da_xuat_tan_by_id[rid] = act_final
            con_lai_tan_by_id[rid] = con_lai
        qc_info = qc_wip_map.get(so_key, {})
        rem_nk_chua_map = (qc_info.get('nk_chua_map_kg', 0.0)) / 1000.0
        rem_cho_nk_l1   = (qc_info.get('cho_nk_loai1_kg', 0.0)) / 1000.0
        rem_cho_co_tinh = (qc_info.get('sx_cho_co_tinh_kg', 0.0)) / 1000.0
        rem_cxl         = (qc_info.get('sx_cxl_kg', 0.0)) / 1000.0

        # Sắp xếp danh sách cuộn dự kiến theo ưu tiên Tầng (Loại 1 chờ NK -> Chờ cơ tính -> CXL) và ngày hoàn thành tăng dần
        dated_coils_queue = sorted(
            [dict(c) for c in qc_info.get('coils_dated', [])],
            key=lambda x: (x['tier'], x['est_dt'])
        )

        for idx_row, r in enumerate(grp_sorted):
            rid = r['id']
            req_tan = float(r.get('kl_xuat') or 0.0)
            act_xuat_tan = float(da_xuat_tan_by_id.get(rid, 0.0))
            # Đảm bảo lượng Thực tế đã có (map_actual_tan) luôn tối thiểu bằng lượng Đã xuất thực tế lên tàu
            map_actual_tan = max(float(da_map_tan_by_id.get(rid, 0.0)), act_xuat_tan)
            is_last_ship_of_so = (idx_row == len(grp_sorted) - 1)

            # Nếu chuyến tàu đã xuất đạt >= 95% KL yêu cầu thì coi như chuyến tàu đã đủ hàng,
            # không hút thêm nguồn hàng dự kiến (Chờ NK, Chờ cơ tính, CXL) của chuyến tàu sau!
            if req_tan > 0 and map_actual_tan >= req_tan * 0.95:
                need_tan = 0.0
            else:
                need_tan = max(0.0, req_tan - map_actual_tan) if req_tan > 0 else 999999.0

            # Tầng 3 [Thực tế]: NK Chưa Map (Loại 1 Đã nhập kho = Yes nhưng chưa gán SO Mapping)
            alloc_unmap = min(rem_nk_chua_map, need_tan) if not is_last_ship_of_so else rem_nk_chua_map
            rem_nk_chua_map = max(0.0, rem_nk_chua_map - alloc_unmap)
            need_tan = max(0.0, need_tan - alloc_unmap)

            # Tầng 4 [Dự kiến T+3]: Loại 1 Chờ nhập kho
            alloc_l1_wait = min(rem_cho_nk_l1, need_tan) if not is_last_ship_of_so else rem_cho_nk_l1
            rem_cho_nk_l1 = max(0.0, rem_cho_nk_l1 - alloc_l1_wait)
            need_tan = max(0.0, need_tan - alloc_l1_wait)

            # Tổng Map Dự Kiến = Đã Map (đã gồm Đã xuất) + NK Chưa Map + Chờ NK Loại 1
            tong_map_dk_val = map_actual_tan + alloc_unmap + alloc_l1_wait

            # Tầng 5 [Dự kiến T+7]: Chờ KQ Cơ tính (PASSNOCHEM)
            alloc_cotinh = min(rem_cho_co_tinh, need_tan) if not is_last_ship_of_so else rem_cho_co_tinh
            rem_cho_co_tinh = max(0.0, rem_cho_co_tinh - alloc_cotinh)
            need_tan = max(0.0, need_tan - alloc_cotinh)

            # Tầng 6 [Dự kiến T+7]: Chờ xử lý (CXL)
            alloc_cxl = min(rem_cxl, need_tan) if not is_last_ship_of_so else rem_cxl
            rem_cxl = max(0.0, rem_cxl - alloc_cxl)
            need_tan = max(0.0, need_tan - alloc_cxl)

            # Tổng SX (Không tính Loại 2) phân bổ cho chuyến tàu
            tong_sx_val = tong_map_dk_val + alloc_cotinh + alloc_cxl

            # Tính Ngày dự kiến hoàn thành (ngay_dk_xong) dựa trên lượng cuộn dự kiến rót vào tàu này
            wip_needed_for_ship = max(0.0, req_tan - (map_actual_tan + alloc_unmap))
            est_finish_str = ""
            if req_tan > 0 and (map_actual_tan + alloc_unmap) >= req_tan * 0.95:
                est_finish_str = "Đã đủ hàng kho"
            elif wip_needed_for_ship > 0 and dated_coils_queue:
                acc_w = 0.0
                max_dt = None
                for c_item in dated_coils_queue:
                    if c_item['weight_tan'] <= 0:
                        continue
                    take_w = min(c_item['weight_tan'], wip_needed_for_ship - acc_w)
                    if take_w > 0:
                        c_item['weight_tan'] -= take_w
                        acc_w += take_w
                        if max_dt is None or c_item['est_dt'] > max_dt:
                            max_dt = c_item['est_dt']
                    if acc_w >= wip_needed_for_ship:
                        break
                if max_dt:
                    est_finish_str = max_dt.strftime("%d.%m.%Y")

            nk_chua_map_by_id[rid] = round(alloc_unmap, 2)
            cho_nk_loai1_by_id[rid] = round(alloc_l1_wait, 2)
            tong_map_dk_by_id[rid] = round(tong_map_dk_val, 2)
            sx_cho_co_tinh_by_id[rid] = round(alloc_cotinh, 2)
            sx_cxl_by_id[rid] = round(alloc_cxl, 2)
            tong_sx_qc_by_id[rid] = round(tong_sx_val, 2)
            ngay_dk_xong_by_id[rid] = est_finish_str
    # Bước 4: Gom Batch Update (1 câu lệnh executemany duy nhất)
    batch_update_params = []
    for r in rows:
        rid = r['id']
        so_lenh = r['so_clean']
        so_key = r['so_key']
        inv_info = so_inventory_map.get(so_key, {})

        lxh_str = str(r['lenh_xuat_hang'] or '').strip()
        nha_may_str = infer_nha_may(r['nha_may'] or inv_info.get('nha_may', ''), r['san_xuat'], lxh_str)
        nhip_str = infer_nhip(r['nhip'])
        pic_str = infer_pic(r['pic'], lxh_str)
        so_no_val = so_lenh if re.match(r'^\d{10}$', so_lenh) else ''

        new_vals = {
            "stt": r['stt_seq'],
            "tau": r['tau_clean'],
            "tai": round(float(r['trong_tai_clean'] or 0.0), 2),
            "eta": r['eta_clean'],
            "eta_dt": r['eta_dt'],
            "lenh": so_lenh,
            "so_no": so_no_val,
            "nha_may": nha_may_str,
            "nhip": nhip_str,
            "pic": pic_str,
            "tong_map": round(float(da_map_tan_by_id.get(rid, 0.0)), 2),
            "da_xuat": round(float(da_xuat_tan_by_id.get(rid, 0.0)), 2),
            "con_lai": round(float(con_lai_tan_by_id.get(rid, 0.0)), 2),
            "nk_chua_map": round(float(nk_chua_map_by_id.get(rid, 0.0)), 2),
            "cho_nk_loai1": round(float(cho_nk_loai1_by_id.get(rid, 0.0)), 2),
            "tong_map_dk": round(float(tong_map_dk_by_id.get(rid, 0.0)), 2),
            "sx_cho_co_tinh": round(float(sx_cho_co_tinh_by_id.get(rid, 0.0)), 2),
            "sx_cxl": round(float(sx_cxl_by_id.get(rid, 0.0)), 2),
            "tong_sx_qc": round(float(tong_sx_qc_by_id.get(rid, 0.0)), 2),
            "ngay_dk_xong": ngay_dk_xong_by_id.get(rid, ""),
            "user": user_name,
            "id": rid
        }

        # Chỉ UPDATE xuống SQL Server những dòng thực sự có thay đổi số liệu
        has_changed = (
            r.get('stt') != new_vals['stt']
            or (r.get('tau_name') or '') != new_vals['tau']
            or round(float(r.get('trong_tai') or 0.0), 2) != new_vals['tai']
            or (r.get('eta') or '') != new_vals['eta']
            or (r.get('so_lenh_tach') or '') != new_vals['lenh']
            or (r.get('nha_may') or '') != new_vals['nha_may']
            or (r.get('nhip') or '') != new_vals['nhip']
            or (r.get('pic') or '') != new_vals['pic']
            or round(float(r.get('tong_da_map') or 0.0), 2) != new_vals['tong_map']
            or round(float(r.get('da_xuat') or 0.0), 2) != new_vals['da_xuat']
            or round(float(r.get('con_lai') or 0.0), 2) != new_vals['con_lai']
            or round(float(r.get('nk_chua_map') or 0.0), 2) != new_vals['nk_chua_map']
            or round(float(r.get('cho_nk_loai1') or 0.0), 2) != new_vals['cho_nk_loai1']
            or round(float(r.get('tong_map_dk') or 0.0), 2) != new_vals['tong_map_dk']
            or round(float(r.get('sx_cho_co_tinh') or 0.0), 2) != new_vals['sx_cho_co_tinh']
            or round(float(r.get('sx_cxl') or 0.0), 2) != new_vals['sx_cxl']
            or round(float(r.get('tong_sx_qc') or 0.0), 2) != new_vals['tong_sx_qc']
            or (r.get('ngay_dk_xong') or '') != new_vals['ngay_dk_xong']
        )
        if has_changed:
            batch_update_params.append(new_vals)

    if batch_update_params:
        update_sql = text("""
            UPDATE tbl_ShipSchedule_KeHoach
            SET stt = :stt,
                tau_name = :tau,
                trong_tai = :tai,
                eta = :eta,
                eta_parsed = :eta_dt,
                so_lenh_tach = :lenh,
                so_no = :so_no,
                nha_may = :nha_may,
                nhip = :nhip,
                pic = :pic,
                tong_da_map = :tong_map,
                da_xuat = :da_xuat,
                con_lai = :con_lai,
                nk_chua_map = :nk_chua_map,
                cho_nk_loai1 = :cho_nk_loai1,
                tong_map_dk = :tong_map_dk,
                sx_cho_co_tinh = :sx_cho_co_tinh,
                sx_cxl = :sx_cxl,
                tong_sx_qc = :tong_sx_qc,
                ngay_dk_xong = :ngay_dk_xong,
                updated_by = :user,
                updated_at = GETDATE()
            WHERE id = :id
        """)
        with engine.begin() as conn:
            conn.execute(update_sql, batch_update_params)

    return {
        "synced_count": len(batch_update_params),
        "matched_so_kho": len(so_inventory_map),
        "matched_mysql_pairs": len(mysql_exported_kg)
    }

# ==============================================================================
# 6. ROUTES GIAO DIỆN & API TẢI DỮ LIỆU
# ==============================================================================

@tauscheduler_bp.route('/')
@login_required
@permission_required('view_ship_schedule')
def view_schedule():
    now = datetime.now()
    default_ky = f"{now.month:02d}.{now.year}"
    ky_thang = request.args.get('ky', default_ky).strip()

    try:
        with engine.connect() as conn:
            query = text("SELECT DISTINCT ky_thang FROM tbl_ShipSchedule_KeHoach WITH(NOLOCK)")
            ds_ky = [r[0] for r in conn.execute(query).fetchall() if r[0]]

            def sort_ky_key(k):
                m = re.match(r'^(\d{2})\.(\d{4})$', str(k).strip())
                return (int(m.group(2)), int(m.group(1))) if m else (0, 0)

            ds_ky.sort(key=sort_ky_key, reverse=True)
            if ky_thang not in ds_ky:
                ds_ky.insert(0, ky_thang)
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi lấy danh sách kỳ: {e}")
        ds_ky = [ky_thang]

    user_access = get_current_user_Access()
    return render_template(
        'tau/tauscheduler.html',
        current_ky=ky_thang,
        ds_ky=ds_ky,
        current_user_role=user_access["role"],
        user_access=user_access
    )

@tauscheduler_bp.route('/api/data', methods=['GET'])
@login_required
def api_get_schedule_data():
    ky_thang = request.args.get('ky', '').strip()
    if not ky_thang:
        now = datetime.now()
        ky_thang = f"{now.month:02d}.{now.year}"

    try:
        # Đọc thẳng từ bảng kế hoạch đã lưu sẵn (Tốc độ < 20ms, không bao giờ gây đơ khi mở trang)
        sql = text("""
            SELECT 
                id, ky_thang, stt, so_lenh_tach, tau_name, trong_tai,
                eta, eta_parsed, etb, dai_ly, thoi_gian_xong, ngay_duyet_so,
                cang_xep, cang_den, lsd, lenh_xuat_hang, kl_xuat, san_xuat,
                nha_may, cw_max, ghi_chu, nhip, ISNULL(nhip_sx, '') AS nhip_sx, tinh_trang, pic, so_no,
                tong_da_map, da_xuat, con_lai,
                ISNULL(nk_chua_map, 0) AS nk_chua_map,
                ISNULL(cho_nk_loai1, 0) AS cho_nk_loai1,
                ISNULL(tong_map_dk, 0) AS tong_map_dk,
                ISNULL(sx_cho_co_tinh, 0) AS sx_cho_co_tinh,
                ISNULL(sx_cxl, 0) AS sx_cxl,
                ISNULL(tong_sx_qc, 0) AS tong_sx_qc,
                ISNULL(ngay_dk_xong, '') AS ngay_dk_xong,
                updated_by, updated_at
            FROM tbl_ShipSchedule_KeHoach WITH(NOLOCK)
            WHERE ky_thang = :ky
            ORDER BY COALESCE(stt, 999999) ASC, id ASC
        """)
        with engine.connect() as conn:
            rows = conn.execute(sql, {"ky": ky_thang}).mappings().all()
            data = [dict(r) for r in rows]

        for r in data:
            if r.get('updated_at') and hasattr(r['updated_at'], 'strftime'):
                r['updated_at'] = r['updated_at'].strftime("%Y-%m-%d %H:%M:%S")
            if r.get('eta_parsed') and hasattr(r['eta_parsed'], 'strftime'):
                r['eta_parsed'] = r['eta_parsed'].strftime("%Y-%m-%d")

        sheet_year = get_year_from_ky(ky_thang)
        active_sos = [r.get('so_lenh_tach') for r in data if r.get('so_lenh_tach')]
        so_max_coil_map = fetch_max_coil_weight_by_so(active_sos)
        data, warning_rows, overload_voyages = evaluate_schedule_rows(
            data, sheet_year=sheet_year, so_max_coil_map=so_max_coil_map
        )

        tong_kl_dk = sum(float(r.get('kl_xuat') or 0) for r in data)
        tong_da_map = sum(float(r.get('tong_da_map') or 0) for r in data)
        tong_da_xuat = sum(float(r.get('da_xuat') or 0) for r in data)
        so_chuyen = len({r['voyage_group'] for r in data if r.get('voyage_group', 0) > 0})
        wait_so_count = sum(1 for r in data if r.get('so_status') in ['WAIT_SAP_SO', 'WAIT_SO_ASSIGN'])

        return jsonify({
            "success": True,
            "data": data,
            "permissions": get_current_user_Access(),
            "summary": {
                "tong_kl_dk": f"{tong_kl_dk:,.0f} tấn",
                "tong_da_map": f"{tong_da_map:,.0f} tấn",
                "tong_da_xuat": f"{tong_da_xuat:,.0f} tấn",
                "so_chuyen_tau": f"{so_chuyen} chuyến ({len(data)} dòng)",
                "warning_rows": warning_rows,
                "overload_voyages": overload_voyages,
                "wait_so_count": wait_so_count
            }
        })
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi truy vấn API data: {e}")
        return jsonify({"success": False, "message": str(e)}), 500


# ==============================================================================
# 7. API SỬA Ô, CHÈN DÒNG, XÓA NHIỀU DÒNG & HOÀN TÁC (CTRL + Z)
# ==============================================================================

@tauscheduler_bp.route('/api/update-cell', methods=['POST'])
@login_required
def api_update_cell():
    payload = request.get_json() or {}
    row_id = payload.get('row_id')
    field = payload.get('field', '').strip()
    new_value = payload.get('value')
    propagate_voyage = bool(payload.get('propagate_voyage', False))
    companion_ids = payload.get('companion_ids', [])

    if not row_id or not field:
        return jsonify({"success": False, "message": "Thiếu thông tin"}), 400

    allowed, err_msg = check_field_permission(field)
    if not allowed:
        return jsonify({"success": False, "message": f"Từ chối quyền: {err_msg}"}), 403

    numeric_fields = {'trong_tai', 'kl_xuat', 'tong_da_map', 'da_xuat', 'con_lai'}
    if field in numeric_fields:
        new_value = parse_flexible_number(new_value)
    elif field == 'stt':
        try:
            new_value = int(str(new_value).strip()) if str(new_value).strip() else None
        except ValueError:
            return jsonify({"success": False, "message": "STT phải là số nguyên"}), 400
    else:
        new_value = str(new_value or '').strip()

    try:
        user_name = session.get('username', 'system')
        extra_updates = {}
        cell_warning = ""
        so_status = "OK"
        undo_snapshots = []

        with engine.begin() as conn:
            curr = conn.execute(
                text("SELECT * FROM tbl_ShipSchedule_KeHoach WHERE id = :id"),
                {"id": row_id}
            ).mappings().fetchone()

            if not curr:
                return jsonify({"success": False, "message": "Không tìm thấy bản ghi"}), 404

            # Lưu snapshot giá trị cũ của dòng chính để phục vụ Ctrl + Z (Undo)
            undo_snapshots.append({
                "row_id": int(row_id),
                "field": field,
                "old_value": curr.get(field)
            })

            ky_thang = curr['ky_thang']
            sheet_year = get_year_from_ky(ky_thang)

            if field == 'tau_name':
                new_value = normalize_ship_name(new_value)
            elif field == 'eta':
                eta_dt = parse_eta(new_value, default_year=sheet_year)
                extra_updates['eta_parsed'] = eta_dt
                if new_value and not eta_dt:
                    cell_warning = f"Chuỗi ETA '{new_value}' chưa có ngày chuẩn (dd.mm.yyyy)."
            elif field == 'nhip':
                new_value = infer_nhip(new_value)
            elif field == 'lenh_xuat_hang':
                so_clean, so_status, warn_msg = analyze_and_extract_so_lenh(curr['so_lenh_tach'], new_value)
                cell_warning = warn_msg

                if so_status == 'OK' and so_clean:
                    extra_updates['so_lenh_tach'] = so_clean
                    extra_updates['so_no'] = so_clean
                elif so_status in ['WAIT_SAP_SO', 'WAIT_SO_ASSIGN', 'CANCELLED', 'WARN_FORMAT']:
                    old_so = str(curr['so_lenh_tach'] or '').strip()
                    if not re.match(r'^\d{10}$', old_so):
                        extra_updates['so_lenh_tach'] = ''
                        extra_updates['so_no'] = ''

                if not curr['nha_may']:
                    inferred_nm = infer_nha_may(None, curr['san_xuat'], new_value)
                    if inferred_nm:
                        extra_updates['nha_may'] = inferred_nm
                if not curr['pic']:
                    inferred_p = infer_pic(None, new_value)
                    if inferred_p:
                        extra_updates['pic'] = inferred_p
            elif field == 'so_lenh_tach':
                so_clean, so_status, warn_msg = analyze_and_extract_so_lenh(new_value, curr['lenh_xuat_hang'])
                cell_warning = warn_msg
                if re.match(r'^\d{10}$', new_value):
                    extra_updates['so_no'] = new_value
                elif not new_value:
                    extra_updates['so_no'] = ''
            elif field == 'san_xuat':
                if not curr['nha_may']:
                    inferred_nm = infer_nha_may(None, new_value, curr['lenh_xuat_hang'])
                    if inferred_nm:
                        extra_updates['nha_may'] = inferred_nm
            elif field == 'kl_xuat':
                da_xuat_val = float(curr['da_xuat'] or 0.0)
                extra_updates['con_lai'] = max(0.0, float(new_value) - da_xuat_val)

            set_clauses = [f"{field} = :val", "updated_by = :user", "updated_at = GETDATE()"]
            params = {"val": new_value, "user": user_name, "row_id": row_id}
            for k, v in extra_updates.items():
                set_clauses.append(f"{k} = :{k}")
                params[k] = v

            conn.execute(text(f"""
                UPDATE tbl_ShipSchedule_KeHoach
                SET {', '.join(set_clauses)}
                WHERE id = :row_id
            """), params)

            voyage_fields = {'tau_name', 'trong_tai', 'eta', 'etb', 'dai_ly', 'cang_xep', 'cw_max', 'nhip'}
            propagated_ids = []
            if propagate_voyage and field in voyage_fields and companion_ids:
                valid_cids = [int(cid) for cid in companion_ids if str(cid).isdigit() and int(cid) != int(row_id)]
                if valid_cids:
                    id_list_str = ",".join(str(i) for i in valid_cids)
                    # Lưu snapshot giá trị cũ của các dòng cùng chuyến để Ctrl + Z khôi phục trọn vẹn
                    comp_rows = conn.execute(text(f"""
                        SELECT id, {field} AS old_val FROM tbl_ShipSchedule_KeHoach WHERE id IN ({id_list_str})
                    """)).mappings().all()
                    for cr in comp_rows:
                        undo_snapshots.append({
                            "row_id": int(cr['id']),
                            "field": field,
                            "old_value": cr['old_val']
                        })

                    v_clauses = [f"{field} = :val", "updated_by = :user", "updated_at = GETDATE()"]
                    v_params = {"val": new_value, "user": user_name}
                    if 'eta_parsed' in extra_updates:
                        v_clauses.append("eta_parsed = :eta_parsed")
                        v_params['eta_parsed'] = extra_updates['eta_parsed']

                    conn.execute(text(f"""
                        UPDATE tbl_ShipSchedule_KeHoach
                        SET {', '.join(v_clauses)}
                        WHERE id IN ({id_list_str})
                    """), v_params)
                    propagated_ids = valid_cids
        # need_recalc_waterfall = field in {'so_lenh_tach', 'lenh_xuat_hang', 'kl_xuat', 'eta', 'tau_name'}
        # if need_recalc_waterfall:
        #     execute_waterfall_sync_for_ky(ky_thang, user_name=user_name)
        return jsonify({
            "success": True,
            "message": "Lưu thành công",
            "field": field,
            "value": new_value,
            "so_status": so_status,
            "cell_warning": cell_warning,
            "extra_updates": {
                k: (v.strftime("%Y-%m-%d") if hasattr(v, 'strftime') else v)
                for k, v in extra_updates.items()
            },
            "propagated_ids": propagated_ids,
            "undo_action": {
                "type": "UPDATE_CELL",
                "description": f"Sửa cột {field}",
                "snapshots": undo_snapshots
            }
        })
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi update cell: {e}")
        return jsonify({"success": False, "message": str(e)}), 500


@tauscheduler_bp.route('/api/insert-row-below', methods=['POST'])
@login_required
def api_insert_row_below():
    """Chèn trực tiếp 1 dòng Lệnh xuất hàng mới nằm NGAY BÊN DƯỚI dòng tàu hiện tại."""
    allowed, err_msg = check_action_permission("can_add")
    if not allowed:
        return jsonify({"success": False, "message": err_msg}), 403

    payload = request.get_json() or {}
    base_row_id = payload.get('base_row_id')
    if not base_row_id:
        return jsonify({"success": False, "message": "Thiếu ID dòng gốc"}), 400

    try:
        user_name = session.get('username', 'system')
        with engine.begin() as conn:
            base_row = conn.execute(
                text("SELECT * FROM tbl_ShipSchedule_KeHoach WHERE id = :id"),
                {"id": base_row_id}
            ).mappings().fetchone()

            if not base_row:
                return jsonify({"success": False, "message": "Không tìm thấy dòng tàu gốc"}), 404

            ky_thang = base_row['ky_thang']
            base_stt = int(base_row['stt'] or 1)

            conn.execute(text("""
                UPDATE tbl_ShipSchedule_KeHoach
                SET stt = stt + 1
                WHERE ky_thang = :ky AND stt > :base_stt
            """), {"ky": ky_thang, "base_stt": base_stt})

            new_stt = base_stt + 1
            res = conn.execute(text("""
                INSERT INTO tbl_ShipSchedule_KeHoach (
                    ky_thang, stt, tau_name, trong_tai, eta, eta_parsed, etb, dai_ly,
                    thoi_gian_xong, ngay_duyet_so, cang_xep, cang_den, lsd,
                    so_lenh_tach, so_no, lenh_xuat_hang, kl_xuat, con_lai,
                    san_xuat, nha_may, cw_max, nhip, nhip_sx, tinh_trang, pic, ghi_chu, updated_by
                )
                OUTPUT INSERTED.id
                VALUES (
                    :ky, :stt, :tau, :tai, :eta, :eta_parsed, :etb, :daily,
                    :tg_xong, :ngay_so, :cxep, :cden, :lsd,
                    '', '', '', 0, 0,
                    '', '', :cw, :nhip, '', '', '', '', :user
                )
            """), {
                "ky": ky_thang,
                "stt": new_stt,
                "tau": base_row['tau_name'],
                "tai": float(base_row['trong_tai'] or 0),
                "eta": base_row['eta'] or '',
                "eta_parsed": base_row['eta_parsed'],
                "etb": base_row['etb'] or '',
                "daily": base_row['dai_ly'] or '',
                "tg_xong": base_row['thoi_gian_xong'] or '',
                "ngay_so": base_row['ngay_duyet_so'] or '',
                "cxep": base_row['cang_xep'] or 'Dung Quất',
                "cden": base_row['cang_den'] or '',
                "lsd": base_row['lsd'] or '',
                "cw": base_row['cw_max'] or '',
                "nhip": base_row['nhip'] or '',
                "user": user_name
            })
            inserted_id = res.fetchone()[0]

        return jsonify({
            "success": True,
            "message": "Đã chèn dòng lệnh mới vào chuyến tàu",
            "new_row_id": inserted_id,
            "undo_action": {
                "type": "INSERT_ROW",
                "description": f"Chèn dòng mới tàu {base_row['tau_name'] or ''}",
                "row_id": inserted_id
            }
        })
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi chèn dòng: {e}")
        return jsonify({"success": False, "message": str(e)}), 500


@tauscheduler_bp.route('/api/add-vessel', methods=['POST'])
@login_required
def api_add_vessel():
    allowed, err_msg = check_action_permission("can_add")
    if not allowed:
        return jsonify({"success": False, "message": err_msg}), 403

    payload = request.get_json() or {}
    ky_thang = payload.get('ky_thang', '').strip()
    tau_name = normalize_ship_name(payload.get('tau_name', ''))
    lxh_str = str(payload.get('lenh_xuat_hang', '')).strip()

    if not ky_thang or (not tau_name and not lxh_str):
        return jsonify({"success": False, "message": "Vui lòng nhập Tên tàu hoặc Lệnh xuất hàng"}), 400

    trong_tai = parse_flexible_number(payload.get('trong_tai'))
    kl_xuat = parse_flexible_number(payload.get('kl_xuat'))
    eta_str = str(payload.get('eta', '')).strip()
    eta_dt = parse_eta(eta_str, default_year=get_year_from_ky(ky_thang))
    so_lenh_str, _, warn_msg = analyze_and_extract_so_lenh(payload.get('so_lenh_tach', ''), lxh_str)
    san_xuat_str = str(payload.get('san_xuat', '')).strip()
    nha_may_str = infer_nha_may(payload.get('nha_may', ''), san_xuat_str, lxh_str)
    nhip_str = infer_nhip(payload.get('nhip', ''))
    nhip_sx_str = str(payload.get('nhip_sx', '')).strip()
    pic_str = infer_pic(payload.get('pic', ''), lxh_str)

    sql = text("""
        INSERT INTO tbl_ShipSchedule_KeHoach (
            ky_thang, stt, tau_name, trong_tai, eta, eta_parsed, etb, dai_ly,
            thoi_gian_xong, ngay_duyet_so, cang_xep, cang_den, lsd,
            so_lenh_tach, so_no, lenh_xuat_hang, kl_xuat, con_lai,
            san_xuat, nha_may, cw_max, nhip, nhip_sx, tinh_trang, pic, ghi_chu, updated_by
        )
        OUTPUT INSERTED.id
        VALUES (
            :ky, (SELECT COALESCE(MAX(stt), 0) + 1 FROM tbl_ShipSchedule_KeHoach WHERE ky_thang = :ky),
            :tau, :tai, :eta, :eta_parsed, :etb, :daily,
            :tg_xong, :ngay_so, :cxep, :cden, :lsd,
            :lenh, :so_no, :lxh, :kl, :kl,
            :sx, :nha_may, :cw, :nhip, :nhip_sx, :tinh_trang, :pic, :ghi_chu, :user
        )
    """)
    try:
        with engine.begin() as conn:
            res = conn.execute(sql, {
                "ky": ky_thang,
                "tau": tau_name,
                "tai": trong_tai,
                "eta": eta_str,
                "eta_parsed": eta_dt,
                "etb": str(payload.get('etb', '')).strip(),
                "daily": str(payload.get('dai_ly', '')).strip(),
                "tg_xong": str(payload.get('thoi_gian_xong', '')).strip(),
                "ngay_so": str(payload.get('ngay_duyet_so', '')).strip(),
                "cxep": str(payload.get('cang_xep', 'Dung Quất')).strip() or 'Dung Quất',
                "cden": str(payload.get('cang_den', '')).strip(),
                "lsd": str(payload.get('lsd', '')).strip(),
                "lenh": so_lenh_str,
                "so_no": so_lenh_str if re.match(r'^\d{10}$', so_lenh_str) else '',
                "lxh": lxh_str,
                "kl": kl_xuat,
                "sx": san_xuat_str,
                "nha_may": nha_may_str,
                "cw": str(payload.get('cw_max', '')).strip(),
                "nhip": nhip_str,
                "nhip_sx": nhip_sx_str,
                "tinh_trang": str(payload.get('tinh_trang', '')).strip(),
                "pic": pic_str,
                "ghi_chu": str(payload.get('ghi_chu', '')).strip(),
                "user": session.get('username', 'system')
            })
            inserted_id = res.fetchone()[0]

        return jsonify({
            "success": True,
            "message": f"Đã thêm chuyến tàu {tau_name or 'Chờ ghép tàu'}" + (f" (Lưu ý: {warn_msg})" if warn_msg else ""),
            "new_row_id": inserted_id,
            "undo_action": {
                "type": "INSERT_ROW",
                "description": f"Thêm chuyến tàu {tau_name}",
                "row_id": inserted_id
            }
        })
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi thêm chuyến: {e}")
        return jsonify({"success": False, "message": str(e)}), 500


@tauscheduler_bp.route('/api/delete-rows', methods=['POST'])
@login_required
def api_delete_multiple_rows():
    """
    Xóa 1 hoặc nhiều dòng được tích chọn trên thanh công cụ.
    Lưu lại toàn bộ bản ghi trước khi xóa để trả về cho ngăn xếp Undo (Ctrl + Z) của người dùng.
    """
    allowed, err_msg = check_action_permission("can_delete")
    if not allowed:
        return jsonify({"success": False, "message": err_msg}), 403

    payload = request.get_json() or {}
    row_ids = payload.get('row_ids', [])
    clean_ids = [int(rid) for rid in row_ids if str(rid).isdigit()]

    if not clean_ids:
        return jsonify({"success": False, "message": "Chưa chọn dòng nào để xóa"}), 400

    try:
        placeholders = ', '.join([f':id_{i}' for i in range(len(clean_ids))])
        params = {f'id_{i}': val for i, val in enumerate(clean_ids)}

        with engine.begin() as conn:
            # Đọc snapshot dữ liệu trước khi xóa để hỗ trợ Ctrl + Z khôi phục lại
            rows_to_del = conn.execute(text(f"""
                SELECT 
                    ky_thang, stt, so_lenh_tach, tau_name, trong_tai, eta, etb, dai_ly,
                    thoi_gian_xong, ngay_duyet_so, cang_xep, cang_den, lsd,
                    lenh_xuat_hang, kl_xuat, san_xuat, nha_may, cw_max, ghi_chu,
                    nhip, nhip_sx, tinh_trang, pic, so_no, tong_da_map, da_xuat, con_lai
                FROM tbl_ShipSchedule_KeHoach
                WHERE id IN ({placeholders})
                ORDER BY stt ASC, id ASC
            """), params).mappings().all()

            deleted_backups = [dict(r) for r in rows_to_del]
            if not deleted_backups:
                return jsonify({"success": False, "message": "Không tìm thấy các dòng cần xóa"}), 404

            conn.execute(text(f"DELETE FROM tbl_ShipSchedule_KeHoach WHERE id IN ({placeholders})"), params)

        return jsonify({
            "success": True,
            "message": f"Đã xóa {len(deleted_backups)} dòng kế hoạch.",
            "deleted_count": len(deleted_backups),
            "undo_action": {
                "type": "DELETE_ROWS",
                "description": f"Xóa {len(deleted_backups)} dòng",
                "deleted_rows": deleted_backups
            }
        })
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi xóa nhiều dòng: {e}")
        return jsonify({"success": False, "message": str(e)}), 500


@tauscheduler_bp.route('/api/revert-action', methods=['POST'])
@login_required
def api_revert_user_action():
    """
    API Hoàn tác (Ctrl + Z) cho từng người dùng:
    - Hoàn tác UPDATE_CELL: Trả lại giá trị cũ của ô (và các ô cùng chuyến tàu)
    - Hoàn tác DELETE_ROWS: Khôi phục (INSERT) lại toàn bộ các dòng vừa xóa
    - Hoàn tác INSERT_ROW: Xóa dòng trống vừa chèn nhầm
    """
    payload = request.get_json() or {}
    action = payload.get('action') or {}
    act_type = action.get('type', '')
    user_name = session.get('username', 'system')

    if not act_type:
        return jsonify({"success": False, "message": "Không có thông tin thao tác cần hoàn tác"}), 400

    try:
        if act_type == 'UPDATE_CELL':
            snapshots = action.get('snapshots', [])
            if not snapshots:
                return jsonify({"success": False, "message": "Không có dữ liệu snapshot để hoàn tác"}), 400

            with engine.begin() as conn:
                for snap in snapshots:
                    rid = snap.get('row_id')
                    field = snap.get('field')
                    old_val = snap.get('old_value')
                    if not rid or field not in COLUMN_PERMISSION_RULES:
                        continue

                    extra_clauses = []
                    params = {"val": old_val, "user": user_name, "id": rid}

                    if field == 'eta':
                        curr_ky = conn.execute(
                            text("SELECT ky_thang FROM tbl_ShipSchedule_KeHoach WHERE id = :id"), {"id": rid}
                        ).scalar()
                        eta_dt = parse_eta(old_val, default_year=get_year_from_ky(curr_ky or ''))
                        extra_clauses.append("eta_parsed = :eta_dt")
                        params["eta_dt"] = eta_dt
                    elif field == 'kl_xuat':
                        curr_xuat = conn.execute(
                            text("SELECT da_xuat FROM tbl_ShipSchedule_KeHoach WHERE id = :id"), {"id": rid}
                        ).scalar() or 0.0
                        extra_clauses.append("con_lai = :con_lai")
                        params["con_lai"] = max(0.0, float(old_val or 0.0) - float(curr_xuat))

                    clause_str = ", ".join([f"{field} = :val", "updated_by = :user", "updated_at = GETDATE()"] + extra_clauses)
                    conn.execute(text(f"UPDATE tbl_ShipSchedule_KeHoach SET {clause_str} WHERE id = :id"), params)

            return jsonify({"success": True, "message": f"Đã hoàn tác: {action.get('description', 'Sửa ô')}"})

        elif act_type == 'DELETE_ROWS':
            deleted_rows = action.get('deleted_rows', [])
            if not deleted_rows:
                return jsonify({"success": False, "message": "Không có dữ liệu dòng đã xóa để khôi phục"}), 400

            insert_sql = text("""
                INSERT INTO tbl_ShipSchedule_KeHoach (
                    ky_thang, stt, so_lenh_tach, tau_name, trong_tai,
                    eta, eta_parsed, etb, dai_ly, thoi_gian_xong, ngay_duyet_so,
                    cang_xep, cang_den, lsd, lenh_xuat_hang, kl_xuat,
                    san_xuat, nha_may, cw_max, ghi_chu, nhip, nhip_sx, tinh_trang,
                    pic, so_no, tong_da_map, da_xuat, con_lai, updated_by
                ) VALUES (
                    :ky_thang, :stt, :so_lenh_tach, :tau_name, :trong_tai,
                    :eta, :eta_parsed, :etb, :dai_ly, :thoi_gian_xong, :ngay_duyet_so,
                    :cang_xep, :cang_den, :lsd, :lenh_xuat_hang, :kl_xuat,
                    :san_xuat, :nha_may, :cw_max, :ghi_chu, :nhip, :nhip_sx, :tinh_trang,
                    :pic, :so_no, :tong_da_map, :da_xuat, :con_lai, :updated_by
                )
            """)
            with engine.begin() as conn:
                for r in deleted_rows:
                    r['eta_parsed'] = parse_eta(r.get('eta'), default_year=get_year_from_ky(r.get('ky_thang', '')))
                    r['updated_by'] = user_name
                conn.execute(insert_sql, deleted_rows)

            return jsonify({"success": True, "message": f"Đã khôi phục lại {len(deleted_rows)} dòng vừa xóa."})

        elif act_type == 'INSERT_ROW':
            rid = action.get('row_id')
            if rid:
                with engine.begin() as conn:
                    conn.execute(text("DELETE FROM tbl_ShipSchedule_KeHoach WHERE id = :id"), {"id": rid})
            return jsonify({"success": True, "message": "Đã hoàn tác thao tác thêm dòng mới."})

        elif act_type == 'BATCH_RESTORE':
            # Hoàn tác toàn diện cho thao tác Ghép / Tách Tàu - SO
            created_ids = [int(x) for x in (action.get('created_ids') or []) if str(x).isdigit()]
            updated_rows = action.get('updated_rows') or []
            deleted_rows = action.get('deleted_rows') or []

            with engine.begin() as conn:
                if created_ids:
                    ph = ", ".join(str(i) for i in created_ids)
                    conn.execute(text(f"DELETE FROM tbl_ShipSchedule_KeHoach WHERE id IN ({ph})"))

                for ur in updated_rows:
                    rid = ur.get('id')
                    fields_data = ur.get('data') or {}
                    if not rid or not fields_data:
                        continue
                    set_parts = []
                    prm = {"id": rid, "user": user_name}
                    for k, v in fields_data.items():
                        set_parts.append(f"{k} = :{k}")
                        prm[k] = v
                    if 'eta' in fields_data:
                        set_parts.append("eta_parsed = :eta_dt")
                        prm['eta_dt'] = parse_eta(fields_data.get('eta'), default_year=get_year_from_ky(fields_data.get('ky_thang') or ''))
                    set_parts.extend(["updated_by = :user", "updated_at = GETDATE()"])
                    conn.execute(text(f"UPDATE tbl_ShipSchedule_KeHoach SET {', '.join(set_parts)} WHERE id = :id"), prm)

                if deleted_rows:
                    insert_sql = text("""
                        INSERT INTO tbl_ShipSchedule_KeHoach (
                            ky_thang, stt, so_lenh_tach, tau_name, trong_tai,
                            eta, eta_parsed, etb, dai_ly, thoi_gian_xong, ngay_duyet_so,
                            cang_xep, cang_den, lsd, lenh_xuat_hang, kl_xuat,
                            san_xuat, nha_may, cw_max, ghi_chu, nhip, nhip_sx, tinh_trang,
                            pic, so_no, tong_da_map, da_xuat, con_lai, updated_by
                        ) VALUES (
                            :ky_thang, :stt, :so_lenh_tach, :tau_name, :trong_tai,
                            :eta, :eta_parsed, :etb, :dai_ly, :thoi_gian_xong, :ngay_duyet_so,
                            :cang_xep, :cang_den, :lsd, :lenh_xuat_hang, :kl_xuat,
                            :san_xuat, :nha_may, :cw_max, :ghi_chu, :nhip, :nhip_sx, :tinh_trang,
                            :pic, :so_no, :tong_da_map, :da_xuat, :con_lai, :updated_by
                        )
                    """)
                    for dr in deleted_rows:
                        dr['eta_parsed'] = parse_eta(dr.get('eta'), default_year=get_year_from_ky(dr.get('ky_thang', '')))
                        dr['updated_by'] = user_name
                    conn.execute(insert_sql, deleted_rows)

            return jsonify({"success": True, "message": f"Đã hoàn tác: {action.get('description', 'Ghép/Tách Tàu - SO')}"})
        return jsonify({"success": False, "message": f"Loại thao tác '{act_type}' không hỗ trợ hoàn tác"}), 400
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi revert action: {e}")
        return jsonify({"success": False, "message": str(e)}), 500


# ==============================================================================
# 8. API IMPORT EXCEL, ĐỒNG BỘ DB, TRA CỨU SO & XUẤT EXCEL
# ==============================================================================

@tauscheduler_bp.route('/api/import-excel', methods=['POST'])
@login_required
def api_import_excel():
    allowed, err_msg = check_action_permission("can_import")
    if not allowed:
        return jsonify({"success": False, "message": err_msg}), 403

    if 'file' not in request.files:
        return jsonify({"success": False, "message": "Không tìm thấy file tải lên"}), 400

    file = request.files['file']
    if not file or file.filename == '':
        return jsonify({"success": False, "message": "Chưa chọn file Excel"}), 400

    import_mode = request.form.get('mode', 'selected_ky').strip()
    target_ky = request.form.get('ky_thang', '').strip()
    user_name = session.get('username', 'system')

    try:
        file_bytes = io.BytesIO(file.read())
        with pd.ExcelFile(file_bytes) as xls:
            candidate_sheets = {}
            for s in xls.sheet_names:
                if 'LỊCH TÀU' not in s.upper():
                    continue
                m = re.search(r'(\d{2})\.(\d{2,4})', s.strip())
                if not m:
                    continue
                mm, yy = m.group(1), m.group(2)
                if len(yy) == 2:
                    yy = "20" + yy
                norm_ky = f"{mm}.{yy}"

                if import_mode == 'selected_ky' and target_ky and norm_ky != target_ky:
                    continue
                if import_mode == 'all_2025_2026' and int(yy) < 2025:
                    continue

                recs = parse_excel_sheet_to_records(xls, s)
                if recs:
                    if norm_ky not in candidate_sheets or len(recs) > len(candidate_sheets[norm_ky]):
                        candidate_sheets[norm_ky] = recs

        if not candidate_sheets:
            return jsonify({
                "success": False,
                "message": f"Không tìm thấy sheet hợp lệ cho kỳ '{target_ky}' trong file Excel."
            }), 400

        total_inserted = 0
        imported_kys = []

        insert_sql = text("""
            INSERT INTO tbl_ShipSchedule_KeHoach (
                ky_thang, stt, so_lenh_tach, tau_name, trong_tai,
                eta, eta_parsed, etb, dai_ly, thoi_gian_xong, ngay_duyet_so,
                cang_xep, cang_den, lsd, lenh_xuat_hang, kl_xuat,
                san_xuat, nha_may, cw_max, ghi_chu, nhip, tinh_trang,
                pic, so_no, tong_da_map, da_xuat, con_lai, updated_by
            ) VALUES (
                :ky_thang, :stt, :so_lenh_tach, :tau_name, :trong_tai,
                :eta, :eta_parsed, :etb, :dai_ly, :thoi_gian_xong, :ngay_duyet_so,
                :cang_xep, :cang_den, :lsd, :lenh_xuat_hang, :kl_xuat,
                :san_xuat, :nha_may, :cw_max, :ghi_chu, :nhip, :tinh_trang,
                :pic, :so_no, :tong_da_map, :da_xuat, :con_lai, :updated_by
            )
        """)

        with engine.begin() as conn:
            for ky, records in candidate_sheets.items():
                conn.execute(text("DELETE FROM tbl_ShipSchedule_KeHoach WHERE ky_thang = :ky"), {"ky": ky})
                for r in records:
                    r['updated_by'] = user_name
                conn.execute(insert_sql, records)
                total_inserted += len(records)
                imported_kys.append(f"{ky} ({len(records)} dòng)")

        # Tự động đồng bộ Waterfall ngay cho kỳ hiện tại sau khi Import
        if target_ky in candidate_sheets:
            execute_waterfall_sync_for_ky(target_ky, user_name=user_name)
        sync_legacy_dashboard_tables_sync()

        return jsonify({
            "success": True,
            "message": f"Đã nạp thành công {total_inserted} dòng dữ liệu cho {len(imported_kys)} kỳ: {', '.join(imported_kys)}",
            "imported_kys": list(candidate_sheets.keys())
        })
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi Import Excel: {e}")
        return jsonify({"success": False, "message": f"Lỗi xử lý file Excel: {e}"}), 500
def sync_legacy_dashboard_tables_sync() -> dict:
    """Đồng bộ trực tiếp 1 luồng sang bảng dbo.lichtau và dbo.testdb để vẽ Dashboard."""
    now = datetime.now()
    if now.day > 10:
        start_dt = datetime(now.year, now.month, 1)
    else:
        first_day_this_month = datetime(now.year, now.month, 1)
        prev_month_date = first_day_this_month - timedelta(days=1)
        start_dt = datetime(prev_month_date.year, prev_month_date.month, 1)

    lichtau_cnt = 0
    testdb_cnt = 0
    with engine.begin() as conn:
        # 1. Đồng bộ bảng dbo.lichtau
        conn.execute(text("""
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'lichtau')
            BEGIN
                SELECT TOP 0 * INTO dbo.lichtau FROM vw_LichTau_Cleaned;
            END
        """))
        conn.execute(text("""
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'lichtau')
            BEGIN
                SELECT TOP 0 * INTO dbo.lichtau FROM vw_LichTau_Cleaned;
            END
            ELSE IF (
                (SELECT COUNT(*) FROM sys.columns WHERE object_id = OBJECT_ID('dbo.lichtau'))
                <>
                (SELECT COUNT(*) FROM sys.columns WHERE object_id = OBJECT_ID('dbo.vw_LichTau_Cleaned'))
            )
            BEGIN
                DROP TABLE dbo.lichtau;
                SELECT TOP 0 * INTO dbo.lichtau FROM vw_LichTau_Cleaned;
            END
            ELSE
            BEGIN
                TRUNCATE TABLE dbo.lichtau;
            END
        """))
        res_lt = conn.execute(text("""
            INSERT INTO dbo.lichtau
            SELECT *
            FROM vw_LichTau_Cleaned
            WHERE [SỐ LỆNH TÁCH] IS NOT NULL
              AND TRY_CONVERT(DATE, '01.' + LTRIM(RTRIM([SheetMonth])), 104) >= :start_dt
        """), {"start_dt": start_dt.strftime('%Y-%m-%d')})
        lichtau_cnt = res_lt.rowcount

        # 2. Đồng bộ bảng dbo.testdb (Không dùng TRY_CAST trên cột View để tận dụng Index)
        has_testdb = conn.execute(text("SELECT 1 FROM sys.tables WHERE name = 'testdb'")).fetchone()
        if has_testdb:
            conn.execute(text("TRUNCATE TABLE dbo.testdb;"))
            res_tdb = conn.execute(text("""
                INSERT INTO dbo.testdb (tau, SheetMonth, ETA_Parsed, shipped_qty, Mapping_kho, khoi_luong_tong)
                SELECT 
                    lt.tau_name AS tau,
                    lt.ky_thang AS SheetMonth,
                    lt.eta_parsed AS ETA_Parsed,
                    ISNULL(s.shipped_qty, 0) AS shipped_qty,
                    ISNULL(s.mapping_kho, 0) AS Mapping_kho,
                    ISNULL(lt.trong_tai, 0) AS khoi_luong_tong
                FROM tbl_ShipSchedule_KeHoach lt WITH (NOLOCK)
                LEFT JOIN (
                    SELECT 
                        CAST([SO Mapping] AS NVARCHAR(50)) AS so_map,
                        SUM(ISNULL([Shipped Quantity (KG)], 0)) AS shipped_qty,
                        SUM(ISNULL([SL Mapping kho], 0)) AS mapping_kho
                    FROM (
                        SELECT DISTINCT [SO Mapping], [Material], [Shipped Quantity (KG)], [SL Mapping kho]
                        FROM vw_so_kho_sumary2 WITH (NOLOCK)
                    ) v
                    GROUP BY [SO Mapping]
                ) s ON s.so_map = lt.so_lenh_tach
                WHERE lt.so_lenh_tach IS NOT NULL AND lt.so_lenh_tach != ''
                  AND lt.tau_name IS NOT NULL AND lt.tau_name != ''
            """))
            testdb_cnt = res_tdb.rowcount

    return {"lichtau_rows": lichtau_cnt, "testdb_rows": testdb_cnt}

@tauscheduler_bp.route('/api/refresh-schedule', methods=['POST'])
@login_required
def api_refresh_schedule():
    """
    API làm mới ngầm bảng tbl_ShipSchedule_KeHoach sau khi người dùng mở trang (F5):
    - Chỉ cập nhật Waterfall vào tbl_ShipSchedule_KeHoach (không đụng đến lichtau/testdb).
    - Trả về cờ updated=True nếu có bất kỳ dòng nào thay đổi số liệu để giao diện tự làm mới êm ái.
    """
    payload = request.get_json() or {}
    ky_thang = payload.get('ky', '').strip()
    if not ky_thang:
        return jsonify({"success": False, "updated": False}), 400

    try:
        stats = execute_waterfall_sync_for_ky(ky_thang, user_name=session.get('username', 'auto_refresh'))
        return jsonify({
            "success": True,
            "updated": stats["synced_count"] > 0,
            "changed_rows": stats["synced_count"]
        })
    except Exception as e:
        logger.warning(f"[SQL Server] Lỗi refresh-schedule: {e}")
        return jsonify({"success": False, "updated": False, "message": str(e)}), 500


@tauscheduler_bp.route('/api/sync-db', methods=['POST'])
@login_required
def api_sync_production():
    """
    Khi bấm nút 'Đồng bộ DB':
    Chạy 1 luồng trực tiếp (Synchronous): Cập nhật tbl_ShipSchedule_KeHoach VÀ đồng bộ sang lichtau & testdb để vẽ Dashboard.
    """
    allowed, err_msg = check_action_permission("can_sync")
    if not allowed:
        return jsonify({"success": False, "message": err_msg}), 403

    payload = request.get_json() or {}
    ky_thang = payload.get('ky', '').strip()
    if not ky_thang:
        return jsonify({"success": False, "message": "Thiếu kỳ tháng"}), 400

    try:
        stats = execute_waterfall_sync_for_ky(ky_thang, user_name=session.get('username', 'system'))
        dash_stats = sync_legacy_dashboard_tables_sync()
        return jsonify({
            "success": True,
            "message": (
                f"Đồng bộ hoàn tất! Đã cập nhật {stats['synced_count']} dòng thay đổi trong kỳ {ky_thang} "
                f"và đồng bộ 1 luồng sang Dashboard (lichtau: {dash_stats['lichtau_rows']} dòng, testdb: {dash_stats['testdb_rows']} dòng)."
            )
        })
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi đồng bộ DB: {e}")
        return jsonify({"success": False, "message": str(e)}), 500


@tauscheduler_bp.route('/api/pkd-so-search', methods=['GET'])
@login_required
def api_pkd_so_search():
    """
    API hỗ trợ P.KD tìm kiếm đa năng theo:
    - Mã SO (10 chữ số hoặc một phần mã SO)
    - Số XNĐH / Hợp đồng / Tên khách hàng (tìm trong cả order_production_rules và tbl_ShipSchedule_KeHoach)
    """
    q = (request.args.get('q') or '').strip()
    if len(q) < 3:
        return jsonify({"success": False, "message": "Vui lòng nhập ít nhất 3 ký tự (Mã SO, Số XNĐH hoặc Tên khách hàng)"}), 400

    try:
        candidates = {}
        with engine.connect() as conn:
            sched_rows = conn.execute(text("""
                SELECT TOP 30
                    so_lenh_tach,
                    MAX(lenh_xuat_hang) AS lxh_sample,
                    COUNT(1) AS ship_count,
                    SUM(ISNULL(kl_xuat, 0)) AS sum_kl_xuat,
                    MAX(ky_thang) AS latest_ky
                FROM tbl_ShipSchedule_KeHoach WITH(NOLOCK)
                WHERE (so_lenh_tach LIKE :kw OR lenh_xuat_hang LIKE :kw OR tau_name LIKE :kw)
                GROUP BY so_lenh_tach
                ORDER BY MAX(ky_thang) DESC
            """), {"kw": f"%{q}%"}).mappings().all()

            for r in sched_rows:
                so_extracted, _, _ = analyze_and_extract_so_lenh(r['so_lenh_tach'], r['lxh_sample'])
                so_k = normalize_so_key(so_extracted)
                if not so_k:
                    continue
                candidates[so_k] = {
                    "so_no": so_k,
                    "xndh": str(r['lxh_sample'] or '').strip()[:120],
                    "ship_count": int(r['ship_count'] or 0),
                    "sum_kl_tau": round(float(r['sum_kl_xuat'] or 0.0), 1),
                    "latest_ky": str(r['latest_ky'] or ''),
                    "source": "Lịch tàu"
                }

            # 2. Tìm trong bảng order_production_rules (Kể cả SO chưa được gắn vào lịch tàu nào)
            has_opr = conn.execute(text("SELECT 1 FROM sys.tables WHERE name = 'order_production_rules'")).fetchone()
            if has_opr:
                opr_rows = conn.execute(text("""
                    SELECT TOP 30
                        CAST([SO_mapping] AS NVARCHAR(50)) AS so_map,
                        MAX([XNDH]) AS xndh,
                        COUNT(DISTINCT [Order]) AS order_cnt,
                        SUM(ISNULL([total_weight], 0)) / 1000.0 AS total_lsx_tan,
                        MAX([KySanXuat]) AS ky_sx
                    FROM dbo.order_production_rules WITH(NOLOCK)
                    WHERE [SO_mapping] IS NOT NULL
                      AND (
                          CAST([SO_mapping] AS NVARCHAR(50)) LIKE :kw
                          OR [XNDH] LIKE :kw
                          OR CAST([Order] AS NVARCHAR(50)) LIKE :kw
                      )
                    GROUP BY [SO_mapping]
                """), {"kw": f"%{q}%"}).mappings().all()

                for r in opr_rows:
                    so_k = normalize_so_key(r['so_map'])
                    if not so_k:
                        continue
                    if so_k in candidates:
                        if r['xndh'] and not candidates[so_k].get('xndh_clean'):
                            candidates[so_k]['xndh'] = str(r['xndh']).strip()
                        candidates[so_k]['total_lsx_tan'] = round(float(r['total_lsx_tan'] or 0.0), 1)
                    else:
                        candidates[so_k] = {
                            "so_no": so_k,
                            "xndh": str(r['xndh'] or '').strip(),
                            "ship_count": 0,
                            "sum_kl_tau": 0.0,
                            "total_lsx_tan": round(float(r['total_lsx_tan'] or 0.0), 1),
                            "latest_ky": str(r['ky_sx'] or 'Chưa xếp tàu'),
                            "source": "Order SX (Chưa xếp tàu)"
                        }

            # 3. Nếu vẫn chưa có và người dùng gõ số SO -> Tìm trong vw_so_kho_sumary2
            if not candidates and re.search(r'\d{5,}', q):
                kho_rows = conn.execute(text("""
                    SELECT TOP 20
                        CAST([SO Mapping] AS NVARCHAR(50)) AS so_map,
                        MAX([Material Description]) AS mat_desc,
                        SUM(ISNULL([Quantity (KG)], 0)) / 1000.0 AS qty_tan
                    FROM vw_so_kho_sumary2 WITH(NOLOCK)
                    WHERE CAST([SO Mapping] AS NVARCHAR(50)) LIKE :kw
                    GROUP BY [SO Mapping]
                """), {"kw": f"%{q}%"}).mappings().all()

                for r in kho_rows:
                    so_k = normalize_so_key(r['so_map'])
                    if so_k:
                        candidates[so_k] = {
                            "so_no": so_k,
                            "xndh": str(r['mat_desc'] or '').strip(),
                            "ship_count": 0,
                            "sum_kl_tau": 0.0,
                            "total_lsx_tan": round(float(r['qty_tan'] or 0.0), 1),
                            "latest_ky": "Chưa xếp tàu",
                            "source": "SAP Kho"
                        }

        results = list(candidates.values())
        return jsonify({
            "success": True,
            "query": q,
            "count": len(results),
            "results": results
        })
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi api_pkd_so_search: {e}")
        return jsonify({"success": False, "message": str(e)}), 500


@tauscheduler_bp.route('/api/so-info/<so_lenh>', methods=['GET'])
@login_required
def api_get_so_info(so_lenh):
    """
    Tra cứu toàn diện thông tin của 1 Số Lệnh Tách (SO Mapping) dành cho P.KD & P.KH:
    1. Tiến độ kho & xuất hàng theo từng kích cỡ vật tư từ vw_so_kho_sumary2
    2. Dữ liệu WIP & Kho chưa Map từ hệ thống QC (nk_chua_map, cho_nk_loai1, sx_cho_co_tinh, sx_cxl)
    3. Danh sách toàn bộ chuyến tàu đang được phân bổ SO này trong tbl_ShipSchedule_KeHoach (mọi kỳ)
    4. Danh sách Order sản xuất & tiến độ cán thực tế từ bảng order_production_rules
    5. Hệ thống Cảnh báo Thông minh (6 kịch bản rủi ro cho P.KD)
    """
    so_clean = re.sub(r'\D', '', str(so_lenh or ''))
    if not so_clean:
        return jsonify({"success": False, "message": "Số lệnh tách không hợp lệ"}), 400

    try:
        so_key = normalize_so_key(so_clean)

        # 1. Lấy dữ liệu tiến độ vật tư từ vw_so_kho_sumary2
        inv_map = fetch_sqlserver_so_inventory([so_clean])
        kho_summary = inv_map.get(so_key, {
            'total_qty_kg': 0.0,
            'shipped_qty_kg': 0.0,
            'mapping_kho_kg': 0.0,
            'available_kg': 0.0,
            'nha_may': '',
            'materials': []
        })

        # 2. Lấy dữ liệu WIP & NK chưa map từ hệ thống QC
        qc_map = fetch_qc_wip_and_unmapped_by_so([so_clean])
        qc_info = qc_map.get(so_key, {
            'nk_chua_map_kg': 0.0,
            'cho_nk_loai1_kg': 0.0,
            'sx_cho_co_tinh_kg': 0.0,
            'sx_cxl_kg': 0.0,
            'total_loai1_qc_kg': 0.0,
            'pending_coils_detail': []
        })

        with engine.connect() as conn:
            # 3. Lấy đầy đủ thông tin các chuyến tàu đang chở SO này trong tbl_ShipSchedule_KeHoach
            ship_Allocations = conn.execute(text("""
                SELECT 
                    id, ky_thang, stt, tau_name, ISNULL(trong_tai, 0) AS trong_tai,
                    eta, eta_parsed, etb, cang_xep, cang_den, lsd,
                    lenh_xuat_hang, ISNULL(kl_xuat, 0) AS kl_xuat,
                    nha_may, nhip, pic, tinh_trang, ghi_chu,
                    ISNULL(tong_da_map, 0) AS tong_da_map,
                    ISNULL(da_xuat, 0) AS da_xuat,
                    ISNULL(con_lai, 0) AS con_lai,
                    ISNULL(nk_chua_map, 0) AS nk_chua_map,
                    ISNULL(cho_nk_loai1, 0) AS cho_nk_loai1,
                    ISNULL(tong_map_dk, 0) AS tong_map_dk,
                    ISNULL(sx_cho_co_tinh, 0) AS sx_cho_co_tinh,
                    ISNULL(sx_cxl, 0) AS sx_cxl,
                    ISNULL(tong_sx_qc, 0) AS tong_sx_qc,
                    ISNULL(ngay_dk_xong, '') AS ngay_dk_xong
                FROM tbl_ShipSchedule_KeHoach WITH(NOLOCK)
                WHERE so_lenh_tach = :so
                ORDER BY ky_thang DESC, COALESCE(eta_parsed, '2099-12-31') ASC, stt ASC
            """), {"so": so_clean}).mappings().all()

            ships_list = []
            for r in ship_Allocations:
                d = dict(r)
                if d.get('eta_parsed') and hasattr(d['eta_parsed'], 'strftime'):
                    d['eta_parsed'] = d['eta_parsed'].strftime("%Y-%m-%d")
                ships_list.append(d)

            # 4. Lấy chi tiết Order sản xuất từ bảng order_production_rules
            has_opr = conn.execute(text("SELECT 1 FROM sys.tables WHERE name = 'order_production_rules'")).fetchone()
            orders_list = []
            if has_opr:
                rows = conn.execute(text("""
                    SELECT 
                        r.[SO_mapping] AS so_mapping,
                        r.[Order] AS order_no,
                        r.[XNDH] AS xndh,
                        COALESCE(sd.grade, r.material_desc, '') AS mac_thep,
                        r.[req_thick] AS do_day,
                        r.[req_width] AS kho_rong,
                        r.[req_min_w] AS min_w_kg,
                        r.[req_max_w] AS max_w_kg,
                        ISNULL(r.[total_weight], 0) AS target_weight_kg,
                        ISNULL(r.[fulfilled_weight], 0) AS fulfilled_weight_kg,
                        r.[KySanXuat] AS ky_sx,
                        ISNULL(r.[is_skin_required], 0) AS is_skin,
                        COALESCE(sd.usage_purpose, r.special_request, '') AS muc_dich
                    FROM dbo.order_production_rules r WITH(NOLOCK)
                    OUTER APPLY (
                        SELECT TOP 1 grade, usage_purpose
                        FROM dbo.so_details s WITH(NOLOCK)
                        WHERE s.so_number = r.SO_mapping
                          AND ISNULL(s.thickness, -1) = ISNULL(r.req_thick, -1)
                          AND ISNULL(s.width, -1) = ISNULL(r.req_width, -1)
                    ) sd
                    WHERE r.[SO_mapping] = TRY_CAST(:so AS BIGINT)
                    ORDER BY r.[Order] ASC
                """), {"so": so_clean}).mappings().all()

                for r in rows:
                    d = dict(r)
                    min_t = round(float(d.get('min_w_kg') or 0) / 1000.0, 1)
                    max_t = round(float(d.get('max_w_kg') or 0) / 1000.0, 1)
                    if min_t > 0 and max_t > 0:
                        d['cw'] = f"{min_t:g}-{max_t:g}"
                    elif max_t > 0:
                        d['cw'] = f"0-{max_t:g}"
                    else:
                        d['cw'] = "-"
                    orders_list.append(d)

        # 5. Tổng hợp các chỉ số Pipeline toàn SO (Đơn vị: Tấn)
        sap_total_tan = round(kho_summary['total_qty_kg'] / 1000.0, 2)
        tong_lsx_tan = round(sum(float(r.get('target_weight_kg') or 0) for r in orders_list) / 1000.0, 2)
        tong_da_can_tan = round(sum(float(r.get('fulfilled_weight_kg') or 0) for r in orders_list) / 1000.0, 2)

        # Tổng khối lượng đơn hàng chuẩn (ưu tiên SAP, nếu SAP chưa đồng bộ thì lấy theo LSX)
        tong_don_chuan_tan = sap_total_tan if sap_total_tan > 0 else tong_lsx_tan

        mapping_kho_tan = round(kho_summary['mapping_kho_kg'] / 1000.0, 2)
        shipped_sap_tan = round(kho_summary['shipped_qty_kg'] / 1000.0, 2)
        avail_sap_tan = round(kho_summary['available_kg'] / 1000.0, 2)

        nk_chua_map_tan = round(qc_info['nk_chua_map_kg'] / 1000.0, 2)
        cho_nk_l1_tan = round(qc_info['cho_nk_loai1_kg'] / 1000.0, 2)
        cho_co_tinh_tan = round(qc_info['sx_cho_co_tinh_kg'] / 1000.0, 2)
        cxl_tan = round(qc_info['sx_cxl_kg'] / 1000.0, 2)
        tong_wip_qc_tan = round(cho_nk_l1_tan + cho_co_tinh_tan + cxl_tan, 2)

        # Tổng thực xuất trên lịch tàu
        sum_kl_xep_tau = round(sum(float(s.get('kl_xuat') or 0) for s in ships_list), 2)
        sum_da_xuat_tau = round(sum(float(s.get('da_xuat') or 0) for s in ships_list), 2)
        thuc_xuat_tan = max(shipped_sap_tan, sum_da_xuat_tau)

        # Tổng nguồn hàng toàn SO (Đã xuất + Đã Map + NK Chưa Map + WIP QC)
        tong_nguon_toan_so = round(max(avail_sap_tan, thuc_xuat_tan + mapping_kho_tan) + nk_chua_map_tan + tong_wip_qc_tan, 2)
        chua_xep_tau_tan = round(max(0.0, tong_don_chuan_tan - sum_kl_xep_tau), 2) if tong_don_chuan_tan > 0 else 0.0
        thieu_nguon_toan_so = round(max(0.0, tong_don_chuan_tan - tong_nguon_toan_so), 2) if tong_don_chuan_tan > 0 else 0.0

        # 6. ĐÁNH GIÁ & TẠO CẢNH BÁO THÔNG MINH CHO P.KD (Alert Engine)
        pkd_alerts = []

        # (A) Kiểm tra tình hình gắn tàu
        if not ships_list:
            pkd_alerts.append({
                "level": "warning",
                "icon": "fa-ship",
                "title": "CHƯA GẮN LỊCH TÀU",
                "msg": f"Mã SO {so_clean} hiện chưa được gắn vào bất kỳ chuyến tàu nào trên hệ thống Kế hoạch Lịch tàu."
            })
        else:
            wait_vessel_rows = [s for s in ships_list if not (s.get('tau_name') or '').strip()]
            if wait_vessel_rows:
                kl_wait = sum(float(s.get('kl_xuat') or 0) for s in wait_vessel_rows)
                pkd_alerts.append({
                    "level": "warning",
                    "icon": "fa-anchor",
                    "title": "ĐANG CHỜ ĐIỀU ĐỘ GHÉP TÀU",
                    "msg": f"Có {len(wait_vessel_rows)} dòng lệnh ({kl_wait:,.0f} tấn) đã lên kế hoạch nhưng chưa có tên tàu (Chờ ghép tàu)."
                })

            if tong_don_chuan_tan > 0 and sum_kl_xep_tau < tong_don_chuan_tan * 0.90:
                pkd_alerts.append({
                    "level": "warning",
                    "icon": "fa-balance-scale",
                    "title": "CHƯA XẾP HẾT KHỐI LƯỢNG ĐƠN HÀNG LÊN TÀU",
                    "msg": f"Tổng đơn {tong_don_chuan_tan:,.0f}T nhưng mới xếp lên lịch tàu {sum_kl_xep_tau:,.0f}T (Còn {chua_xep_tau_tan:,.0f}T chưa gắn tàu)."
                })

        # (B) Kiểm tra từng chuyến tàu: Rủi ro trễ LSD, trễ ETA, và thiếu nguồn cán bù
        for s in ships_list:
            tau_label = (s.get('tau_name') or '').strip() or "Chờ ghép tàu"
            ky_str = s.get('ky_thang') or ''
            sheet_yr = get_year_from_ky(ky_str)
            kl_yc = float(s.get('kl_xuat') or 0)
            da_xuat_s = float(s.get('da_xuat') or 0)
            tong_sx_s = max(float(s.get('tong_sx_qc') or 0), float(s.get('tong_da_map') or 0), da_xuat_s)
            is_ship_done = (kl_yc > 0 and da_xuat_s >= kl_yc * 0.95)

            eta_dt = parse_eta(s.get('eta'), default_year=sheet_yr)
            lsd_dt = parse_eta(s.get('lsd'), default_year=sheet_yr)
            ngay_xong_str = (s.get('ngay_dk_xong') or '').strip()
            est_dt = parse_eta(ngay_xong_str, default_year=sheet_yr) if (ngay_xong_str and ngay_xong_str != 'Đã đủ hàng kho') else None

            # Cảnh báo 1: ETA trễ hơn LSD
            if eta_dt and lsd_dt and eta_dt > lsd_dt and not is_ship_done:
                pkd_alerts.append({
                    "level": "danger",
                    "icon": "fa-calendar-times",
                    "title": f"RỦI RO TRỄ HẠN GIAO HÀNG (LSD) - TÀU {tau_label} (Kỳ {ky_str})",
                    "msg": f"Ngày tàu đến dự kiến (ETA: {eta_dt.strftime('%d/%m/%Y')}) đang trễ hơn hạn LSD tối đa ({lsd_dt.strftime('%d/%m/%Y')})!"
                })

            # Cảnh báo 2: Ngày dự kiến xong hàng trễ hơn LSD hoặc ETA
            if est_dt and not is_ship_done:
                if lsd_dt and est_dt > lsd_dt:
                    pkd_alerts.append({
                        "level": "danger",
                        "icon": "fa-clock",
                        "title": f"HÀNG SẢN XUẤT XONG TRỄ HƠN LSD - TÀU {tau_label}",
                        "msg": f"Ngày dự kiến xong hàng ({est_dt.strftime('%d/%m/%Y')}) trễ hơn hạn giao hàng LSD ({lsd_dt.strftime('%d/%m/%Y')})!"
                    })
                elif eta_dt and est_dt > eta_dt:
                    pkd_alerts.append({
                        "level": "danger",
                        "icon": "fa-exclamation-triangle",
                        "title": f"CẢNH BÁO TÀU CHỜ HÀNG - TÀU {tau_label} (Kỳ {ky_str})",
                        "msg": f"Ngày dự kiến hoàn thành nhập kho ({est_dt.strftime('%d/%m/%Y')}) trễ hơn ngày ETA của tàu ({eta_dt.strftime('%d/%m/%Y')})!"
                    })

            # Cảnh báo 3: Chuyến tàu thiếu nguồn hàng (Cần cán bù)
            if kl_yc > 0 and not is_ship_done and tong_sx_s < kl_yc * 0.95:
                thieu_tau = round(kl_yc - tong_sx_s, 1)
                pkd_alerts.append({
                    "level": "danger",
                    "icon": "fa-industry",
                    "title": f"THIẾU NGUỒN HÀNG CHO TÀU {tau_label} (Kỳ {ky_str})",
                    "msg": f"Yêu cầu xuất {kl_yc:,.0f}T nhưng tổng nguồn phân bổ (Kho + Dự kiến QC) mới đạt {tong_sx_s:,.0f}T. Thiếu {thieu_tau:,.0f}T cần đôn đốc cán bù!"
                })

        # (C) Kiểm tra hàng NK Chưa Map & Hàng vướng CXL / Chờ cơ tính
        if nk_chua_map_tan > 0:
            pkd_alerts.append({
                "level": "warning",
                "icon": "fa-boxes",
                "title": f"CÓ {nk_chua_map_tan:,.1f} TẤN ĐÃ NHẬP KHO NHƯNG CHƯA MAP SO",
                "msg": f"Hàng Loại 1 đã nhập kho thực tế nhưng chưa được gán SO Mapping trên SAP. Đề nghị P.KH / Kho thực hiện Map SO để giữ hàng cho tàu!"
            })

        if cxl_tan > 0 or cho_co_tinh_tan > 0:
            pkd_alerts.append({
                "level": "info",
                "icon": "fa-flask",
                "title": "CÓ LƯỢNG HÀNG ĐANG CHỜ KẾT QUẢ QC / XỬ LÝ (T+7)",
                "msg": f"Đang có {cho_co_tinh_tan:,.1f}T chờ kết quả Cơ tính và {cxl_tan:,.1f}T chờ xử lý (CXL) tại nhà máy."
            })

        # (D) Nếu không có bất kỳ cảnh báo nguy hiểm/vàng nào -> Báo trạng thái An toàn
        if not any(a['level'] in ['danger', 'warning'] for a in pkd_alerts):
            pkd_alerts.insert(0, {
                "level": "success",
                "icon": "fa-check-circle",
                "title": "TIẾN ĐỘ ĐƠN HÀNG & LỊCH TÀU AN TOÀN",
                "msg": f"Đơn hàng SO {so_clean} đã được bố trí lịch tàu và nguồn hàng đáp ứng tốt tiến độ."
            })

        xndh_title = next((str(o.get('xndh') or '').strip() for o in orders_list if o.get('xndh')), "")
        if not xndh_title and ships_list:
            xndh_title = str(ships_list[0].get('lenh_xuat_hang') or '').strip()[:100]

        return jsonify({
            "success": True,
            "so_lenh": so_clean,
            "xndh_info": xndh_title,
            "kho_summary": {
                "total_qty_tan": sap_total_tan,
                "tong_don_chuan_tan": tong_don_chuan_tan,
                "mapping_kho_tan": mapping_kho_tan,
                "shipped_tan": shipped_sap_tan,
                "thuc_xuat_tan": thuc_xuat_tan,
                "available_tan": avail_sap_tan,
                "nk_chua_map_tan": nk_chua_map_tan,
                "cho_nk_l1_tan": cho_nk_l1_tan,
                "cho_co_tinh_tan": cho_co_tinh_tan,
                "cxl_tan": cxl_tan,
                "tong_wip_qc_tan": tong_wip_qc_tan,
                "tong_nguon_toan_so": tong_nguon_toan_so,
                "sum_kl_xep_tau": sum_kl_xep_tau,
                "chua_xep_tau_tan": chua_xep_tau_tan,
                "thieu_nguon_toan_so": thieu_nguon_toan_so,
                "nha_may": kho_summary['nha_may'],
                "materials": kho_summary['materials']
            },
            "pkd_alerts": pkd_alerts,
            "ship_allocations": ships_list,
            "pending_coils_count": len(qc_info.get('pending_coils_detail', [])),
            "count": len(orders_list),
            "tong_lsx_tan": tong_lsx_tan,
            "tong_da_can_tan": tong_da_can_tan,
            "items": orders_list
        })
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi tra cứu tổng hợp SO: {e}")
        return jsonify({"success": False, "message": str(e)}), 500


@tauscheduler_bp.route('/export-excel')
@login_required
@permission_required('view_ship_schedule')
def export_excel():
    now = datetime.now()
    default_ky = f"{now.month:02d}.{now.year}"
    ky_thang = request.args.get('ky', default_ky).strip()
    try:
        with engine.connect() as conn:
            query = text("""
                SELECT 
                    stt, so_lenh_tach, tau_name, trong_tai, eta, dai_ly, etb,
                    thoi_gian_xong, ngay_duyet_so, cang_xep, cang_den, lsd,
                    lenh_xuat_hang, kl_xuat, san_xuat, nhip_sx, cw_max,
                    nhip, tinh_trang, nha_may, pic, so_no,
                    tong_da_map, da_xuat, nk_chua_map, cho_nk_loai1, tong_map_dk,
                    sx_cho_co_tinh, sx_cxl, tong_sx_qc, con_lai, ngay_dk_xong, ghi_chu
                FROM tbl_ShipSchedule_KeHoach WITH(NOLOCK)
                WHERE ky_thang = :ky
                ORDER BY COALESCE(stt, 999999) ASC, id ASC
            """)
            df = pd.read_sql(query, conn, params={"ky": ky_thang})

        col_name_mapping = {
            'stt': 'STT',
            'so_lenh_tach': 'SỐ LỆNH TÁCH',
            'tau_name': 'TÀU/PHƯƠNG TIỆN VẬN TẢI',
            'trong_tai': 'KHỐI LƯỢNG TỔNG TÀU',
            'eta': 'ETA DUNG QUẤT',
            'dai_ly': 'ĐẠI LÝ',
            'etb': 'ETB DUNG QUẤT',
            'thoi_gian_xong': 'THỜI GIAN LÀM XONG HÀNG',
            'ngay_duyet_so': 'NGÀY DK DUYỆT SO',
            'cang_xep': 'Cảng xếp',
            'cang_den': 'CẢNG ĐẾN',
            'lsd': 'LSD',
            'lenh_xuat_hang': 'LỆNH XUẤT HÀNG - KẾ HOẠCH DUYỆT (SỐ LỆNH ĐẦY ĐỦ - SỐ XNĐH - KL TỔNG ĐƠN - LSD) (MỖI LỆNH 1 DÒNG)',
            'kl_xuat': 'KHỐI LƯỢNG HÀNG XUẤT LÊN TÀU',
            'san_xuat': 'SẢN XUẤT (HRC 1/2-TÌNH TRẠNG)',
            'cw_max': 'C.W MAX TÀU NHẬN ĐƯỢC',
            'ghi_chu': 'GHI CHÚ',
            'nhip': 'NHỊP',
            'nhip_sx': 'NHỊP SẢN XUẤT',
            'tinh_trang': 'TÌNH TRẠNG',
            'nha_may': 'HRC1/2',
            'pic': 'PIC',
            'so_no': 'SO',
            'tong_da_map': 'TỔNG ĐÃ MAP',
            'da_xuat': 'ĐÃ XUẤT',
            'con_lai': 'CÒN LẠI',
            'nk_chua_map': 'NK CHƯA MAP (LOẠI 1)',
            'cho_nk_loai1': 'CHỜ NK (LOẠI 1)',
            'tong_map_dk': 'TỔNG MAP DỰ KIẾN',
            'sx_cho_co_tinh': 'CHỜ KQ CƠ TÍNH',
            'sx_cxl': 'CHỜ XỬ LÝ (CXL)',
            'tong_sx_qc': 'TỔNG SX (KHÔNG TÍNH LOẠI 2)',
            'ngay_dk_xong': 'NGÀY DK HOÀN THÀNH',
        }
        df.rename(columns=col_name_mapping, inplace=True)

        output = io.BytesIO()
        sheet_label = f"LỊCH TÀU - {ky_thang}"[:31]
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_label, startrow=2)
            ws = writer.sheets[sheet_label]
            ws.cell(1, 1, f"LỊCH XUẤT HRC THÁNG {ky_thang.replace('.', '/')}")
            ws.cell(2, 1, "TỔNG XUẤT DỰ KIẾN")
            ws.cell(2, 14, float(df['KHỐI LƯỢNG HÀNG XUẤT LÊN TÀU'].fillna(0).sum()))

        output.seek(0)
        file_name = f"SO_HOA_LICH_TAU_{ky_thang.replace('.', '_')}.xlsx"

        return send_file(
            output,
            as_attachment=True,
            download_name=file_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi xuất file: {e}")
        flash(f"Lỗi xuất file: {e}", "danger")
        return redirect(url_for('tauscheduler_bp.view_schedule', ky=ky_thang))
# ==============================================================================
# 9. TRÍCH XUẤT COIL_LIST (ĐÃ MAPPING KHO & ĐÃ XUẤT LÊN TÀU THEO TÀU - SO)
# ==============================================================================

def fetch_coil_list_by_ship_and_so(tau_name: str = "", so_list: list = None) -> dict:
    """
    Lấy danh sách chi tiết từng cuộn (COIL_LIST):
    1. exported_coils: Cuộn thực tế đã xuất lên tàu từ MySQL (v_phieuxuathang_hrc) theo (Tàu, SO).
    2. mapped_coils: Cuộn đã mapping trong kho từ SQL Server theo danh sách SO.
       Hỗ trợ tự động dò cột linh hoạt (Auto-detect columns) để không phụ thuộc cứng vào tên cột.
    """
    clean_tau = normalize_ship_name(tau_name)
    clean_sos = [normalize_so_key(s) for s in (so_list or []) if normalize_so_key(s)]
    clean_sos = list(dict.fromkeys(clean_sos))

    exported_coils = []
    mapped_coils = []

    # 1. Truy vấn danh sách cuộn ĐÃ XUẤT LÊN TÀU trên MySQL (v_phieuxuathang_hrc)
    if engine_mysql is not None and (clean_tau or clean_sos):
        try:
            with engine_mysql.connect() as mconn:
                where_clauses = ["SO IS NOT NULL", "SO != ''"]
                params = {}

                if clean_tau and clean_tau not in ['ĐƯỜNG BỘ', 'DUONG BO', 'ALL']:
                    where_clauses.append("UPPER(TRIM(Transporter)) = :tau")
                    params['tau'] = clean_tau

                if clean_sos:
                    so_placeholders = []
                    for idx, s_val in enumerate(clean_sos):
                        k1, k2 = f"so_{idx}", f"so0_{idx}"
                        so_placeholders.extend([f":{k1}", f":{k2}"])
                        params[k1] = s_val
                        params[k2] = s_val.zfill(10)
                    where_clauses.append(f"TRIM(SO) IN ({', '.join(so_placeholders)})")

                sql_mysql = text(f"""
                    SELECT *
                    FROM v_phieuxuathang_hrc
                    WHERE {' AND '.join(where_clauses)}
                    ORDER BY IssueDate DESC
                """)
                rows = mconn.execute(sql_mysql, params).mappings().all()

                for r in rows:
                    d = dict(r)
                    # Tự động nhận diện tên cột Mã cuộn / Phiếu / Trọng lượng / Ngày xuất trong v_phieuxuathang_hrc
                    coil_id = (
                        d.get('CoilID') or d.get('ProductName')
                    )
                    material_desc = d.get('SAPDescription') or ''
                    weight_kg = float(d.get('Weight') or 0.0)
                    issue_dt = d.get('IssueDate')
                    if hasattr(issue_dt, 'strftime'):
                        issue_dt = issue_dt.strftime('%d/%m/%Y %H:%M')

                    exported_coils.append({
                        "coil_id": str(coil_id).strip(),
                        "so_no": normalize_so_key(d.get('SO')),
                        "tau_name": normalize_ship_name(d.get('Transporter')),
                        "material": str(material_desc).strip(),
                        "weight_kg": round(weight_kg, 1),
                        "weight_tan": round(weight_kg / 1000.0, 3),
                        "issue_date": str(issue_dt),
                        "nha_may": str(d.get('Plant') or d.get('NhaMay') or d.get('Factory') or '').strip(),
                        "phieu_xuat": str(d.get('TicketName') or '').strip()
                    })
        except Exception as e:
            logger.warning(f"[MySQL] Lỗi truy vấn chi tiết cuộn v_phieuxuathang_hrc: {e}")

   # 2. Truy vấn danh sách cuộn ĐÃ MAPPING KHO trên SQL Server (Trực tiếp từ bảng kho & so)
    # 2. Truy vấn danh sách cuộn ĐÃ MAPPING KHO trên SQL Server (Khớp chuẩn cấu trúc 23 cột của dbo.kho)
    if clean_sos:
        try:
            import unicodedata

            def norm_nfc(val: str) -> str:
                """Chuẩn hóa chuỗi về Unicode NFC chữ thường để so khớp không sợ lệch bảng mã NFC/NFD."""
                return unicodedata.normalize('NFC', str(val or '')).strip().lower()

            with engine.connect() as sconn:
                # Chuyển danh sách SO/Order về kiểu số nguyên (BIGINT) đúng với kiểu cột [SO Mapping] trong dbo.kho
                so_int_vals = list(dict.fromkeys(int(s) for s in clean_sos if s.isdigit()))
                if so_int_vals:
                    so_params = {f"s_{idx}": val for idx, val in enumerate(so_int_vals)}
                    in_clause = ", ".join(f":{k}" for k in so_params.keys())

                    # Đọc tên cột nguyên bản từ sys.columns và map qua khóa đã chuẩn hóa NFC
                    raw_cols = [
                        row[0] for row in sconn.execute(text("""
                            SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('dbo.kho')
                        """)).fetchall()
                    ]
                    kho_cols_nfc = {norm_nfc(c): c for c in raw_cols}

                    if kho_cols_nfc:
                        # Lấy đúng tên cột nguyên bản trong SQL Server (bất kể DB lưu NFC hay NFD)
                        col_id_cuon = next((kho_cols_nfc[k] for k in ['id cuộn bó', 'id cuộn', 'id_cuon_bo'] if k in kho_cols_nfc), None)
                        col_kl = next((kho_cols_nfc[k] for k in ['khối lượng', 'khoi luong', 'weight'] if k in kho_cols_nfc), None)
                        col_mat_desc = next((kho_cols_nfc[k] for k in ['material description', 'item description'] if k in kho_cols_nfc), None)
                        col_vi_tri = next((kho_cols_nfc[k] for k in ['vị trí', 'vi tri'] if k in kho_cols_nfc), None)
                        col_mac_thep = next((kho_cols_nfc[k] for k in ['mác thép', 'mac thep'] if k in kho_cols_nfc), None)

                        expr_coil = f"k.[{col_id_cuon}]" if col_id_cuon else "CAST(CAST(k.[Batch] AS BIGINT) AS NVARCHAR(50))"
                        expr_weight = f"ISNULL(k.[{col_kl}], 0)" if col_kl else "0"
                        expr_desc = f"k.[{col_mat_desc}]" if col_mat_desc else "NULL"
                        expr_vitri = f"k.[{col_vi_tri}]" if col_vi_tri else "''"
                        expr_macthep = f"k.[{col_mac_thep}]" if col_mac_thep else "''"

                        # Tìm đồng thời trên cả [SO Mapping] (2000...) và [Order] (2400...)
                        q_kho = text(f"""
                            SELECT 
                                {expr_coil} AS coil_id,
                                CAST(CAST(k.[Order] AS BIGINT) AS NVARCHAR(50)) AS order_no,
                                CAST(k.[SO Mapping] AS NVARCHAR(50)) AS so_no,
                                CAST(k.[Material] AS NVARCHAR(50)) AS material_code,
                                COALESCE({expr_desc}, s.[Item Description], {expr_macthep}, CAST(k.[Material] AS NVARCHAR(50))) AS material_desc,
                                ISNULL(k.[Khối lượng], 0) AS weight_kg,
                                {expr_vitri} AS vi_tri,
                                k.[status] AS kho_status,
                                CASE 
                                    WHEN s.[Factory] = N'Hòa phát dung quất' THEN 'HRC1'
                                    WHEN s.[Factory] = N'Hòa phát dung quất 2' THEN 'HRC2'
                                    WHEN k.[Plant] = 1000 THEN 'HRC1'
                                    WHEN k.[Plant] = 1002 THEN 'HRC2'
                                    ELSE COALESCE(s.[Factory], CAST(k.[Plant] AS NVARCHAR(50)), '')
                                END AS nha_may
                            FROM dbo.kho k WITH(NOLOCK)
                            LEFT JOIN (
                                SELECT DISTINCT 
                                    TRY_CAST([Sales Document] AS BIGINT) AS so_bigint,
                                    TRY_CAST([Material] AS BIGINT) AS mat_bigint,
                                    [Item Description],
                                    [Factory]
                                FROM dbo.so WITH(NOLOCK)
                            ) s ON k.[SO Mapping] = s.so_bigint AND k.[Material] = s.mat_bigint
                            WHERE (
                                k.[SO Mapping] IN ({in_clause})
                            )
                            ORDER BY k.[SO Mapping], k.[Order], coil_id
                        """)

                        c_rows = sconn.execute(q_kho, so_params).mappings().all()
                        for r in c_rows:
                            w_val = float(r['weight_kg'] or 0.0)
                            # Nếu cột Khối lượng lưu theo KG (ví dụ 24500) thì chia 1000 ra Tấn; nếu lưu Tấn (<100) thì nhân 1000 ra KG
                            w_kg = w_val * 1000.0 if (0 < w_val < 100) else w_val
                            w_tan = w_kg / 1000.0

                            order_str = str(r['order_no'] or '').strip()
                            vi_tri_str = str(r['vi_tri'] or '').strip()
                            status_parts = []
                            if order_str and order_str != '0':
                                status_parts.append(f"Order: {order_str}")
                            if vi_tri_str:
                                status_parts.append(f"Vị trí: {vi_tri_str}")
                            if r.get('kho_status'):
                                status_parts.append(str(r['kho_status']).strip())

                            mapped_coils.append({
                                "coil_id": str(r['coil_id'] or '').strip(),
                                "so_no": normalize_so_key(r['so_no']) or order_str,
                                "material": str(r['material_desc'] or '').strip(),
                                "thick": "-",
                                "width": "-",
                                "weight_kg": round(w_kg, 1),
                                "weight_tan": round(w_tan, 3),
                                "nha_may": str(r['nha_may'] or '').strip(),
                                "status": " | ".join(status_parts) if status_parts else "Đã Map Kho"
                            })

                    # Fallback sang vw_so_kho_sumary2 nếu bảng kho không còn tồn của SO này
                    if not mapped_coils:
                        q_sum = text(f"""
                            SELECT DISTINCT
                                TRY_CAST([SO Mapping] AS BIGINT) AS so_no,
                                [Material],
                                [Material Description] AS material,
                                ISNULL([SL Mapping kho], 0) AS map_kg,
                                NhaMay AS nha_may
                            FROM vw_so_kho_sumary2 WITH(NOLOCK)
                            WHERE TRY_CAST([SO Mapping] AS BIGINT) IN ({in_clause})
                        """)
                        s_rows = sconn.execute(q_sum, so_params).mappings().all()
                        for idx, r in enumerate(s_rows, start=1):
                            map_kg = float(r['map_kg'] or 0.0)
                            if map_kg > 0:
                                mapped_coils.append({
                                    "coil_id": f"LOT-{idx:02d} (Gom theo quy cách)",
                                    "so_no": normalize_so_key(r['so_no']),
                                    "material": str(r['material'] or '').strip(),
                                    "thick": "-",
                                    "width": "-",
                                    "weight_kg": round(map_kg, 1),
                                    "weight_tan": round(map_kg / 1000.0, 2),
                                    "nha_may": str(r['nha_may'] or '').strip(),
                                    "status": f"Map kho: {map_kg/1000:,.2f}T"
                                })
        except Exception as e:
            logger.error(f"[SQL Server] Lỗi truy vấn danh sách cuộn mapping từ bảng kho: {e}")

    total_exp_kg = sum(c['weight_kg'] for c in exported_coils)
    total_map_kg = sum(c['weight_kg'] for c in mapped_coils)
    return {
        "tau_name": clean_tau,
        "so_list": clean_sos,
        "exported_coils": exported_coils,
        "mapped_coils": mapped_coils,
        "total_exported_coils": len(exported_coils),
       "total_exported_tan": round(total_exp_kg / 1000.0, 2),
        "total_mapped_items": len(mapped_coils),
        "total_mapped_tan": round(total_map_kg / 1000.0, 2)
    }


@tauscheduler_bp.route('/api/coil-list', methods=['GET'])
@login_required
def api_get_coil_list():
    """API trả về danh sách COIL_LIST đã mapping & đã xuất lên tàu theo Tàu - SO."""
    tau_name = request.args.get('tau', '').strip()
    so_param = request.args.get('so', '').strip()
    voyage_group = request.args.get('vg', '').strip()
    ky_thang = request.args.get('ky', '').strip()

    so_list = [s.strip() for s in so_param.split(',') if s.strip()]

    # Nếu truyền vào voyage_group hoặc chỉ truyền tên tàu, tự động gom tất cả SO của tàu đó trong kỳ
    if not so_list and tau_name and ky_thang:
        with engine.connect() as conn:
            rows = conn.execute(text("""
                SELECT DISTINCT so_lenh_tach
                FROM tbl_ShipSchedule_KeHoach WITH(NOLOCK)
                WHERE ky_thang = :ky AND tau_name = :tau AND LEN(ISNULL(so_lenh_tach, '')) >= 8
            """), {"ky": ky_thang, "tau": normalize_ship_name(tau_name)}).fetchall()
            so_list = [r[0] for r in rows if r[0]]

    data = fetch_coil_list_by_ship_and_so(tau_name=tau_name, so_list=so_list)
    return jsonify({"success": True, **data})


@tauscheduler_bp.route('/export-coil-list', methods=['GET'])
@login_required
def export_coil_list_excel():
    """Xuất file Excel COIL_LIST (gồm 2 Sheet: Đã xuất lên tàu & Đã Mapping kho) theo Tàu - SO."""
    tau_name = request.args.get('tau', '').strip()
    so_param = request.args.get('so', '').strip()
    ky_thang = request.args.get('ky', '').strip()

    so_list = [s.strip() for s in so_param.split(',') if s.strip()]
    if not so_list and tau_name and ky_thang:
        with engine.connect() as conn:
            rows = conn.execute(text("""
                SELECT DISTINCT so_lenh_tach
                FROM tbl_ShipSchedule_KeHoach WITH(NOLOCK)
                WHERE ky_thang = :ky AND tau_name = :tau AND LEN(ISNULL(so_lenh_tach, '')) >= 8
            """), {"ky": ky_thang, "tau": normalize_ship_name(tau_name)}).fetchall()
            so_list = [r[0] for r in rows if r[0]]

    data = fetch_coil_list_by_ship_and_so(tau_name=tau_name, so_list=so_list)

    df_exported = pd.DataFrame(data['exported_coils'])
    if not df_exported.empty:
        df_exported.rename(columns={
            'coil_id': 'MÃ CUỘN (COIL ID)',
            'so_no': 'SỐ LỆNH (SO)',
            'tau_name': 'TÀU / PHƯƠNG TIỆN',
            'material': 'QUY CÁCH / VẬT TƯ',
            'weight_kg': 'KHỐI LƯỢNG (KG)',
            'weight_tan': 'KHỐI LƯỢNG (TẤN)',
            'issue_date': 'NGÀY XUẤT',
            'nha_may': 'NHÀ MÁY',
            'phieu_xuat': 'SỐ PHIẾU XUẤT'
        }, inplace=True)
    else:
        df_exported = pd.DataFrame(columns=[
            'MÃ CUỘN (COIL ID)', 'SỐ LỆNH (SO)', 'TÀU / PHƯƠNG TIỆN',
            'QUY CÁCH / VẬT TƯ', 'KHỐI LƯỢNG (KG)', 'KHỐI LƯỢNG (TẤN)', 'NGÀY XUẤT'
        ])

    df_mapped = pd.DataFrame(data['mapped_coils'])
    if not df_mapped.empty:
        df_mapped.rename(columns={
            'coil_id': 'MÃ CUỘN / LÔ MAP',
            'so_no': 'SỐ LỆNH (SO)',
            'material': 'MÔ TẢ VẬT TƯ (MATERIAL)',
            'thick': 'DÀY',
            'width': 'RỘNG',
            'weight_kg': 'KL MAPPING (KG)',
            'weight_tan': 'KL MAPPING (TẤN)',
            'nha_may': 'NHÀ MÁY',
            'status': 'TRẠNG THÁI KHO'
        }, inplace=True)
    else:
        df_mapped = pd.DataFrame(columns=[
            'MÃ CUỘN / LÔ MAP', 'SỐ LỆNH (SO)', 'MÔ TẢ VẬT TƯ (MATERIAL)',
            'KL MAPPING (KG)', 'KL MAPPING (TẤN)', 'NHÀ MÁY', 'TRẠNG THÁI KHO'
        ])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_exported.to_excel(writer, index=False, sheet_name='1_CUON_DA_XUAT_TAU')
        df_mapped.to_excel(writer, index=False, sheet_name='2_CUON_DA_MAP_KHO')

    output.seek(0)
    safe_tau = re.sub(r'[^A-Z0-9_-]', '_', normalize_ship_name(tau_name) or 'ALL')
    safe_so = re.sub(r'[^0-9_-]', '_', "_".join(data['so_list'][:2]) or 'ALL')
    fname = f"COIL_LIST_{safe_tau}_{safe_so}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
@tauscheduler_bp.route('/api/qc-coils/<so_lenh>', methods=['GET'])
@login_required
def api_get_qc_coils_by_so(so_lenh):
    """API trả về danh sách cuộn (ID) của 4 cột: NK Chưa Map, Chờ NK L1, Chờ Cơ Tính, CXL cho 1 mã SO."""
    so_key = normalize_so_key(so_lenh)
    if not so_key:
        return jsonify({"success": False, "message": "Mã SO không hợp lệ"}), 400

    qc_map = fetch_qc_wip_and_unmapped_by_so([so_key])
    so_data = qc_map.get(so_key, {})
    coils = so_data.get('pending_coils_detail', [])

    return jsonify({
        "success": True,
        "so_no": so_key,
        "coils": coils
    })
# ==============================================================================
# 10. API GHÉP TÀU - SO (STAGING DOCK MATCHING) & TÁCH KHỎI TÀU
# ==============================================================================

SNAPSHOT_COLS = [
    'ky_thang', 'stt', 'so_lenh_tach', 'tau_name', 'trong_tai', 'eta', 'etb', 'dai_ly',
    'thoi_gian_xong', 'ngay_duyet_so', 'cang_xep', 'cang_den', 'lsd',
    'lenh_xuat_hang', 'kl_xuat', 'san_xuat', 'nha_may', 'cw_max', 'ghi_chu',
    'nhip', 'nhip_sx', 'tinh_trang', 'pic', 'so_no', 'tong_da_map', 'da_xuat', 'con_lai'
]


@tauscheduler_bp.route('/api/match-vessel-so', methods=['POST'])
@login_required
def api_match_vessel_so():
    """
    Ghép 1 dòng SO chờ tàu (so_row_id) vào 1 dòng Chuyến tàu đích (vessel_row_id):
    - Tự động nhận diện nếu dòng tàu đích đang trống SO (WAIT_SO_ASSIGN) -> Gộp thẳng vào dòng tàu đích.
    - Nếu chuyến tàu đích đã có SO -> Chèn dòng SO vào ngay cuối Cụm chuyến tàu đó và kế thừa thông số tàu.
    - Hỗ trợ tách khối lượng (split_kl): Ví dụ SO có 3,500T nhưng chỉ ghép 2,000T lên tàu này, giữ lại 1,500T trong Giỏ chờ.
    """
    allowed, err_msg = check_action_permission("can_add")
    if not allowed:
        return jsonify({"success": False, "message": err_msg}), 403

    payload = request.get_json() or {}
    so_row_id = payload.get('so_row_id')
    vessel_row_id = payload.get('vessel_row_id')
    split_kl = parse_flexible_number(payload.get('split_kl'))

    if not so_row_id or not vessel_row_id or int(so_row_id) == int(vessel_row_id):
        return jsonify({"success": False, "message": "Dữ liệu dòng Tàu hoặc SO không hợp lệ"}), 400

    user_name = session.get('username', 'system')

    try:
        with engine.begin() as conn:
            so_row = conn.execute(
                text("SELECT * FROM tbl_ShipSchedule_KeHoach WHERE id = :id"),
                {"id": so_row_id}
            ).mappings().fetchone()
            vessel_row = conn.execute(
                text("SELECT * FROM tbl_ShipSchedule_KeHoach WHERE id = :id"),
                {"id": vessel_row_id}
            ).mappings().fetchone()

            if not so_row or not vessel_row:
                return jsonify({"success": False, "message": "Không tìm thấy dòng SO hoặc dòng Tàu trên hệ thống"}), 404

            ky_thang = vessel_row['ky_thang']
            sheet_year = get_year_from_ky(ky_thang)
            tau_name = normalize_ship_name(vessel_row['tau_name'])
            if not tau_name:
                return jsonify({"success": False, "message": "Dòng đích chưa có Tên tàu hợp lệ để ghép!"}), 400

            # Lấy toàn bộ dòng trong kỳ để xác định đúng Cụm chuyến tàu (voyage_group) của vessel_row
            # Lấy toàn bộ dòng trong kỳ để xác định đúng Cụm chuyến tàu (voyage_group) của vessel_row
            all_rows = conn.execute(text("""
                SELECT id, stt, tau_name, trong_tai, eta, etb, dai_ly, cang_xep, cang_den,
                       cw_max, nhip, nhip_sx, so_lenh_tach, lenh_xuat_hang, kl_xuat, da_xuat
                FROM tbl_ShipSchedule_KeHoach WITH(NOLOCK)
                WHERE ky_thang = :ky
                ORDER BY COALESCE(stt, 999999) ASC, id ASC
            """), {"ky": ky_thang}).mappings().all()
            all_rows_eval, _, _ = evaluate_schedule_rows([dict(r) for r in all_rows], sheet_year=sheet_year)

            target_vg = 0
            for r in all_rows_eval:
                if int(r['id']) == int(vessel_row_id):
                    target_vg = r.get('voyage_group', 0)
                    break

            voyage_rows = [r for r in all_rows_eval if target_vg > 0 and r.get('voyage_group') == target_vg]
            if not voyage_rows:
                voyage_rows = [dict(vessel_row)]

            # Thông số chuẩn của chuyến tàu đích
            v_trong_tai = next((float(r['trong_tai']) for r in voyage_rows if float(r.get('trong_tai') or 0) > 0), float(vessel_row['trong_tai'] or 0))
            v_eta = next((str(r['eta']).strip() for r in voyage_rows if str(r.get('eta') or '').strip()), str(vessel_row['eta'] or '').strip())
            v_eta_dt = parse_eta(v_eta, default_year=sheet_year)
            v_etb = str(vessel_row['etb'] or '').strip()
            v_daily = str(vessel_row['dai_ly'] or '').strip()
            v_cxep = str(vessel_row['cang_xep'] or 'Dung Quất').strip() or 'Dung Quất'
            v_nhip = infer_nhip(vessel_row['nhip'])
            v_cw_max = next((str(r.get('cw_max') or '').strip() for r in voyage_rows if str(r.get('cw_max') or '').strip()), str(vessel_row['cw_max'] or '').strip())
            so_nhip_sx = str(so_row.get('nhip_sx') or '').strip()
            last_stt_in_voyage = max(int(r.get('stt') or 1) for r in voyage_rows)

            # Kiểm tra dòng tàu đích có đang là "Vỏ tàu trống chờ SO" không
            v_so_clean, v_so_status, _ = analyze_and_extract_so_lenh(vessel_row['so_lenh_tach'], vessel_row['lenh_xuat_hang'])
            is_empty_vessel_slot = (v_so_status == 'WAIT_SO_ASSIGN' and float(vessel_row['kl_xuat'] or 0) == 0)

            orig_kl = float(so_row['kl_xuat'] or 0.0)
            assign_kl = orig_kl
            remain_kl = 0.0
            if split_kl > 0 and orig_kl > split_kl:
                assign_kl = round(split_kl, 2)
                remain_kl = round(orig_kl - assign_kl, 2)

            so_clean, _, _ = analyze_and_extract_so_lenh(so_row['so_lenh_tach'], so_row['lenh_xuat_hang'])
            so_no_val = so_clean if re.match(r'^\d{10}$', so_clean) else (so_row['so_no'] or '')
            nha_may_val = infer_nha_may(so_row['nha_may'] or vessel_row['nha_may'], so_row['san_xuat'], so_row['lenh_xuat_hang'])
            pic_val = infer_pic(so_row['pic'] or vessel_row['pic'], so_row['lenh_xuat_hang'])
            cang_den_val = str(so_row['cang_den'] or '').strip() or str(vessel_row['cang_den'] or '').strip()

            undo_updated = []
            undo_deleted = []
            undo_created = []
            target_highlight_id = int(vessel_row_id)

            if is_empty_vessel_slot:
                # KỊCH BẢN 1: Dòng tàu đích đang trống SO -> Đổ dữ liệu SO vào thẳng dòng vessel_row_id
                undo_updated.append({
                    "id": int(vessel_row_id),
                    "data": {col: vessel_row[col] for col in SNAPSHOT_COLS}
                })
                conn.execute(text("""
                    UPDATE tbl_ShipSchedule_KeHoach
                    SET so_lenh_tach = :so,
                        so_no = :so_no,
                        lenh_xuat_hang = :lxh,
                        kl_xuat = :kl,
                        con_lai = :kl,
                        cang_den = :cden,
                        lsd = :lsd,
                        ngay_duyet_so = :ngay_so,
                        san_xuat = :sx,
                        nha_may = :nm,
                        cw_max = :cw,
                        nhip = :nhip,
                        nhip_sx = :nhip_sx,
                        tinh_trang = :tt,
                        pic = :pic,
                        ghi_chu = :gc,
                        updated_by = :user,
                        updated_at = GETDATE()
                    WHERE id = :vid
                """), {
                    "so": so_clean,
                    "so_no": so_no_val,
                    "lxh": so_row['lenh_xuat_hang'] or '',
                    "kl": assign_kl,
                    "cden": cang_den_val,
                    "lsd": so_row['lsd'] or vessel_row['lsd'] or '',
                    "ngay_so": so_row['ngay_duyet_so'] or vessel_row['ngay_duyet_so'] or '',
                    "sx": so_row['san_xuat'] or vessel_row['san_xuat'] or '',
                    "nm": nha_may_val,
                    "cw": v_cw_max or so_row['cw_max'] or '',
                    "nhip": v_nhip,
                    "nhip_sx": so_nhip_sx or vessel_row.get('nhip_sx') or '',
                    "tt": so_row['tinh_trang'] or vessel_row['tinh_trang'] or '',
                    "pic": pic_val,
                    "gc": so_row['ghi_chu'] or vessel_row['ghi_chu'] or '',
                    "user": user_name,
                    "vid": vessel_row_id
                })

                if remain_kl > 0:
                    undo_updated.append({
                        "id": int(so_row_id),
                        "data": {"kl_xuat": orig_kl, "con_lai": float(so_row['con_lai'] or orig_kl)}
                    })
                    conn.execute(text("""
                        UPDATE tbl_ShipSchedule_KeHoach
                        SET kl_xuat = :rem_kl, con_lai = :rem_kl, updated_by = :user, updated_at = GETDATE()
                        WHERE id = :sid
                    """), {"rem_kl": remain_kl, "user": user_name, "sid": so_row_id})
                else:
                    undo_deleted.append({col: so_row[col] for col in SNAPSHOT_COLS})
                    conn.execute(text("DELETE FROM tbl_ShipSchedule_KeHoach WHERE id = :sid"), {"sid": so_row_id})

                target_highlight_id = int(vessel_row_id)

            else:
                # KỊCH BẢN 2: Chuyến tàu đích đã có SO -> Ghép nối tiếp vào cuối Cụm chuyến tàu đó
                new_stt = last_stt_in_voyage + 1
                conn.execute(text("""
                    UPDATE tbl_ShipSchedule_KeHoach
                    SET stt = stt + 1
                    WHERE ky_thang = :ky AND stt > :last_stt AND id != :sid
                """), {"ky": ky_thang, "last_stt": last_stt_in_voyage, "sid": so_row_id})

                if remain_kl > 0:
                    undo_updated.append({
                        "id": int(so_row_id),
                        "data": {"kl_xuat": orig_kl, "con_lai": float(so_row['con_lai'] or orig_kl)}
                    })
                    conn.execute(text("""
                        UPDATE tbl_ShipSchedule_KeHoach
                        SET kl_xuat = :rem_kl, con_lai = :rem_kl, updated_by = :user, updated_at = GETDATE()
                        WHERE id = :sid
                    """), {"rem_kl": remain_kl, "user": user_name, "sid": so_row_id})

                    res_ins = conn.execute(text("""
                        INSERT INTO tbl_ShipSchedule_KeHoach (
                            ky_thang, stt, tau_name, trong_tai, eta, eta_parsed, etb, dai_ly,
                            thoi_gian_xong, ngay_duyet_so, cang_xep, cang_den, lsd,
                            so_lenh_tach, so_no, lenh_xuat_hang, kl_xuat, con_lai,
                            san_xuat, nha_may, cw_max, nhip, nhip_sx, tinh_trang, pic, ghi_chu, updated_by
                        )
                        OUTPUT INSERTED.id
                        VALUES (
                            :ky, :stt, :tau, :tai, :eta, :eta_dt, :etb, :daily,
                            :tg_xong, :ngay_so, :cxep, :cden, :lsd,
                            :so, :so_no, :lxh, :kl, :kl,
                            :sx, :nm, :cw, :nhip, :nhip_sx, :tt, :pic, :gc, :user
                        )
                    """), {
                        "ky": ky_thang, "stt": new_stt, "tau": tau_name, "tai": v_trong_tai,
                        "eta": v_eta, "eta_dt": v_eta_dt, "etb": v_etb, "daily": v_daily,
                        "tg_xong": vessel_row['thoi_gian_xong'] or '', "ngay_so": so_row['ngay_duyet_so'] or '',
                        "cxep": v_cxep, "cden": cang_den_val, "lsd": so_row['lsd'] or '',
                        "so": so_clean, "so_no": so_no_val, "lxh": so_row['lenh_xuat_hang'] or '',
                        "kl": assign_kl, "sx": so_row['san_xuat'] or '', "nm": nha_may_val,
                        "cw": v_cw_max or so_row['cw_max'] or '', "nhip": v_nhip, "nhip_sx": so_nhip_sx,
                        "tt": so_row['tinh_trang'] or '', "pic": pic_val, "gc": so_row['ghi_chu'] or '', "user": user_name
                    })
                    new_id = res_ins.fetchone()[0]
                    undo_created.append(int(new_id))
                    target_highlight_id = int(new_id)
                else:
                    undo_updated.append({
                        "id": int(so_row_id),
                        "data": {col: so_row[col] for col in SNAPSHOT_COLS}
                    })
                    conn.execute(text("""
                        UPDATE tbl_ShipSchedule_KeHoach
                        SET ky_thang = :ky,
                            stt = :stt,
                            tau_name = :tau,
                            trong_tai = :tai,
                            eta = :eta,
                            eta_parsed = :eta_dt,
                            etb = :etb,
                            dai_ly = :daily,
                            cang_xep = :cxep,
                            cang_den = :cden,
                            cw_max = :cw,
                            nhip = :nhip,
                            nhip_sx = :nhip_sx,
                            nha_may = :nm,
                            pic = :pic,
                            so_lenh_tach = :so,
                            so_no = :so_no,
                            updated_by = :user,
                            updated_at = GETDATE()
                        WHERE id = :sid
                    """), {
                        "ky": ky_thang, "stt": new_stt, "tau": tau_name, "tai": v_trong_tai,
                        "eta": v_eta, "eta_dt": v_eta_dt, "etb": v_etb, "daily": v_daily,
                        "cxep": v_cxep, "cden": cang_den_val, "cw": v_cw_max or so_row['cw_max'] or '',
                        "nhip": v_nhip, "nhip_sx": so_nhip_sx,
                        "nm": nha_may_val, "pic": pic_val, "so": so_clean, "so_no": so_no_val,
                        "user": user_name, "sid": so_row_id
                    })
                    target_highlight_id = int(so_row_id)

        # Chạy đồng bộ Waterfall ngay cho kỳ để phân bổ lại kho & QC cho tàu vừa ghép
        execute_waterfall_sync_for_ky(ky_thang, user_name=user_name)

        # Kiểm tra nhanh các cảnh báo sau khi ghép để thông báo cho P.KD
        match_warnings = []
        lsd_dt = parse_eta(so_row['lsd'], default_year=sheet_year)
        if v_eta_dt and lsd_dt and v_eta_dt > lsd_dt:
            match_warnings.append(f"ETA tàu ({v_eta_dt.strftime('%d/%m')}) đang trễ hơn hạn LSD của SO ({lsd_dt.strftime('%d/%m')})!")

        sum_kl_after = sum(float(r.get('kl_xuat') or 0) for r in voyage_rows) + assign_kl
        if v_trong_tai > 0 and sum_kl_after > v_trong_tai:
            match_warnings.append(f"Tổng KL các lệnh trên tàu ({sum_kl_after:,.0f}T) vượt tổng tải tàu ({v_trong_tai:,.0f}T) +{sum_kl_after - v_trong_tai:,.0f}T!")

        cw_ton = parse_cw_max_ton(v_cw_max)
        if cw_ton > 0 and so_clean:
            coil_map = fetch_max_coil_weight_by_so([so_clean])
            c_info = coil_map.get(normalize_so_key(so_clean), {})
            max_c_tan = float(c_info.get('max_coil_tan') or 0.0)
            if max_c_tan > cw_ton:
                over_c = sum(1 for w in c_info.get('coil_weights', []) if w > cw_ton)
                match_warnings.append(
                    f"CẢNH BÁO C.W MAX: Tàu {tau_name} có C.W MAX = {cw_ton:g}T nhưng SO {so_clean} có {over_c} cuộn Map trong kho vượt tải (lớn nhất: {max_c_tan:g}T)!"
                )

        msg = f"Đã ghép SO {so_clean or ''} ({assign_kl:,.0f}T) vào tàu {tau_name}"
        if remain_kl > 0:
            msg += f" (Đã tách giữ lại {remain_kl:,.0f}T trong Giỏ chờ)."

        return jsonify({
            "success": True,
            "message": msg,
            "warnings": match_warnings,
            "highlight_row_id": target_highlight_id,
            "undo_action": {
                "type": "BATCH_RESTORE",
                "description": f"Ghép SO {so_clean or ''} vào tàu {tau_name}",
                "created_ids": undo_created,
                "updated_rows": undo_updated,
                "deleted_rows": undo_deleted
            }
        })
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi api_match_vessel_so: {e}")
        return jsonify({"success": False, "message": str(e)}), 500


@tauscheduler_bp.route('/api/unassign-vessel', methods=['POST'])
@login_required
def api_unassign_vessel():
    """
    Tách 1 dòng SO ra khỏi chuyến tàu hiện tại để đưa về 'Giỏ SO chờ ghép tàu' (xóa thông tin tàu trên dòng đó).
    Nếu dòng đó là dòng DUY NHẤT của chuyến tàu và người dùng chọn keep_vessel_slot=True -> Giữ lại 1 dòng vỏ tàu trống.
    """
    allowed, err_msg = check_action_permission("can_add")
    if not allowed:
        return jsonify({"success": False, "message": err_msg}), 403

    payload = request.get_json() or {}
    row_id = payload.get('row_id')
    keep_vessel_slot = bool(payload.get('keep_vessel_slot', True))

    if not row_id:
        return jsonify({"success": False, "message": "Thiếu ID dòng cần tách"}), 400

    user_name = session.get('username', 'system')

    try:
        with engine.begin() as conn:
            curr = conn.execute(
                text("SELECT * FROM tbl_ShipSchedule_KeHoach WHERE id = :id"),
                {"id": row_id}
            ).mappings().fetchone()
            if not curr:
                return jsonify({"success": False, "message": "Không tìm thấy bản ghi"}), 404

            ky_thang = curr['ky_thang']
            tau_old = curr['tau_name'] or ''
            so_old = curr['so_lenh_tach'] or ''

            # Kiểm tra xem trong cùng chuyến tàu còn dòng nào khác không
            all_rows = conn.execute(text("""
                SELECT id, stt, tau_name, eta FROM tbl_ShipSchedule_KeHoach WITH(NOLOCK)
                WHERE ky_thang = :ky ORDER BY COALESCE(stt, 999999) ASC, id ASC
            """), {"ky": ky_thang}).mappings().all()
            eval_rows, _, _ = evaluate_schedule_rows([dict(r) for r in all_rows])
            curr_vg = next((r.get('voyage_group', 0) for r in eval_rows if int(r['id']) == int(row_id)), 0)
            same_vg_count = sum(1 for r in eval_rows if curr_vg > 0 and r.get('voyage_group') == curr_vg)

            undo_updated = [{
                "id": int(row_id),
                "data": {col: curr[col] for col in SNAPSHOT_COLS}
            }]
            undo_created = []

            # Nếu đây là dòng duy nhất của chuyến tàu, tạo 1 dòng giữ chỗ vỏ tàu trống tại vị trí cũ
            if keep_vessel_slot and tau_old and same_vg_count == 1:
                res_v = conn.execute(text("""
                    INSERT INTO tbl_ShipSchedule_KeHoach (
                        ky_thang, stt, tau_name, trong_tai, eta, eta_parsed, etb, dai_ly,
                        thoi_gian_xong, cang_xep, cang_den, cw_max, nhip, so_lenh_tach, lenh_xuat_hang,
                        kl_xuat, con_lai, updated_by
                    )
                    OUTPUT INSERTED.id
                    VALUES (
                        :ky, :stt, :tau, :tai, :eta, :eta_dt, :etb, :daily,
                        :tg_xong, :cxep, :cden, :cw, :nhip, '', '', 0, 0, :user
                    )
                """), {
                    "ky": ky_thang, "stt": curr['stt'], "tau": tau_old,
                    "tai": float(curr['trong_tai'] or 0), "eta": curr['eta'] or '',
                    "eta_dt": curr['eta_parsed'], "etb": curr['etb'] or '',
                    "daily": curr['dai_ly'] or '', "tg_xong": curr['thoi_gian_xong'] or '',
                    "cxep": curr['cang_xep'] or 'Dung Quất', "cden": curr['cang_den'] or '',
                    "cw": curr['cw_max'] or '', "nhip": curr['nhip'] or '', "user": user_name
                })
                undo_created.append(int(res_v.fetchone()[0]))

            # Xóa thông tin tàu trên dòng SO và đưa xuống cuối bảng (vào Giỏ SO chờ ghép tàu)
            max_stt = conn.execute(
                text("SELECT COALESCE(MAX(stt), 0) + 1 FROM tbl_ShipSchedule_KeHoach WHERE ky_thang = :ky"),
                {"ky": ky_thang}
            ).scalar()

            conn.execute(text("""
                UPDATE tbl_ShipSchedule_KeHoach
                SET stt = :max_stt,
                    tau_name = '',
                    trong_tai = 0,
                    eta = '',
                    eta_parsed = NULL,
                    etb = '',
                    dai_ly = '',
                    thoi_gian_xong = '',
                    da_xuat = 0,
                    con_lai = kl_xuat,
                    updated_by = :user,
                    updated_at = GETDATE()
                WHERE id = :id
            """), {"max_stt": max_stt, "user": user_name, "id": row_id})

        execute_waterfall_sync_for_ky(ky_thang, user_name=user_name)

        return jsonify({
            "success": True,
            "message": f"Đã tách SO {so_old} khỏi tàu {tau_old} và đưa vào Giỏ SO chờ ghép tàu.",
            "highlight_row_id": int(row_id),
            "undo_action": {
                "type": "BATCH_RESTORE",
                "description": f"Tách SO {so_old} khỏi tàu {tau_old}",
                "created_ids": undo_created,
                "updated_rows": undo_updated,
                "deleted_rows": []
            }
        })
    except Exception as e:
        logger.error(f"[SQL Server] Lỗi api_unassign_vessel: {e}")
        return jsonify({"success": False, "message": str(e)}), 500