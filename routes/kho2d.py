from flask import Blueprint, jsonify, render_template, request
import pandas as pd
import re
import numpy as np
import io
import math
from db import engine
from flask import send_file

kho2d_bp = Blueprint('kho2d_bp', __name__, template_folder='templates', static_folder='static')

# =============================================================================================
# 1. CẤU HÌNH HỆ THỐNG
# =============================================================================================
ZONE_CONFIG = {
    # --- KHO HSPM ---
    "HA1": { "valid_lines": list(range(1, 10)), "lines": { "DEFAULT": {"cap": 48, "tiers": 3} } },
    "HA2": { "valid_lines": [91, 92, 93, 94] + list(range(1, 10)), "lines": { 91: {"cap": 23, "tiers": 2, "desc": "RAY 1"}, 92: {"cap": 23, "tiers": 2, "desc": "RAY 2"}, 93: {"cap": 23, "tiers": 2, "desc": "RAY 3"}, 94: {"cap": 23, "tiers": 2, "desc": "RAY 4"}, "DEFAULT": {"cap": 45, "tiers": 3} } },
    "HA3": { "valid_lines": list(range(1, 5)), "lines": { "DEFAULT": {"cap": 45, "tiers":2} } },
    "HB1": { "valid_lines": list(range(1, 10)), "lines": { "DEFAULT": {"cap": 48, "tiers": 3} } },
    "HB2": { "valid_lines": list(range(1, 10)), "lines": { "DEFAULT": {"cap": 45, "tiers": 3} } },
    "HB3": { "valid_lines": list(range(1, 10)), "lines": { "DEFAULT": {"cap": 39, "tiers": 3} } },
    "HC1": { "valid_lines": list(range(1, 11)), "lines": { "DEFAULT": {"cap": 48, "tiers": 3} } },
    "HC2": { "valid_lines": list(range(1, 11)), "lines": { "DEFAULT": {"cap": 45, "tiers": 3} } },
    "HC3": { "valid_lines": list(range(1, 11)), "lines": { "DEFAULT": {"cap": 39, "tiers": 3} } },
    "HD1": { "valid_lines": list(range(1, 11)), "lines": { "DEFAULT": {"cap": 48, "tiers": 3} } },
    "HD2": { "valid_lines": list(range(1, 11)), "lines": { "DEFAULT": {"cap": 45, "tiers": 3} } },
    "HD3": { "valid_lines": list(range(1, 11)), "lines": { "DEFAULT": {"cap": 39, "tiers": 3} } },
    # --- KHO HRC1 ---
    "KA1": { "valid_lines": list(range(1, 10)), "lines": { 1: {"cap": 4, "tiers": 3}, 2: {"cap": 14, "tiers": 3}, 3: {"cap": 15, "tiers": 3}, 4: {"cap": 18, "tiers": 3}, "DEFAULT": {"cap": 20, "tiers": 3} } },
    "KA2": { "valid_lines": [0], "lines": { 0: {"cap": 18, "tiers": 2, "desc": "RAY"} } },
    "KA3": { "valid_lines": [0] + list(range(1, 10)), "lines": { 0: {"cap": 30, "tiers": 1, "desc": "RAY"},  9: {"cap": 11, "tiers": 3}, "DEFAULT": {"cap": 23, "tiers": 3} } },
    "KB1": { "valid_lines": list(range(1, 10)), "lines": { 1: {"cap": 4, "tiers": 3}, 2: {"cap": 14, "tiers": 3}, 3: {"cap": 15, "tiers": 3}, 4: {"cap": 18, "tiers": 3}, "DEFAULT": {"cap": 20, "tiers": 3} } },
    "KB2": { "valid_lines": [0] + list(range(1, 10)), "lines": { 0: {"cap": 18, "tiers": 1, "desc": "RAY"}, "DEFAULT": {"cap": 21, "tiers": 3} } },
    "KB3": { "valid_lines": [0] + list(range(1, 10)), "lines": { 0: {"cap": 18, "tiers": 2, "desc": "RAY"}, 6: {"cap": 8, "tiers": 3}, 7: {"cap": 8, "tiers": 3}, 8: {"cap": 8, "tiers": 3}, 9: {"cap": 8, "tiers": 3}, "DEFAULT": {"cap": 20, "tiers": 3} } },
    "KC1": { "valid_lines": list(range(1, 10)), "lines": { 1: {"cap": 4, "tiers": 3}, 2: {"cap": 14, "tiers": 3}, 3: {"cap": 15, "tiers": 3}, 4: {"cap": 18, "tiers": 3}, "DEFAULT": {"cap": 20, "tiers": 3} } },
    "KC2": { "valid_lines": [0], "lines": { 0: {"cap": 18, "tiers": 2, "desc": "RAY"} } },
    "KC3": { "valid_lines": [0] + list(range(1, 10)), "lines": { 0: {"cap": 18, "tiers": 2, "desc": "RAY"}, "DEFAULT": {"cap": 20, "tiers": 3} } },
    # --- KHO HRC2 ---
    "NF": { "valid_lines": [99] + list(range(1, 10)), "lines": { 99: {"cap": 25, "tiers": 1, "desc": "RAY"}, 1: {"cap": 34, "tiers": 2}, 2: {"cap": 34, "tiers": 2}, "DEFAULT": {"cap": 46, "tiers": 2} } },
    "ND": { "valid_lines": [99] + list(range(1, 7)), "lines": { 99: {"cap": 50, "tiers": 1, "desc": "RAY"}, "DEFAULT": {"cap": 66, "tiers": 2} } },
    "NG": { "valid_lines": [99] + list(range(1, 10)), "lines": { 99: {"cap": 30, "tiers": 1, "desc": "RAY"}, 7: {"cap": 37, "tiers": 2}, 8: {"cap": 34, "tiers": 2}, 9: {"cap": 34, "tiers": 2}, "DEFAULT": {"cap": 46, "tiers": 2} } },
    "NA": { "valid_lines": list(range(1, 12)), "lines": { 1: {"cap": 48, "tiers": 2}, 2: {"cap": 48, "tiers": 2}, 3: {"cap": 48, "tiers": 2}, 4: {"cap": 51, "tiers": 2}, 5: {"cap": 51, "tiers": 2}, 6: {"cap": 45, "tiers": 2}, "DEFAULT": {"cap": 27, "tiers": 2} } },
    "NB": { "valid_lines": list(range(1, 12)), "lines": { "DEFAULT": {"cap": 33, "tiers": 2} } },
    "NC": { "valid_lines": list(range(1, 7)), "lines": { "DEFAULT": {"cap": 32, "tiers": 2} } },
    "NE": { "valid_lines": list(range(1, 8)), "lines": { "DEFAULT": {"cap": 23, "tiers": 2} } },
    
    # --- BÃI E (Baby Coils - Cho phép quá tải) ---
    "E01": { "valid_lines": list(range(2, 37)), "lines": { "DEFAULT": {"cap": 22, "tiers": 2} } },
    "E02": { "valid_lines": list(range(2, 36)), "lines": { "DEFAULT": {"cap": 22, "tiers": 2} } },
    # --- BÃI D (Cho phép quá tải) ---
    "D05": { "valid_lines": list(range(1, 34)), "lines": { "DEFAULT": {"cap": 23, "tiers": 2} } },
    "D06": { "valid_lines": list(range(1, 30)), "lines": { "DEFAULT": {"cap": 23, "tiers": 2} } },
    "D07": { "valid_lines": list(range(1, 30)), "lines": { "DEFAULT": {"cap": 23, "tiers": 2} } },
    
    "GLOBAL": {"cap": 30, "tiers": 3} 
}

SHIFT_MAPPING_HRC1 = {
    "AUTO_FILL": [("KA3", 1), ("KA3", 2), ("KA3", 3), ("KA3", 4),("KA3", 5), ("KA3", 6),("KA3", 7), ("KA3", 8),("KA3", 9), ("KB3", 1), ("KB3", 2), ("KB3", 3), ("KB3", 4),("KB3", 5), ("KB3", 6),("KB3", 7), ("KB3", 8),("KB3", 9), ("KC3", 1), ("KC3", 2), ("KC3", 3), ("KC3", 4),("KC3", 5), ("KC3", 6),("KC3", 7), ("KC3", 8),("KC3", 9)],
    "A": [("KA3", 1), ("KA3", 2), ("KA3", 7),("KB3", 1), ("KB3", 4),("KC3", 1),("KC3", 4),("KC3", 5)],
    "A1": [("KA3", 1), ("KA3", 2), ("KA3", 7),("KB3", 1), ("KB3", 4),("KC3", 1),("KC3", 4),("KC3", 5)],
    "A2": [("KA3", 1), ("KA3", 2), ("KA3", 7),("KB3", 1), ("KB3", 4),("KC3", 1),("KC3", 4),("KC3", 5)],
    "B": [("KA3", 3), ("KA3", 4), ("KA3", 8),("KB3", 2),("KB3", 5),("KC3", 2),("KC3", 6),("KC3", 7)],
    "B1": [("KA3", 3), ("KA3", 4), ("KA3", 8),("KB3", 2),("KB3", 5),("KC3", 2),("KC3", 6),("KC3", 7)],
    "B2": [("KA3", 3), ("KA3", 4), ("KA3", 8),("KB3", 2),("KB3", 5),("KC3", 2),("KC3", 6),("KC3", 7)],
    "C": [("KA3", 5), ("KA3", 6), ("KA3", 9),("KB3", 3),("KB3", 6),("KB3", 7),("KB3", 8),("KC3", 3),("KC3", 8),("KC3", 9)],
    "C1": [("KA3", 5), ("KA3", 6), ("KA3", 9),("KB3", 3),("KB3", 6),("KB3", 7),("KB3", 8),("KC3", 3),("KC3", 8),("KC3", 9)],
    "C2": [("KA3", 5), ("KA3", 6), ("KA3", 9),("KB3", 3),("KB3", 6),("KB3", 7),("KB3", 8),("KC3", 3),("KC3", 8),("KC3", 9)],
}

SHIFT_MAPPING_HRC2 = {
    "AUTO_FILL" : [("ND", 1), ("ND", 2), ("ND", 3), ("ND", 4),("ND", 5), ("ND", 6),("NC", 1), ("NC", 2),("NC", 3), ("NC", 4),("NC", 5), ("NC", 6),("NE", 1), ("NE", 2),("NE", 3), ("NE", 4),("NE", 5), ("NE", 6),("NG", 1), ("NG", 2),("NG", 3), ("NG", 4),("NG", 5), ("NG", 6),("NG", 7), ("NG", 8),("NG", 9),("NF", 1), ("NF", 2),("NF", 3), ("NF", 4),("NF", 5), ("NF", 6),("NF", 7), ("NF", 8),("NF", 9)],
    "A": [("ND", 3), ("ND", 6), ("NC", 2),("NC", 5), ("NE", 3),("NE", 6),("NG", 3),("NG", 6),("NG", 9),("NF", 3),("NF", 6),("NF", 9)],
    "A1": [("ND", 3), ("ND", 6), ("NC", 2),("NC", 5), ("NE", 3),("NE", 6),("NG", 3),("NG", 6),("NG", 9),("NF", 3),("NF", 6),("NF", 9)],
    "A2": [("ND", 3), ("ND", 6), ("NC", 2),("NC", 5), ("NE", 3),("NE", 6),("NG", 3),("NG", 6),("NG", 9),("NF", 3),("NF", 6),("NF", 9)],
    "B": [("ND", 2), ("ND", 5), ("NC", 1),("NC", 4), ("NE", 2),("NE", 5),("NG", 2),("NG", 5),("NG", 8),("NF", 2),("NF", 5),("NF", 8)],
    "B1": [("ND", 2), ("ND", 5), ("NC", 1),("NC", 4), ("NE", 2),("NE", 5),("NG", 2),("NG", 5),("NG", 8),("NF", 2),("NF", 5),("NF", 8)],
    "B2": [("ND", 2), ("ND", 5), ("NC", 1),("NC", 4), ("NE", 2),("NE", 5),("NG", 2),("NG", 5),("NG", 8),("NF", 2),("NF", 5),("NF", 8)],
    "C": [("ND", 1), ("ND", 4), ("NC", 3),("NC", 6), ("NE", 1),("NE", 4),("NE", 7),("NG", 1),("NG", 4),("NG", 7),("NF", 1),("NF", 4),("NF", 7)],
    "C1": [("ND", 1), ("ND", 4), ("NC", 3),("NC", 6), ("NE", 1),("NE", 4),("NE", 7),("NG", 1),("NG", 4),("NG", 7),("NF", 1),("NF", 4),("NF", 7)],
    "C2": [("ND", 1), ("ND", 4), ("NC", 3),("NC", 6), ("NE", 1),("NE", 4),("NE", 7),("NG", 1),("NG", 4),("NG", 7),("NF", 1),("NF", 4),("NF", 7)],
}
WAREHOUSE_LIMITS = {
    "HSPM": 326968.0,
    "HRC1": 79258.0,
    "HRC2": 107272.0,
    "BAIE": 68241.0, # Giả sử Bãi E và D riêng biệt, nếu chung thì cộng lại
    "BAID": 94185.0,
    "OTHER": 100000.0
}
# =============================================================================================
# 2. HELPER FUNCTIONS
# =============================================================================================

@kho2d_bp.route('/kho2d')
def kho2d():
    return render_template('kho2d.html')

def normalize_pos_key(suffix_str):
    if not suffix_str: return "UNKNOWN"
    s_clean = suffix_str.strip('.-_ ').upper()
    if re.match(r"^\d+$", s_clean):
        try: return str(int(s_clean))
        except: return s_clean
    match_x = re.match(r"^(\d+)X$", s_clean)
    if match_x:
        try: return f"{int(match_x.group(1))}X"
        except: return s_clean
    try:
        val = float(s_clean)
        if val.is_integer(): return str(int(val))
        return str(val)
    except:
        pass
    return s_clean

def clean_obj(obj):
    if isinstance(obj, (np.integer, np.int64)): return int(obj)
    elif isinstance(obj, (np.floating, np.float64)): return float(obj)
    elif isinstance(obj, np.ndarray): return obj.tolist()
    elif isinstance(obj, dict): return {str(k): clean_obj(v) for k, v in obj.items()}
    elif isinstance(obj, list): return [clean_obj(i) for i in obj]
    else: return obj

def get_zone_info(zone, line):
    z_conf = ZONE_CONFIG.get(zone, {})
    l_lines = z_conf.get("lines", {})
    info = l_lines.get(line, l_lines.get("DEFAULT", ZONE_CONFIG["GLOBAL"]))
    return { "cap": info.get("cap", 30), "tiers": info.get("tiers", 3), "desc": info.get("desc", str(line)) }

def calculate_max_slots(cap, tiers):
    total = cap
    if tiers >= 2: total += (cap - 1)
    if tiers >= 3: total += (cap - 2)
    return total

def validate_capacity(suffix_str, max_cap, max_tiers):
    if not suffix_str: return True 
    s_clean = suffix_str.strip('.-_ ').upper()
    try:
        if s_clean.endswith('X'):
            if max_tiers < 3: return False 
            match = re.search(r'\d+', s_clean)
            if not match: return False
            num_part = int(match.group())
            # Vị trí X cũng phải bắt đầu từ 1 (ví dụ 1X, 3X...)
            if num_part < 1: return False 
            
            real_index = (num_part - 1) / 2
            return real_index <= (max_cap - 2)
        else:
            match = re.match(r"^(\d+(\.\d+)?)", s_clean)
            if not match: return False
            pos_val = float(match.group(1))
            pos_num = int(pos_val)
            
            # [FIX QUAN TRỌNG]: Vị trí phải bắt đầu từ 1. Số 0 là không hợp lệ.
            if pos_num < 1: return False 

            if max_tiers == 1: return pos_num <= (max_cap * 2)
            
            if pos_num % 2 != 0: return ((pos_num + 1) / 2) <= max_cap
            else: 
                if max_tiers < 2: return False 
                return (pos_num / 2) <= (max_cap - 1)
    except: return False
    return True

def extract_suffix_only(raw_str):
    match = re.search(r'[\.\-_ ]+(\d+[X]?)$', raw_str.strip())
    if match: return match.group(1)
    return None

def parse_pos_flexible(pos_str):
    if not pos_str: return None, None, None, False
    pos_str = pos_str.upper().strip()
    match = re.match(r"^([HK][A-Z]\d)[^0-9]*(\d{1,2})(.*)$", pos_str)
    if match:
        try:
            zone = match.group(1); line = int(match.group(2)); raw_suffix = match.group(3); suffix = None
            if raw_suffix: clean = raw_suffix.strip('.-_ '); suffix = clean if clean else None
            return zone, line, suffix, True
        except: pass
    return None, None, None, False

def parse_pos_DE(raw_pos):
    if not raw_pos: return None, None, None, False
    clean_pos = re.sub(r'[\-_ ]+', '.', raw_pos.strip().upper())
    match = re.match(r"^([DE]0?\d+)(\.|)(\d+)(.*)$", clean_pos)
    if not match: return None, None, None, False
    raw_zone = match.group(1)
    z_match = re.match(r"^([DE])0?(\d+)$", raw_zone)
    if not z_match: return None, None, None, False
    zone = f"{z_match.group(1)}{int(z_match.group(2)):02d}"
    try: line = int(match.group(3))
    except: return None, None, None, False
    raw_rest = match.group(4)
    suffix = None
    if raw_rest:
        s_clean = raw_rest.strip('.')
        if re.match(r"^\d+$", s_clean): suffix = s_clean
        elif s_clean.endswith('X') and re.match(r"^\d+X$", s_clean): suffix = s_clean
    return zone, line, suffix, True

def get_warehouse_name(zone):
    if zone.startswith('H'): return "HSPM"
    if zone.startswith('K'): return "HRC1"
    if zone.startswith('N'): return "HRC2 (Kho Nóng)"
    if zone.startswith('E'): return "Bãi E"
    if zone.startswith('D'): return "Bãi D"
    return "Khác"

def generate_structure():
    structure = {}
    for zone, conf in ZONE_CONFIG.items():
        if zone == "GLOBAL": continue
        structure[zone] = []
        valid_lines = conf.get("valid_lines", [])
        for line in valid_lines:
            info = get_zone_info(zone, line)
            structure[zone].append({ "line": line, "desc": info['desc'], "max_capacity": info['cap'], "max_tiers": info['tiers'] })
    return structure
def calculate_priority(val):
    s = str(val).strip().upper()
    
    # Ưu tiên 0: RAY (Luôn xếp đầu tiên - Hàng đang cẩu)
    if "RAY" in s:
        return 0
    
    # Ưu tiên 1: Mã đích danh (VD: ND6, HA1, E01, D05...) - Hàng đã có chỗ cố định
    if re.match(r"^[HKNED][A-Z]*\d+", s):
        return 1
        
    if not s or s in ['NONE', 'NAN', 'NULL']:
        return 3

    # Ưu tiên 2: Mã Mapping chung (A, B, C, A1, B2...)
    return 2
# =============================================================================================
# 3. API CHÍNH
# =============================================================================================

@kho2d_bp.route('/api/data')
def get_data():
    if not engine: return jsonify({"status": "error", "message": "Lỗi kết nối DB"}), 500
    try:
        query = """select s.[ID Cuộn Bó], s.[Vị trí], s.[Khối lượng], s.Nhóm, 
        0 as [SO Mapping]
        from sanluong s
        LEFT JOIN kho k ON s.[ID Cuộn Bó] = k.[ID Cuộn Bó] -- Đặt tên tắt cho bảng kho là 'k'
        WHERE 
            s.[ID Cuộn Bó] IS NOT NULL AND s.[ID Cuộn Bó] <> ''
            AND (s.[Đã nhập kho] = 'No' OR s.[Đã nhập kho] IS NULL)
            AND k.[ID Cuộn Bó] IS NULL
        
        UNION ALL
        SELECT [ID Cuộn Bó], [Vị trí], [Khối lượng], Nhóm, [SO Mapping] FROM kho"""
        df = pd.read_sql(query, engine)
        df = df.fillna('') 
        df['pos_len'] = df['Vị trí'].astype(str).str.len()
        df['priority'] = df['Vị trí'].apply(calculate_priority)
        # df = df.sort_values(by='pos_len', ascending=True) 
        df = df.sort_values(by=['priority', 'pos_len'], ascending=[True, True])
        
        total_capacity_slots = 0
        warehouse_structure = generate_structure() 
        data_hspm = {}; data_hrc1 = {}; data_hrc2 = {}; data_baie = {}; data_baid = {} 
        errors_by_wh = { "ALL": [], "HSPM": [], "HRC1": [], "HRC2": [], "BAI_DE": [], "OTHER": [] }
        auto_list = [] 
        
        for zone, z_conf in ZONE_CONFIG.items():
            if zone == 'GLOBAL': continue
            for line in z_conf.get("valid_lines", []):
                info = get_zone_info(zone, line)
                slots = calculate_max_slots(info['cap'], info['tiers'])
                total_capacity_slots += slots   

        stats = { 
            "total": 0, "valid": 0, "invalid": 0, "auto_assigned": 0, 
            "hspm_count": 0, "hrc1_count": 0, "hrc2_count": 0, "baie_count": 0, "baid_count": 0,
            "total_capacity": total_capacity_slots, "total_weight": 0,
            
            # --- THÊM MỚI ---
            "limits": WAREHOUSE_LIMITS, # Gửi cấu hình limit xuống client
            
            # 1. Thống kê LỖI (Số lượng & Khối lượng)
            "err_count_by_wh":  {"HSPM": 0, "HRC1": 0, "HRC2": 0, "BAIE": 0, "BAID": 0, "OTHER": 0},
            "err_weight_by_wh": {"HSPM": 0.0, "HRC1": 0.0, "HRC2": 0.0, "BAIE": 0.0, "BAID": 0.0, "OTHER": 0.0},

            # 2. Thống kê HỢP LỆ (Khối lượng - để tính %)
            "valid_weight_by_wh": {"HSPM": 0.0, "HRC1": 0.0, "HRC2": 0.0, "BAIE": 0.0, "BAID": 0.0}
        }
        
        line_usage_counter = {}
        total_weight_kg = 0.0

        for _, row in df.iterrows():
            stats["total"] += 1
            coil_id = str(row['ID Cuộn Bó']).strip()
            raw_pos = str(row['Vị trí']).strip().upper()
            if raw_pos in ['NAN', 'NONE', 'NULL', 'nan', 'None']:
                raw_pos = ""
            raw_so = str(row['SO Mapping']).strip()
            so_val = raw_so if raw_so not in ['0', 'NONE', 'NAN', '', 'None'] else ""
            try: 
                w_val_float = float(row['Khối lượng']) if row['Khối lượng'] else 0.0
                total_weight_kg += w_val_float
                w_val = f"{w_val_float:,.0f}"
            except: w_val = "0"
            w_ton = w_val_float / 1000.0
            group_val = str(row['Nhóm'])
            if not coil_id: continue

            target_zone = None; target_line = None; target_suffix = None; target_repo_type = None

            if "RAY" in raw_pos:
                target_suffix = extract_suffix_only(raw_pos)
                if "HA2" in raw_pos: 
                    target_repo_type = 'HSPM'; target_zone = "HA2"; 
                    try: target_line = 91 + (int(coil_id[-1]) % 4)
                    except: target_line = 91
                elif any(z in raw_pos for z in ["KA2", "KA3", "KB2", "KB3", "KC2", "KC3"]):
                    target_repo_type = 'HRC1'
                    for z in ["KA2", "KA3", "KB2", "KB3", "KC2", "KC3"]: 
                        if z in raw_pos: target_zone = z; break
                    target_line = 0
                elif any(z in raw_pos for z in ["NF", "ND", "NG"]):
                    target_repo_type = 'HRC2'
                    for z in ["NF", "ND", "NG"]: 
                        if z in raw_pos: target_zone = z; break
                    target_line = 99
            elif raw_pos == "" or raw_pos in SHIFT_MAPPING_HRC1 or raw_pos in SHIFT_MAPPING_HRC2:
                target_suffix = None 
                target_list = []
                lookup_key = "AUTO_FILL" if raw_pos == "" else raw_pos
                if coil_id.startswith('8'): 
                    target_list = SHIFT_MAPPING_HRC2.get(lookup_key, []); target_repo_type = 'HRC2'
                else: 
                    target_list = SHIFT_MAPPING_HRC1.get(lookup_key, []); target_repo_type = 'HRC1'
                if isinstance(target_list, tuple): target_list = [target_list]
                chosen_zone, chosen_line = None, None
                if target_list:
                    for (cz, cl) in target_list:
                        c_info = get_zone_info(cz, cl)
                        max_limit = calculate_max_slots(c_info['cap'], c_info['tiers'])
                        curr = line_usage_counter.get((cz, cl), 0)
                        if curr < max_limit: chosen_zone, chosen_line = cz, cl; break 
                    if not chosen_zone: chosen_zone, chosen_line = target_list[-1]
                target_zone = chosen_zone; target_line = chosen_line
            elif raw_pos.startswith(('D', 'E')):
                z, l, s, v = parse_pos_DE(raw_pos)
                if v: target_zone = z; target_line = l; target_suffix = s; target_repo_type = 'BAI_DE' 
            elif raw_pos.startswith(('H', 'K')):
                normalized_pos = re.sub(r'[\.\-\_ ]', '', raw_pos)[:8]
                zone, line, suffix, is_valid = parse_pos_flexible(normalized_pos)
                if is_valid: 
                    target_zone = zone
                    target_line = line
                    target_suffix = suffix
                    target_repo_type = 'HSPM' if zone.startswith('H') else 'HRC1'
            elif raw_pos.startswith('N'):
                m = re.match(r"^([N][A-G])[\.\s\-_]*0*(\d+)(.*)$", raw_pos)
                if m: target_zone = m.group(1); target_line = int(m.group(2)); s = m.group(3); target_suffix = s.strip('.-_ ') if s else None; target_repo_type = 'HRC2'

            is_valid_config = False
            error_reason = ""
            if target_zone and target_line is not None and target_repo_type:
                z_conf = ZONE_CONFIG.get(target_zone)
                if z_conf and target_line in z_conf.get("valid_lines", []):
                    is_valid_config = True
                else: error_reason = f"Line {target_line} chưa cấu hình trong {target_zone}"

                if is_valid_config:
                    info = get_zone_info(target_zone, target_line)
                    max_slots = calculate_max_slots(info['cap'], info['tiers'])
                    current_count = line_usage_counter.get((target_zone, target_line), 0)
                    
                    is_bai_de = target_repo_type == 'BAI_DE' or (target_zone and target_zone.startswith(('E', 'D')))
                    is_fixed = (target_suffix is not None)

                    valid_to_add = True
                    normalized_key = None

                    if is_fixed: normalized_key = normalize_pos_key(target_suffix)
                    if is_fixed:
                        if not is_bai_de and not validate_capacity(target_suffix, info['cap'], info['tiers']):
                            is_fixed = False; target_suffix = None; normalized_key = None
                    
                    dest_dict = None; stats_key = ""
                    if target_repo_type == 'HSPM': dest_dict = data_hspm; stats_key = 'hspm_count'
                    elif target_repo_type == 'HRC1': dest_dict = data_hrc1; stats_key = 'hrc1_count'
                    elif target_repo_type == 'HRC2': dest_dict = data_hrc2; stats_key = 'hrc2_count'
                    elif target_repo_type == 'BAI_DE':
                        if target_zone.startswith('E'): dest_dict = data_baie; stats_key = 'baie_count'
                        elif target_zone.startswith('D'): dest_dict = data_baid; stats_key = 'baid_count'

                    if dest_dict is not None:
                        if target_zone not in dest_dict: dest_dict[target_zone] = {}
                        if target_line not in dest_dict[target_zone]: dest_dict[target_zone][target_line] = {"fixed": {}, "pending": []}
                        
                        cell_data = [coil_id, raw_pos, w_val, group_val, so_val]
                        added_as_pending = False

                        if normalized_key and is_fixed:
                            pos_key = normalized_key
                            if pos_key in dest_dict[target_zone][target_line]["fixed"]:
                                if is_bai_de:
                                    cell_data.append(True)
                                    dest_dict[target_zone][target_line]["pending"].append(cell_data)
                                    added_as_pending = True
                                else:
                                    existing_id = dest_dict[target_zone][target_line]["fixed"][pos_key][0]
                                    valid_to_add = False
                                    error_reason = f"Vị trí {target_suffix} bị trùng với {existing_id}"
                            else:
                                dest_dict[target_zone][target_line]["fixed"][pos_key] = cell_data
                        else:
                            cell_data.append(True)
                            dest_dict[target_zone][target_line]["pending"].append(cell_data)
                            added_as_pending = True

                        if valid_to_add:
                            # [FIX REALITY]: Nếu KHÔNG phải Auto (tức là có vị trí Fixed), cho phép tràn.
                            # Chỉ check Full nếu là hàng Auto (Shift) hoặc đã bị đẩy xuống Pending.
                            if not is_bai_de and not is_fixed:
                                if current_count >= max_slots:
                                    valid_to_add = False
                                    if added_as_pending: dest_dict[target_zone][target_line]["pending"].pop()
                                    error_reason = f"Dư thừa - Line đã đầy ({current_count}/{max_slots})"

                        if valid_to_add:
                            line_usage_counter[(target_zone, target_line)] = current_count + 1
                            if added_as_pending:
                                auto_list.append({"id": coil_id, "pos": raw_pos, "reason": f"Gợi ý {target_zone}-{info['desc']}", "so": so_val})
                                stats["auto_assigned"] += 1
                            if stats_key: stats[stats_key] += 1
                            stats["valid"] += 1
                            continue 

            stats["invalid"] += 1
            if not error_reason: error_reason = "Sai định dạng/Chưa Mapping"
            err_item = {"id": coil_id, "pos": raw_pos, "so": so_val, "reason": error_reason}
            errors_by_wh["ALL"].append(err_item)
            
            # [MODIFIED] Logic phân loại lỗi chính xác hơn cho BAI_DE (tách E và D)
            determined_repo = "OTHER"
            if target_repo_type: 
                if target_repo_type == 'BAI_DE': 
                    if target_zone and target_zone.startswith('D'): determined_repo = "BAID"
                    else: determined_repo = "BAIE" # Mặc định E nếu không rõ
                else: determined_repo = target_repo_type
            else: 
                if raw_pos.startswith('H'): determined_repo = "HSPM"
                elif raw_pos.startswith('K'): determined_repo = "HRC1"
                elif raw_pos.startswith('N'): determined_repo = "HRC2"
                elif raw_pos.startswith('D'): determined_repo = "BAID"
                elif raw_pos.startswith('E'): determined_repo = "BAIE"
                else: determined_repo = "OTHER"

            # Cộng khối lượng và số lượng LỖI
            if determined_repo in stats["err_weight_by_wh"]:
                stats["err_weight_by_wh"][determined_repo] += w_ton
                stats["err_count_by_wh"][determined_repo] += 1
            else:
                stats["err_weight_by_wh"]["OTHER"] += w_ton
                stats["err_count_by_wh"]["OTHER"] += 1
            
            # List chi tiết lỗi (để giữ tương thích code cũ, vẫn gộp BAI_DE vào chung 1 list)
            if determined_repo in ["BAIE", "BAID"]: errors_by_wh["BAI_DE"].append(err_item)
            elif determined_repo == "HSPM": errors_by_wh["HSPM"].append(err_item)
            elif determined_repo == "HRC1": errors_by_wh["HRC1"].append(err_item)
            elif determined_repo == "HRC2": errors_by_wh["HRC2"].append(err_item)
            else: errors_by_wh["OTHER"].append(err_item)

        stats["total_weight"] = total_weight_kg / 1000.0
        
        response_data = {
            "status": "success", "structure": warehouse_structure,
            "data_hspm": data_hspm, "data_hrc1": data_hrc1, "data_hrc2": data_hrc2, "data_baie": data_baie, "data_baid": data_baid, 
            "errors_by_wh": errors_by_wh, "auto_list": auto_list, "stats": stats
        }
        return jsonify(clean_obj(response_data))
    except Exception as e: return jsonify({"status": "error", "message": str(e)}), 500
def simulate_warehouse_state():
    df = pd.read_sql("""select s.[ID Cuộn Bó], s.[Vị trí]
                from sanluong s
                LEFT JOIN kho k ON s.[ID Cuộn Bó] = k.[ID Cuộn Bó] -- Đặt tên tắt cho bảng kho là 'k'
                WHERE 
                    s.[ID Cuộn Bó] IS NOT NULL AND s.[ID Cuộn Bó] <> ''
                    AND (s.[Đã nhập kho] = 'No' OR s.[Đã nhập kho] IS NULL)
                    AND k.[ID Cuộn Bó] IS NULL

                UNION ALL
                SELECT [ID Cuộn Bó], [Vị trí] FROM kho""", engine)
    df = df.fillna('')
    df['pos_len'] = df['Vị trí'].astype(str).str.len()
    df['priority'] = df['Vị trí'].apply(calculate_priority)
        # df = df.sort_values(by='pos_len', ascending=True) 
    df = df.sort_values(by=['priority', 'pos_len'], ascending=[True, True])

    warehouse_state = {}
    used_positions_map = {} 
    line_usage_counter = {} 

    for zone, z_conf in ZONE_CONFIG.items():
        if zone == 'GLOBAL': continue
        used_positions_map[zone] = {}
        for line in z_conf.get("valid_lines", []):
            warehouse_state[(zone, line)] = { 
                "occupied_indices": {1: set(), 2: set(), 3: set()}, 
                "pending_count": 0, 
                "pending_items": [] 
            }
            used_positions_map[zone][line] = set()
            line_usage_counter[(zone, line)] = 0

    for _, row in df.iterrows():
        coil_id = str(row['ID Cuộn Bó']).strip()
        raw_pos = str(row['Vị trí']).strip().upper()
        if raw_pos in ['NAN', 'NONE', 'NULL', 'nan', 'None']:
                raw_pos = ""
        if not coil_id: continue
        
        target_zone = None; target_line = None; target_suffix = None; target_repo_type = None
        
        if "RAY" in raw_pos:
            target_suffix = extract_suffix_only(raw_pos)
            if "HA2" in raw_pos: 
                target_zone = "HA2"; target_repo_type='HSPM'
                try: target_line = 91 + (int(coil_id[-1]) % 4)
                except: target_line = 91
            elif any(z in raw_pos for z in ["KA2", "KA3", "KB2", "KB3", "KC2", "KC3"]):
                target_repo_type='HRC1'
                for z in ["KA2", "KA3", "KB2", "KB3", "KC2", "KC3"]: 
                    if z in raw_pos: target_zone = z; break
                target_line = 0
            elif any(z in raw_pos for z in ["NF", "ND", "NG"]):
                target_repo_type='HRC2'
                for z in ["NF", "ND", "NG"]:
                    if z in raw_pos: target_zone = z; break
                target_line = 99
        elif raw_pos == "" or raw_pos in SHIFT_MAPPING_HRC1 or raw_pos in SHIFT_MAPPING_HRC2:
            target_suffix = None 
            target_list = []
            lookup_key = "AUTO_FILL" if raw_pos == "" else raw_pos
            if coil_id.startswith('8'): 
                target_list = SHIFT_MAPPING_HRC2.get(lookup_key, []); target_repo_type = 'HRC2'
            else: 
                target_list = SHIFT_MAPPING_HRC1.get(lookup_key, []); target_repo_type = 'HRC1'
            if isinstance(target_list, tuple): target_list = [target_list]
            chosen_zone, chosen_line = None, None
            if target_list:
                for (cz, cl) in target_list:
                    c_info = get_zone_info(cz, cl)
                    max_limit = calculate_max_slots(c_info['cap'], c_info['tiers'])
                    curr = line_usage_counter.get((cz, cl), 0)
                    if curr < max_limit: chosen_zone, chosen_line = cz, cl; break 
                if not chosen_zone: chosen_zone, chosen_line = target_list[-1]
            target_zone = chosen_zone; target_line = chosen_line
        elif raw_pos.startswith(('D', 'E')):
            z, l, s, v = parse_pos_DE(raw_pos)
            if v: target_zone = z; target_line = l; target_suffix = s; target_repo_type='BAI_DE'
        elif raw_pos.startswith(('H', 'K')):
            normalized_pos = re.sub(r'[\.\-\_ ]', '', raw_pos)[:8]
            zone, line, suffix, is_valid = parse_pos_flexible(normalized_pos)
            if is_valid: 
                target_zone = zone
                target_line = line
                target_suffix = suffix
                target_repo_type = 'HSPM' if zone.startswith('H') else 'HRC1'
        elif raw_pos.startswith('N'):
            m = re.match(r"^([N][A-G])[\.\s\-_]*0*(\d+)(.*)$", raw_pos)
            if m: target_zone = m.group(1); target_line = int(m.group(2)); s = m.group(3); target_suffix = s.strip('.-_ ') if s else None; target_repo_type='HRC2'

        if target_zone and target_line is not None:
            state_item = warehouse_state.get((target_zone, target_line))
            if state_item:
                info = get_zone_info(target_zone, target_line)
                is_bai_de = target_zone.startswith(('E', 'D'))
                max_slots = calculate_max_slots(info['cap'], info['tiers'])
                current_count = line_usage_counter.get((target_zone, target_line), 0)

                valid_add = True
                is_duplicate = False
                is_fixed = (target_suffix is not None)
                normalized_key = None

                if is_fixed:
                    normalized_key = normalize_pos_key(target_suffix)
                    if not is_bai_de and not validate_capacity(target_suffix, info['cap'], info['tiers']):
                        is_fixed = False; target_suffix = None; normalized_key = None
                    if is_fixed: 
                        pos_key = normalized_key
                        if pos_key in used_positions_map[target_zone][target_line]:
                            if is_bai_de: is_duplicate = True 
                            else: valid_add = False 
                        else:
                            used_positions_map[target_zone][target_line].add(pos_key)
                
                # [FIX REALITY]: Check Capacity chỉ áp dụng cho Auto/Pending
                if valid_add:
                    if not is_bai_de and not is_fixed:
                        if current_count >= max_slots:
                            valid_add = False 

                if valid_add:
                    line_usage_counter[(target_zone, target_line)] += 1
                    if is_fixed and not is_duplicate:
                        s_clean = normalized_key 
                        match_x = re.search(r'(\d+)X', s_clean)
                        if match_x: 
                             state_item["occupied_indices"][3].add(int(match_x.group(1)))
                        else:
                            try:
                                num = int(float(s_clean))
                                if info['tiers'] == 1: state_item["occupied_indices"][1].add(num)
                                else:
                                    if num % 2 != 0: state_item["occupied_indices"][1].add(num)
                                    else: state_item["occupied_indices"][2].add(num)
                            except: pass
                    else:
                        state_item["pending_count"] += 1
                        state_item["pending_items"].append(coil_id)

    return warehouse_state

def calculate_layout_allocation(zone, line, data, info):
    max_tiers = info['tiers']
    config_max = info['cap']
    real_max_t1 = config_max
    
    # Dữ liệu đầu vào
    occupied = data["occupied_indices"]
    pending_queue = list(data["pending_items"]) # Copy danh sách chờ
    total_pending_initial = list(pending_queue) # Lưu bản gốc để so sánh
    
    virtual_map = {1: {}, 2: {}, 3: {}}
    max_fixed_pos = 0
    
    # 1. Load hàng Fixed
    for t in [1, 2, 3]:
        for idx in occupied[t]:
            virtual_map[t][idx] = "FIXED"
            if idx > max_fixed_pos: max_fixed_pos = idx
            
    step = 1 if max_tiers == 1 else 2
    req_len = math.ceil(max_fixed_pos / step) if step > 0 else 0
    if req_len > real_max_t1: real_max_t1 = req_len
    
    is_bai_de = zone.startswith(('E', 'D'))
    allocated_items = []
    allocated_ids = [] # Danh sách ID đã xếp thành công
    
    # VÁ CHÂN ĐẾ
    if pending_queue and max_tiers >= 2:
        t2_positions = sorted(list(virtual_map[2].keys()))
        for t2_loc in t2_positions:
            for leg in [t2_loc-1, t2_loc+1]:
                if leg > 0 and leg not in virtual_map[1] and pending_queue:
                    pid = pending_queue.pop(0)
                    virtual_map[1][leg] = pid
                    allocated_items.append({"tier": 1, "id": leg, "val": pid, "type": "suggest"})
                    allocated_ids.append(pid)

    # VÒNG LẶP CHÍNH
    k = 0
    while True:
        if not is_bai_de and k >= real_max_t1: break
        if is_bai_de and k >= real_max_t1 and not pending_queue: break
        if k > 500: break 
        
        # T1
        t1_num = 1 + k * step
        if t1_num not in virtual_map[1] and pending_queue:
            pid = pending_queue.pop(0)
            virtual_map[1][t1_num] = pid
            allocated_items.append({"tier": 1, "id": t1_num, "val": pid, "type": "suggest"})
            allocated_ids.append(pid)
            
        if t1_num in virtual_map[1]:
            if k + 1 > real_max_t1: real_max_t1 = k + 1
            
        # T2
        if max_tiers >= 2 and k > 0:
            t2_num = 2 + (k - 1) * step
            if t2_num not in virtual_map[2] and pending_queue:
                if (t2_num-1) in virtual_map[1] and (t2_num+1) in virtual_map[1]:
                    pid = pending_queue.pop(0)
                    virtual_map[2][t2_num] = pid
                    allocated_items.append({"tier": 2, "id": t2_num, "val": pid, "type": "suggest"})
                    allocated_ids.append(pid)
                    
        # T3
        if max_tiers >= 3 and k > 1:
            t3_num = 3 + (k - 2) * step
            if t3_num not in virtual_map[3] and pending_queue:
                if (t3_num-1) in virtual_map[2] and (t3_num+1) in virtual_map[2]:
                    pid = pending_queue.pop(0)
                    virtual_map[3][t3_num] = pid
                    allocated_items.append({"tier": 3, "id": t3_num, "val": pid, "type": "suggest"})
                    allocated_ids.append(pid)
        k += 1
        
    # 3. TÍNH TOÁN CUỘN THỪA (UNPLACED)
    # Logic: Lấy tổng pending ban đầu trừ đi những cái đã xếp được
    unplaced_items = []
    # Lưu ý: ID trong pending có thể trùng nhau (Bãi E), nên phải remove từng cái một
    temp_allocated = list(allocated_ids)
    for item in total_pending_initial:
        if item in temp_allocated:
            temp_allocated.remove(item)
        else:
            unplaced_items.append(item)
            
    return allocated_items, real_max_t1, unplaced_items

@kho2d_bp.route('/api/stats/capacity')
def get_capacity_stats():
    try: wh_state = simulate_warehouse_state()
    except Exception as e: return jsonify({"status": "error", "message": str(e)}), 500
    
    stats_data = []
    for (zone, line), data in wh_state.items():
        info = get_zone_info(zone, line)
        
        # 1. Chạy xếp hình
        allocated, final_len, unplaced_items = calculate_layout_allocation(zone, line, data, info)
        
        # 2. Đếm số lượng đã xếp theo Geometry (Vật lý)
        count_t1 = sum(1 for x in allocated if x['tier'] == 1)
        count_t2 = sum(1 for x in allocated if x['tier'] == 2)
        count_t3 = sum(1 for x in allocated if x['tier'] == 3)
        
        fixed_t1 = len(data["occupied_indices"][1])
        fixed_t2 = len(data["occupied_indices"][2])
        fixed_t3 = len(data["occupied_indices"][3])
        
        suggest_map = { (i['tier'], i['id']): i['val'] for i in allocated }
        
        # 3. [QUAN TRỌNG] MÔ PHỎNG LẤP ĐẦY ĐỂ TÍNH CHÍNH XÁC CUỘN THỪA RƠI VÀO TẦNG NÀO
        # (Logic này copy y hệt từ export_capacity_excel)
        
        fill_t1, fill_t2, fill_t3 = 0, 0, 0
        rem_unplaced = len(unplaced_items) # Số lượng cần lấp
        
        step = 1 if info['tiers'] == 1 else 2
        
        # --- Mô phỏng lấp T1 ---
        for k in range(final_len):
            if rem_unplaced <= 0: break
            t1_id = 1 + k * step
            # Nếu vị trí này chưa có Fixed và chưa có Suggest -> Nó là Trống -> Lấp được
            if t1_id not in data["occupied_indices"][1] and (1, t1_id) not in suggest_map:
                fill_t1 += 1
                rem_unplaced -= 1
        
        # --- Mô phỏng lấp T2 ---
        if info['tiers'] >= 2 and rem_unplaced > 0:
            count_loop_t2 = final_len - 1
            for k in range(count_loop_t2):
                if rem_unplaced <= 0: break
                t2_id = 2 + k * step
                if t2_id not in data["occupied_indices"][2] and (2, t2_id) not in suggest_map:
                    fill_t2 += 1
                    rem_unplaced -= 1
                    
        # --- Mô phỏng lấp T3 ---
        if info['tiers'] >= 3 and rem_unplaced > 0:
            count_loop_t3 = final_len - 2
            for k in range(count_loop_t3):
                if rem_unplaced <= 0: break
                t3_id = 3 + k * step
                if t3_id not in data["occupied_indices"][3] and (3, t3_id) not in suggest_map:
                    fill_t3 += 1
                    rem_unplaced -= 1
        
        # Nếu vẫn còn dư (Tràn kho) -> Cộng vào T1 để báo động
        if rem_unplaced > 0:
            fill_t1 += rem_unplaced

        # 4. Tổng hợp số liệu
        limit_t1 = info['cap']
        limit_t2 = info['cap'] - 1 if info['tiers'] >= 2 else 0
        limit_t3 = info['cap'] - 2 if info['tiers'] >= 3 else 0
        
        # Total Used = Fixed + Allocated + Filled
        total_used_t1 = fixed_t1 + count_t1 + fill_t1
        total_used_t2 = fixed_t2 + count_t2 + fill_t2
        total_used_t3 = fixed_t3 + count_t3 + fill_t3
        
        stats_data.append({
            "warehouse": get_warehouse_name(zone), "zone": zone, "line": line, "desc": info['desc'], "tiers": info['tiers'],
            "stats": {
                "t1": { 
                    "max": limit_t1, 
                    "total_used": total_used_t1, 
                    "suggested": count_t1 + fill_t1, # Gợi ý = Xếp hình + Lấp chỗ
                    "final_empty": max(0, limit_t1 - total_used_t1) 
                },
                "t2": { "max": limit_t2, "total_used": total_used_t2, "suggested": count_t2 + fill_t2, "final_empty": max(0, limit_t2 - total_used_t2) },
                "t3": { "max": limit_t3, "total_used": total_used_t3, "suggested": count_t3 + fill_t3, "final_empty": max(0, limit_t3 - total_used_t3) }
            }
        })
    return jsonify({"status": "success", "data": stats_data})
@kho2d_bp.route('/export/capacity')
def export_capacity_excel():
    try: wh_state = simulate_warehouse_state()
    except Exception as e: return str(e), 500
    export_rows = []
    sorted_keys = sorted(wh_state.keys())
    
    for (zone, line) in sorted_keys:
        data = wh_state[(zone, line)]; info = get_zone_info(zone, line)
        wh_name = get_warehouse_name(zone)
        
        # Gọi hàm xếp hình mới
        allocated, final_len, unplaced_items = calculate_layout_allocation(zone, line, data, info)
        
        suggest_map = { (i['tier'], i['id']): i['val'] for i in allocated }
        
        # Hàng đợi để lấp vào chỗ trống
        unplaced_queue = list(unplaced_items)
        
        # Hàm hỗ trợ lấy 1 cuộn từ hàng đợi
        def fill_gap():
            if len(unplaced_queue) > 0:
                return "Gợi ý (Lỗi chân)", unplaced_queue.pop(0)
            return "Trống", ""

        # --- VẼ TẦNG 1 ---
        step = 1 if info['tiers'] == 1 else 2
        for k in range(final_len):
            t1_id = 1 + k * step
            status = "Trống"; val = ""
            
            if t1_id in data["occupied_indices"][1]: status = "Đúng vị trí"
            elif (1, t1_id) in suggest_map: status = "Gợi ý"; val = suggest_map[(1, t1_id)]
            else: status, val = fill_gap() # Lấp chỗ trống
            
            if status == "Đúng vị trí": continue 
            if "Gợi ý" in status or k < info['cap']:
                export_rows.append({ "Kho Tổng": wh_name, "Zone": zone, "Line": f"{line}", "Tầng": 1, "Vị trí": f"{zone}.{line:02d}{t1_id:03d}", "Trạng thái": status, "ID Gợi ý": val })
        
        # --- VẼ TẦNG 2 ---
        if info['tiers'] >= 2:
            count_t2 = final_len - 1
            for k in range(count_t2):
                t2_id = 2 + k * step
                status = "Trống"; val = ""
                if t2_id in data["occupied_indices"][2]: status = "Đúng vị trí"
                elif (2, t2_id) in suggest_map: status = "Gợi ý"; val = suggest_map[(2, t2_id)]
                else: status, val = fill_gap() # Lấp chỗ trống
                
                if status == "Đúng vị trí": continue
                if "Gợi ý" in status or k < (info['cap'] - 1):
                    export_rows.append({ "Kho Tổng": wh_name, "Zone": zone, "Line": f"{line}", "Tầng": 2, "Vị trí": f"{zone}.{line:02d}{t2_id:03d}", "Trạng thái": status, "ID Gợi ý": val })

        # --- VẼ TẦNG 3 ---
        if info['tiers'] >= 3:
            count_t3 = final_len - 2
            for k in range(count_t3):
                t3_id = 3 + k * step
                status = "Trống"; val = ""
                if t3_id in data["occupied_indices"][3]: status = "Đúng vị trí"
                elif (3, t3_id) in suggest_map: status = "Gợi ý"; val = suggest_map[(3, t3_id)]
                else: status, val = fill_gap() # Lấp chỗ trống
                
                if status == "Đúng vị trí": continue
                if "Gợi ý" in status or k < (info['cap'] - 2):
                    export_rows.append({ "Kho Tổng": wh_name, "Zone": zone, "Line": f"{line}", "Tầng": 3, "Vị trí": f"{zone}.{line:02d}{t3_id:03d}X", "Trạng thái": status, "ID Gợi ý": val })

        # --- NẾU VẪN CÒN DƯ (KHO QUÁ TẢI) ---
        while len(unplaced_queue) > 0:
            rem_id = unplaced_queue.pop(0)
            export_rows.append({ 
                "Kho Tổng": wh_name, "Zone": zone, "Line": f"{line}", "Tầng": 1, 
                "Vị trí": "Tràn kho (Overflow)", "Trạng thái": "Gợi ý (Dư thừa)", "ID Gợi ý": rem_id 
            })

    df_export = pd.DataFrame(export_rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, sheet_name='ViTriTrong', index=False)
        workbook = writer.book; worksheet = writer.sheets['ViTriTrong']
        red_fmt = workbook.add_format({'font_color': '#d63384', 'bold': True})
        worksheet.set_column('G:G', 20, red_fmt)
    output.seek(0)
    return send_file(output, download_name="BaoCao_ViTri_GoiY.xlsx", as_attachment=True)

@kho2d_bp.route('/export/warehouse_list')
def export_warehouse_list():
    if not engine: return "Lỗi kết nối DB", 500
    try:
        repo_request = request.args.get('repo', 'ALL') 
        query = """select s.[ID Cuộn Bó], s.[Vị trí], s.[Khối lượng], s.Nhóm, 
                    0 as [SO Mapping]
                    from sanluong s
                    LEFT JOIN kho k ON s.[ID Cuộn Bó] = k.[ID Cuộn Bó] -- Đặt tên tắt cho bảng kho là 'k'
                    WHERE 
                        s.[ID Cuộn Bó] IS NOT NULL AND s.[ID Cuộn Bó] <> ''
                        AND (s.[Đã nhập kho] = 'No' OR s.[Đã nhập kho] IS NULL)
                        AND k.[ID Cuộn Bó] IS NULL
 
                    UNION ALL
                    SELECT [ID Cuộn Bó], [Vị trí], [Khối lượng], Nhóm, [SO Mapping] FROM kho"""
        df = pd.read_sql(query, engine)
        df = df.fillna('')
        
        df['pos_len'] = df['Vị trí'].astype(str).str.len()
        df['priority'] = df['Vị trí'].apply(calculate_priority)
        # df = df.sort_values(by='pos_len', ascending=True) 
        df = df.sort_values(by=['priority', 'pos_len'], ascending=[True, True])

        data_hspm = {}; data_hrc1 = {}; data_hrc2 = {}; data_baie = {}; data_baid = {}
        line_usage_counter = {} 
        final_export_rows = []

        for _, row in df.iterrows():
            coil_id = str(row['ID Cuộn Bó']).strip()
            raw_pos = str(row['Vị trí']).strip().upper()
            if raw_pos in ['NAN', 'NONE', 'NULL', 'nan', 'None']:
                raw_pos = ""
            raw_so = str(row['SO Mapping']).strip()
            so_val = raw_so if raw_so not in ['0', 'NONE', 'NAN', '', 'None'] else ""
            try: w_val = float(row['Khối lượng']) if row['Khối lượng'] else 0
            except: w_val = 0
            group_val = str(row['Nhóm'])

            if not coil_id: continue

            # --- PARSE VỊ TRÍ ---
            target_zone = None; target_line = None; target_suffix = None; target_repo_type = None

            if "RAY" in raw_pos:
                target_suffix = extract_suffix_only(raw_pos)
                if "HA2" in raw_pos: 
                    target_repo_type = 'HSPM'; target_zone = "HA2"; 
                    try: target_line = 91 + (int(coil_id[-1]) % 4)
                    except: target_line = 91
                elif any(z in raw_pos for z in ["KA2", "KA3", "KB2", "KB3", "KC2", "KC3"]):
                    target_repo_type = 'HRC1'
                    for z in ["KA2", "KA3", "KB2", "KB3", "KC2", "KC3"]: 
                        if z in raw_pos: target_zone = z; break
                    target_line = 0
                elif any(z in raw_pos for z in ["NF", "ND", "NG"]):
                    target_repo_type = 'HRC2'
                    for z in ["NF", "ND", "NG"]: 
                        if z in raw_pos: target_zone = z; break
                    target_line = 99
            elif raw_pos == "" or raw_pos in SHIFT_MAPPING_HRC1 or raw_pos in SHIFT_MAPPING_HRC2:
                target_suffix = None 
                target_list = []
                lookup_key = "AUTO_FILL" if raw_pos == "" else raw_pos
                if coil_id.startswith('8'): 
                    target_list = SHIFT_MAPPING_HRC2.get(lookup_key, []); target_repo_type = 'HRC2'
                else: 
                    target_list = SHIFT_MAPPING_HRC1.get(lookup_key, []); target_repo_type = 'HRC1'
                if isinstance(target_list, tuple): target_list = [target_list]
                chosen_zone, chosen_line = None, None
                if target_list:
                    for (cz, cl) in target_list:
                        c_info = get_zone_info(cz, cl)
                        max_limit = calculate_max_slots(c_info['cap'], c_info['tiers'])
                        curr = line_usage_counter.get((cz, cl), 0)
                        if curr < max_limit: chosen_zone, chosen_line = cz, cl; break 
                    if not chosen_zone: chosen_zone, chosen_line = target_list[-1]
                target_zone = chosen_zone; target_line = chosen_line
            elif raw_pos.startswith(('D', 'E')):
                z, l, s, v = parse_pos_DE(raw_pos)
                if v: target_zone = z; target_line = l; target_suffix = s; target_repo_type = 'BAI_DE' 
            elif raw_pos.startswith(('H', 'K')):
                normalized_pos = re.sub(r'[\.\-\_ ]', '', raw_pos)[:8]
                zone, line, suffix, is_valid = parse_pos_flexible(normalized_pos)
                if is_valid: 
                    target_zone = zone
                    target_line = line
                    target_suffix = suffix
                    target_repo_type = 'HSPM' if zone.startswith('H') else 'HRC1'
            elif raw_pos.startswith('N'):
                m = re.match(r"^([N][A-G])[\.\s\-_]*0*(\d+)(.*)$", raw_pos)
                if m: target_zone = m.group(1); target_line = int(m.group(2)); s = m.group(3); target_suffix = s.strip('.-_ ') if s else None; target_repo_type = 'HRC2'

            # --- VALIDATION ---
            is_valid_config = False
            
            if target_zone and target_line is not None and target_repo_type:
                z_conf = ZONE_CONFIG.get(target_zone)
                if z_conf and target_line in z_conf.get("valid_lines", []):
                    is_valid_config = True

                if is_valid_config:
                    info = get_zone_info(target_zone, target_line)
                    max_slots = calculate_max_slots(info['cap'], info['tiers'])
                    current_count = line_usage_counter.get((target_zone, target_line), 0)
                    is_bai_de = target_repo_type == 'BAI_DE' or (target_zone and target_zone.startswith(('E', 'D')))
                    is_fixed = (target_suffix is not None)

                    valid_to_add = True
                    status_text = "Đúng vị trí"
                    
                    # Logic sai format -> Chuyển thành Gợi ý
                    if is_fixed:
                        if not is_bai_de and not validate_capacity(target_suffix, info['cap'], info['tiers']):
                            is_fixed = False
                            target_suffix = None
                            # [FIX QUAN TRỌNG]: Cập nhật trạng thái để người dùng biết
                            status_text = "Gợi ý (Sai định dạng)"
                    
                    dest_dict = None
                    if target_repo_type == 'HSPM': dest_dict = data_hspm
                    elif target_repo_type == 'HRC1': dest_dict = data_hrc1
                    elif target_repo_type == 'HRC2': dest_dict = data_hrc2
                    elif target_repo_type == 'BAI_DE':
                        if target_zone.startswith('E'): dest_dict = data_baie
                        elif target_zone.startswith('D'): dest_dict = data_baid

                    if dest_dict is not None:
                        if target_zone not in dest_dict: dest_dict[target_zone] = {}
                        if target_line not in dest_dict[target_zone]: dest_dict[target_zone][target_line] = {"fixed": {}, "pending": []}
                        
                        if is_fixed:
                            pos_key = normalize_pos_key(target_suffix)
                            # Check trùng
                            if pos_key in dest_dict[target_zone][target_line]["fixed"]:
                                if not is_bai_de:
                                    valid_to_add = False 
                                else:
                                    status_text = "Gợi ý (Chờ xếp)"
                            else:
                                dest_dict[target_zone][target_line]["fixed"][pos_key] = [coil_id]
                        
                    # Check capacity (Chỉ check nếu là Gợi ý/Auto)
                    if valid_to_add:
                        if not is_bai_de and not is_fixed:
                            if current_count >= max_slots:
                                valid_to_add = False 

                    # --- QUYẾT ĐỊNH XUẤT ---
                    if valid_to_add:
                        line_usage_counter[(target_zone, target_line)] = current_count + 1
                        
                        should_add = False
                        if repo_request == 'ALL': should_add = True
                        elif repo_request == 'BAI_DE' and target_repo_type == 'BAI_DE': should_add = True 
                        elif repo_request == 'BAIE' and target_repo_type == 'BAI_DE' and target_zone.startswith('E'): should_add = True
                        elif repo_request == 'BAID' and target_repo_type == 'BAI_DE' and target_zone.startswith('D'): should_add = True
                        elif repo_request == target_repo_type: should_add = True
                        
                        if should_add:
                            final_export_rows.append({
                                "ID Cuộn": coil_id,
                                "Vị trí": raw_pos,
                                "Khối lượng": w_val,
                                "Nhóm": group_val,
                                "SO Mapping": so_val,
                                "Trạng thái": status_text,
                                "Zone-Line": f"{target_zone}-{target_line}"
                            })

        if not final_export_rows: return "Không có dữ liệu hợp lệ cho kho này", 404

        df_export = pd.DataFrame(final_export_rows)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            sheet_name = f'ChiTiet_{repo_request}'
            df_export.to_excel(writer, sheet_name=sheet_name[:30], index=False)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name[:30]]
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFF2CC', 'border': 1})
            for i, col in enumerate(df_export.columns):
                worksheet.write(0, i, col, header_fmt)
                worksheet.set_column(i, i, 18)

        output.seek(0)
        return send_file(output, download_name=f"ChiTiet_{repo_request}.xlsx", as_attachment=True)

    except Exception as e:
        print(e)
        return str(e), 500
@kho2d_bp.route('/export/errors')
def export_errors_excel():
    if not engine: return "Lỗi kết nối DB", 500
    try:
        query = """select s.[ID Cuộn Bó], s.[Vị trí], s.[Khối lượng], s.Nhóm, 
0 as [SO Mapping]
from sanluong s
LEFT JOIN kho k ON s.[ID Cuộn Bó] = k.[ID Cuộn Bó] -- Đặt tên tắt cho bảng kho là 'k'
WHERE 
    s.[ID Cuộn Bó] IS NOT NULL AND s.[ID Cuộn Bó] <> ''
    AND (s.[Đã nhập kho] = 'No' OR s.[Đã nhập kho] IS NULL)
    AND k.[ID Cuộn Bó] IS NULL
UNION ALL
SELECT [ID Cuộn Bó], [Vị trí], [Khối lượng], Nhóm, [SO Mapping] FROM kho"""
        df = pd.read_sql(query, engine)
        df = df.fillna('') 
        df['pos_len'] = df['Vị trí'].astype(str).str.len()
        df['priority'] = df['Vị trí'].apply(calculate_priority)
        df = df.sort_values(by=['priority', 'pos_len'], ascending=[True, True])

        data_hspm = {}; data_hrc1 = {}; data_hrc2 = {}; data_baie = {}; data_baid = {}
        line_usage_counter = {} 
        export_list = [] 

        for _, row in df.iterrows():
            coil_id = str(row['ID Cuộn Bó']).strip()
            raw_pos = str(row['Vị trí']).strip().upper()
            if raw_pos in ['NAN', 'NONE', 'NULL', 'nan', 'None']:
                raw_pos = ""
            raw_so = str(row['SO Mapping']).strip()
            so_val = raw_so if raw_so not in ['0', 'NONE', 'NAN', '', 'None'] else ""
            
            if not coil_id: continue

            target_zone = None; target_line = None; target_suffix = None; target_repo_type = None

            if "RAY" in raw_pos:
                target_suffix = extract_suffix_only(raw_pos)
                if "HA2" in raw_pos: 
                    target_repo_type = 'HSPM'; target_zone = "HA2"; 
                    try: target_line = 91 + (int(coil_id[-1]) % 4)
                    except: target_line = 91
                elif any(z in raw_pos for z in ["KA2", "KA3", "KB2", "KB3", "KC2", "KC3"]):
                    target_repo_type = 'HRC1'
                    for z in ["KA2", "KA3", "KB2", "KB3", "KC2", "KC3"]: 
                        if z in raw_pos: target_zone = z; break
                    target_line = 0
                elif any(z in raw_pos for z in ["NF", "ND", "NG"]):
                    target_repo_type = 'HRC2'
                    for z in ["NF", "ND", "NG"]: 
                        if z in raw_pos: target_zone = z; break
                    target_line = 99
            elif raw_pos == "" or raw_pos in SHIFT_MAPPING_HRC1 or raw_pos in SHIFT_MAPPING_HRC2:
                target_suffix = None 
                target_list = []
                lookup_key = "AUTO_FILL" if raw_pos == "" else raw_pos
                if coil_id.startswith('8'): 
                    target_list = SHIFT_MAPPING_HRC2.get(lookup_key, []); target_repo_type = 'HRC2'
                else: 
                    target_list = SHIFT_MAPPING_HRC1.get(lookup_key, []); target_repo_type = 'HRC1'
                if isinstance(target_list, tuple): target_list = [target_list]
                chosen_zone, chosen_line = None, None
                if target_list:
                    for (cz, cl) in target_list:
                        c_info = get_zone_info(cz, cl)
                        max_limit = calculate_max_slots(c_info['cap'], c_info['tiers'])
                        curr = line_usage_counter.get((cz, cl), 0)
                        if curr < max_limit: chosen_zone, chosen_line = cz, cl; break 
                    if not chosen_zone: chosen_zone, chosen_line = target_list[-1]
                target_zone = chosen_zone; target_line = chosen_line
            elif raw_pos.startswith(('D', 'E')):
                z, l, s, v = parse_pos_DE(raw_pos)
                if v: target_zone = z; target_line = l; target_suffix = s; target_repo_type = 'BAI_DE' 
            elif raw_pos.startswith(('H', 'K')):
                normalized_pos = re.sub(r'[\.\-\_ ]', '', raw_pos)[:8]
                zone, line, suffix, is_valid = parse_pos_flexible(normalized_pos)
                if is_valid: 
                    target_zone = zone
                    target_line = line
                    target_suffix = suffix
                    target_repo_type = 'HSPM' if zone.startswith('H') else 'HRC1'
            elif raw_pos.startswith('N'):
                m = re.match(r"^([N][A-G])[\.\s\-_]*0*(\d+)(.*)$", raw_pos)
                if m: target_zone = m.group(1); target_line = int(m.group(2)); s = m.group(3); target_suffix = s.strip('.-_ ') if s else None; target_repo_type = 'HRC2'

            is_valid_config = False
            error_reason = ""
            
            if target_zone and target_line is not None and target_repo_type:
                z_conf = ZONE_CONFIG.get(target_zone)
                if z_conf and target_line in z_conf.get("valid_lines", []):
                    is_valid_config = True
                else: error_reason = f"Line {target_line} chưa cấu hình trong {target_zone}"

                if is_valid_config:
                    info = get_zone_info(target_zone, target_line)
                    max_slots = calculate_max_slots(info['cap'], info['tiers'])
                    current_count = line_usage_counter.get((target_zone, target_line), 0)
                    is_bai_de = target_repo_type == 'BAI_DE' or (target_zone and target_zone.startswith(('E', 'D')))
                    is_fixed = (target_suffix is not None)

                    valid_to_add = True
                    
                    if is_fixed:
                        if not is_bai_de and not validate_capacity(target_suffix, info['cap'], info['tiers']):
                            is_fixed = False; target_suffix = None
                    
                    dest_dict = None
                    if target_repo_type == 'HSPM': dest_dict = data_hspm
                    elif target_repo_type == 'HRC1': dest_dict = data_hrc1
                    elif target_repo_type == 'HRC2': dest_dict = data_hrc2
                    elif target_repo_type == 'BAI_DE':
                        if target_zone.startswith('E'): dest_dict = data_baie
                        elif target_zone.startswith('D'): dest_dict = data_baid

                    if dest_dict is not None:
                        if target_zone not in dest_dict: dest_dict[target_zone] = {}
                        if target_line not in dest_dict[target_zone]: dest_dict[target_zone][target_line] = {"fixed": {}, "pending": []}
                        
                        if is_fixed:
                            pos_key = normalize_pos_key(target_suffix)
                            if pos_key in dest_dict[target_zone][target_line]["fixed"]:
                                if not is_bai_de:
                                    valid_to_add = False
                                    existing = dest_dict[target_zone][target_line]["fixed"][pos_key][0]
                                    error_reason = f"Vị trí {target_suffix} bị trùng với {existing}"
                                else: pass
                            else:
                                dest_dict[target_zone][target_line]["fixed"][pos_key] = [coil_id] 
                        
                    # [FIX REALITY]: Check Capacity chỉ áp dụng cho Auto/Pending
                    if valid_to_add:
                        if not is_bai_de and not is_fixed:
                            if current_count >= max_slots:
                                valid_to_add = False
                                error_reason = f"Dư thừa - Line đã đầy ({current_count}/{max_slots})"

                    if valid_to_add:
                        line_usage_counter[(target_zone, target_line)] = current_count + 1
                        continue 

            if not error_reason: error_reason = "Sai định dạng/Chưa Mapping"
            export_list.append({
                "ID Cuộn": coil_id, "Vị trí (Gốc)": raw_pos, "SO Mapping": so_val, "Thông tin lỗi": error_reason
            })

        df_export = pd.DataFrame(export_list)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_export.to_excel(writer, sheet_name='DanhSachLoi', index=False)
            workbook = writer.book
            worksheet = writer.sheets['DanhSachLoi']
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
            for col_num, value in enumerate(df_export.columns.values):
                worksheet.write(0, col_num, value, header_fmt)
                worksheet.set_column(col_num, col_num, 20) 
            red_fmt = workbook.add_format({'font_color': 'red'})
            worksheet.set_column('D:D', 35, red_fmt)

        output.seek(0)
        return send_file(output, download_name="DanhSach_Loi.xlsx", as_attachment=True)

    except Exception as e: return str(e), 500