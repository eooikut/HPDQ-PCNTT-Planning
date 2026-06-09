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
    "HA": {
        "valid_lines": list(range(1, 10)) + list(range(10, 19)) + [81, 82, 83, 84, 85, 86] ,
        "lines": {
            **{i: {"cap": 48, "tiers": 3} for i in range(1, 10)},
            **{i: {"cap": 45, "tiers": 3} for i in range(10, 19)},
            81: {"cap": 48, "tiers": 2, "desc": "HAR 1"},
            82: {"cap": 48, "tiers": 2, "desc": "HAR 2"},
            83: {"cap": 48, "tiers": 2, "desc": "HAR 3"},
            84: {"cap": 45, "tiers": 2, "desc": "HAR 4"},
            85: {"cap": 45, "tiers": 2, "desc": "HAR 5"},
            86: {"cap": 45, "tiers": 2, "desc": "HAR 6"},
        }
    },
    "HB": { "valid_lines": list(range(1, 28)), "lines": { **{i: {"cap": 48, "tiers": 3} for i in range(1, 10)}, **{i: {"cap": 45, "tiers": 3} for i in range(10, 19)}, **{i: {"cap": 39, "tiers": 3} for i in range(19, 28)} } },
    "HC": { "valid_lines": list(range(1, 31)), "lines": { **{i: {"cap": 48, "tiers": 3} for i in range(1, 11)}, **{i: {"cap": 45, "tiers": 3} for i in range(11, 21)}, **{i: {"cap": 39, "tiers": 3} for i in range(21, 31)} } },
    "HD": { "valid_lines": list(range(1, 31)), "lines": { **{i: {"cap": 48, "tiers": 3} for i in range(1, 11)}, **{i: {"cap": 45, "tiers": 3} for i in range(11, 21)}, **{i: {"cap": 39, "tiers": 3} for i in range(21, 31)} } },

    # --- KHO HRC1 ---
    "KA": {
        "valid_lines": [0] + list(range(1, 19)), 
        "lines": {
            0: {"cap": 18, "tiers": 2, "desc": "RAY"},
            **{i: {"cap": 20, "tiers": 3} for i in range(1, 10)}, 
            1: {"cap": 4, "tiers": 3}, 2: {"cap": 14, "tiers": 3}, 3: {"cap": 15, "tiers": 3}, 4: {"cap": 18, "tiers": 3},
            **{i: {"cap": 23, "tiers": 3} for i in range(10, 19)}, 18: {"cap": 11, "tiers": 3}
        }
    },
    "KB": {
        "valid_lines": [0] + list(range(1, 28)),
        "lines": {
            0: {"cap": 18, "tiers": 1, "desc": "RAY"},
            **{i: {"cap": 20, "tiers": 3} for i in range(1, 10)}, 
            1: {"cap": 4, "tiers": 3}, 2: {"cap": 14, "tiers": 3}, 3: {"cap": 15, "tiers": 3}, 4: {"cap": 18, "tiers": 3},
            **{i: {"cap": 21, "tiers": 3} for i in range(10, 19)},
            **{i: {"cap": 20, "tiers": 3} for i in range(19, 28)},
            24: {"cap": 8, "tiers": 3}, 25: {"cap": 8, "tiers": 3}, 26: {"cap": 8, "tiers": 3}, 27: {"cap": 8, "tiers": 3}
        }
    },
    "KC": {
        "valid_lines": [0] + list(range(1, 19)),
        "lines": {
            0: {"cap": 18, "tiers": 2, "desc": "RAY"},
            **{i: {"cap": 20, "tiers": 3} for i in range(1, 10)}, 
            1: {"cap": 4, "tiers": 3}, 2: {"cap": 14, "tiers": 3}, 3: {"cap": 15, "tiers": 3}, 4: {"cap": 18, "tiers": 3},
            **{i: {"cap": 20, "tiers": 3} for i in range(10, 19)}
        }
    },

    # --- KHO HRC2 ---
    "NF": { "valid_lines": [99] + list(range(1, 10)), "lines": { 99: {"cap": 25, "tiers": 1, "desc": "RAY"}, 1: {"cap": 34, "tiers": 2}, 2: {"cap": 34, "tiers": 2}, "DEFAULT": {"cap": 46, "tiers": 2} } },
    "ND": { "valid_lines": [99] + list(range(1, 7)), "lines": { 99: {"cap": 50, "tiers": 1, "desc": "RAY"}, "DEFAULT": {"cap": 66, "tiers": 2} } },
    "NG": { "valid_lines": [99] + list(range(1, 10)), "lines": { 99: {"cap": 30, "tiers": 1, "desc": "RAY"}, 7: {"cap": 37, "tiers": 2}, 8: {"cap": 34, "tiers": 2}, 9: {"cap": 34, "tiers": 2}, "DEFAULT": {"cap": 46, "tiers": 2} } },
    "NA": { "valid_lines": list(range(1, 12)), "lines": { 1: {"cap": 48, "tiers": 2}, 2: {"cap": 48, "tiers": 2}, 3: {"cap": 48, "tiers": 2}, 4: {"cap": 51, "tiers": 2}, 5: {"cap": 51, "tiers": 2}, 6: {"cap": 45, "tiers": 2}, "DEFAULT": {"cap": 27, "tiers": 2} } },
    "NB": { "valid_lines": list(range(1, 12)), "lines": { "DEFAULT": {"cap": 33, "tiers": 2} } },
    "NC": { "valid_lines": list(range(1, 7)), "lines": { "DEFAULT": {"cap": 32, "tiers": 2} } },
    "NE": { "valid_lines": list(range(1, 8)), "lines": { "DEFAULT": {"cap": 23, "tiers": 2} } },
    
    # --- BÃI E ---
    "E01": { "valid_lines": list(range(2, 37)), "lines": { "DEFAULT": {"cap": 22, "tiers": 2} } },
    "E02": { "valid_lines": list(range(2, 36)), "lines": { "DEFAULT": {"cap": 22, "tiers": 2} } },
    
    # --- BÃI D ---
    "D05": { "valid_lines": list(range(1, 34)), "lines": { "DEFAULT": {"cap": 23, "tiers": 2} } },
    "D06": { "valid_lines": list(range(1, 30)), "lines": { "DEFAULT": {"cap": 23, "tiers": 2} } },
    "D07": { "valid_lines": list(range(1, 30)), "lines": { "DEFAULT": {"cap": 23, "tiers": 2} } },
    
    "GLOBAL": {"cap": 30, "tiers": 3} 
}

SHIFT_MAPPING_HRC1 = {
    "AUTO_FILL": [
        # KA3 cũ (1-9) -> KA mới (10-18)
        ("KA", 10), ("KA", 11), ("KA", 12), ("KA", 13), ("KA", 14), ("KA", 15), ("KA", 16), ("KA", 17), ("KA", 18),
        # KB3 cũ (1-9) -> KB mới (19-27) (Do KB1=1-9, KB2=10-18)
        ("KB", 19), ("KB", 20), ("KB", 21), ("KB", 22), ("KB", 23), ("KB", 24), ("KB", 25), ("KB", 26), ("KB", 27),
        # KC3 cũ (1-9) -> KC mới (10-18)
        ("KC", 10), ("KC", 11), ("KC", 12), ("KC", 13), ("KC", 14), ("KC", 15), ("KC", 16), ("KC", 17), ("KC", 18)
    ],
    "A":  [("KA", 10), ("KA", 11), ("KA", 16), ("KB", 19), ("KB", 22), ("KC", 10), ("KC", 13), ("KC", 14)],
    "A1": [("KA", 10), ("KA", 11), ("KA", 16), ("KB", 19), ("KB", 22), ("KC", 10), ("KC", 13), ("KC", 14)],
    "A2": [("KA", 10), ("KA", 11), ("KA", 16), ("KB", 19), ("KB", 22), ("KC", 10), ("KC", 13), ("KC", 14)],
    
    "B":  [("KA", 12), ("KA", 13), ("KA", 17), ("KB", 20), ("KB", 23), ("KC", 11), ("KC", 15), ("KC", 16)],
    "B1": [("KA", 12), ("KA", 13), ("KA", 17), ("KB", 20), ("KB", 23), ("KC", 11), ("KC", 15), ("KC", 16)],
    "B2": [("KA", 12), ("KA", 13), ("KA", 17), ("KB", 20), ("KB", 23), ("KC", 11), ("KC", 15), ("KC", 16)],
    
    "C":  [("KA", 14), ("KA", 15), ("KA", 18), ("KB", 21), ("KB", 24), ("KB", 25), ("KB", 26), ("KC", 12), ("KC", 17), ("KC", 18)],
    "C1": [("KA", 14), ("KA", 15), ("KA", 18), ("KB", 21), ("KB", 24), ("KB", 25), ("KB", 26), ("KC", 12), ("KC", 17), ("KC", 18)],
    "C2": [("KA", 14), ("KA", 15), ("KA", 18), ("KB", 21), ("KB", 24), ("KB", 25), ("KB", 26), ("KC", 12), ("KC", 17), ("KC", 18)],
}

# HRC2 giữ nguyên số Line (Vì cấu hình Zone HRC2 không gộp line)
SHIFT_MAPPING_HRC2 = {
    "AUTO_FILL" : [
        ("ND", 1), ("ND", 2), ("ND", 3), ("ND", 4),("ND", 5), ("ND", 6),
        ("NC", 1), ("NC", 2),("NC", 3), ("NC", 4),("NC", 5), ("NC", 6),
        ("NE", 1), ("NE", 2),("NE", 3), ("NE", 4),("NE", 5), ("NE", 6),
        ("NG", 1), ("NG", 2),("NG", 3), ("NG", 4),("NG", 5), ("NG", 6),("NG", 7), ("NG", 8),("NG", 9),
        ("NF", 1), ("NF", 2),("NF", 3), ("NF", 4),("NF", 5), ("NF", 6),("NF", 7), ("NF", 8),("NF", 9)
    ],
    "A":  [("ND", 3), ("ND", 6), ("NC", 2), ("NC", 5), ("NE", 3), ("NE", 6), ("NG", 3), ("NG", 6), ("NG", 9), ("NF", 3), ("NF", 6), ("NF", 9)],
    "A1": [("ND", 3), ("ND", 6), ("NC", 2), ("NC", 5), ("NE", 3), ("NE", 6), ("NG", 3), ("NG", 6), ("NG", 9), ("NF", 3), ("NF", 6), ("NF", 9)],
    "A2": [("ND", 3), ("ND", 6), ("NC", 2), ("NC", 5), ("NE", 3), ("NE", 6), ("NG", 3), ("NG", 6), ("NG", 9), ("NF", 3), ("NF", 6), ("NF", 9)],
    
    "B":  [("ND", 2), ("ND", 5), ("NC", 1), ("NC", 4), ("NE", 2), ("NE", 5), ("NG", 2), ("NG", 5), ("NG", 8), ("NF", 2), ("NF", 5), ("NF", 8)],
    "B1": [("ND", 2), ("ND", 5), ("NC", 1), ("NC", 4), ("NE", 2), ("NE", 5), ("NG", 2), ("NG", 5), ("NG", 8), ("NF", 2), ("NF", 5), ("NF", 8)],
    "B2": [("ND", 2), ("ND", 5), ("NC", 1), ("NC", 4), ("NE", 2), ("NE", 5), ("NG", 2), ("NG", 5), ("NG", 8), ("NF", 2), ("NF", 5), ("NF", 8)],
    
    "C":  [("ND", 1), ("ND", 4), ("NC", 3), ("NC", 6), ("NE", 1), ("NE", 4), ("NE", 7), ("NG", 1), ("NG", 4), ("NG", 7), ("NF", 1), ("NF", 4), ("NF", 7)],
    "C1": [("ND", 1), ("ND", 4), ("NC", 3), ("NC", 6), ("NE", 1), ("NE", 4), ("NE", 7), ("NG", 1), ("NG", 4), ("NG", 7), ("NF", 1), ("NF", 4), ("NF", 7)],
    "C2": [("ND", 1), ("ND", 4), ("NC", 3), ("NC", 6), ("NE", 1), ("NE", 4), ("NE", 7), ("NG", 1), ("NG", 4), ("NG", 7), ("NF", 1), ("NF", 4), ("NF", 7)],
}

WAREHOUSE_LIMITS = { "HSPM": 326968.0, "HRC1": 79258.0, "HRC2": 107272.0, "BAIE": 68241.0, "BAID": 94185.0, "OTHER": 100000.0 }

ZONE_MAPPING_OLD_TO_NEW = {
    "HA1":"HA", "HA2":"HA", "HA3":"HA", "HB1":"HB", "HB2":"HB", "HB3":"HB",
    "HC1":"HC", "HC2":"HC", "HC3":"HC", "HD1":"HD", "HD2":"HD", "HD3":"HD",
    "KA1":"KA", "KA2":"KA", "KA3":"KA", "KB1":"KB", "KB2":"KB", "KB3":"KB", "KC1":"KC", "KC2":"KC", "KC3":"KC"
}

# =============================================================================================
# 2. HELPER FUNCTIONS
# =============================================================================================

@kho2d_bp.route('/kho2d')
def kho2d(): return render_template('kho2d.html')

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

def validate_capacity(tier, index_val, max_cap, max_tiers):
    if isinstance(index_val, str): return True 
    if index_val < 1: return False
    if index_val > (max_cap * 2 + 2): return False
    if tier == 'Y' and max_tiers < 2: return False
    if tier == 'Z' and max_tiers < 3: return False
    return True

def extract_suffix_only(raw_str):
    match = re.search(r'[\.\-_ ]+(\d+[X]?)$', raw_str.strip())
    if match: return match.group(1)
    return None

def normalize_pos_key(suffix_val, tier):
    if tier == 'Z': return f"{suffix_val}X"
    return str(suffix_val)

def parse_pos_DE(raw_pos):
    if not raw_pos: return None, None, None, None, False
    clean_pos = re.sub(r'[\-_ ]+', '.', raw_pos.strip().upper())
    match = re.match(r"^([DE]0?\d+)(\.|)(\d+)(.*)$", clean_pos)
    if match:
        raw_zone = match.group(1)
        z_match = re.match(r"^([DE])0?(\d+)$", raw_zone)
        if z_match:
            zone = f"{z_match.group(1)}{int(z_match.group(2)):02d}"
            try: line = int(match.group(3))
            except: return None, None, None, None, False
            return zone, line, match.group(4).strip('.'), "DE", True
    return None, None, None, None, False

def parse_pos_flexible(pos_str):
    if not pos_str: return None, None, None, None, False
    pos_str = pos_str.upper().strip()
    match = re.match(r"^([HK][A-Z]\d)[^0-9]*(\d{1,2})(.*)$", pos_str)
    if match:
        try:
            zone = match.group(1); line = int(match.group(2))
            raw_suffix = match.group(3).strip('.-_ ')
            return zone, line, raw_suffix, None, True
        except: pass
    return None, None, None, None, False

# [ĐÃ SỬA]: Logic đọc Z giữ nguyên số (Z03 là 3)
def parse_pos_unified(pos_str):
    if not pos_str: return None, None, None, None, False
    pos_str = pos_str.strip().upper()
    if "HAR" in pos_str:
        # Regex giải thích:
        # ^HAR       : Bắt đầu bằng HAR
        # (?:AY)?    : Có thể có chữ AY hoặc không (Non-capturing group)
        # [\.\-_ ]* : Có thể có dấu chấm, gạch ngang, khoảng trắng
        # 0* : Có thể có số 0 ở đầu (để bắt 01 thành 1)
        # (\d)       : BẮT BUỘC phải có 1 chữ số -> Đây là số Line (1-9)
        # (.*)       : Phần còn lại (Suffix)
        
        m_har = re.match(r"^HAR(?:AY)?[\.\-_ ]*0*(\d)(.*)$", pos_str)
        
        if m_har:
            har_num = int(m_har.group(1)) # Lấy được số 1 từ "HARAY01"
            raw_suffix = m_har.group(2).strip('.-_ ')
            
            # Map sang Line nội bộ (HAR1 -> 81)
            internal_line = 80 + har_num
            
            # Xử lý phần đuôi để xem là Fixed hay Pending
            m_tier = re.match(r"^([XYZ])(\d+)$", raw_suffix)
            m_tier_rev = re.match(r"^(\d+)([XYZ])$", raw_suffix)
            
            if m_tier:
                return "HA", internal_line, int(m_tier.group(2)), m_tier.group(1), True
            elif m_tier_rev:
                return "HA", internal_line, int(m_tier_rev.group(1)), m_tier_rev.group(2), True
            elif re.match(r"^\d+$", raw_suffix):
                return "HA", internal_line, int(raw_suffix), "LOOSE_IDX", True
            else:
                # Nếu chuỗi là "HARAY01" -> Regex ăn hết "HARAY01" vào phần đầu -> Suffix rỗng
                return "HA", internal_line, raw_suffix, "PENDING", True
    # 1. Format Chuẩn: HA01X01, HA01Z03
    is_standard = re.match(r"^([HK][A-Z])(\d{2})([XYZ])(\d{2})$", pos_str)
    if is_standard:
        zone = is_standard.group(1)
        line = int(is_standard.group(2))
        tier = is_standard.group(3)
        suffix_val = int(is_standard.group(4))
        
        # [SỬA]: Giữ nguyên giá trị, không cộng 2 nữa
        final_index = suffix_val 
        
        return zone, line, final_index, tier, True

    # 2. Format Ray
    if "RAY" in pos_str:
        for old_z, new_z in ZONE_MAPPING_OLD_TO_NEW.items():
            if old_z in pos_str:
                line_match = re.search(r'(\d+)RAY', pos_str)
                line = int(line_match.group(1)) if line_match else 0
                if "KA" in old_z and line==0: line=0 
                elif "HA" in old_z and line==0: line=91
                return new_z, line, pos_str, "RAY", True
        if any(z in pos_str for z in ["NF", "ND", "NG"]):
             for z in ["NF", "ND", "NG"]:
                 if z in pos_str: return z, 99, pos_str, "RAY", True

    # 3. Format Bãi E/D
    if pos_str.startswith(('D', 'E')):
        return parse_pos_DE(pos_str)
    
    # 4. Format Lỏng (Loose)
    m_loose = re.match(r"^([HKN][A-Z])[\.\s\-_]*0*(\d+)(.*)$", pos_str)
    if m_loose:
        zone = m_loose.group(1); line = int(m_loose.group(2)); raw_suffix = m_loose.group(3).strip('.-_ ')
        m_tier = re.match(r"^([XYZ])(\d+)$", raw_suffix)     # Dạng Z03
        m_tier_rev = re.match(r"^(\d+)([XYZ])$", raw_suffix) # Dạng 03Z
        
        if m_tier:
            t = m_tier.group(1); idx = int(m_tier.group(2))
            final_idx = idx # [SỬA]: Không cộng 2
            return zone, line, final_idx, t, True
        elif m_tier_rev:
            idx = int(m_tier_rev.group(1)); t = m_tier_rev.group(2)
            final_idx = idx # [SỬA]: Không cộng 2
            return zone, line, final_idx, t, True
            
        if re.match(r"^\d+$", raw_suffix): return zone, line, int(raw_suffix), "LOOSE_IDX", True
        return zone, line, raw_suffix, "PENDING", True

    return None, None, None, None, False
def calculate_priority(val):
    s = str(val).strip().upper()
    if "RAY" in s: return 0
    if re.match(r"^[HKNED][A-Z]*\d+", s): return 1
    if not s or s in ['NONE', 'NAN', 'NULL']: return 3
    return 2

def get_warehouse_name(zone):
    if zone.startswith('H'): return "HSPM"
    if zone.startswith('K'): return "HRC1"
    if zone.startswith('N'): return "HRC2"
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

def translate_legacy_coords(old_zone, old_line):
    new_zone = ZONE_MAPPING_OLD_TO_NEW.get(old_zone, old_zone)
    new_line = old_line
    if old_zone.endswith('2'): 
        if old_zone.startswith(('HC', 'HD')): new_line = old_line + 10
        else: new_line = old_line + 9
    elif old_zone.endswith('3'): 
        if old_zone.startswith(('HC', 'HD')): new_line = old_line + 20
        elif old_zone.startswith('HA'): new_line = old_line + 18
        else: new_line = old_line + 9 
        if old_zone.startswith('KB'): new_line = old_line + 18
    return new_zone, new_line

# =============================================================================================
# 3. CORE LOGIC PROCESSOR
# =============================================================================================
def get_warehouse_core_data():
    if not engine: raise Exception("Lỗi kết nối DB")
    
    query = """select s.[ID Cuộn Bó], s.[Vị trí], s.[Khối lượng], s.Nhóm, 0 as [SO Mapping]
        from sanluong s WITH (NOLOCK) LEFT JOIN kho k WITH (NOLOCK) ON s.[ID Cuộn Bó] = k.[ID Cuộn Bó]
        WHERE s.[ID Cuộn Bó] IS NOT NULL AND s.[ID Cuộn Bó] <> '' AND (s.[Đã nhập kho] = 'No' OR s.[Đã nhập kho] IS NULL) AND k.[ID Cuộn Bó] IS NULL
        UNION ALL SELECT [ID Cuộn Bó], [Vị trí], [Khối lượng], Nhóm, [SO Mapping] FROM kho WITH (NOLOCK)"""
    
    df = pd.read_sql(query, engine).fillna('')
    df['pos_len'] = df['Vị trí'].astype(str).str.len()
    df['priority'] = df['Vị trí'].apply(calculate_priority)
    df = df.sort_values(by=['priority', 'pos_len'], ascending=[True, True])

    data_repos = {"HSPM": {}, "HRC1": {}, "HRC2": {}, "BAIE": {}, "BAID": {}}
    capacity_state = {}
    for zone, z_conf in ZONE_CONFIG.items():
        if zone == 'GLOBAL': continue
        for line in z_conf.get("valid_lines", []):
            capacity_state[(zone, line)] = { "occupied_indices": {1: set(), 2: set(), 3: set()}, "pending_count": 0, "pending_items": [] }

    errors_by_wh = { "ALL": [], "HSPM": [], "HRC1": [], "HRC2": [], "BAI_DE": [], "OTHER": [] }
    valid_export_list = []; auto_list = [] 
    
    total_capacity_slots = 0
    for zone, z_conf in ZONE_CONFIG.items():
        if zone == 'GLOBAL': continue
        for line in z_conf.get("valid_lines", []):
            info = get_zone_info(zone, line)
            total_capacity_slots += calculate_max_slots(info['cap'], info['tiers'])

    stats = { 
        "total": 0, "valid": 0, "invalid": 0, "auto_assigned": 0, 
        "hspm_count": 0, "hrc1_count": 0, "hrc2_count": 0, "baie_count": 0, "baid_count": 0,
        "total_capacity": total_capacity_slots, "total_weight": 0, "limits": WAREHOUSE_LIMITS,
        "err_count_by_wh":  {"HSPM": 0, "HRC1": 0, "HRC2": 0, "BAIE": 0, "BAID": 0, "OTHER": 0},
        "err_weight_by_wh": {"HSPM": 0.0, "HRC1": 0.0, "HRC2": 0.0, "BAIE": 0.0, "BAID": 0.0, "OTHER": 0.0},
        "valid_weight_by_wh": {"HSPM": 0.0, "HRC1": 0.0, "HRC2": 0.0, "BAIE": 0.0, "BAID": 0.0}
    }
    
    line_usage_counter = {}
    total_weight_kg = 0.0

    for _, row in df.iterrows():
        stats["total"] += 1
        coil_id = str(row['ID Cuộn Bó']).strip(); raw_pos = str(row['Vị trí']).strip().upper()
        if raw_pos in ['NAN', 'NONE', 'NULL', 'nan', 'None']: raw_pos = ""
        raw_so = str(row['SO Mapping']).strip(); so_val = raw_so if raw_so not in ['0', 'NONE', 'NAN', '', 'None'] else ""
        try: w_val_float = float(row['Khối lượng']) if row['Khối lượng'] else 0.0
        except: w_val_float = 0.0
        total_weight_kg += w_val_float; w_ton = w_val_float / 1000.0; w_val = f"{w_val_float:,.0f}"
        group_val = str(row['Nhóm'])
        if not coil_id: continue

        target_zone = None; target_line = None; target_repo_type = None
        idx_suffix = None; target_tier = None
        is_fixed = False

        # --- LOGIC ƯU TIÊN 1: MAPPING (A, B, C...) & AUTO FILL ---
        mapping_found = False
        lookup_key = "AUTO_FILL" if raw_pos == "" else raw_pos
        
        target_list_old = []
        is_mapping_code = (raw_pos == "" or raw_pos in SHIFT_MAPPING_HRC1 or raw_pos in SHIFT_MAPPING_HRC2 or len(raw_pos) <= 3)
        if is_mapping_code and re.match(r"^[HKNE][A-Z0-9]+$", raw_pos): is_mapping_code = False

        if is_mapping_code:
            # [FIX LỖI PHÂN LOẠI]: Không gán cứng target_repo_type ngay lập tức
            # Chỉ gán khi tìm thấy mapping thực sự
            temp_repo = None
            if coil_id.startswith('8'): 
                target_list_old = SHIFT_MAPPING_HRC2.get(lookup_key, []); temp_repo = 'HRC2'
            else: 
                target_list_old = SHIFT_MAPPING_HRC1.get(lookup_key, []); temp_repo = 'HRC1'
            
            if target_list_old:
                mapping_found = True
                is_fixed = False
                target_repo_type = temp_repo # Chỉ gán khi có dữ liệu
                
                if isinstance(target_list_old, tuple): target_list_old = [target_list_old]
                
                target_list_new = []
                for (oz, ol) in target_list_old:
                    nz, nl = translate_legacy_coords(oz, ol)
                    target_list_new.append((nz, nl))
                
                chosen_zone, chosen_line = None, None
                for (cz, cl) in target_list_new:
                    c_info = get_zone_info(cz, cl)
                    max_limit = calculate_max_slots(c_info['cap'], c_info['tiers'])
                    curr = line_usage_counter.get((cz, cl), 0)
                    if curr < max_limit: 
                        chosen_zone, chosen_line = cz, cl; break 
                if not chosen_zone: chosen_zone, chosen_line = target_list_new[-1]
                target_zone = chosen_zone; target_line = chosen_line

        # --- LOGIC ƯU TIÊN 2: BÃI E/D ---
        if not mapping_found and not target_zone and raw_pos.startswith(('E', 'D')):
             z_ed, l_ed, s_ed, t_ed, v_ed = parse_pos_DE(raw_pos)
             if v_ed:
                 target_zone = z_ed; target_line = l_ed; idx_suffix = s_ed; target_tier = t_ed
                 target_repo_type = 'BAI_DE'
                 # [FIX LỖI BBC/ABC VÀO FIXED]: Chỉ số mới là Fixed, chữ là Pending
                 if idx_suffix and idx_suffix.isdigit(): is_fixed = True
                 else: is_fixed = False

        # --- LOGIC ƯU TIÊN 3: PARSER MỚI ---
        if not mapping_found and not target_zone:
            p_zone, p_line, p_idx, p_tier, is_valid_new = parse_pos_unified(raw_pos)
            matched_unified = False
            
            if is_valid_new and p_zone:
                is_standard = re.match(r"^([HK][A-Z])(\d{2})([XYZ])(\d{2})$", raw_pos)
                if is_standard or "RAY" in raw_pos or p_tier in ["LOOSE_IDX", "PENDING"] or p_tier in ["X", "Y", "Z"]:
                    target_zone = p_zone; target_line = p_line; idx_suffix = p_idx; target_tier = p_tier
                    
                    if target_tier == "RAY": is_fixed = False
                    elif target_tier == "PENDING": is_fixed = False 
                    elif not idx_suffix: is_fixed = False 
                    else: is_fixed = True
                        
                    if p_zone.startswith('H'): target_repo_type = 'HSPM'
                    elif p_zone.startswith('K'): target_repo_type = 'HRC1'
                    elif p_zone.startswith('N'): target_repo_type = 'HRC2'
                    elif p_zone.startswith(('E', 'D')): target_repo_type = 'BAI_DE'
                    matched_unified = True
            
            if not matched_unified and raw_pos.startswith(('H', 'K')):
                normalized_pos = re.sub(r'[\.\-\_ ]', '', raw_pos)[:8]
                old_z, old_l, old_s, _, old_valid = parse_pos_flexible(normalized_pos)
                if old_valid and old_z in ZONE_MAPPING_OLD_TO_NEW:
                    target_zone = ZONE_MAPPING_OLD_TO_NEW[old_z]
                    _, target_line = translate_legacy_coords(old_z, old_l)
                    idx_suffix = old_s
                    target_repo_type = 'HSPM' if target_zone.startswith('H') else 'HRC1'
                    if not idx_suffix: is_fixed = False
                    else: is_fixed = True

        # Validation & Insert
        is_valid_config = False; error_reason = ""
        
        if target_zone and target_line is not None and target_repo_type:
            z_conf = ZONE_CONFIG.get(target_zone)
            if z_conf and target_line in z_conf.get("valid_lines", []): is_valid_config = True
            else: error_reason = f"Line {target_line} chưa cấu hình trong {target_zone}"

            if is_valid_config:
                info = get_zone_info(target_zone, target_line)
                max_slots = calculate_max_slots(info['cap'], info['tiers'])
                current_count = line_usage_counter.get((target_zone, target_line), 0)
                is_bai_de = (target_repo_type == 'BAI_DE') or (target_zone.startswith(('E','D')))
                is_ray = (target_tier == "RAY") or ("RAY" in raw_pos)

                valid_to_add = True
                
                # Check Capacity cho Fixed (Bãi E/D bỏ qua check này)
                if is_fixed and not is_ray and not is_bai_de:
                    if not validate_capacity(target_tier, idx_suffix, info['cap'], info['tiers']):
                        valid_to_add = False; error_reason = f"Vị trí vượt quá giới hạn"

                dest_dict = None; stats_key = ""
                if target_repo_type == 'HSPM': dest_dict = data_repos["HSPM"]; stats_key = 'hspm_count'
                elif target_repo_type == 'HRC1': dest_dict = data_repos["HRC1"]; stats_key = 'hrc1_count'
                elif target_repo_type == 'HRC2': dest_dict = data_repos["HRC2"]; stats_key = 'hrc2_count'
                elif target_repo_type == 'BAI_DE':
                    if target_zone.startswith('E'): dest_dict = data_repos["BAIE"]; stats_key = 'baie_count'
                    elif target_zone.startswith('D'): dest_dict = data_repos["BAID"]; stats_key = 'baid_count'
                
                if dest_dict is not None and valid_to_add:
                    if target_zone not in dest_dict: dest_dict[target_zone] = {}
                    if target_line not in dest_dict[target_zone]: dest_dict[target_zone][target_line] = {"fixed": {}, "pending": []}
                    
                    cell_data = [coil_id, raw_pos, w_val, group_val, so_val]
                    added_as_pending = False
                    status_text = "Đúng vị trí"

                    if is_fixed:
                        pos_key = str(idx_suffix)
                        if target_tier == 'Z': pos_key = f"{idx_suffix}X"
                        
                        if pos_key in dest_dict[target_zone][target_line]["fixed"]:
                            if is_bai_de:
                                dest_dict[target_zone][target_line]["pending"].append(cell_data + [True])
                                added_as_pending = True; status_text = "Gợi ý (Chờ xếp)"
                            else:
                                existing_item = dest_dict[target_zone][target_line]["fixed"][pos_key]
                                existing_id = existing_item[0] 
                                valid_to_add = False
                                error_reason = f"Trùng vị trí {pos_key} với ID {existing_id}" 
                        else:
                            dest_dict[target_zone][target_line]["fixed"][pos_key] = cell_data

                            cap_entry = capacity_state.get((target_zone, target_line))
                            if cap_entry and not is_ray and not is_bai_de and (isinstance(idx_suffix, int) or str(idx_suffix).isdigit()):
                                num_idx = int(idx_suffix)
                                t_idx = 1
                                if target_tier == 'X': t_idx = 1
                                elif target_tier == 'Y': t_idx = 2
                                elif target_tier == 'Z': t_idx = 3
                                elif num_idx % 2 == 0: t_idx = 2 
                                cap_entry["occupied_indices"][t_idx].add(num_idx)
                    else:
                        dest_dict[target_zone][target_line]["pending"].append(cell_data + [True])
                        added_as_pending = True
                        if target_tier == "RAY": status_text = "Gợi ý (Ray)"
                        else: status_text = "Gợi ý (Auto)"
                        cap_entry = capacity_state.get((target_zone, target_line))
                        if cap_entry: cap_entry["pending_items"].append(coil_id)

                    if valid_to_add:
                        if not is_bai_de and not is_fixed and current_count >= max_slots:
                            valid_to_add = False; error_reason = f"Line đầy ({current_count}/{max_slots})"
                            if added_as_pending: dest_dict[target_zone][target_line]["pending"].pop()
                            if cap_entry: cap_entry["pending_items"].pop()

                    if valid_to_add:
                        line_usage_counter[(target_zone, target_line)] = current_count + 1
                        if added_as_pending:
                             auto_list.append({"id": coil_id, "pos": raw_pos, "reason": f"{target_zone}-{info['desc']}", "so": so_val})
                             stats["auto_assigned"] += 1
                        if stats_key: stats[stats_key] += 1
                        stats["valid"] += 1
                        valid_export_list.append({
                            "ID Cuộn": coil_id, "Vị trí": raw_pos, "Khối lượng": w_val, 
                            "Nhóm": group_val, "SO Mapping": so_val, "Trạng thái": status_text,
                            "Zone-Line": f"{target_zone}-{target_line}", "Repo": target_repo_type
                        })
                        continue

        # [FIX LỖI MẤT DỮ LIỆU]: Nếu xuống đây, nghĩa là có Lỗi (bao gồm cả Overflow)
        stats["invalid"] += 1
        if not error_reason: error_reason = "Sai format/Chưa Mapping"
        err_item = {"id": coil_id, "pos": raw_pos, "so": so_val, "reason": error_reason}
        errors_by_wh["ALL"].append(err_item)
        
        # [FIX LỖI PHÂN LOẠI]: Logic phân loại chính xác hơn
        det_repo = "OTHER"
        if target_repo_type == 'BAI_DE': 
            if target_zone and target_zone.startswith('D'): det_repo = "BAID" 
            else: det_repo = "BAIE"
        elif target_repo_type: det_repo = target_repo_type
        # Fallback nếu target_repo_type None
        elif raw_pos.startswith('H'): det_repo = "HSPM"
        elif raw_pos.startswith('K'): det_repo = "HRC1"
        elif raw_pos.startswith('N'): det_repo = "HRC2"
        elif raw_pos.startswith('D'): det_repo = "BAID"
        elif raw_pos.startswith('E'): det_repo = "BAIE"
        
        if det_repo in stats["err_weight_by_wh"]:
            stats["err_weight_by_wh"][det_repo] += w_ton
            stats["err_count_by_wh"][det_repo] += 1
        else:
            stats["err_weight_by_wh"]["OTHER"] += w_ton
            stats["err_count_by_wh"]["OTHER"] += 1
            
        if det_repo in ["BAIE", "BAID"]: errors_by_wh["BAI_DE"].append(err_item)
        elif det_repo in errors_by_wh: errors_by_wh[det_repo].append(err_item)
        else: errors_by_wh["OTHER"].append(err_item)

    stats["total_weight"] = total_weight_kg / 1000.0
    return { "frontend_data": data_repos, "capacity_state": capacity_state, "valid_export_list": valid_export_list, "errors_by_wh": errors_by_wh, "auto_list": auto_list, "stats": stats, "structure": generate_structure() }

# =============================================================================================
# 4. ENDPOINTS & UTILS
# =============================================================================================

@kho2d_bp.route('/api/data')
def get_data():
    try:
        core_data = get_warehouse_core_data()
        return jsonify(clean_obj({
            "status": "success", "structure": core_data["structure"], "data_hspm": core_data["frontend_data"]["HSPM"], 
            "data_hrc1": core_data["frontend_data"]["HRC1"], "data_hrc2": core_data["frontend_data"]["HRC2"], 
            "data_baie": core_data["frontend_data"]["BAIE"], "data_baid": core_data["frontend_data"]["BAID"], 
            "errors_by_wh": core_data["errors_by_wh"], "auto_list": core_data["auto_list"], "stats": core_data["stats"]
        }))
    except Exception as e: return jsonify({"status": "error", "message": str(e)}), 500

@kho2d_bp.route('/export/warehouse_list')
def export_warehouse_list():
    try:
        core_data = get_warehouse_core_data()
        full_list = core_data["valid_export_list"]; repo_request = request.args.get('repo', 'ALL'); filtered_list = []
        for item in full_list:
            should_add = False; r_type = item["Repo"]
            if repo_request == 'ALL': should_add = True
            elif repo_request == 'BAI_DE' and r_type == 'BAI_DE': should_add = True
            elif repo_request == 'BAIE' and r_type == 'BAI_DE' and "E" in item["Zone-Line"]: should_add = True
            elif repo_request == 'BAID' and r_type == 'BAI_DE' and "D" in item["Zone-Line"]: should_add = True
            elif repo_request == r_type: should_add = True
            if should_add:
                item_copy = item.copy(); del item_copy["Repo"]; filtered_list.append(item_copy)
        if not filtered_list: return "Không có dữ liệu", 404
        df = pd.DataFrame(filtered_list); output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='ChiTiet', index=False); worksheet = writer.sheets['ChiTiet']; worksheet.set_column(0, 6, 18)
        output.seek(0)
        return send_file(output, download_name=f"ChiTiet_{repo_request}.xlsx", as_attachment=True)
    except Exception as e: return str(e), 500

@kho2d_bp.route('/export/errors')
def export_errors_excel():
    try:
        core_data = get_warehouse_core_data()
        err_list = [{"ID Cuộn": err["id"], "Vị trí (Gốc)": err["pos"], "SO Mapping": err["so"], "Thông tin lỗi": err["reason"]} for err in core_data["errors_by_wh"]["ALL"]]
        df = pd.DataFrame(err_list); output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Loi', index=False)
        output.seek(0)
        return send_file(output, download_name="DanhSach_Loi.xlsx", as_attachment=True)
    except Exception as e: return str(e), 500

def calculate_layout_allocation(zone, line, data, info):
    max_tiers = info['tiers']; config_max = info['cap']; real_max_t1 = config_max
    occupied = data["occupied_indices"]; pending_queue = list(data["pending_items"])
    total_pending_initial = list(pending_queue); virtual_map = {1: {}, 2: {}, 3: {}}; max_fixed_pos = 0
    for t in [1, 2, 3]:
        for idx in occupied[t]:
            virtual_map[t][idx] = "FIXED"; 
            if idx > max_fixed_pos: max_fixed_pos = idx
    step = 1 if max_tiers == 1 else 2
    req_len = math.ceil(max_fixed_pos / step) if step > 0 else 0
    if req_len > real_max_t1: real_max_t1 = req_len
    is_bai_de = zone.startswith(('E', 'D')); allocated_items = []; allocated_ids = []
    
    if pending_queue and max_tiers >= 2:
        t2_positions = sorted(list(virtual_map[2].keys()))
        for t2_loc in t2_positions:
            for leg in [t2_loc-1, t2_loc+1]:
                if leg > 0 and leg not in virtual_map[1] and pending_queue:
                    pid = pending_queue.pop(0); virtual_map[1][leg] = pid
                    allocated_items.append({"tier": 1, "id": leg, "val": pid, "type": "suggest"}); allocated_ids.append(pid)
    k = 0
    while True:
        if not is_bai_de and k >= real_max_t1: break
        if is_bai_de and k >= real_max_t1 and not pending_queue: break
        if k > 500: break 
        t1_num = 1 + k * step
        if t1_num not in virtual_map[1] and pending_queue:
            pid = pending_queue.pop(0); virtual_map[1][t1_num] = pid
            allocated_items.append({"tier": 1, "id": t1_num, "val": pid, "type": "suggest"}); allocated_ids.append(pid)
        if t1_num in virtual_map[1]:
            if k + 1 > real_max_t1: real_max_t1 = k + 1
        if max_tiers >= 2 and k > 0:
            t2_num = 2 + (k - 1) * step
            if t2_num not in virtual_map[2] and pending_queue:
                if (t2_num-1) in virtual_map[1] and (t2_num+1) in virtual_map[1]:
                    pid = pending_queue.pop(0); virtual_map[2][t2_num] = pid
                    allocated_items.append({"tier": 2, "id": t2_num, "val": pid, "type": "suggest"}); allocated_ids.append(pid)
        if max_tiers >= 3 and k > 1:
            t3_num = 3 + (k - 2) * step
            if t3_num not in virtual_map[3] and pending_queue:
                if (t3_num-1) in virtual_map[2] and (t3_num+1) in virtual_map[2]:
                    pid = pending_queue.pop(0); virtual_map[3][t3_num] = pid
                    allocated_items.append({"tier": 3, "id": t3_num, "val": pid, "type": "suggest"}); allocated_ids.append(pid)
        k += 1
    unplaced_items = []
    temp_allocated = list(allocated_ids)
    for item in total_pending_initial:
        if item in temp_allocated: temp_allocated.remove(item)
        else: unplaced_items.append(item)
    return allocated_items, real_max_t1, unplaced_items

@kho2d_bp.route('/api/stats/capacity')
def get_capacity_stats():
    try:
        core_data = get_warehouse_core_data(); wh_state = core_data["capacity_state"]; stats_data = []
        for (zone, line), data in wh_state.items():
            info = get_zone_info(zone, line); allocated, final_len, unplaced_items = calculate_layout_allocation(zone, line, data, info)
            count_t1 = sum(1 for x in allocated if x['tier'] == 1); count_t2 = sum(1 for x in allocated if x['tier'] == 2); count_t3 = sum(1 for x in allocated if x['tier'] == 3)
            fixed_t1 = len(data["occupied_indices"][1]); fixed_t2 = len(data["occupied_indices"][2]); fixed_t3 = len(data["occupied_indices"][3])
            fill_t1, fill_t2, fill_t3 = 0, 0, 0; rem = len(unplaced_items)
            suggest_map = { (i['tier'], i['id']): i['val'] for i in allocated }; step = 1 if info['tiers'] == 1 else 2
            for k in range(final_len):
                if rem <= 0: break
                tid = 1 + k*step
                if tid not in data["occupied_indices"][1] and (1, tid) not in suggest_map: fill_t1+=1; rem-=1
            if info['tiers']>=2:
                 for k in range(final_len-1):
                    if rem <= 0: break
                    tid = 2 + k*step
                    if tid not in data["occupied_indices"][2] and (2, tid) not in suggest_map: fill_t2+=1; rem-=1
            if info['tiers']>=3:
                 for k in range(final_len-2):
                    if rem <= 0: break
                    tid = 3 + k*step
                    if tid not in data["occupied_indices"][3] and (3, tid) not in suggest_map: fill_t3+=1; rem-=1
            if rem > 0: fill_t1 += rem 
            stats_data.append({
                "warehouse": get_warehouse_name(zone), "zone": zone, "line": line, "desc": info['desc'], "tiers": info['tiers'],
                "stats": {
                    "t1": { "max": info['cap'], "total_used": fixed_t1+count_t1+fill_t1, "suggested": count_t1+fill_t1, "final_empty": max(0, info['cap']-(fixed_t1+count_t1+fill_t1)) },
                    "t2": { "max": info['cap']-1, "total_used": fixed_t2+count_t2+fill_t2, "suggested": count_t2+fill_t2, "final_empty": max(0, (info['cap']-1)-(fixed_t2+count_t2+fill_t2)) },
                    "t3": { "max": info['cap']-2, "total_used": fixed_t3+count_t3+fill_t3, "suggested": count_t3+fill_t3, "final_empty": max(0, (info['cap']-2)-(fixed_t3+count_t3+fill_t3)) }
                }
            })
        return jsonify({"status": "success", "data": stats_data})
    except Exception as e: return jsonify({"status": "error", "message": str(e)}), 500

# [ĐÃ SỬA]: Xuất Excel Z03 thay vì Z01
@kho2d_bp.route('/export/capacity')
def export_capacity_excel():
    try:
        core_data = get_warehouse_core_data(); wh_state = core_data["capacity_state"]; export_rows = []; sorted_keys = sorted(wh_state.keys())
        for (zone, line) in sorted_keys:
            data = wh_state[(zone, line)]; info = get_zone_info(zone, line); wh_name = get_warehouse_name(zone)
            allocated, final_len, unplaced_items = calculate_layout_allocation(zone, line, data, info)
            suggest_map = { (i['tier'], i['id']): i['val'] for i in allocated }; unplaced_queue = list(unplaced_items)
            def fill_gap(): return ("Gợi ý (Lỗi chân)", unplaced_queue.pop(0)) if len(unplaced_queue) > 0 else ("Trống", "")
            step = 1 if info['tiers'] == 1 else 2
            
            # --- T1 ---
            for k in range(final_len):
                t1_id = 1 + k * step; status = "Trống"; val = ""
                pos_str = f"{zone}{line:02d}X{t1_id:02d}" 
                if t1_id in data["occupied_indices"][1]: status = "Đúng vị trí"
                elif (1, t1_id) in suggest_map: status = "Gợi ý"; val = suggest_map[(1, t1_id)]
                else: status, val = fill_gap()
                if status == "Đúng vị trí": continue 
                if "Gợi ý" in status or k < info['cap']: export_rows.append({ "Kho Tổng": wh_name, "Zone": zone, "Line": f"{line}", "Tầng": 1, "Vị trí": pos_str, "Trạng thái": status, "ID Gợi ý": val })
            
            # --- T2 ---
            if info['tiers'] >= 2:
                count_t2 = final_len - 1
                for k in range(count_t2):
                    t2_id = 2 + k * step; status = "Trống"; val = ""
                    pos_str = f"{zone}{line:02d}Y{t2_id:02d}"
                    if t2_id in data["occupied_indices"][2]: status = "Đúng vị trí"
                    elif (2, t2_id) in suggest_map: status = "Gợi ý"; val = suggest_map[(2, t2_id)]
                    else: status, val = fill_gap()
                    if status == "Đúng vị trí": continue
                    if "Gợi ý" in status or k < (info['cap'] - 1): export_rows.append({ "Kho Tổng": wh_name, "Zone": zone, "Line": f"{line}", "Tầng": 2, "Vị trí": pos_str, "Trạng thái": status, "ID Gợi ý": val })
            
            # --- T3 (CHỈNH SỬA TẠI ĐÂY) ---
            if info['tiers'] >= 3:
                count_t3 = final_len - 2
                for k in range(count_t3):
                    t3_id = 3 + k * step; status = "Trống"; val = ""
                    
                    # [FIX]: Dùng thẳng t3_id (3, 5, 7...) thay vì trừ 2
                    pos_str = f"{zone}{line:02d}Z{t3_id:02d}" 
                    
                    if t3_id in data["occupied_indices"][3]: status = "Đúng vị trí"
                    elif (3, t3_id) in suggest_map: status = "Gợi ý"; val = suggest_map[(3, t3_id)]
                    else: status, val = fill_gap()
                    if status == "Đúng vị trí": continue
                    if "Gợi ý" in status or k < (info['cap'] - 2): export_rows.append({ "Kho Tổng": wh_name, "Zone": zone, "Line": f"{line}", "Tầng": 3, "Vị trí": pos_str, "Trạng thái": status, "ID Gợi ý": val })
            
            while len(unplaced_queue) > 0:
                rem_id = unplaced_queue.pop(0); export_rows.append({ "Kho Tổng": wh_name, "Zone": zone, "Line": f"{line}", "Tầng": 1, "Vị trí": "Tràn kho (Overflow)", "Trạng thái": "Gợi ý (Dư thừa)", "ID Gợi ý": rem_id })
        
        df_export = pd.DataFrame(export_rows); output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_export.to_excel(writer, sheet_name='ViTriTrong', index=False); workbook = writer.book; worksheet = writer.sheets['ViTriTrong']; red_fmt = workbook.add_format({'font_color': '#d63384', 'bold': True}); worksheet.set_column('G:G', 20, red_fmt)
        output.seek(0); return send_file(output, download_name="BaoCao_ViTri_GoiY.xlsx", as_attachment=True)
    except Exception as e: return str(e), 500