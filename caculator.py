import pandas as pd
import numpy as np
import re
FACTORY_CONFIGS = {
    # ======================================================
    # NHÀ MÁY SỐ 1 (HRC1) - Logic cũ của bạn
    # ======================================================
        'HRC1': 
{
    'LOW_CARBON': {
        'name': 'Low Carbon (Thép mềm)',
        # Định nghĩa các khoảng: [Cận dưới, Cận trên, Label hiển thị]
        'ranges': [
            (1.20, 1.30, '1.20<=T<1.30'), # Index 0
            (1.30, 1.40, '1.30<=T<1.40'), # Index 1
            (1.40, 1.50, '1.40<=T<1.50'),
            (1.50, 1.65, '1.50<=T<1.65'),
            (1.65, 1.80, '1.65<=T<1.80'),
            (1.80, 2.00, '1.80<=T<2.00'),
            (2.00, 2.20, '2.00<=T<2.20'),
            (2.20, 2.40, '2.20<=T<2.40'),
            (2.40, 2.75, '2.40<=T<2.75'),
            (2.75, 2.90, '2.75<=T<2.90'),
            (2.90, 12.00, '2.90>=T')       # Index 10
        ],
        # Tỷ lệ % 
        'ratios': {
                '900-1000':  [0.0, 0.0, 0.0, 0.02, 0.03, 0.10, 0.50, 0.20, 0.10, 0.05, 0.0],
                '1000-1100': [0.0, 0.0, 0.0, 0.02, 0.03, 0.10, 0.50, 0.20, 0.10, 0.05, 0.0],
                '1100-1200': [0.0, 0.0, 0.0, 0.02, 0.03, 0.05, 0.55, 0.20, 0.10, 0.05, 0.0],
                '1200-1300': [0.0, 0.0, 0.0, 0.02, 0.03, 0.05, 0.55, 0.20, 0.10, 0.05, 0.0],
                '1300-1400': [0.0, 0.0, 0.0, 0.0, 0.0, 0.05, 0.45, 0.25, 0.15, 0.05, 0.05],
                '1400-1500': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.30, 0.35, 0.20, 0.10, 0.05],
                '1500-1524': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.30, 0.35, 0.20, 0.10, 0.05]
        }
    },

    'MEDIUM_CARBON': {
        'name': 'Medium Carbon (Thép trung)',
        # Khoảng độ dày KHÁC (Cập nhật theo ẢNH 2)
        # Ví dụ: Medium thường ko chạy mỏng dưới 1.6
        'ranges': [
            (2.00, 2.25, '2.00<=T<2.25'), # Index 0
            (2.25, 2.45, '2.25<=T<2.45'), # Index 1
            (2.45, 2.75, '2.45<=T<2.75'),
            (2.75, 3.00, '2.75<=T<3.00'),
            (3.00, 3.50, '3.00<=T<3.50'),
            (3.50, 12.00, '3.50>=T')
        ],
        # Tỷ lệ % của Medium (Cập nhật theo ẢNH 2)
        # Lưu ý: Số lượng phần tử trong list này phải bằng số lượng ranges ở trên (ví dụ ở đây là 6)
        'ratios': {
            '950-1000':  [0.05, 0.50, 0.25, 0.10, 0.10, 0.0], 
            '1000-1100': [0.05, 0.45, 0.30, 0.10, 0.10, 0.0],
            '1100-1200': [0.05, 0.40, 0.45, 0.05, 0.05, 0.0],
            '1200-1300': [0.05, 0.30, 0.50, 0.05, 0.05, 0.05],
            '1300-1400': [0.0, 0.20, 0.50, 0.15, 0.10, 0.05],
            '1400-1500': [0.0, 0.05, 0.45, 0.25, 0.15, 0.10],
            '1500-1524': [0.0, 0.0, 0.40, 0.30, 0.20, 0.10]
        }
    },

    'WEATHER_RESISTANT': {
        'name': 'Kháng thời tiết',
        # Khoảng độ dày KHÁC (Cập nhật theo ẢNH 3)
        'ranges': [
            (1.50, 1.65, '1.50<=T<1.65'),
            (1.65, 1.80, '1.65<=T<1.80'),
            (1.80, 2.00, '1.80<=T<2.00'),
            (2.00, 2.20, '2.00<=T<2.20'),
            (2.20, 2.40, '2.20<=T<2.40'),
            (2.40, 2.75, '2.40<=T<2.75'),
            (2.75, 2.90, '2.75<=T<2.90'),
            (2.90, 6.00, '2.90>=T')
        ],
        # Tỷ lệ % (Cập nhật theo ẢNH 3)
        'ratios': {
            '950-1000':  [0.0, 0.10, 0.15, 0.40, 0.20, 0.05, 0.05, 0.05],
            '1000-1100': [0.05, 0.10, 0.10, 0.45, 0.15, 0.05, 0.05, 0.05],
            '1100-1200': [0.05, 0.10, 0.15, 0.15, 0.20, 0.15, 0.10, 0.10],
            '1200-1300': [0.05, 0.05, 0.10, 0.25, 0.20, 0.15, 0.10, 0.10],
            '1300-1400': [0.0, 0.0, 0.0, 0.0, 0.0, 0.20, 0.30, 0.50],
            '1400-1500': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
            '1500-1524': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
        }
    }
},

    # ======================================================
    # NHÀ MÁY SỐ 2 (HRC2) - Mác riêng & Khổ rộng riêng
    # ======================================================
        'HRC2': 
{
        'LOW_CARBON': {  # Ví dụ mác thép mềm của HRC2
            'name': 'Low Carbon (Thép mềm)',
            # HRC2 có thể chạy dải dày rộng hơn hoặc khác
            'ranges': [
                (1.20, 1.35, '1.20<=T<1.35'), # Index 0
                (1.35, 1.50, '1.35<=T<1.50'), # Index 1
                (1.50, 1.65, '1.50<=T<1.65'),
                (1.65, 2.00, '1.65<=T<2.00'),
                (2.00, 2.15, '2.00<=T<2.15'),
                (2.15, 2.30, '2.15<=T<2.30'),
                (2.30, 2.50, '2.30<=T<2.50'),
                (2.50, 2.75, '2.50<=T<2.75'),
                (2.75, 3.00, '2.75<=T<3.00'),
                (3.00, 99.9, '3.00>=T'), 
            ],
            # HRC2 chỉ có 3 khoảng khổ rộng (Ví dụ) -> Key khác HRC1
            'ratios': {
                '900-1150':     [0.03, 0.03, 0.03, 0.03, 0.72, 0.05, 0.05, 0.02, 0.02, 0.02], # List 4 phần tử tương ứng 4 ranges
                '1150-1400 ':   [0.02, 0.02, 0.02, 0.03, 0.70, 0.05, 0.04, 0.04, 0.04, 0.04],
                '1400-1650 ':   [0.0, 0.0, 0.0, 0.00, 0.05, 0.10, 0.24, 0.25, 0.20, 0.16]
            }
        },
        'MEDIUM_CARBON': {  # Ví dụ mác thép mềm của HRC2
            'name': 'Medium Carbon (Thép trung)',
            # HRC2 có thể chạy dải dày rộng hơn hoặc khác
            'ranges': [
                (1.20, 1.35, '1.20<=T<1.35'), # Index 0
                (1.35, 1.50, '1.35<=T<1.50'), # Index 1
                (1.50, 1.65, '1.50<=T<1.65'),
                (1.65, 2.00, '1.65<=T<2.00'),
                (2.00, 2.15, '2.00<=T<2.15'),
                (2.15, 2.30, '2.15<=T<2.30'),
                (2.30, 2.50, '2.30<=T<2.50'),
                (2.50, 2.75, '2.50<=T<2.75'),
                (2.75, 3.00, '2.75<=T<3.00'),
                (3.00, 99.9, '3.00>=T'),  # Index 3
            ],
            # HRC2 chỉ có 3 khoảng khổ rộng (Ví dụ) -> Key khác HRC1
            'ratios': {
                '900-1150':     [0.00, 0.00, 0.04, 0.08, 0.54, 0.08, 0.08, 0.08, 0.07, 0.03], # List 4 phần tử tương ứng 4 ranges
                '1150-1400 ':   [0.00, 0.00, 0.04, 0.04, 0.40, 0.10, 0.10, 0.10, 0.10, 0.12],
                '1400-1650 ':   [0.0, 0.0, 0.0, 0.0, 0.02, 0.02, 0.06, 0.15, 0.40, 0.35]
            }
        },
        'WEATHER_RESISTANT': {  # Ví dụ mác thép mềm của HRC2
            'name': 'Kháng thời tiết SPA-H',
            # HRC2 có thể chạy dải dày rộng hơn hoặc khác
            'ranges': [
                (1.40, 1.65, '1.40<=T<1.65'),
                (1.65, 1.80, '1.65<=T<1.80'),
                (1.80, 2.00, '1.80<=T<2.00'),
                (2.00, 2.20, '2.00<=T<2.20'),
                (2.20, 2.45, '2.20<=T<2.45'),
                (2.45, 2.75, '2.45<=T<2.75'),
                (2.75, 3.00, '2.75<=T<3.00'),
                (3.00, 6.00, '3.00<=T<6.00') # Index 3
            ],
            # HRC2 chỉ có 3 khoảng khổ rộng (Ví dụ) -> Key khác HRC1
            'ratios': {
                '900-1150':     [0.05, 0.15, 0.20, 0.20, 0.15, 0.10, 0.10, 0.05], # List 4 phần tử tương ứng 4 ranges
                '1150-1400 ':   [0.03, 0.15, 0.15, 0.27, 0.15, 0.10, 0.10, 0.05],
                '1400-1650 ':   [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.15, 0.85]
            }
        },
        'DP600/DP780': {  # Ví dụ mác thép mềm của HRC2
            'name': 'Thép Dual-phase-DP600/DP780',
            # HRC2 có thể chạy dải dày rộng hơn hoặc khác
            'ranges': [
                (1.40, 1.65, '1.40<=T<1.65'),
                (1.65, 1.80, '1.65<=T<1.80'),
                (1.80, 2.00, '1.80<=T<2.00'),
                (2.00, 2.20, '2.00<=T<2.20'),
                (2.20, 2.45, '2.20<=T<2.45'),
                (2.45, 2.75, '2.45<=T<2.75'),
                (2.75, 3.00, '2.75<=T<3.00'),
                (3.00, 6.00, '3.00<=T<6.00')  # Index 3
            ],
            # HRC2 chỉ có 3 khoảng khổ rộng (Ví dụ) -> Key khác HRC1
            'ratios': {
                '900-1150':     [0.0, 0.0, 0.0, 0.45, 0.10, 0.15, 0.15, 0.15], # List 4 phần tử tương ứng 4 ranges
                '1150-1400 ':   [0.0, 0.0, 0.0, 0.40, 0.15, 0.15, 0.15, 0.15],
                '1400-1650 ':   [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 1.00]
            }
        },
        'API X70/X80': {  # Ví dụ mác thép mềm của HRC2
            'name': 'Thép API X70/X80',
            # HRC2 có thể chạy dải dày rộng hơn hoặc khác
            'ranges': [
                (6.00, 12.00, '6.00<=T<12.0')  # Index 3
            ],
            # HRC2 chỉ có 3 khoảng khổ rộng (Ví dụ) -> Key khác HRC1
            'ratios': {
               '900-1150':     [1.00], # List 4 phần tử tương ứng 4 ranges
                '1150-1400 ':   [1.00],
                '1400-1650 ':   [1.00]
            }
        },
        'ULC': {  # Ví dụ mác thép mềm của HRC2
            'name': 'Thép ULC',
            # HRC2 có thể chạy dải dày rộng hơn hoặc khác
            'ranges': [
                (2.00, 2.20, '2.00<=T<2.20'),
                (2.20, 2.45, '2.20<=T<2.45'),
                (2.45, 2.75, '2.45<=T<2.75'),
                (2.75, 3.00, '2.75<=T<3.00'),
                (3.00, 6.00, '3.00<=T<6.00') # Index 3
            ],
            # HRC2 chỉ có 3 khoảng khổ rộng (Ví dụ) -> Key khác HRC1
            'ratios': {
                '900-1150':     [0.10, 0.15, 0.20, 0.30, 0.25], # List 4 phần tử tương ứng 4 ranges
                '1150-1400 ':   [0.10, 0.15, 0.40, 0.20, 0.15],
                '1400-1650 ':   [0.00, 0.15, 0.35, 0.30, 0.20]  
            }
        },
        'High_Cacbon': {  # Ví dụ mác thép mềm của HRC2
            'name': 'High Cacbon',
            # HRC2 có thể chạy dải dày rộng hơn hoặc khác
            'ranges': [
                (2.00, 2.20, '2.00<=T<2.20'),
                (2.20, 2.45, '2.20<=T<2.45'),
                (2.45, 2.75, '2.45<=T<2.75'),
                (2.75, 3.00, '2.75<=T<3.00'),
                (3.00, 6.00, '3.00<=T<6.00') # Index 3
            ],
            # HRC2 chỉ có 3 khoảng khổ rộng (Ví dụ) -> Key khác HRC1
            'ratios': {
                '900-1150':     [0.20, 0.20, 0.30, 0.20, 0.10], # List 4 phần tử tương ứng 4 ranges
                '1150-1400 ':   [0.10, 0.10, 0.35, 0.30, 0.15],
                '1400-1650 ':   [0.0, 0.0, 0.0, 0.0, 0.0]
            }
        },
        'HSLA_S315MC-S700MC': {  # Ví dụ mác thép mềm của HRC2
            'name': 'Thép độ bền cao',
            # HRC2 có thể chạy dải dày rộng hơn hoặc khác
            'ranges': [
                (2.00, 2.20, '2.00<=T<2.20'),
                (2.20, 2.45, '2.20<=T<2.45'),
                (2.45, 2.75, '2.45<=T<2.75'),
                (2.75, 3.00, '2.75<=T<3.00'),
                (3.00, 12.00, '3.00<=T<12.00')  # Index 3
            ],
            # HRC2 chỉ có 3 khoảng khổ rộng (Ví dụ) -> Key khác HRC1
            'ratios': {
                '900-1150':     [0.05, 0.10, 0.10, 0.10, 0.65], # List 4 phần tử tương ứng 4 ranges
                '1150-1400 ':   [0.03, 0.03, 0.03, 0.05, 0.86],
                '1400-1650 ':   [0.0, 0.0, 0.0, 0.05, 0.95]   
            }
        },
    }
}
# ==========================================
# 2. HÀM TRA CỨU GIÁ CHÍNH XÁC
# ==========================================
def get_exact_surcharge(width_val, thick_val):
    # 1. Xác định Cột
    col_idx = -1
    if 900 <= width_val <= 1199: col_idx = 0
    elif 1200 <= width_val <= 1500: col_idx = 1
    elif 1501 <= width_val <= 1650: col_idx = 2
    if col_idx == -1: return 0

    # 2. Ma trận giá
    matrix = {
        (1.20, 1.34): [35, 35, 0],
        (1.35, 1.54): [22, 25, 55],
        (1.55, 1.74): [20, 15, 45],
        (1.75, 1.99): [15, 10, 20],
        (2.00, 2.54): [10, 0, 6],
        (2.55, 3.99): [10, 0, 6],
        (4.00, 8.99): [10, 0, 7],
        (9.00, 15.99): [20, 0, 5],
        (16.00, 25.40): [22, 0, 5]
    }

    # 3. Tra cứu
    for (t_min, t_max), rates in matrix.items():
        if t_min <= thick_val <= t_max + 0.001:
            return rates[col_idx]
    return 0

# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================
def normalize_columns(df):
    df.columns = [str(c).strip().lower() for c in df.columns]
    col_mapping = {
        'Khổ rộng': ['width', 'khổ rộng', 'kho rong', 'k/r', 'rong'],
        'Chiều dày': ['thickness', 'chiều dày', 'chieu day', 'dày', 'day', 'thick'],
        'Khối lượng': ['mass', 'weight', 'khối lượng', 'khoi luong', 'kl', 'qty', 'tấn', 'tan', 'kg']
    }
    new_cols = {}
    for standard_col, variations in col_mapping.items():
        for col in df.columns:
            if col in variations:
                new_cols[col] = standard_col
                break
    if new_cols: df = df.rename(columns=new_cols)
    return df

def get_width_label(width, factory_code='HRC1'):
    if pd.isna(width): return None
    w = float(width)
    
    # === LOGIC CŨ CHO HRC1 ===
    if factory_code == 'HRC1':
        if 900 <= w < 1000: return '900-1000'
        if 1000 <= w < 1100: return '1000-1100'
        if 1100 <= w < 1200: return '1100-1200'
        if 1200 <= w < 1300: return '1200-1300'
        if 1300 <= w < 1400: return '1300-1400'
        if 1400 <= w < 1500: return '1400-1500'
        if 1500 <= w <= 1524: return '1500-1524'

    # === LOGIC MỚI CHO HRC2 (Dựa theo key trong Config bạn gửi) ===
    elif factory_code == 'HRC2':
        # Lưu ý: Các string trả về phải KHỚP 100% với key trong dictionary 'ratios'
        if 900 <= w < 1150: return '900-1150'
        if 1150 <= w < 1400: return '1150-1400 ' # Chú ý: Config của bạn có dấu cách cuối
        if 1400 <= w <= 1650: return '1400-1650 '
        
    return None
def get_thickness_index_dynamic(thick_val, ranges):
    """
    Trả về (index, label) dựa trên giá trị độ dày và cấu hình ranges của Mác.
    """
    if pd.isna(thick_val): return -1, None
    t = float(thick_val)
    
    for i, (low, high, label) in enumerate(ranges):
        # Kiểm tra: low <= t < high
        if low <= t < high:
            return i, label
            
    return -1, None # Không thuộc khoảng nào
# def get_thickness_index_full(thick):
#     if pd.isna(thick): return -1
#     t = float(thick)
#     if 1.20 <= t < 1.30: return 0
#     if 1.30 <= t < 1.40: return 1
#     if 1.40 <= t < 1.50: return 2
#     if 1.50 <= t < 1.65: return 3
#     if 1.65 <= t < 1.80: return 4
#     if 1.80 <= t < 2.00: return 5
#     if 2.00 <= t < 2.20: return 6
#     if 2.20 <= t < 2.40: return 7
#     if 2.40 <= t < 2.75: return 8
#     if 2.75 <= t < 2.90: return 9
#     if t >= 2.90: return 10
#     return -1

def validate_spec(width, thick, factory_code='HRC1'):
    try: w, t = float(width), float(thick)
    except: return False, "Lỗi số liệu"
    
    # === LUẬT CỦA HRC1 (GIỮ NGUYÊN) ===
    if factory_code == 'HRC1':
        if 1.20 <= t < 1.30: return False, f"Độ dày {t}mm chưa hỗ trợ (Vùng đỏ)."
        if 1.30 <= t < 1.40: return False, f"Độ dày {t}mm chưa hỗ trợ (Vùng đỏ)."
        if 1.40 <= t < 1.50 and w >= 1200: return False, f"Độ dày {t}mm cấm khổ >= 1200."
        if 1.50 <= t < 1.65 and w >= 1400: return False, f"Độ dày {t}mm cấm khổ >= 1400."
        if 1.65 <= t < 1.80 and w >= 1500: return False, f"Độ dày {t}mm cấm khổ >= 1500."
    
    # === LUẬT CỦA HRC2 (NẾU CHƯA CÓ THÌ CHO QUA HẾT) ===
    elif factory_code == 'HRC2':
        # Ví dụ: HRC2 máy khỏe hơn, chạy được hết -> Luôn True
        return True, ""
        
    return True, ""

# ==========================================
# 4. LOGIC TÍNH TOÁN (ĐÃ UPDATE HIỂN THỊ)
# ==========================================
def calculate_production_status_dynamic(demand_data, width_label, ratios_dict, ranges_list):
    if width_label not in ratios_dict: return []
    ratios = ratios_dict[width_label]
    
    results = []
    current_mass_sum = 0.0   
    current_ratio_sum = 0.0
    
    for i, ratio in enumerate(ratios):
        label = ranges_list[i][2]
        
        # Lấy dữ liệu từ bước gộp nhóm
        item_data = demand_data.get(i, {'mass': 0, 'money': 0, 'details': ''})
        actual_demand = item_data['mass']
        total_surcharge_amount = item_data['money']
        detail_html = item_data['details'] # Chuỗi HTML đã nối sẵn
        
        # Logic Supply
        current_ratio_sum += ratio
        prev_ratio_sum = current_ratio_sum - ratio
        if prev_ratio_sum == 0:
            generated_supply = current_mass_sum * ratio if ratio > 0 else 0
        else:
            generated_supply = current_mass_sum * (ratio / prev_ratio_sum)
            
        final_production = max(generated_supply, actual_demand)
        diff = generated_supply - actual_demand
        status_text = f"Dư {int(diff):,} T".replace(',', '.') if diff > 0 else ""
        current_mass_sum += final_production
        
        if final_production > 0 or actual_demand > 0:
            results.append({
                'Khổ rộng': width_label,
                'Độ dày': label,
                'Chốt (Sản xuất)': int(round(final_production, 0)),
                'Đơn hàng nhập vào': int(round(actual_demand, 0)),
                'Trạng thái': status_text,
                'Phụ thu (Thành tiền)': total_surcharge_amount,
                # 🟢 Dữ liệu chi tiết dạng text HTML (đã được tạo ở bước run_calculation_tool)
                'Phụ thu (Chi tiết)': detail_html 
            })
            
    return results

def run_calculation_tool(df_input, selected_grade='LOW_CARBON', factory_code='HRC1'):
    # 1. Lấy Config
    factory_config = FACTORY_CONFIGS.get(factory_code)
    if not factory_config: 
        return {'error': f"Nhà máy '{factory_code}' chưa được cấu hình."}
    config = factory_config.get(selected_grade)
    if not config: return {'error': f"Loại mác thép '{selected_grade}' không tồn tại."}
    
    grade_ranges = config['ranges']
    grade_ratios = config['ratios'] # Dictionary chứa tỷ lệ

    # 2. Xử lý Input & Chuẩn hóa
    df = normalize_columns(df_input)
    required_cols = ['Khổ rộng', 'Chiều dày', 'Khối lượng']
    if any(c not in df.columns for c in required_cols):
        return {'error': "Thiếu cột dữ liệu bắt buộc (Khổ rộng, Chiều dày, Khối lượng)."}

    for col in required_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna(subset=required_cols)
    if df.empty: return {'error': "Dữ liệu không hợp lệ."}

    # 3. Validate & Pre-calculation
    errors = []     # Lỗi nghiêm trọng (sai format, vùng đỏ cấm tuyệt đối)
    warnings = []   # Cảnh báo (nhập vào vùng 0% năng lực)
    valid_rows = [] # Danh sách các dòng hợp lệ để tính toán
    
    # Định nghĩa hàm lấy index độ dày
    def get_thick_idx(val):
        return get_thickness_index_dynamic(val, grade_ranges)[0]

    for idx, row in df.iterrows():
        w_val = row['Khổ rộng']
        t_val = row['Chiều dày']
        
        # 3.1 Check Validate cứng (Vùng đỏ kỹ thuật)
        is_valid, msg = validate_spec(w_val, t_val)
        if not is_valid:
            errors.append(f"Dòng {idx+2}: {msg}")
            continue # Bỏ qua dòng lỗi
            
        # 3.2 Check Tỷ trọng 0% (Ngoài khung năng lực)
        # Lấy nhãn khổ rộng (ví dụ: '1200-1300')
        w_label = get_width_label(w_val,factory_code=factory_code)
        # Lấy index độ dày (ví dụ: 2)
        t_idx = get_thick_idx(t_val)
        
        # Nếu không xác định được khổ hoặc dày -> Lỗi
        if w_label is None or t_idx == -1:
             warnings.append(f"Dòng {idx+2}: Khổ {w_val} x Dày {t_val} không nằm trong phạm vi quy định.")
             continue

        # Kiểm tra tỷ trọng trong cấu hình
        # grade_ratios[w_label] là list các %
        # Nếu grade_ratios[w_label][t_idx] == 0.0 -> Cảnh báo
        if w_label in grade_ratios:
            ratio_val = grade_ratios[w_label][t_idx]
            if ratio_val == 0.0:
                warnings.append(f"Dòng {idx+2}: Khổ {w_val} x Dày {t_val} thuộc vùng tỷ trọng 0% (Ngoài khung năng lực). Đã bỏ qua.")
                continue # Bỏ qua, không tính toán dòng này
        else:
             # Trường hợp hiếm: Khổ rộng có nhưng không có trong bảng ratio
             warnings.append(f"Dòng {idx+2}: Khổ {w_val} chưa có cấu hình tỷ lệ.")
             continue

        # Nếu vượt qua hết -> Thêm vào danh sách hợp lệ
        valid_rows.append(row)

    # Nếu có lỗi nghiêm trọng (như sai format file), trả về lỗi ngay
    if errors: return {'error': "<br>".join(errors)}
    
    # Nếu không còn dòng nào hợp lệ (do bị warning hết)
    if not valid_rows:
        msg = "Không có dòng dữ liệu nào hợp lệ để tính toán.<br><b>Chi tiết:</b><br>" + "<br>".join(warnings)
        return {'error': msg}

    # Tạo DataFrame mới chỉ chứa các dòng hợp lệ
    df_valid = pd.DataFrame(valid_rows)

    # 4. Tính tiền chi tiết (Chỉ tính cho df_valid)
    detail_list = []
    surcharge_list = []
    
    for _, row in df_valid.iterrows():
        rate = get_exact_surcharge(row['Khổ rộng'], row['Chiều dày'])
        mass_ton = row['Khối lượng']
        money = mass_ton * rate
        
        if rate > 0:
            info_str = f"{row['Chiều dày']}mm - {mass_ton:g}T - ${rate} - <b>${money:,.0f}</b>"
        else:
            info_str = f"{row['Chiều dày']}mm - {mass_ton:g}T - Không phụ thu"
            
        surcharge_list.append(money)
        detail_list.append(info_str)

    df_valid['Surcharge_Amount'] = surcharge_list
    df_valid['Detail_Info'] = detail_list

    # 5. Gộp nhóm & Tính toán (Logic cũ)
    df_valid['Width_Label'] = df_valid['Khổ rộng'].apply(lambda x: get_width_label(x, factory_code=factory_code))
    
    # Lưu ý: Apply lại hàm lấy index vì df_valid là dataframe mới
    df_valid['Thickness_Index'] = df_valid['Chiều dày'].apply(lambda x: get_thick_idx(x))
    
    df_agg = df_valid.dropna(subset=['Width_Label']).groupby(['Width_Label', 'Thickness_Index']).agg({
        'Khối lượng': 'sum',
        'Surcharge_Amount': 'sum',
        'Detail_Info': lambda x: '\n'.join(x)
    }).reset_index()

    final_report = []
    for width in df_agg['Width_Label'].unique():
        group_data = df_agg[df_agg['Width_Label'] == width]
        
        demand_data = {}
        for _, row in group_data.iterrows():
            idx = row['Thickness_Index']
            if idx != -1: 
                demand_data[idx] = {
                    'mass': row['Khối lượng'],
                    'money': row['Surcharge_Amount'],
                    'details': row['Detail_Info']
                }
        
        final_report.extend(calculate_production_status_dynamic(
            demand_data, 
            width, 
            grade_ratios, 
            grade_ranges
        ))

    # 🟢 TRẢ VỀ CẢ KẾT QUẢ VÀ CẢNH BÁO
    # Thay vì trả về list, ta trả về dict để chứa cả 2 thông tin
    return {
        'data': final_report,
        'warnings': warnings
    }