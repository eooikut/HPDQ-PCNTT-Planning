from flask import Blueprint, render_template, request, jsonify
from sqlalchemy import text
from db import engine,engine_mysql
from datetime import datetime, timedelta
from dateutil import parser
from auth.decorator import permission_required
import io
import pandas as pd
from flask import send_file
from collections import OrderedDict
import os
import requests
from dotenv import load_dotenv
dashboard_bp = Blueprint("dashboard_bp", __name__)
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Load các biến môi trường từ file .env
load_dotenv()

def get_date_from_iso(iso_str):
    """Hàm phụ trợ parse chuỗi ISO sang Date"""
    if not iso_str: return None
    try:
        return datetime.fromisoformat(iso_str).date()
    except ValueError:
        return None

def get_vessel_times(v):
    """Lấy ngày bắt đầu và kết thúc làm hàng dựa vào các mốc thời gian"""
    start_date = get_date_from_iso(v.get('atb')) or get_date_from_iso(v.get('etb')) or get_date_from_iso(v.get('ata')) or get_date_from_iso(v.get('eta'))
    end_date = get_date_from_iso(v.get('atc')) or get_date_from_iso(v.get('etc')) or get_date_from_iso(v.get('atd')) or get_date_from_iso(v.get('etd'))
    return start_date, end_date

import time
import concurrent.futures
def get_actual_exported_from_mysql(ship_names):
    """Truy vấn tổng khối lượng thực tế đã xuất (kg) từ MySQL theo Tàu và SO"""
    if not ship_names:
        return {}

    # Chèn params an toàn để tránh SQL Injection
    placeholders = ', '.join([':ship_' + str(i) for i in range(len(ship_names))])
    params = {f'ship_{i}': name for i, name in enumerate(ship_names)}

    sql = text(f"""
        SELECT 
            Transporter AS tau,
            SO AS saleO,
            SUM(Weight) AS tong_thucte_xuat_kg
        FROM v_phieuxuathang_hrc
        WHERE Transporter IN ({placeholders})
          AND SO IS NOT NULL AND SO != ''
        GROUP BY Transporter, SO
    """)

    try:
        with engine_mysql.connect() as conn:
            rows = conn.execute(sql, params).mappings().all()
            
        # Trả về dict format: { ('TÊN TÀU', 'SO'): kg } (Upper text để match chuẩn xác)
        return { 
            (str(r['tau']).upper().strip(), str(r['saleO']).lstrip('0').upper().strip()): float(r['tong_thucte_xuat_kg'] or 0)
            for r in rows 
        }
    except Exception as e:
        print(f"Lỗi truy vấn MySQL: {e}")
        return {}
# --- CẤU HÌNH CACHE TOÀN CỤC (Lưu 10 phút) ---
_cached_vessels = []
_last_fetch_time = 0
CACHE_DURATION = 600 

def fetch_single_page(page, from_date, to_date, api_user, api_pass):
    """Hàm phụ trợ gọi 1 trang đơn lẻ (Dùng cho đa luồng)"""
    url = "https://apiplcos.hoaphatdungquat.vn:60524/api/hpdq_ship_sync/GetScheduleBerth"
    params = {'fromDate': from_date, 'toDate': to_date, 'offset': page, 'limit': 10}
    try:
        res = requests.get(url, auth=(api_user, api_pass), params=params, verify=False, timeout=10)
        return res.json().get('data', {}) if res.status_code == 200 else {}
    except: return {}

def get_hrc_planned_schedule(selected_ship=None):
    """Logic tính Kế hoạch HRC: Tối ưu Parallel + Cache + Chia đều khi có Ngày ra"""
    global _cached_vessels, _last_fetch_time
    current_time = time.time()
    
    # 1. PHẦN TẢI DỮ LIỆU (Sử dụng Cache và Gọi song song)
    if not _cached_vessels or (current_time - _last_fetch_time > CACHE_DURATION):
        api_user, api_pass = os.environ.get("CRM_API_USER"), os.environ.get("CRM_API_PASS")
        now = datetime.now()
        f_date, t_date = (now - timedelta(days=30)).strftime('%Y-%m-%d'), (now + timedelta(days=30)).strftime('%Y-%m-%d')

        # Gọi trang 1 để lấy tổng số trang
        d1 = fetch_single_page(1, f_date, t_date, api_user, api_pass)
        t_pages = d1.get('totalPages', 1)
        all_v = d1.get('vessels', [])

        # Gọi song song các trang còn lại để tăng tốc (Parallel Requests)
        if t_pages > 1:
            with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
                fs = [executor.submit(fetch_single_page, p, f_date, t_date, api_user, api_pass) for p in range(2, t_pages + 1)]
                for f in concurrent.futures.as_completed(fs):
                    all_v.extend(f.result().get('vessels', []))
        
        _cached_vessels, _last_fetch_time = all_v, current_time
    else:
        all_v = _cached_vessels

    # 2. PHẦN XỬ LÝ THUẬT TOÁN PHÂN BỔ
    MAX_CAPACITY = 25000
    hrc_planned_daily = {}
    hrc_avg_vessels = []   # Nhóm biết rõ ngày vào - ngày ra
    hrc_spill_vessels = [] # Nhóm chỉ biết ngày vào

    for v in all_v:
        weight = float(v.get('cargO_WEIGHT') or 0)
        if weight <= 0: continue
        if weight > 200000: weight /= 1000.0 # Chuẩn hóa đơn vị
            
        start_date, end_date = get_vessel_times(v) # Hàm này đã ưu tiên Thực tế > Dự kiến
        if not start_date: continue
        
        if 'HRC' in str(v.get('cargO_TYPE', '')).upper():
            # Lọc theo tên tàu
            v_id, v_name = str(v.get('vesseL_ID', '')).upper(), str(v.get('vesseL_NAME', '')).upper()
            if selected_ship:
                s_up = str(selected_ship).upper()
                if s_up not in v_id and s_up not in v_name: continue

            # --- LOGIC PHÂN LOẠI MỚI THEO Ý BẠN ---
            if end_date and end_date >= start_date:
                # Nếu có cả ngày Vào và Ra (Dự kiến hoặc Thực tế) -> Chia trung bình
                hrc_avg_vessels.append({'start': start_date, 'end': end_date, 'weight': weight})
            else:
                # Nếu thiếu ngày Ra -> Dùng logic dồn ngày và lùi dần
                hrc_spill_vessels.append({'start': start_date, 'weight': weight})

    # A. Ưu tiên xử lý Nhóm chia trung bình trước
    for item in hrc_avg_vessels:
        days = (item['end'] - item['start']).days + 1
        avg_w = item['weight'] / days
        for i in range(days):
            d = item['start'] + timedelta(days=i)
            hrc_planned_daily[d] = hrc_planned_daily.get(d, 0) + avg_w

    # B. Xử lý Nhóm dồn ngày dựa trên dung lượng còn lại
    hrc_spill_vessels.sort(key=lambda x: x['start'])
    for item in hrc_spill_vessels:
        w_left, curr_d = item['weight'], item['start']
        while w_left > 0:
            used = hrc_planned_daily.get(curr_d, 0)
            avail = max(0, MAX_CAPACITY - used)
            if avail > 0:
                alloc = min(w_left, avail)
                hrc_planned_daily[curr_d] = used + alloc
                w_left -= alloc
            if w_left > 0: curr_d += timedelta(days=1)

    return hrc_planned_daily
def get_rows_from_db():
    """Lấy dữ liệu tổng hợp cho thanh tiến độ chính của tàu."""
    sql = text("""SELECT
        tau,
        SheetMonth,
        ETA_Parsed,
        -- Tính tổng khối lượng đã giao (shipped + mapped)
        CAST(SUM(ISNULL(shipped_qty, 0) + ISNULL(Mapping_kho, 0)) / 1000.0 AS INT) AS tongkhoiluong,
        -- Lấy khối lượng tổng của tàu, MAX() để đảm bảo chỉ có 1 giá trị trên mỗi nhóm
        MAX(khoi_luong_tong) AS khoi_luong_tong,
        -- Thêm ETA_Parsed vào đây để join với chi tiết SO
        ETA_Parsed AS eta_date_key 
    FROM testdb with (nolock)
    WHERE
        tau != N'ĐƯỜNG BỘ' AND tau != ''
    GROUP BY
            tau, SheetMonth, ETA_Parsed
    """)
    with engine.connect() as conn:
        rows = conn.execute(sql).mappings().all()
        
    # --- BỘ LỌC THÉP: XỬ LÝ DỮ LIỆU BẨN TỪ GỐC ---
    result = []
    for r in rows:
        d = dict(r)
        eta = d.get('ETA_Parsed')
        if isinstance(eta, str):
            try:
                # Ép chuỗi thành date, nếu lỗi hoặc năm < 2000 (như 0202) thì gán None
                parsed_date = parser.parse(eta[:10]).date() if len(eta) >= 10 else parser.parse(eta).date()
                d['ETA_Parsed'] = parsed_date if parsed_date.year > 2000 else None
            except Exception:
                d['ETA_Parsed'] = None
        elif isinstance(eta, datetime):
            d['ETA_Parsed'] = eta.date() if eta.year > 2000 else None
        
        result.append(d)
        
    return result

def get_so_details_for_dashboard():
    """Lấy dữ liệu chi tiết của từng SO để hiển thị trong tooltip."""
    sql = text("""
      WITH data AS (
        SELECT
            lt.[TÀU/PHƯƠNG TIỆN VẬN TẢI] AS tau,
            lt.ETA_Parsed,
            s.[SO Mapping] AS saleO,
            s.[Material Description] AS material,
            ISNULL(s.[Shipped Quantity (KG)],0) AS shipped_qty,
            ISNULL(s.[Quantity (KG)],0) AS qty,
            ISNULL(s.[SL Mapping kho],0) AS Mapping_kho,
            s.NhaMay AS nhamay,
            ISNULL(lt.[KHỐI LƯỢNG HÀNG XUẤT LÊN TÀU], 0) as klyeucau,
            lt.SheetMonth
        FROM vw_so_kho_sumary2 s WITH (NOLOCK)
        JOIN dbo.lichtau lt WITH (NOLOCK)
            ON s.[SO Mapping] = TRY_CAST(lt.[SỐ LỆNH TÁCH] AS BIGINT)
        WHERE lt.ETA_Parsed IS NOT NULL
    ),
    process AS (
        SELECT
            d.*,
            ROUND(
                CASE
                    WHEN d.qty <= 100000 AND (d.Mapping_kho + d.shipped_qty) > 1.25 * d.qty THEN
                        (CASE WHEN ABS(d.shipped_qty - d.qty) < ABS(d.Mapping_kho - d.qty) THEN (d.shipped_qty * 100.0 / NULLIF(d.qty,0)) ELSE (d.Mapping_kho * 100.0 / NULLIF(d.qty,0)) END)
                    WHEN d.qty > 100000 AND (d.Mapping_kho + d.shipped_qty) > 1.1 * d.qty THEN
                        (CASE WHEN ABS(d.shipped_qty - d.qty) < ABS(d.Mapping_kho - d.qty) THEN (d.shipped_qty * 100.0 / NULLIF(d.qty,0)) ELSE (d.Mapping_kho * 100.0 / NULLIF(d.qty,0)) END)
                    ELSE ((d.Mapping_kho + d.shipped_qty) * 100.0 / NULLIF(d.qty,0))
                END, 2
            ) AS process_value
        FROM data d
    )
    SELECT
        p.tau,
        p.SheetMonth,
        p.ETA_Parsed,
        p.saleO,
        p.material,
        p.qty,
        p.nhamay,
        p.shipped_qty,
        p.Mapping_kho,
        p.process_value,
        p.klyeucau
    FROM process p
    """)
    with engine.connect() as conn:
        rows = conn.execute(sql).mappings().all()

    # --- BỘ LỌC THÉP: XỬ LÝ DỮ LIỆU BẨN TỪ GỐC ---
    result = []
    for r in rows:
        d = dict(r)
        eta = d.get('ETA_Parsed')
        if isinstance(eta, str):
            try:
                parsed_date = parser.parse(eta[:10]).date() if len(eta) >= 10 else parser.parse(eta).date()
                d['ETA_Parsed'] = parsed_date if parsed_date.year > 2000 else None
            except Exception:
                d['ETA_Parsed'] = None
        elif isinstance(eta, datetime):
            d['ETA_Parsed'] = eta.date() if eta.year > 2000 else None
            
        result.append(d)
        
    return result

def calculate_chart_data(so_details, summary_records):
    """Tính toán dữ liệu tổng hợp cho các biểu đồ từ chi tiết SO."""
    if not so_details:
        return {
            "factory_production": {"labels": [], "datasets": []},
            "ship_status": {"labels": ["Đủ dung sai", "Chưa đủ dung sai"], "data": [0, 0]},
            "delivery_trend": {"labels": [], "data": []}
        }

    # 1. Biểu đồ cột xếp chồng: Tình trạng vật tư theo Nhà máy (Đủ dung sai vs. Chưa đủ dung sai)
    factory_data = {
        "HRC1": {'delivered': 0, 'remaining': 0},
        "HRC2": {'delivered': 0, 'remaining': 0}
    }
    processed_so_materials = set() # Set để theo dõi các cặp (SO, Material) đã xử lý

    for detail in so_details:
        factory = detail.get('nhamay', 'Khác')
        so_number = detail.get('saleO')
        material_name = detail.get('material')
        unique_key = (so_number, material_name)

        if unique_key not in processed_so_materials:
            if factory not in factory_data:
                continue

            total_qty_kg = detail.get('qty', 0)
            delivered_kg = detail.get('shipped_qty', 0) + detail.get('Mapping_kho', 0)
            remaining_kg = total_qty_kg - delivered_kg
            process_val = detail.get('process_value', 0)

            # Xác định ngưỡng dung sai: 80% cho vật tư <= 100 tấn, 90% cho vật tư > 100 tấn
            success_threshold = 80 if total_qty_kg <= 100000 else 90

            # Luôn cộng tổng khối lượng đã sản xuất vào 'delivered'
            factory_data[factory]['delivered'] += delivered_kg
            
            # Chỉ cộng khối lượng còn thiếu nếu vật tư chưa đạt ngưỡng dung sai
            if process_val < success_threshold and remaining_kg > 0:
                factory_data[factory]['remaining'] += remaining_kg
                
            processed_so_materials.add(unique_key)

    factory_production_chart = {
        "labels": list(factory_data.keys()),
        "datasets": [
            {"label": "Đã có hàng", "data": [int(d['delivered'] / 1000) for d in factory_data.values()]},
            {"label": "Còn thiếu", "data": [int(d['remaining'] / 1000) for d in factory_data.values()]}
        ]
    }

    # 2. Biểu đồ cột: Tình trạng các Tàu & 3. Biểu đồ đường: Xu hướng giao hàng
    ship_status_data = {"Đủ dung sai": set(), "Chưa đủ dung sai": set()}

    for detail in so_details:
        ship_name = detail.get('tau')
        if ship_name:
            total_qty_kg = detail.get('qty', 0)
            process_val = detail.get('process_value', 0)
            success_threshold = 80 if total_qty_kg <= 100000 else 90

            if process_val < success_threshold:
                ship_status_data["Chưa đủ dung sai"].add(ship_name)
            else:
                ship_status_data["Đủ dung sai"].add(ship_name)

    # --- LOGIC MỚI CHO BIỂU ĐỒ ĐƯỜNG ---
    # Tính tổng khối lượng tàu theo ngày từ dữ liệu summary đã có
    delivery_trend_data = {} # Dùng dict để nhóm theo ngày
    for record in summary_records:
        eta_date = record.get('ETA_Parsed')
        if eta_date:
            date_str = eta_date.strftime('%Y-%m-%d')
            total_ship_tons = record.get('khoi_luong_tong', 0) # Lấy tổng khối lượng tàu
            delivery_trend_data[date_str] = delivery_trend_data.get(date_str, 0) + total_ship_tons

    # Đảm bảo tàu "Chưa đủ" sẽ ghi đè "Đủ"
    ships_ok = ship_status_data["Đủ dung sai"] - ship_status_data["Chưa đủ dung sai"]
    ship_status_chart = {
        "labels": ["Đủ dung sai", "Chưa đủ dung sai"],
        "data": [len(ships_ok), len(ship_status_data["Chưa đủ dung sai"])]
    }

    # Sắp xếp dữ liệu biểu đồ đường theo ngày
    sorted_delivery_trend = sorted(delivery_trend_data.items())
    delivery_trend_chart = {
        "labels": [item[0] for item in sorted_delivery_trend],
        "data": [int(item[1]) for item in sorted_delivery_trend] # Dữ liệu đã là tấn và số nguyên
    }

    return {
        "factory_production": factory_production_chart,
        "ship_status": ship_status_chart,
        "delivery_trend": delivery_trend_chart
    }

@dashboard_bp.route("/dashboard")
@permission_required('view_ship_schedule') # Đảm bảo người dùng có quyền xem lịch tàu
def dashboard():
    # 1. Lấy dữ liệu tổng hợp cho thanh tiến độ chính
    all_records_summary = get_rows_from_db()

    # Tạo danh sách các tháng duy nhất để lọc, sắp xếp giảm dần
    sheetmonth_list = sorted(list(set(r['SheetMonth'] for r in all_records_summary if r.get('SheetMonth'))), reverse=True)
    selected_status = request.args.get("status", "")
    # Lấy tháng được chọn từ URL. Nếu không có (lần đầu truy cập), mặc định là "Tất cả tháng" (chuỗi rỗng).
    selected_month = request.args.get("sheetmonth")
    if selected_month is None: # Chỉ đặt mặc định khi không có tham số trên URL (lần đầu tải trang)
        selected_month = "" # Mặc định là "Tất cả tháng"

    # Lấy bộ lọc ngày từ URL
    selected_start_date_str = request.args.get("start_date")
    selected_end_date_str = request.args.get("end_date")

    # Lấy tàu được chọn từ URL
    selected_ship = request.args.get("tau")

    # Lọc các bản ghi theo SheetMonth đã chọn
    if selected_month: # Nếu selected_month có giá trị (không phải chuỗi rỗng)
        records_by_month = [r for r in all_records_summary if r.get('SheetMonth') == selected_month]
    else: # Nếu người dùng chọn "Tất cả tháng" (selected_month là chuỗi rỗng)
        records_by_month = all_records_summary
        
    # Tạo danh sách tàu dựa trên các bản ghi của tháng đã chọn
    tau_list = sorted(list(set(r['tau'] for r in records_by_month if r.get('tau'))))
    # Đảm bảo selected_ship luôn là một chuỗi để so sánh nhất quán
    selected_ship = str(selected_ship) if selected_ship is not None else ""

    # Lọc các bản ghi theo ngày trước khi lọc theo tàu
    records_filtered_by_date = records_by_month
    if selected_start_date_str:
        filter_start_date = datetime.strptime(selected_start_date_str, '%Y-%m-%d').date()
        records_filtered_by_date = [r for r in records_filtered_by_date if r.get('ETA_Parsed') and r.get('ETA_Parsed').date() >= filter_start_date]

    if selected_end_date_str:
        filter_end_date = datetime.strptime(selected_end_date_str, '%Y-%m-%d').date()
        records_filtered_by_date = [r for r in records_filtered_by_date if r.get('ETA_Parsed') and r.get('ETA_Parsed').date() <= filter_end_date]

    # Sau khi đã lọc theo ngày, mới lọc theo tàu (nếu có)
    # Điều này đảm bảo chỉ các tàu trong khoảng ngày đã chọn mới được hiển thị
    if selected_ship:
        # Đảm bảo r.get('tau') cũng được xử lý dưới dạng chuỗi cho việc so sánh
        records = [r for r in records_filtered_by_date if str(r.get('tau', '')) == selected_ship]
    else: # Nếu selected_ship là chuỗi rỗng (ví dụ: "Tất cả Tàu" được chọn)
        records = records_filtered_by_date

    # Sắp xếp lại các bản ghi theo ETA để thứ tự tàu được sắp xếp đúng
    records.sort(key=lambda r: r.get('ETA_Parsed') or datetime.max.date())
    # Xác định dải ngày cho các cột của bảng
    date_range = []
    if records:
        valid_dates = []
        for r in records:
            eta = r.get('ETA_Parsed')
            if not eta: 
                continue
            
            # Xử lý ép kiểu triệt để, loại bỏ năm lỗi (như 0202)
            if isinstance(eta, str):
                try:
                    # Ép thành datetime.date
                    dt = datetime.strptime(eta[:10], '%Y-%m-%d').date()
                    if dt.year > 2000:  # Chỉ lấy những năm hợp lý (bỏ qua năm 0202)
                        valid_dates.append(dt)
                except ValueError:
                    pass
            elif isinstance(eta, datetime):
                if eta.year > 2000:
                    valid_dates.append(eta.date())
            else:
                # Nếu đã là kiểu date sẵn
                if getattr(eta, 'year', 0) > 2000:
                    valid_dates.append(eta)

        if valid_dates:
            start_date = min(valid_dates)
            end_date = max(valid_dates)
            
            current_date = start_date
            while current_date <= end_date:
                date_range.append(current_date)
                current_date += timedelta(days=1)

    # Áp dụng bộ lọc ngày nếu có
    # if selected_start_date_str:
    #     filter_start_date = datetime.strptime(selected_start_date_str, '%Y-%m-%d')
    #     date_range = [d for d in date_range if d >= filter_start_date]
    # 
    # if selected_end_date_str:
    #     filter_end_date = datetime.strptime(selected_end_date_str, '%Y-%m-%d')
    #     date_range = [d for d in date_range if d <= filter_end_date]


    # Cấu trúc lại dữ liệu để template dễ dàng render
    # Dạng: { 'Tên Tàu': { 'YYYY-MM-DD': { data }, 'YYYY-MM-DD': { data } }, ... }
    ships_data = OrderedDict()
    for r in records: # Dùng `records` đã được lọc đầy đủ
        ship_name = r.get('tau')
        eta = r.get('ETA_Parsed')
        
        # Bỏ qua nếu thiếu thông tin cần thiết
        if not ship_name or not eta:
            continue

        # Khởi tạo dictionary cho tàu nếu chưa có
        if ship_name not in ships_data:
            ships_data[ship_name] = {}

        # Khởi tạo dữ liệu cho ngày nếu chưa có
        eta_str = eta.strftime('%Y-%m-%d')
        
        # Tính toán tiến độ và màu sắc
        tong_kl = r.get('tongkhoiluong') or 0
        kl_tong = r.get('khoi_luong_tong') or 0
        percentage = (tong_kl / kl_tong * 100) if kl_tong > 0 else 0

        color = "bg-danger" # Mặc định là màu đỏ (tiến độ thấp)
        if percentage >= 95:
            color = "bg-success" # Xanh lá (gần hoàn thành)
        elif percentage >= 75:
            color = "bg-warning" # Vàng (tiến độ khá)

        # Gán dữ liệu vào đúng ngày ETA của tàu
        ships_data[ship_name][eta_str] = {
            "tongkhoiluong": int(tong_kl),
            "khoi_luong_tong": int(kl_tong),
            "percentage": percentage,
            "color": color,
            "so_details": OrderedDict(), # Chuẩn bị chỗ để chứa chi tiết SO
            "has_underperforming_item": False # Flag để kiểm tra có item nào < 90% không
        }

    # 2. Lấy dữ liệu chi tiết SO và mapping vào ships_data
    all_so_details = get_so_details_for_dashboard()
    # Lọc chi tiết SO theo tháng đã chọn để tối ưu
    if selected_month:
        so_details_by_month = [r for r in all_so_details if r.get('SheetMonth') == selected_month]
    else:
        so_details_by_month = all_so_details
        
    # ==============================================================================
    # 🔹 BƯỚC 2.1: LOGIC PHÂN BỔ LẠI SẢN LƯỢNG SO CHO CÁC TÀU THEO ETA
    # ==============================================================================
    # 🔹 BƯỚC 2.1: LOGIC PHÂN BỔ LẠI SẢN LƯỢNG SO CHO CÁC TÀU THEO ETA
    # ==============================================================================
    so_total_available = {} # { so_id: total_qty }
    so_to_ships_map = {}  # { so_id: [ {ship_name, eta, klyeucau, allocated_qty} ] }

    # --- 2.1.0: TÍNH TỔNG KLYEUCAU ---
    so_klyeucau_total_per_ship = {}
    for so_detail in so_details_by_month:
        so_id = so_detail.get('saleO')
        ship_name = so_detail.get('tau')
        eta = so_detail.get('ETA_Parsed')
        if not all([so_id, ship_name, eta]): continue
        
        key = (so_id, ship_name, eta)
        if key not in so_klyeucau_total_per_ship:
            so_klyeucau_total_per_ship[key] = so_detail.get('klyeucau', 0) * 1000

    # --- THUẬT TOÁN WATERFALL (MYSQL) ---
    active_ships = list(set(str(r.get('tau', '')).strip() for r in records if r.get('tau')))
    mysql_exported_data = get_actual_exported_from_mysql(active_ships)
    actual_waterfall_allocation = {}
    trips_map = {}

    for (so_id, ship_name, eta), req_kg in so_klyeucau_total_per_ship.items():
        trip_key = (ship_name.upper().strip(), str(so_id).lstrip('0').upper().strip())
        if trip_key not in trips_map: trips_map[trip_key] = {}
        trips_map[trip_key][eta] = req_kg

    for (tau, so), trips in trips_map.items():
        total_actual_kg = mysql_exported_data.get((tau, so), 0.0)
        sorted_etas = sorted([e for e in trips.keys() if e is not None])
        for eta in sorted_etas:
            req_kg = trips[eta]
            allocated = min(total_actual_kg, req_kg)
            actual_waterfall_allocation[(so, tau, eta)] = allocated
            total_actual_kg -= allocated
            if total_actual_kg <= 0: break

    # --- 2.1.1: GÁN DỮ LIỆU VÀO TOOLTIP VÀ PHÂN BỔ KHO ---
    unique_so_materials = set()
    
    for so_detail in so_details_by_month:
        ship_name = so_detail.get('tau')
        eta = so_detail.get('ETA_Parsed')
        if not ship_name or not eta: continue
        
        eta_str = eta.strftime('%Y-%m-%d')
        if ship_name in ships_data and eta_str in ships_data[ship_name]:
            sale_order = so_detail.get('saleO')
            material_name = so_detail.get('material')
            key = (sale_order, ship_name, eta)
            total_klyeucau_for_so_kg = so_klyeucau_total_per_ship.get(key, 0)
            
            # --- KHỞI TẠO SO ---
            if sale_order not in ships_data[ship_name][eta_str]['so_details']:
                waterfall_key = (str(sale_order).upper().strip(), ship_name.upper().strip(), eta)
                thucte_xuat_kg = actual_waterfall_allocation.get(waterfall_key, 0.0)

                is_exported_finished = False
                if total_klyeucau_for_so_kg > 0:
                    if thucte_xuat_kg >= (total_klyeucau_for_so_kg * 0.9):
                        is_exported_finished = True

                ships_data[ship_name][eta_str]['so_details'][sale_order] = {
                    'has_underperforming_material': False,
                    'is_exported_finished': is_exported_finished,
                    'processed_materials': set(),
                    'summary': {
                        'total_qty_kg': 0,
                        'delivered_kg': 0,
                        'progress_percent': 0,
                        'progress_text': '0 / 0',
                        'klyeucau': total_klyeucau_for_so_kg,
                        'so_phan_bo': 0,
                        'thucte_xuat_tan': int(thucte_xuat_kg / 1000)
                    },
                    'materials': []
                }
            
            current_so_data = ships_data[ship_name][eta_str]['so_details'][sale_order]

            # --- XỬ LÝ VẬT TƯ ---
            if material_name not in current_so_data['processed_materials']:
                delivered_kg = so_detail.get('shipped_qty', 0) + so_detail.get('Mapping_kho', 0)
                total_qty_kg = so_detail.get('qty', 0)
                process_val = so_detail.get('process_value', 0)
                success_threshold = 80 if total_qty_kg <= 100000 else 90

                if process_val < success_threshold:
                    current_so_data['has_underperforming_material'] = True
                    ships_data[ship_name][eta_str]['has_underperforming_item'] = True

                current_so_data['materials'].append({
                    "name": material_name,
                    "progress_percent": process_val,
                    "progress_text": f"{int(delivered_kg / 1000)} / {int(total_qty_kg / 1000)}",
                    "qty_kg": total_qty_kg
                })

                so_summary = current_so_data['summary']
                so_summary['total_qty_kg'] += total_qty_kg
                so_summary['delivered_kg'] += delivered_kg
                current_so_data['processed_materials'].add(material_name)

            # --- GOM NHÓM ĐỂ PHÂN BỔ KHO ---
            if (sale_order, material_name) not in unique_so_materials:
                available_qty = so_detail.get('shipped_qty', 0) + so_detail.get('Mapping_kho', 0)
                so_total_available[sale_order] = so_total_available.get(sale_order, 0) + available_qty
                unique_so_materials.add((sale_order, material_name))

            if sale_order not in so_to_ships_map:
                so_to_ships_map[sale_order] = []
            
            if not any(s['ship_name'] == ship_name for s in so_to_ships_map[sale_order]):
                 so_to_ships_map[sale_order].append({
                    "ship_name": ship_name,
                    "eta": eta,
                    "klyeucau": total_klyeucau_for_so_kg,
                    "allocated_qty": 0
                })

    # --- 2.1.2: THỰC HIỆN PHÂN BỔ KHO ---
    for so_id, ships in so_to_ships_map.items():
        remaining_qty = so_total_available.get(so_id, 0)
        
        # BƯỚC A: Trích xuất quỹ hàng trong kho để "trả nợ" cho các hàng đã thực tế lên tàu
        for ship_info in ships:
            # Lấy số lượng đã thực sự xuất lên tàu này (từ MySQL)
            waterfall_key = (str(so_id).upper().strip(), ship_info['ship_name'].upper().strip(), ship_info['eta'])
            thucte_xuat_kg = actual_waterfall_allocation.get(waterfall_key, 0.0)
            
            # Gán cứng: Phân bổ kho ít nhất phải bằng số thực tế đã xuất
            giao_truoc = min(remaining_qty, thucte_xuat_kg)
            ship_info['allocated_qty'] = giao_truoc
            remaining_qty -= giao_truoc
            
            # Tính lại khoảng trống còn lại của tàu (nếu tàu cần 1800, đã xuất 1796 -> còn cần 4)
            ship_info['klyeucau_con_lai'] = max(0, ship_info['klyeucau'] - thucte_xuat_kg)

        # BƯỚC B: Nếu kho vẫn còn dư hàng, tiếp tục phân bổ phần dư đó theo ETA (Waterfall truyền thống)
        ships.sort(key=lambda x: x['eta'])
        for ship_info in ships:
            if remaining_qty <= 0:
                break
            
            allocated_them = min(remaining_qty, ship_info['klyeucau_con_lai'])
            ship_info['allocated_qty'] += allocated_them
            remaining_qty -= allocated_them
    # 🔹 BƯỚC 2.2: TÍNH LẠI TỔNG KHỐI LƯỢNG TÀU DỰA TRÊN SẢN LƯỢNG ĐÃ PHÂN BỔ
    # ==============================================================================
    for ship_name, dates in ships_data.items():
        for eta_str, data in dates.items():
            new_total_delivered_for_ship_kg = 0
            # Lặp qua tất cả các SO trong tooltip của tàu này
            for so_number in data['so_details']:
                # Tìm sản lượng đã phân bổ cho cặp SO-Tàu này
                if so_number in so_to_ships_map:
                    for ship_info in so_to_ships_map[so_number]:
                        if ship_info['ship_name'] == ship_name:
                            new_total_delivered_for_ship_kg += ship_info['allocated_qty']
                            break # Đã tìm thấy, chuyển sang SO tiếp theo
            
            # Cập nhật lại tổng khối lượng tàu (tấn) và phần trăm
            data['tongkhoiluong'] = int(new_total_delivered_for_ship_kg / 1000)
            data['percentage'] = (data['tongkhoiluong'] / data['khoi_luong_tong'] * 100) if data['khoi_luong_tong'] > 0 else 0

    # --- Vòng lặp cuối để tính toán % tổng hợp cho SO và màu sắc cho Tàu ---
    for ship_name, dates in ships_data.items():
        for eta_str, data in dates.items():
            all_so_finished = True
            has_any_so = False
            
            # 1. Tính toán chi tiết từng SO
            for so_number, so_data in data['so_details'].items():
                has_any_so = True
                if not so_data['is_exported_finished']:
                    all_so_finished = False
                    
                summary = so_data['summary']
                allocated_for_this_ship = 0
                if so_number in so_to_ships_map:
                    for ship_info in so_to_ships_map[so_number]:
                        if ship_info['ship_name'] == ship_name:
                            allocated_for_this_ship = ship_info['allocated_qty']
                            break 
                
                summary['so_phan_bo'] = int(allocated_for_this_ship / 1000)
                summary['progress_percent'] = (summary['delivered_kg'] * 100.0 / summary['total_qty_kg']) if summary['total_qty_kg'] > 0 else 0
                summary['progress_text'] = f"{int(summary['delivered_kg'] / 1000)} / {int(summary['total_qty_kg'] / 1000)}"
                
            # 2. XÉT MÀU CHO TÀU (Nằm NGOÀI vòng lặp SO, ngang hàng với for so_number)
            if has_any_so and all_so_finished:
                data['color'] = 'bg-info'
            elif data['has_underperforming_item']:
                data['color'] = 'bg-warning'
            else:
                data['color'] = 'bg-success'
    
    # Trả về template với các dữ liệu đã xử lý
    filtered_ships_data = OrderedDict()
    
    for ship_name, dates in ships_data.items():
        filtered_dates = {}
        for eta_str, data in dates.items():
            # Xác định tàu đã xong chưa dựa vào màu bg-info mà ta vừa gán ở trên
            is_done = (data.get('color') == 'bg-info')

            # Kiểm tra điều kiện lọc
            if selected_status == 'done' and not is_done:
                continue # Nếu user chọn Đã xong, bỏ qua các tàu chưa xong
            if selected_status == 'pending' and is_done:
                continue # Nếu user chọn Chưa xong, bỏ qua các tàu đã xong

            # Nếu thỏa mãn điều kiện lọc, đưa vào dict mới
            filtered_dates[eta_str] = data

        # Nếu tàu này còn ít nhất 1 chuyến thỏa mãn điều kiện thì mới giữ lại hiển thị
        if filtered_dates:
            filtered_ships_data[ship_name] = filtered_dates

    # Truyền thêm selected_status và dùng filtered_ships_data thay cho ships_data
    return render_template("dashboard.html",
                           sheetmonth_list=sheetmonth_list,
                           selected_month=selected_month,
                           tau_list=tau_list,
                           selected_ship=selected_ship,
                           selected_start_date=selected_start_date_str,
                           selected_end_date=selected_end_date_str,
                           selected_status=selected_status, # <--- Biến mới
                           date_range=date_range,
                           ships_data=filtered_ships_data)

@dashboard_bp.route("/api/dashboard-charts")
@permission_required('view_ship_schedule') # Đảm bảo người dùng có quyền xem lịch tàu
def api_dashboard_charts():
    """API cung cấp dữ liệu cho biểu đồ, có áp dụng bộ lọc."""
    selected_month = request.args.get("sheetmonth")
    selected_ship = request.args.get("tau") # Sửa lỗi typo gert -> get
    selected_start_date_str = request.args.get("start_date")
    selected_end_date_str = request.args.get("end_date")

    # Lấy cả hai nguồn dữ liệu
    all_so_details = get_so_details_for_dashboard()
    all_records_summary = get_rows_from_db()
    
    # Áp dụng các bộ lọc tương tự như route dashboard
    if selected_month:
        all_so_details = [r for r in all_so_details if r.get('SheetMonth') == selected_month]
        all_records_summary = [r for r in all_records_summary if r.get('SheetMonth') == selected_month]

    if selected_ship:
        all_so_details = [r for r in all_so_details if str(r.get('tau', '')) == selected_ship]
        all_records_summary = [r for r in all_records_summary if str(r.get('tau', '')) == selected_ship]

    if selected_start_date_str:
        filter_start_date = datetime.strptime(selected_start_date_str, '%Y-%m-%d').date()
        all_so_details = [r for r in all_so_details if r.get('ETA_Parsed') and r.get('ETA_Parsed').date() >= filter_start_date]
        all_records_summary = [r for r in all_records_summary if r.get('ETA_Parsed') and r.get('ETA_Parsed').date() >= filter_start_date]

    if selected_end_date_str:
        filter_end_date = datetime.strptime(selected_end_date_str, '%Y-%m-%d').date()
        all_so_details = [r for r in all_so_details if r.get('ETA_Parsed') and r.get('ETA_Parsed').date() <= filter_end_date]
        all_records_summary = [r for r in all_records_summary if r.get('ETA_Parsed') and r.get('ETA_Parsed').date() <= filter_end_date]

    # Gán `records` để truyền vào hàm tính toán
    records = all_records_summary
    chart_data = calculate_chart_data(all_so_details, records)
     # Truyền `records` đã lọc vào hàm
    hrc_planned_daily = get_hrc_planned_schedule(selected_ship)
    
    actual_labels = chart_data["delivery_trend"]["labels"]
    actual_dict = dict(zip(actual_labels, chart_data["delivery_trend"]["data"] if "data" in chart_data["delivery_trend"] else []))
    
    all_dates_set = set(actual_labels)
    for d in hrc_planned_daily.keys():
        all_dates_set.add(d.strftime('%Y-%m-%d'))
        
    all_dates_sorted = sorted(list(all_dates_set))
    final_labels, final_actual, final_planned = [], [], []
    
    for date_str in all_dates_sorted:
        date_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
        
        # 2. KIỂM TRA BỘ LỌC NGÀY TỪ/ĐẾN
        if selected_start_date_str and date_obj < datetime.strptime(selected_start_date_str, '%Y-%m-%d').date():
            continue
        if selected_end_date_str and date_obj > datetime.strptime(selected_end_date_str, '%Y-%m-%d').date():
            continue
            
        # 3. KIỂM TRA BỘ LỌC THÁNG (Format: "03.2026")
        if selected_month:
            try:
                # Cắt chuỗi "03.2026" thành tháng 3, năm 2026
                m, y = selected_month.split('.')
                if date_obj.month != int(m) or date_obj.year != int(y):
                    continue
            except ValueError:
                pass # Nếu lỡ format tháng bị sai thì bỏ qua việc cắt

        actual_val = actual_dict.get(date_str, 0)
        planned_val = int(hrc_planned_daily.get(date_obj, 0))
        
        # Chỉ hiển thị ngày nào có làm hàng hoặc có kế hoạch (xóa các ngày trống làm biểu đồ bị loãng)
        if actual_val > 0 or planned_val > 0 or date_str in actual_labels:
            final_labels.append(date_str)
            final_actual.append(actual_val)
            final_planned.append(planned_val)
        
    chart_data["delivery_trend"]["labels"] = final_labels
    chart_data["delivery_trend"]["actual_data"] = final_actual
    chart_data["delivery_trend"]["planned_data"] = final_planned
    
    if "data" in chart_data["delivery_trend"]:
        del chart_data["delivery_trend"]["data"]
    if "datasets" in chart_data["delivery_trend"]:
        del chart_data["delivery_trend"]["datasets"]

    return jsonify(chart_data)

@dashboard_bp.route("/api/dashboard/missing-details")
@permission_required('view_ship_schedule')
def api_missing_details():
    """API để lấy chi tiết các vật tư còn thiếu cho một nhà máy cụ thể."""
    factory = request.args.get("factory")
    if not factory:
        return jsonify({"error": "Factory parameter is required"}), 400

    # Lấy các bộ lọc khác
    selected_month = request.args.get("sheetmonth")
    selected_ship = request.args.get("tau")
    selected_start_date_str = request.args.get("start_date")
    selected_end_date_str = request.args.get("end_date")

    all_so_details = get_so_details_for_dashboard()

    # Áp dụng các bộ lọc tương tự như các API khác
    if selected_month:
        all_so_details = [r for r in all_so_details if r.get('SheetMonth') == selected_month]
    if selected_ship:
        all_so_details = [r for r in all_so_details if str(r.get('tau', '')) == selected_ship]
    if selected_start_date_str:
        filter_start_date = datetime.strptime(selected_start_date_str, '%Y-%m-%d').date()
        all_so_details = [r for r in all_so_details if r.get('ETA_Parsed') and r.get('ETA_Parsed').date() >= filter_start_date]
    if selected_end_date_str:
        filter_end_date = datetime.strptime(selected_end_date_str, '%Y-%m-%d').date()
        all_so_details = [r for r in all_so_details if r.get('ETA_Parsed') and r.get('ETA_Parsed').date() <= filter_end_date]

    # Lọc theo nhà máy và điều kiện "còn thiếu"
    missing_items = []
    for r in all_so_details:
        if r.get('nhamay') == factory:
            total_qty_kg = r.get('qty', 0)
            process_val = r.get('process_value', 0)
            success_threshold = 80 if total_qty_kg <= 100000 else 90
            
            if process_val < success_threshold:
                remaining_kg = total_qty_kg - (r.get('shipped_qty', 0) + r.get('Mapping_kho', 0))
                if remaining_kg > 0:
                    missing_items.append({
                        "tau": r.get('tau'),
                        "eta": r.get('ETA_Parsed').strftime('%d/%m/%Y') if r.get('ETA_Parsed') else 'N/A',
                        "so": r.get('saleO'),
                        "material": r.get('material'),
                        "missing_tons": int(remaining_kg / 1000)
                    })
    missing_items.sort(key=lambda x: (        
        datetime.strptime(x['eta'], '%d/%m/%Y') if x['eta'] != 'N/A' else datetime.max,
        x.get('tau', ''),
        x.get('so', 0)
    ))
    
    return jsonify(missing_items)
