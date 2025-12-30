from flask import Blueprint, render_template, request, jsonify
from sqlalchemy import text
from db import engine
from datetime import datetime, timedelta
from dateutil import parser
from auth.decorator import permission_required
import io
import pandas as pd
from flask import send_file
from collections import OrderedDict

dashboard_bp = Blueprint("dashboard_bp", __name__)
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
    FROM testdb
    WHERE
        tau != N'ĐƯỜNG BỘ' AND tau != ''
    GROUP BY
            tau, SheetMonth, ETA_Parsed
    """)
    with engine.connect() as conn:
        rows = conn.execute(sql).mappings().all()
    return [dict(r) for r in rows]

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
        FROM vw_so_kho_sumary2 s
        JOIN dbo.lichtau lt
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
    return [dict(r) for r in rows]

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
    records.sort(key=lambda r: r.get('ETA_Parsed') or datetime.max)
    # Xác định dải ngày cho các cột của bảng
    date_range = []
    if records:
        # Lấy tất cả các ngày ETA_Parsed hợp lệ từ các bản ghi đã lọc
        eta_dates = [r['ETA_Parsed'] for r in records if r.get('ETA_Parsed')]
        
        if eta_dates:
            # Tìm ngày ETA sớm nhất và muộn nhất. Điều này sẽ tự động xử lý trường hợp
            # ETA thuộc tháng trước (ví dụ: tháng 9) trong khi SheetMonth là tháng 10.
            start_date = min(eta_dates)
            end_date = max(eta_dates)
            
            # Tạo danh sách các ngày liên tục từ ngày bắt đầu đến ngày kết thúc
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
    so_total_available = {} # { so_id: total_qty }
    so_to_ships_map = {}  # { so_id: [ {ship_name, eta, klyeucau, allocated_qty} ] }

    # --- 2.1.0: TÍNH TỔNG KLYEUCAU CHO MỖI (TÀU, SO) TRƯỚC TIÊN ---
    # Điều này cực kỳ quan trọng để đảm bảo klyeucau được tính đúng trước khi phân bổ
    so_klyeucau_total_per_ship = {} # {(so, ship, eta): total_klyeucau_in_kg}
    
    for so_detail in so_details_by_month:
        so_id = so_detail.get('saleO')
        ship_name = so_detail.get('tau')
        eta = so_detail.get('ETA_Parsed')
        if not all([so_id, ship_name, eta]):
            continue
        key = (so_id, ship_name, eta)
        # Chỉ lấy giá trị klyeucau đại diện, không cộng dồn.
        # Nếu key chưa tồn tại, gán giá trị klyeucau (đã đổi sang kg).
        if key not in so_klyeucau_total_per_ship:
            so_klyeucau_total_per_ship[key] = so_detail.get('klyeucau', 0) * 1000

    # --- 2.1.1: Tính tổng sản lượng đã có cho mỗi SO và gom nhóm tàu theo SO ---
    unique_so_materials = set()
    for so_detail in so_details_by_month:
        so_id = so_detail.get('saleO')
        material = so_detail.get('material')
        ship_name = so_detail.get('tau')
        eta = so_detail.get('ETA_Parsed')

        if not so_id or not ship_name or not eta:
            continue

        # Tính tổng sản lượng có sẵn (chỉ cộng một lần cho mỗi cặp SO-Material)
        if (so_id, material) not in unique_so_materials:
            available_qty = so_detail.get('shipped_qty', 0) + so_detail.get('Mapping_kho', 0)
            so_total_available[so_id] = so_total_available.get(so_id, 0) + available_qty
            unique_so_materials.add((so_id, material))

        # Gom nhóm các tàu cho mỗi SO
        if so_id not in so_to_ships_map:
            so_to_ships_map[so_id] = []
        
        # Lấy tổng klyeucau đã tính toán cho cặp (Tàu, SO) này
        ship_so_key = (so_id, ship_name, eta)
        total_klyeucau_for_this_ship_so = so_klyeucau_total_per_ship.get(ship_so_key, 0)

        # Đảm bảo không thêm trùng lặp tàu cho một SO
        if not any(s['ship_name'] == ship_name for s in so_to_ships_map[so_id]):
             so_to_ships_map[so_id].append({
                "ship_name": ship_name,
                "eta": eta,
                "klyeucau": total_klyeucau_for_this_ship_so, # Sử dụng tổng đã tính (đơn vị là KG)
                "allocated_qty": 0 # Khởi tạo
            })

    # --- 2.1.2: Thực hiện phân bổ ---
    for so_id, ships in so_to_ships_map.items():
        remaining_qty = so_total_available.get(so_id, 0)
        # Sắp xếp các tàu theo ETA tăng dần
        ships.sort(key=lambda x: x['eta'])
        # `remaining_qty` và `ship_info['klyeucau']` đều đang ở đơn vị KG
        for ship_info in ships:
            allocated = min(remaining_qty, ship_info['klyeucau'])
            ship_info['allocated_qty'] = allocated
            remaining_qty -= allocated

    for so_detail in so_details_by_month:
        ship_name = so_detail.get('tau')
        eta = so_detail.get('ETA_Parsed')

        if not ship_name or not eta:
            continue
        
        eta_str = eta.strftime('%Y-%m-%d')

        # Kiểm tra xem tàu và ngày có tồn tại trong ships_data không
        if ship_name in ships_data and eta_str in ships_data[ship_name]:
            sale_order = so_detail.get('saleO')
            material_name = so_detail.get('material')
            
            # Lấy tổng klyeucau đã tính toán
            key = (sale_order, ship_name, eta)
            # Giá trị này đã được tính tổng và chuyển sang KG ở bước 2.1.0
            total_klyeucau_for_so_kg = so_klyeucau_total_per_ship.get(key, 0)
            # Khởi tạo SO trong tooltip nếu chưa có
            if sale_order not in ships_data[ship_name][eta_str]['so_details']:
                ships_data[ship_name][eta_str]['so_details'][sale_order] = {
                    # Flag để kiểm tra SO này có material nào < 90% không
                    'has_underperforming_material': False,
                    'summary': {
                        'total_qty_kg': 0,
                        'delivered_kg': 0,
                        'progress_percent': 0,
                        'progress_text': '0 / 0',
                        'klyeucau': total_klyeucau_for_so_kg, # Lưu tổng KL yêu cầu của SO trên tàu này (đơn vị KG)
                        'so_phan_bo': 0 # Khởi tạo số SO phân bổ
                    },
                    'materials': []
                }
            
            current_so_data = ships_data[ship_name][eta_str]['so_details'][sale_order]

            # Sử dụng một set để theo dõi các material đã được cộng vào summary
            if 'processed_materials' not in current_so_data:
                current_so_data['processed_materials'] = set()
            
            # --- LOGIC MỚI: Chỉ xử lý (thêm vào list và cộng dồn) cho các material duy nhất ---
            if material_name not in current_so_data['processed_materials']:
                delivered_kg = so_detail.get('shipped_qty', 0) + so_detail.get('Mapping_kho', 0)
                total_qty_kg = so_detail.get('qty', 0)
                process_val = so_detail.get('process_value', 0)

                # Xác định ngưỡng thành công dựa trên khối lượng của material
                success_threshold = 80 if total_qty_kg <= 100000 else 90

                # Cập nhật flag nếu tiến độ material chưa đạt ngưỡng
                if process_val < success_threshold:
                    current_so_data['has_underperforming_material'] = True
                    ships_data[ship_name][eta_str]['has_underperforming_item'] = True

                # Thêm chi tiết material vào danh sách để hiển thị (chỉ một lần)
                current_so_data['materials'].append({
                    "name": material_name,
                    "progress_percent": process_val,
                    "progress_text": f"{int(delivered_kg / 1000)} / {int(total_qty_kg / 1000)}",
                    "qty_kg": total_qty_kg
                })

                # Cộng dồn khối lượng vào summary (chỉ một lần)
                so_summary = current_so_data['summary']
                so_summary['total_qty_kg'] += total_qty_kg
                so_summary['delivered_kg'] += delivered_kg
                current_so_data['processed_materials'].add(material_name)

    # ==============================================================================
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
            # Cập nhật màu cho TÀU dựa trên flag
            if data['has_underperforming_item']:
                data['color'] = 'bg-warning' # Nếu có item < 90%, tàu sẽ màu vàng
            else:
                data['color'] = 'bg-success' # Nếu tất cả item >= 90%, tàu màu xanh

            for so_number, so_data in data['so_details'].items():
                summary = so_data['summary']
                # Lấy sản lượng đã được phân bổ cho cặp SO-Tàu này
                allocated_for_this_ship = 0
                if so_number in so_to_ships_map:
                    for ship_info in so_to_ships_map[so_number]:
                        if ship_info['ship_name'] == ship_name:
                            allocated_for_this_ship = ship_info['allocated_qty']
                            break # Đã tìm thấy, thoát vòng lặp
                
                # Cập nhật số SO đã phân bổ (tấn)
                summary['so_phan_bo'] = int(allocated_for_this_ship / 1000)

                # Giữ nguyên cách tính tiến độ ban đầu của SO
                summary['progress_percent'] = (summary['delivered_kg'] * 100.0 / summary['total_qty_kg']) if summary['total_qty_kg'] > 0 else 0
                summary['progress_text'] = f"{int(summary['delivered_kg'] / 1000)} / {int(summary['total_qty_kg'] / 1000)}"

    # Trả về template với các dữ liệu đã xử lý
    return render_template("dashboard.html",
                           sheetmonth_list=sheetmonth_list,
                           selected_month=selected_month,
                           tau_list=tau_list,
                           selected_ship=selected_ship,
                           selected_start_date=selected_start_date_str,
                           selected_end_date=selected_end_date_str,
                           date_range=date_range,
                           ships_data=ships_data)

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
    chart_data = calculate_chart_data(all_so_details, records) # Truyền `records` đã lọc vào hàm
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
