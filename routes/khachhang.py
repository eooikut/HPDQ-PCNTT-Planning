from flask import Blueprint, render_template, request, jsonify, send_file
from sqlalchemy import text
from collections import OrderedDict
from db import engine
from storage_utils import load_metadata
from auth.decorator import permission_required
import pandas as pd
import io
from datetime import datetime

so_bp = Blueprint("so", __name__)

def get_rows_from_db():
    sql = text("""
WITH Data_with_Process AS (
    SELECT DISTINCT
        s.[Document Date],
        s.[PO.],
        s.Material,
        s.[Customer Number] AS customer_num,
        s.[Customer] AS customer_name,
        s.[Sales Document] AS so_mapping,
        t.[Material Description] AS material_desc, -- Đã đổi tên (tránh trùng lặp)
        s.[Reqd Deliv Date] AS req_date,
        s.[Shipped Quantity (KG)] AS shipped_qty,
        s.[Quantity (KG)] AS qty,
        
        -- Chuyển giá trị Factory
        CASE 
            WHEN s.[Factory] = 'Hòa Phát Dung Quất 2' THEN 'HRC2'
            WHEN s.[Factory] = 'Hòa Phát Dung Quất' THEN 'HRC1'
            ELSE s.[Factory]
        END AS Factory,

        -- Nếu mapping_kho NULL thì đặt = 0
        ISNULL(t.[SL Mapping kho], 0) AS mapping_kho,

        -- Tính lại process_value MỘT LẦN ở đây
        CASE 
            WHEN s.[Quantity (KG)] > 0 THEN
                ROUND(
                    CASE
                        -- Nếu qty nhỏ và vượt 120%
                        WHEN s.[Quantity (KG)] <= 100000 
                            AND (ISNULL(t.[SL Mapping kho],0) + ISNULL(s.[Shipped Quantity (KG)],0)) > 1.2 * s.[Quantity (KG)] THEN
                            CASE 
                                WHEN ABS(ISNULL(s.[Shipped Quantity (KG)],0) - s.[Quantity (KG)]) 
                                    < ABS(ISNULL(t.[SL Mapping kho],0) - s.[Quantity (KG)]) 
                                THEN ISNULL(s.[Shipped Quantity (KG)],0) * 100.0 / s.[Quantity (KG)]
                                ELSE ISNULL(t.[SL Mapping kho],0) * 100.0 / s.[Quantity (KG)]
                            END

                        -- Nếu qty lớn và vượt 110%
                        WHEN s.[Quantity (KG)] > 100000 
                            AND (ISNULL(t.[SL Mapping kho],0) + ISNULL(s.[Shipped Quantity (KG)],0)) > 1.1 * s.[Quantity (KG)] THEN
                            CASE 
                                WHEN ABS(ISNULL(s.[Shipped Quantity (KG)],0) - s.[Quantity (KG)]) 
                                    < ABS(ISNULL(t.[SL Mapping kho],0) - s.[Quantity (KG)]) 
                                THEN ISNULL(s.[Shipped Quantity (KG)],0) * 100.0 / s.[Quantity (KG)]
                                ELSE ISNULL(t.[SL Mapping kho],0) * 100.0 / s.[Quantity (KG)]
                            END

                        -- Trường hợp bình thường
                        ELSE (ISNULL(t.[SL Mapping kho],0) + ISNULL(s.[Shipped Quantity (KG)],0)) * 100.0 / s.[Quantity (KG)]
                    END
                , 2)
            ELSE 0
        END AS process_value
        
    FROM so s
    LEFT JOIN vw_so_kho_sumary2 t
        ON s.[Sales Document] = t.[SO Mapping]
        AND s.[Material] = t.[Material]
)

-- BƯỚC 2: SELECT từ CTE và áp dụng logic MÀU SẮC ĐƠN GIẢN
SELECT
    d.*, -- Lấy tất cả các cột đã tính từ CTE
    
    -- *** LOGIC MÀU SẮC MỚI (Đơn giản và Đúng) ***
    CASE
        -- SO nhỏ (qty < 100T) đạt >= 80% là XANH
        WHEN (d.qty <= 100000 AND d.process_value >= 80) THEN 'bg-success'
        
        -- SO lớn (qty >= 100T) đạt >= 90% là XANH
        WHEN (d.qty > 100000 AND d.process_value >= 90) THEN 'bg-success'
        
        -- SO đạt >= 75% là VÀNG (Bạn có thể đổi ngưỡng này nếu muốn)
        
        -- Còn lại là ĐỎ
        ELSE 'bg-warning'
    END AS process_color

FROM Data_with_Process d;

    """)
    with engine.connect() as conn:
        rows = conn.execute(sql).mappings().all()
    return [dict(r) for r in rows]
#
def filter_and_group(records, keyword="", filter_customer="", filter_factory=""):
    keywords = [k.strip().lower() for k in keyword.split(",") if k.strip()]
    grouped = OrderedDict()

    for r in records:
        so_str  = str(r.get("so_mapping") or "").strip()
        mat_str = (r.get("material_desc") or "").lower()
        cust_str = (r.get("customer_name") or "").lower()
        factory_str = r.get("Factory") or ""

        # lọc theo keyword: nếu keyword không có trong material và cũng không trong so_mapping và customer thì bỏ
        if keywords:
            if not any(
                (
                    k in so_str
                    or k in mat_str
                    or k in cust_str
                )
                for k in keywords
            ):
                continue
        if filter_customer and r.get("customer_name") != filter_customer:
            continue
        if filter_factory and factory_str != filter_factory:
            continue

        cust_key = r["customer_name"]
        so_key   = so_str
        grouped.setdefault(cust_key, OrderedDict()).setdefault(so_key, {"factory": factory_str, "materials": []})["materials"].append(r)

    rows_with_flags = []
    for cust, so_map in grouped.items():
        cust_printed = False
        all_materials = []
        for so, so_data in so_map.items():
            so_printed = False
            factory_name = so_data["factory"]
            mats = so_data["materials"]
            all_materials.extend(so_data["materials"])
            for m in mats:
                is_last_of_customer = (m == all_materials[-1])
                rows_with_flags.append({
                    "customer_num": m["customer_num"] if not cust_printed else None,
                    "customer_name": m["customer_name"] if not cust_printed else None,
                    "cust_rowspan": sum(len(v["materials"]) for v in so_map.values()) if not cust_printed else None,
                    "so_mapping": so if not so_printed else None,
                    "so_rowspan": len(mats) if not so_printed else None,
                    "factory": factory_name if not so_printed else None,  # gắn Factory ở SO-level
                    "material": m["material_desc"],
                    "req_date": "/".join(m["req_date"][:10].split("-")) if m.get("req_date") else None,
                    "mapping_kho": "{:,}".format(int(m["mapping_kho"] or 0) // 1000),
                    "shipped_qty": "{:,}".format(int(m["shipped_qty"] or 0)//1000),
                    "qty": "{:,}".format(int(m["qty"] or 0)//1000),
                    "process_value": m["process_value"],
                    "process_color": m["process_color"],
                    "is_last_row_of_customer": is_last_of_customer,
                })
                cust_printed = True
                so_printed   = True
    return rows_with_flags


@so_bp.route("/so_all")
@permission_required('view_customer')
def so_all():    
    process_name_map = {
    "bg-success": "Hoàn thành",
    "bg-warning": "Chưa hoàn thành",
    "bg-danger": "Thiếu nhiều"
    }

    # Lấy các màu duy nhất từ DB
    
    records = get_rows_from_db()
    factory_map = {
    "Hòa Phát Dung Quất 2": "HRC2",
    "Hòa Phát Dung Quất": "HRC1"
}

# Áp dụng map cho tất cả record
    for r in records:
        f = r.get("Factory")
        if f:
            # loại bỏ khoảng trắng thừa trước khi map
            f_clean = f.strip()
            r["Factory"] = factory_map.get(f_clean, f_clean)
    rows = filter_and_group(records)
    # customer_list = sorted({r["customer_name"] for r in records})
    customer_list = list(OrderedDict.fromkeys([r["customer_name"] for r in rows if r["customer_name"]]))

    unique_colors = sorted({r["process_color"] for r in records})  # thêm
    factory_list  = sorted({r["Factory"] for r in records if r["Factory"]}) 
    process_list = [(color, process_name_map.get(color, color)) for color in unique_colors]
    return render_template(
        "khachhang.html",
        rows=rows,
        customer_list=customer_list,
        factory_list=factory_list,
        process_list=process_list
    )

@so_bp.route("/tiendo_search")
@permission_required('view_customer')
def tiendo_search():
    # 1. Lấy tham số
    keyword = request.args.get("keyword", "")
    customer = request.args.get("customer", "")
    process_color = request.args.get("process_color", "")
    factory = request.args.get("factory", "")

    # 2. Lấy dữ liệu thô & Map nhà máy
    records = get_rows_from_db()
    factory_map = {
        "Hòa Phát Dung Quất 2": "HRC2",
        "Hòa Phát Dung Quất": "HRC1"
    }
    for r in records:
        f = r.get("Factory")
        if f:
            f_clean = f.strip()
            r["Factory"] = factory_map.get(f_clean, f_clean)

    # 3. LỌC (FILTER)
    filtered_records = []
    keywords = [k.strip().lower() for k in keyword.split(",") if k.strip()]

    for r in records:
        # Lọc Keyword
        if keywords:
            so_str = str(r.get("so_mapping") or "").strip().lower()
            mat_str = (r.get("material_desc") or "").lower()
            cust_str = (r.get("customer_name") or "").lower()
            if not any((k in so_str or k in mat_str or k in cust_str) for k in keywords):
                continue
        
        if customer and r.get("customer_name") != customer:
            continue
        if factory and r.get("Factory") != factory:
            continue
        if process_color and r.get("process_color") != process_color:
            continue
            
        filtered_records.append(r)

    # 4. TÍNH TOÁN THẺ DỮ LIỆU (MỚI)
    total_qty = 0
    total_shipped = 0
    total_mapping = 0 # Biến mới tính tổng mapping

    for r in filtered_records:
        r_qty = float(r.get("qty") or 0)
        r_shipped = float(r.get("shipped_qty") or 0)
        r_mapping = float(r.get("mapping_kho") or 0) # Lấy giá trị mapping

        total_qty += r_qty
        total_shipped += r_shipped
        total_mapping += r_mapping

    # Công thức: Chưa sản xuất = Quantity - Shipped - Mapping
    total_unproduced = total_qty - total_shipped - total_mapping

    # Hàm format số liệu: "X / Y Tấn (Z%)"
    def format_metric(value, total):
        if total == 0:
            return "0 Tấn"
        # Đổi đơn vị sang Tấn (chia 1000)
        val_ton = value / 1000
        total_ton = total / 1000
        percent = (value / total) * 100
        return "{:,.0f} / {:,.0f} Tấn ({:.1f}%)".format(val_ton, total_ton, percent)

    # Format riêng cho Tổng đơn hàng (chỉ hiện 1 số)
    total_display = "{:,.0f} Tấn".format(total_qty / 1000)

    summary = {
        "total_qty_display": total_display,
        "shipped_display": format_metric(total_shipped, total_qty),
        "mapping_display": format_metric(total_mapping, total_qty),
        "unproduced_display": format_metric(total_unproduced, total_qty),
        "count": len(filtered_records)
    }

    # 5. Tạo rows hiển thị
    rows = filter_and_group(filtered_records, keyword="", filter_customer="", filter_factory="")
    
    return jsonify({"rows": rows, "summary": summary})
@so_bp.route("/customers_by_process")
@permission_required('view_customer')
def customers_by_process():
    process_color = request.args.get("process_color", "")
    records = get_rows_from_db()
    if process_color:
        records = [r for r in records if r["process_color"] == process_color]
    customers = sorted({r["customer_name"] for r in records})
    return jsonify({"customers": customers})

@so_bp.route("/export_khachhang")
@permission_required('view_customer')
def export_khachhang():
    # 1. Lấy các tham số lọc từ URL
    keyword = request.args.get("keyword", "")
    customer = request.args.get("customer", "")
    process_color = request.args.get("process_color", "")
    factory = request.args.get("factory", "")

    # 2. Lấy toàn bộ dữ liệu gốc từ DB
    records = get_rows_from_db()

    # 3. Áp dụng logic lọc tương tự như hàm tìm kiếm
    filtered_records = []
    keywords = [k.strip().lower() for k in keyword.split(",") if k.strip()]

    for r in records:
        # Lọc theo keyword
        if keywords:
            if not any(
                (
                    k in str(r.get("so_mapping") or "").lower()
                    or k in (r.get("material_desc") or "").lower()
                    or k in (r.get("customer_name") or "").lower()
                )
                for k in keywords
            ):
                continue
        
        # Lọc theo khách hàng
        if customer and r.get("customer_name") != customer:
            continue

        # Lọc theo nhà máy
        if factory and r.get("Factory") != factory:
            continue

        # Lọc theo màu tiến độ
        if process_color and r.get("process_color") != process_color:
            continue

        filtered_records.append(r)

    if not filtered_records:
        return "Không có dữ liệu để xuất.", 404

    # 4. Tạo DataFrame từ dữ liệu đã lọc
    df = pd.DataFrame(filtered_records)

    # 5. Thêm cột 'Tình trạng' được tính toán
    def get_status(row):
        qty_kg = row.get('qty', 0) or 0
        process_value = row.get('process_value', 0) or 0
        if qty_kg <= 100000: # <= 100 Tấn
            if process_value >= 80:
                return "Hoàn thành"
        elif qty_kg > 100000: # > 100 Tấn
            if process_value >= 90:
                return "Hoàn thành"
        return "Chưa hoàn thành"

    df['Tình trạng'] = df.apply(get_status, axis=1)

    # Thêm ký tự '%' vào cột process_value
    df['process_value'] = df['process_value'].apply(lambda x: f"{x}%" if pd.notna(x) else "")
    df['Document Date'] = pd.to_datetime(df['Document Date'], errors='coerce')
    # Định dạng lại cột 'Document Date' thành 'dd/mm/YYYY'
    # .fillna('') để thay thế các giá trị rỗng/lỗi (NaT) bằng chuỗi rỗng
    df['Document Date'] = df['Document Date'].dt.strftime('%d/%m/%Y').fillna('')
    # 5. Chọn, sắp xếp và đổi tên các cột để xuất file
    columns_to_export = {
        'Document Date': 'Document Date',
        'PO.': 'PO.',
        'so_mapping': 'Sales Document',
        'customer_name': 'Customer',
        'Material': 'Material',
        'material': 'Material Description',
        'qty': 'Quantity (KG)',
        'shipped_qty': 'Shipped Quantity (KG)',
        'mapping_kho': 'Đã MAP (KG)',
        'process_value': 'Tỷ lệ',
        'Tình trạng': 'Tình trạng'
    }
    df_export = df[list(columns_to_export.keys())]
    df_export = df_export.rename(columns=columns_to_export)

    # 6. Tạo file Excel trong bộ nhớ
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='TienDoKhachHang')
        # Tự động điều chỉnh độ rộng cột
        for column in df_export:
            column_length = max(df_export[column].astype(str).map(len).max(), len(column))
            col_idx = df_export.columns.get_loc(column)
            writer.sheets['TienDoKhachHang'].set_column(col_idx, col_idx, column_length + 2)

    output.seek(0)

    # 7. Trả file về cho người dùng
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'TienDoKhachHang_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    )