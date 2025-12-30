from flask import Blueprint, render_template, request, jsonify
from collections import OrderedDict
from db import engine
from sqlalchemy import text
from auth.decorator import permission_required

report_bp = Blueprint("report", __name__)

def get_tonkho_from_db():
    sql = text("""
SELECT
    [Order],
    [Material],
    [Material Description],
    [Tồn kho chưa Mapping SO],
    [Tồn kho Mapping SO],
    [Tổng tồn kho],
    [Tổng Loại 1],
    [Tổng Loại 2],
    [Số lượng chờ nhập kho],
    CASE [Plant]
        WHEN 1000 THEN 'HRC1'
        WHEN 1600 THEN 'HRC2'
        ELSE CAST([Plant] AS NVARCHAR(50))
    END AS [Plant],
    [SO Mapping],
    [SL Mapping kho],
    [Shipped Quantity (KG)],
    [Quantity (KG)],
    [Process]
FROM Order_mapping_so
ORDER BY [Order], [SO Mapping];
    """)
    with engine.connect() as conn:
        rows = conn.execute(sql).mappings().all()
    return [dict(r) for r in rows]

def format_number(val, divide_1000=False):
    if val is None:
        return None
    try:
        num = float(val)
        if divide_1000:
            num = num / 1000
        return f"{int(num):,}"
    except:
        return val

def filter_and_group_Order(records, keyword="", filter_nhamay="", filter_process_color=""):
    grouped = OrderedDict()
    for r in records:
        nha_may_val = r.get("Plant", "")
        process_val = float(r.get("Process") or 0)
        process_color = (
            "bg-success" if process_val >= 90 else
            "bg-warning" if process_val >= 75 else
            "bg-danger"
        )

        # --- 🔹 FILTER KEYWORD: tách theo dấu phẩy (,) ---
        keywords = [k.strip() for k in keyword.split(",") if k.strip()]

        if not all(
            (
                k.lower() in str(r.get("SO Mapping", "")).lower() or
                k.lower() in str(r.get("Order", "")).lower() or
                k.lower() in str(r.get("Material Description", "")).lower() or
                k.lower() in str(r.get("Material", "")).lower()
            )
            for k in keywords
        ):
            continue

        # --- 🔹 FILTER NHÀ MÁY ---
        if filter_nhamay and nha_may_val != filter_nhamay:
            continue

        # --- 🔹 FILTER PROCESS COLOR ---
        if filter_process_color and process_color != filter_process_color:
            continue

        order_id = str(int(r.get("Order"))) if r.get("Order") is not None else ""
        grouped.setdefault(order_id, []).append({
            **r,
            "NhaMay": nha_may_val,
            "process_color": process_color,
            "process_value": process_val
        })

    # --- 🔹 GHÉP NHÓM & TÍNH TỔNG ---
    rows_with_flags = []
    for order, mats in grouped.items():
        if not order:
            continue
        
        # --- LOGIC TÍNH TỔNG MỚI ---
        # Khởi tạo các biến tổng
        total_tonkho_chua_map = 0
        total_tonkho_map = 0
        total_tonkho = 0
        total_loai1 = 0
        total_loai2 = 0
        total_cho_nhapkho = 0
        
        # Set để theo dõi các material đã được tính tổng cho order này
        seen_materials = set()

        # Lặp qua các material để tính tổng duy nhất
        for m in mats:
            material_id = m.get("Material")
            # Chỉ cộng nếu material_id tồn tại và chưa được tính trước đó
            if material_id and material_id not in seen_materials:
                total_tonkho_chua_map += float(m.get("Tồn kho chưa Mapping SO") or 0)
                total_tonkho_map += float(m.get("Tồn kho Mapping SO") or 0)
                total_tonkho += float(m.get("Tổng tồn kho") or 0)
                total_loai1 += float(m.get("Tổng Loại 1") or 0)
                total_loai2 += float(m.get("Tổng Loại 2") or 0)
                total_cho_nhapkho += float(m.get("Số lượng chờ nhập kho") or 0)
                seen_materials.add(material_id)
        # --- KẾT THÚC LOGIC TÍNH TỔNG MỚI ---

        order_printed = False
        order_rowspan = len(mats)
        
        for m in mats:
            rows_with_flags.append({
                "order": order if not order_printed else None,
                "nha_may": m["NhaMay"],
                "material": m.get("Material") if not order_printed else None,
                "material_desc": m.get("Material Description"),
                "tonkho_chua_map": format_number(total_tonkho_chua_map, True) if not order_printed else None,
                "tonkho_map": format_number(total_tonkho_map, True) if not order_printed else None,
                "tong_tonkho": format_number(total_tonkho, True) if not order_printed else None,
                "tong_loai1": format_number(total_loai1, True) if not order_printed else None,
                "tong_loai2": format_number(total_loai2, True) if not order_printed else None,
                "cho_nhapkho": format_number(total_cho_nhapkho, True) if not order_printed else None,
                "so_mapping": str(m.get("SO Mapping") or ""),
                "sl_mapping_kho": format_number(m.get("SL Mapping kho"), True),
                "shipped_qty": format_number(m.get("Shipped Quantity (KG)"), True),
                "qty": format_number(m.get("Quantity (KG)"), True),
                "process_value": m["process_value"],
                "process_color": m["process_color"],
                "order_rowspan": order_rowspan
            })
            order_printed = True
            
    return rows_with_flags

# mapping giữa CSS class và tên dễ đọc
PROCESS_NAME_MAP = {
    "bg-success": "Hoàn thành",
    "bg-warning": "Chưa hoàn thành",
    "bg-danger": "Thiếu nhiều"
}
NAME_TO_CLASS = {v: k for k, v in PROCESS_NAME_MAP.items()}

@report_bp.route("/tiendo_order")
@permission_required('view_order')
def order():
    records = get_tonkho_from_db()
    
    # gắn process_color vào mỗi record
    for r in records:
        process_val = float(r.get("Process") or 0)
        r["process_color"] = (
            "bg-success" if process_val >= 90 else
            "bg-warning" if process_val >= 75 else
            "bg-danger"
        )

    nha_may_list = sorted({r.get("Plant") for r in records if r.get("Plant")})
    process_color_list = list(PROCESS_NAME_MAP.values())  # dùng tên dễ đọc

    selected_nhamay = request.args.get("nhamay","")
    selected_process_color_name = request.args.get("process_color","")
    filter_kw = request.args.get("keyword","")

    # convert tên dễ đọc sang CSS class để filter
    filter_process_color_class = NAME_TO_CLASS.get(selected_process_color_name, "")

    filtered_records = filter_and_group_Order(
        records,
        keyword=filter_kw,
        filter_nhamay=selected_nhamay,
        filter_process_color=filter_process_color_class
    )

    return render_template(
        "tiendo_order.html",
        rows=filtered_records,
        nha_may_list=nha_may_list,
        process_color_list=process_color_list,
        selected_nhamay=selected_nhamay,
        selected_process_color=selected_process_color_name,
        filter_kw=filter_kw
    )

@report_bp.route("/tonkho_search")
@permission_required('view_order')
def tonkho_search():
    keyword = request.args.get("keyword", "")
    filter_nhamay = request.args.get("nhamay", "")
    filter_process_color_name = request.args.get("process_color", "")

    filter_process_color_class = NAME_TO_CLASS.get(filter_process_color_name, "")

    records = get_tonkho_from_db()
    rows_with_flags = filter_and_group_Order(
        records,
        keyword=keyword,
        filter_nhamay=filter_nhamay,
        filter_process_color=filter_process_color_class
    )

    # ----------- Sort ----------


    return jsonify({"rows": rows_with_flags})   