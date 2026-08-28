from flask import Blueprint, render_template, request, jsonify
from sqlalchemy import text
from collections import OrderedDict
from db import engine
from storage_utils import load_metadata
from datetime import datetime, timedelta
from auth.decorator import permission_required
reportlsx_bp = Blueprint("reportlsx", __name__)

def get_rows_from_db(lsx_id: str | None = None, only_with_material: bool = True):
    """
    - only_with_material=True: giữ điều kiện (material IS NULL OR material <> '') 
      để **không vô tình loại các hàng lsx không có bản ghi trong ton_kho**.
    - only_with_material=False: bỏ hoàn toàn điều kiện material (dùng khi bạn muốn xem mọi hàng).
    """
    params: dict = {}
    conditions: list[str] = []

    if only_with_material:
        # giữ LEFT JOIN behavior: nếu t là NULL (không có ton_kho) thì vẫn bao gồm
        conditions.append("(t.[Material Description] IS NULL OR LTRIM(RTRIM(t.[Material Description])) <> '')")

    if lsx_id:
        conditions.append("l.lsx_id = :lsx_id")
        params["lsx_id"] = lsx_id

    where_clause = (" WHERE " + " AND ".join(conditions)) if conditions else ""

    sql = f"""
        SELECT DISTINCT
            l.[Order],
            s.[Customer]          AS customer_name, -- Lấy tên khách hàng từ bảng 'so'
            l.[Ngày bắt đầu block]   AS start_date,
            l.[Ngày kết thúc block]  AS end_date,
            l.[Sản lượng 1A]         AS prod_1a,
            l.[Sản lượng 1B]         AS prod_1b,
            l.[Mac thep]            AS macthep,
            l.[Yêu cầu đặc biệt]     AS yeucau,
            l.[KL_Cuộn_(Tấn)]       AS klcuon,
            l.[Mục đích sử dụng]    AS mucdich,
            l.[Khối lượng cuộn trung bình] AS klcuontb,
            l.[Tổng yêu cầu]        AS total_req,
            t.[SO Mapping]          AS so_mapping,
            t.[SL Mapping kho]      AS mapping_kho,
            t.[Quantity (KG)]       AS qty_kg,
            t.Process               AS process_value,
            t.[Material Description] AS material,
            l.lsx_id
        FROM lsx l
        LEFT JOIN Order_mapping_so t ON l.[Order] = t.[Order]
        LEFT JOIN so s ON t.[SO Mapping] = s.[Sales Document] 
            {where_clause}
        ORDER BY l.[Order], t.[SO Mapping]

    """

    # Debug: in case bạn muốn kiểm tra SQL và params trên server
    # print("DEBUG SQL:", sql)
    # print("DEBUG PARAMS:", params)

    with engine.connect() as conn:
        result = conn.execute(text(sql), params)
        rows = result.mappings().all()

    return [dict(r) for r in rows]


def filter_and_group(records, keyword: str = "", filter_customer: str = ""):
    # giữ nguyên hàm của bạn (bỏ lại như cũ)
    keyword = keyword.lower().strip()
    grouped = OrderedDict()

    for r in records:
        order_str = str(r.get("Order") or "").strip()
        mat_str   = str(r.get("material") or "").strip()
        cust_str  = str(r.get("customer_name") or "").strip()
        so_str    = str(r.get("so_mapping") or "").strip()

        if keyword and not any(keyword in x.lower() for x in (order_str, mat_str, cust_str, so_str)):
            continue

        start_date_str = r["start_date"].strftime("%d/%m/%Y") if r.get("start_date") else ""
        end_date_str   = r["end_date"].strftime("%d/%m/%Y")   if r.get("end_date")   else ""
        date_range_str = f"Từ {start_date_str} đến {end_date_str}" if start_date_str or end_date_str else ""

        grouped.setdefault(date_range_str, OrderedDict()).setdefault(order_str, []).append(r)

    rows_with_flags = []
    for date_range, order_map in grouped.items():
        date_printed = False
        for order, mats in order_map.items():
            order_printed = False
            for m in mats:
                rows_with_flags.append({
                    "date_range": date_range if not date_printed else None,
                    "date_rowspan": sum(len(v) for v in order_map.values()) if not date_printed else None,
                    "order": order if not order_printed else None,
                    "order_rowspan": len(mats) if not order_printed else None, # order_rowspan vẫn giữ nguyên cho các cột của Order
                    "prod_1a": m["prod_1a"],
                    "prod_1b": m["prod_1b"],
                    "macthep": m["macthep"],
                    "yeucau": m["yeucau"],
                    "klcuon": m["klcuon"],
                    "mucdich": m["mucdich"],
                    "klcuontb": m["klcuontb"],
                    "total_req": m["total_req"],
                    "customer_name": m["customer_name"], # customer_name giờ sẽ xuất hiện ở mỗi dòng material
                    "material": m["material"],
                    "so_mapping": m["so_mapping"],
                    "mapping_kho": "{:,}".format(int((m["mapping_kho"] or 0) // 1000)),
                    "qty_kg": "{:,}".format(int((m["qty_kg"] or 0) // 1000)),
                    "process_value": m["process_value"] or 0,
                    "process_color": (
                        "bg-success" if (m["process_value"] or 0) >= 95 else
                        "bg-warning" if (m["process_value"] or 0) >= 75 else
                        "bg-danger"
                    ),
                })
                date_printed = True
                order_printed = True
    return rows_with_flags


# Routes
@reportlsx_bp.route("/lsx")
@permission_required('view_lsx_report')
def lsx_all():
    metadata = load_metadata()
    if not metadata:
        return render_template("xem_theo_ngay.html")

    # 🟢 Chỉ lấy các item có type = "lsx"
    lsx_items = [item for item in metadata if item.get("type") == "lsx"]

    # Nếu không có bản ghi LSX nào
    if not lsx_items:
        return render_template(
            "lsx.html",
            rows=[],
            customer_list=[],
            lsx_list=[],
        )

    # Lấy dữ liệu chính từ DB
    records = get_rows_from_db(only_with_material=True)
    rows = filter_and_group(records)
    customer_list = sorted({r["customer_name"] for r in records if r["customer_name"]})

    # 🟢 Tạo danh sách LSX hiển thị
    lsx_list = []
    for item in lsx_items:
        base_name = item.get("name") or item.get("id", "Không có ID")
        uploaded_at = item.get("uploaded_at")

        if uploaded_at:
            try:
                # Nếu uploaded_at là chuỗi ISO: "2025-09-29T10:15:30"
                dt = datetime.fromisoformat(str(uploaded_at))
                uploaded_display = dt.strftime("%d/%m/%Y %H:%M:%S")
            except Exception:
                uploaded_display = str(uploaded_at)
            display_name = f"{base_name} - {uploaded_display}"
        else:
            display_name = base_name

        lsx_list.append({
            "id": item.get("id", ""),  # tránh KeyError
            "name": display_name
        })

    return render_template(
        "lsx.html",
        rows=rows,
        customer_list=customer_list,
        lsx_list=lsx_list
    )

@reportlsx_bp.route("/lsx_search")
@permission_required('view_lsx_report')
def lsx_search():
    keyword = request.args.get("keyword", "").strip()
    lsx_id  = request.args.get("lsx_id", "").strip() or None

    # thêm param tùy chọn để debug: only_with_material=0 => bỏ filter material
    only_with_material = request.args.get("only_with_material", "1") == "1"

    records = get_rows_from_db(lsx_id=lsx_id, only_with_material=only_with_material)
    rows = filter_and_group(records, keyword)
    return jsonify({"rows": rows})