from flask import Blueprint, render_template, request, jsonify
from sqlalchemy import text
from collections import OrderedDict
from db import engine   # ✅ Sử dụng engine SQL Server đã tạo
import unicodedata
from auth.decorator import permission_required
from storage_utils import load_metadata
tau_bp = Blueprint("tau_bp", __name__)


def normalize_text(s: str) -> str:
    """Chuyển sang chữ thường và loại bỏ dấu Unicode."""
    if not s:
        return ""
    s = s.strip().lower()
    s = unicodedata.normalize('NFKD', s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s


def get_rows_from_db():
    """
    Lấy toàn bộ dữ liệu kết hợp 3 bảng: so, lichtau, ton_kho.
    """
    sql = text("""
WITH data AS (
    SELECT
        lt.[TÀU/PHƯƠNG TIỆN VẬN TẢI] AS tau,
        s.[SO Mapping] AS saleO,
        s.[Material Description] AS material,
        lt.[KHỐI LƯỢNG HÀNG XUẤT LÊN TÀU] AS so_khoi_luong,
        lt.[KHỐI LƯỢNG TỔNG TÀU] AS khoi_luong_tong,
        lt.[ĐẠI LÝ] AS daily,
        lt.[Cảng xếp] AS cangxep,
        lt.[CẢNG ĐẾN] AS cangden,
        ISNULL(s.[Shipped Quantity (KG)],0) AS shipped_qty,
        ISNULL(s.[Quantity (KG)],0) AS qty,
        ISNULL(s.[SL Mapping kho],0) AS Mapping_kho,
        lt.[NGÀY DK DUYỆT SO] AS duyetso,
        lt.[NHỊP] AS nhip,
        lt.SheetMonth,
        s.NhaMay AS nhamay
    FROM vw_so_kho_sumary2 s
    JOIN dbo.lichtau lt
        ON s.[SO Mapping] = TRY_CAST(lt.[SỐ LỆNH TÁCH] AS BIGINT)
),
process AS (
    -- CTE này BÂY GIỜ CHỈ TÍNH process_value
    SELECT
        d.*,
        ROUND(
            CASE
                WHEN d.qty <= 100000 AND (d.Mapping_kho + d.shipped_qty) > 1.25 * d.qty THEN
                    (CASE 
                        WHEN ABS(d.shipped_qty - d.qty) < ABS(d.Mapping_kho - d.qty)
                            THEN (d.shipped_qty * 100.0 / NULLIF(d.qty,0))
                        ELSE (d.Mapping_kho * 100.0 / NULLIF(d.qty,0))
                    END)
                WHEN d.qty > 100000 AND (d.Mapping_kho + d.shipped_qty) > 1.1 * d.qty THEN
                    (CASE 
                        WHEN ABS(d.shipped_qty - d.qty) < ABS(d.Mapping_kho - d.qty)
                            THEN (d.shipped_qty * 100.0 / NULLIF(d.qty,0))
                        ELSE (d.Mapping_kho * 100.0 / NULLIF(d.qty,0))
                    END)
                ELSE ((d.Mapping_kho + d.shipped_qty) * 100.0 / NULLIF(d.qty,0))
            END, 2
        ) AS process_value
    FROM data d
),
process_with_color AS (
    -- *** CTE MỚI ***
    -- Gán màu cho từng dòng dựa trên logic 80/90
    SELECT
        p.*,
        CASE
            -- Logic MỚI cho process_color theo yêu cầu
            WHEN (p.qty <= 100000 AND p.process_value >= 80) THEN 'bg-success'
            WHEN (p.qty >  100000 AND p.process_value >= 90) THEN 'bg-success'
            
            ELSE 'bg-warning'
        END AS process_color
    FROM process p
),
tau_status AS (
    SELECT
        tau,
        SheetMonth,
        CASE
            -- Logic tô màu tàu (vẫn như cũ, đã đúng)
            WHEN MIN(CASE 
                    WHEN d.qty <= 100000 AND d.process_value < 80 THEN 0 -- Material <= 100T, ngưỡng là 80%
                    WHEN d.qty > 100000 AND d.process_value < 90 THEN 0  -- Material > 100T, ngưỡng là 90%
                    ELSE 1 -- Đạt ngưỡng
                END) = 0
            THEN 'bg-warning'
        ELSE 'bg-success'
        END AS tau_status_color
    FROM process_with_color d -- <-- Cập nhật: Lấy từ CTE mới
    GROUP BY tau, SheetMonth
)
SELECT DISTINCT
    p.tau,
    p.SheetMonth,
    p.saleO,
    p.material,
    p.so_khoi_luong,
    p.khoi_luong_tong,
    p.daily,
    p.cangxep,
    p.cangden,
    p.shipped_qty,
    p.qty,
    p.Mapping_kho,
    p.duyetso,
    p.nhip,
    p.nhamay,
    p.process_value,
    p.process_color, -- <-- Màu này đã được tính theo logic mới
    ts.tau_status_color,
    fr.Kết_Quả_Giao_Hàng AS Dự_Đoán_Hoàn_Thành
FROM process_with_color p -- <-- Cập nhật: Lấy từ CTE mới
LEFT JOIN tau_status ts 
    ON p.tau = ts.tau 
    AND p.SheetMonth = ts.SheetMonth 
LEFT JOIN dbo.tbl_SO_Forecast_Result fr 
    ON p.saleO = fr.[SO Mapping] 
    AND p.material = fr.Material 
    AND p.nhamay = fr.NhaMay 
    """)

    with engine.connect() as conn:
        rows = conn.execute(sql).mappings().all()

    return [dict(r) for r in rows]


def filter_and_group(records, keyword="", filter_tau="", filter_cangxep="",
                     filter_cangden="", filter_daily="", filter_tau_color="", filter_sheetmonth="", filter_nhamay="", filter_process_color=""):
    keywords = [k.strip().lower() for k in keyword.split(",") if k.strip()]
    grouped = OrderedDict()

    for r in records:
        if filter_sheetmonth and r.get("SheetMonth") != filter_sheetmonth:
            continue
        if filter_nhamay and r.get("nhamay", "") != filter_nhamay:
            continue

        tau_name = (r.get("tau") or "").strip()
        if tau_name == "":
            continue
        so_name = str(r.get("saleO") or "").strip()

        if keywords:
            if not all(
                (
                    k in (r.get("material") or "").lower()
                    or k in so_name.lower()
                    or k in tau_name.lower()
                )
                for k in keywords
            ):
                continue

        tau_name_norm = normalize_text(tau_name)
        filter_tau_norm = normalize_text(filter_tau)
        if filter_tau and tau_name_norm != filter_tau_norm:
            continue
        if filter_cangxep and r.get("cangxep", "") != filter_cangxep:
            continue
        if filter_cangden and r.get("cangden", "") != filter_cangden:
            continue
        if filter_daily and r.get("daily", "") != filter_daily:
            continue
        if filter_tau_color and r.get("process_color", "") != filter_tau_color:
            continue
        if filter_process_color:
            if filter_process_color == "bg-success":
                if r.get("process_color") != "bg-success":
                    continue
            elif filter_process_color == "not_success":
                if r.get("process_color") == "bg-success":
                    continue

        grouped.setdefault(tau_name, OrderedDict()).setdefault(so_name, []).append(r)

    rows_with_flags = []
    for tau, so_map in grouped.items():
        tau_printed = False
        for so, mats in so_map.items():
            so_len = len(mats)
            so_printed = False

            so_khoi_luong = int(mats[0].get("so_khoi_luong") or 0)

            sum_shipped = int(sum(m.get("Mapping_kho") or 0 for m in mats) / 1000)
            sum_abc = int(sum(m.get("shipped_qty") or 0 for m in mats) / 1000)
            sum_qty = int(sum(m.get("qty") or 0 for m in mats) / 1000)

            # Tổng hiện tại
            tong = sum_shipped + sum_abc

            # Xác định ngưỡng vượt
            nguong = 1.2 if sum_qty < 100 else 1.1  # đơn vị 100 nghìn kg → 100 tấn

            # Nếu tổng vượt quá nhiều so với định lượng thực tế
            if tong > sum_qty * nguong:
                # So sánh cái nào gần sum_qty hơn thì chọn cái đó
                diff_shipped = abs(sum_shipped - sum_qty)
                diff_abc = abs(sum_abc - sum_qty)
                if diff_shipped <= diff_abc:
                    gan_dung = sum_shipped
                else:
                    gan_dung = sum_abc
                dinh_luong_text = f"{gan_dung}/{sum_qty}"
            else:
                # Bình thường thì cộng lại
                dinh_luong_text = f"{tong}/{sum_qty}"
            for mat in mats:
                rows_with_flags.append({
                    "tau": tau,
                    "daily": mat.get("daily") if not tau_printed else None,
                    "cangxep": mat.get("cangxep") if not tau_printed else None,
                    "cangden": mat.get("cangden") if not tau_printed else None,
                    "tau_rowspan": sum(len(v) for v in so_map.values()) if not tau_printed else None,
                    "print_tau": not tau_printed,
                    "so": so,
                    "nhamay": mat.get("nhamay") if not so_printed else None,
                    "so_khoi_luong": "{:,}".format(so_khoi_luong),
                    "duyetso": mat.get("duyetso") if not so_printed else None,
                    "dinh_luong": dinh_luong_text if not so_printed else None,
                    "so_rowspan": so_len if not so_printed else None,
                    "print_so": not so_printed,
                    "material": mat.get("material"),
                    "mapping_kho": "{:,}".format(int((mat.get("Mapping_kho") or 0)/1000)),
                    "shipped_qty": "{:,}".format(int((mat.get("shipped_qty") or 0) / 1000)),
                    "qty": "{:,}".format(int((mat.get("qty") or 0) / 1000)),
                    "process_value": mat.get("process_value", 0),
                    "process_color": mat.get("process_color", ""),
                    "Dự_Đoán_Hoàn_Thành": mat.get("Dự_Đoán_Hoàn_Thành") or "",
                    "tau_status_color": mats[0].get("tau_status_color", "")
                })
                tau_printed = True
                so_printed = True

    return rows_with_flags


@tau_bp.route("/lichtau")
@permission_required('view_ship_schedule')
def lichtau():
    records = get_rows_from_db()
    sheetmonth_list = sorted({r["SheetMonth"] for r in records if r.get("SheetMonth")})
    selected_month = request.args.get("sheetmonth", "")
    
    if not selected_month and sheetmonth_list:
        selected_month = sheetmonth_list[-1]
    filtered_records = [r for r in records if r.get("SheetMonth") == selected_month]
    rows_with_flags = filter_and_group(filtered_records, filter_sheetmonth=selected_month)

    tau_list = sorted({r["tau"] for r in records if r.get("tau")})
    cangxep_list = sorted({r["cangxep"] for r in records if r.get("cangxep")})
    cangden_list = sorted({r["cangden"] for r in records if r.get("cangden")})
    daily_list = sorted({r["daily"] for r in records if r.get("daily")})
    tau_color_list = sorted({r["process_color"] for r in records if r.get("process_color")})
    nhamay_list = sorted({r["nhamay"] for r in records if r.get("nhamay")})
    process_color_list = sorted({r["process_color"] for r in records if r.get("process_color")})
    total_tau = len({r["tau"] for r in records if r.get("tau")})
    total_khoi_luong = sum(int(r.get("so_khoi_luong") or 0) for r in records)

    return render_template(
        "lichtau.html",
        rows=rows_with_flags,
        sheetmonth_list=sheetmonth_list,
        selected_month=selected_month,
        tau_list=tau_list,
        cangxep_list=cangxep_list,
        cangden_list=cangden_list,
        daily_list=daily_list,
        tau_color_list=tau_color_list,
        nhamay_list=nhamay_list,
        filter_tau="",
        filter_tau_color="",
        filter_cangxep="",
        filter_cangden="",
        filter_daily="",
        filter_nhamay="",
        total_tau=total_tau,
        process_color_list=process_color_list,
        total_khoi_luong=total_khoi_luong
    )


@tau_bp.route("/lichtau_search")
@permission_required('view_ship_schedule')
def lichtau_search():
    keyword = request.args.get("keyword", "")
    filter_tau = request.args.get("tau", "")
    filter_cangxep = request.args.get("cangxep", "")
    filter_cangden = request.args.get("cangden", "")
    filter_daily = request.args.get("daily", "")
    filter_tau_color = request.args.get("tau_color", "")
    filter_sheetmonth = request.args.get("sheetmonth", "")
    filter_nhamay = request.args.get("nhamay", "")
    filter_process_color = request.args.get("process_color", "")
    records = get_rows_from_db()
    rows_with_flags = filter_and_group(
        records,
        keyword,
        filter_tau,
        filter_cangxep,
        filter_cangden,
        filter_daily,
        filter_tau_color,
        filter_sheetmonth,
        filter_nhamay,
        filter_process_color
    )

    tau_set = set(r['tau'] for r in rows_with_flags if r.get('tau'))
    total_tau = len(tau_set)
    total_khoi = sum(int(r.get('so_khoi_luong', '0').replace(',', '')) for r in rows_with_flags if r.get('so_khoi_luong'))

    return jsonify({
        "rows": rows_with_flags,
        "total_tau": total_tau,
        "total_khoi": total_khoi
    })
