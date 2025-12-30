from flask import Blueprint, render_template, request, jsonify
from sqlalchemy import text
from db import engine
from datetime import datetime
from dateutil import parser
from auth.decorator import permission_required
import io
import pandas as pd
from flask import send_file
idcuonbo_bp = Blueprint("idcuonbo_bp", __name__)

def parse_date_str(s):
    """
    Chuyển chuỗi NVARCHAR ngày tháng sang datetime.
    Hỗ trợ cả ngày 1 chữ số và 2 chữ số.
    Trả về None nếu lỗi.
    """
    if not s:
        return None
    try:
        s = str(s).strip()
        dt = parser.parse(s, fuzzy=True)
        return dt
    except Exception:
        return None


def format_date(d):
    """Chuyển datetime sang dd/mm/yyyy"""
    if not d:
        return ""
    return d.strftime("%d/%m/%Y")

def get_sanluong_kho():
    sql = text("""
-- Phần 1: Lấy các cuộn "Chờ nhập kho"
WITH Raw_Data AS (
    /* --- KHỐI 1: Sản lượng (Đã lọc trùng với kho) --- */
    SELECT DISTINCT
    s.[ID Cuộn Bó],
    s.[Material Description],
    s.[Nhóm],
    s.[Vị trí],
    s.[Lô phôi],
    s.[Khối lượng],
    s.[Ngày sản xuất],
    s.Ca,
    s.[Order],
    s.Batch,
    s.[Mác thép],
    '' AS [Customer N],
    CASE WHEN s.[Tp loại 2] = 1 THEN 1 ELSE 0 END AS TpLoai2,
    N'Chờ nhập kho' AS TrangThai,
    CASE 
        WHEN s.[NhaMay] = N'HRC1' THEN DATEADD(DAY, 7, s.[Ngày sản xuất])
        WHEN s.[NhaMay] = N'HRC2' THEN DATEADD(DAY, 8, s.[Ngày sản xuất])
        ELSE DATEADD(DAY, 7, s.[Ngày sản xuất])
    END AS NgayDuKien,
    0 AS [SO Mapping],  -- SỬA LẠI ĐÚNG: Gán giá trị mặc định là 0
    ISNULL(so_detail.[SO_Mapping], 0) AS [SO Mapping dự kiến],
    s.[NhaMay],
    2 AS NguonDuLieu
FROM sanluong s
LEFT JOIN tbl_SO_Allocation_Detail so_detail ON s.[ID Cuộn Bó] = so_detail.[Roll_ID]
WHERE 
    s.[ID Cuộn Bó] IS NOT NULL
    AND s.[ID Cuộn Bó] <> ''
    AND (s.[Đã nhập kho] = 'No' OR s.[Đã nhập kho] IS NULL)
    AND NOT EXISTS (SELECT 1 FROM kho k WHERE k.[ID Cuộn Bó] = s.[ID Cuộn Bó])

UNION ALL

-- Phần 2: Lấy các cuộn đã ở trong kho (phần này vẫn giữ nguyên)
SELECT
    k.[ID Cuộn Bó],
    k.[Material Description],
    k.[Nhóm],
    k.[Vị trí],
    k.[Lô phôi],
    k.[Khối lượng],
    k.[Ngày sản xuất],
    k.Ca,
    k.[Order],
    k.Batch,
    k.[Mác thép],
    K.[Customer N],
    CASE WHEN k.[Tp loại 2] = 'X' THEN 1 ELSE 0 END AS TpLoai2,
    CASE 
        WHEN k.[SO Mapping] = 0 OR k.[SO Mapping] IS NULL THEN N'Nhập kho chưa mapping'
        ELSE N'Nhập kho đã mapping'
    END AS TrangThai,
    CAST(NULL AS DATE) AS NgayDuKien,
    k.[SO Mapping],
    CASE 
        WHEN k.[SO Mapping] = 0 OR k.[SO Mapping] IS NULL THEN ISNULL(so_detail.[SO_Mapping], 0) 
        ELSE 0 
    END AS [SO Mapping dự kiến],
    CASE 
        WHEN k.[Plant] = 1000 THEN N'HRC1'
        WHEN k.[Plant] = 1600 THEN N'HRC2'
        ELSE N'Khác'
    END AS NhaMay,
     1 AS NguonDuLieu
FROM kho k
LEFT JOIN tbl_SO_Allocation_Detail so_detail ON k.[ID Cuộn Bó] = so_detail.[Roll_ID]
),

/* --- BƯỚC LỌC TRÙNG (QUAN TRỌNG NHẤT) --- */
Deduped_Data AS (
    SELECT *,
           ROW_NUMBER() OVER(
               PARTITION BY [ID Cuộn Bó] 
               ORDER BY 
                   NguonDuLieu ASC,       -- QUAN TRỌNG: Số 1 (Kho) sẽ được xếp trước số 2 (Sanluong)
                   [Ngày sản xuất] DESC   -- Nếu cùng nguồn (vd lỗi mapping nhân đôi dòng kho), lấy ngày mới nhất
           ) as rn
    FROM Raw_Data
)
/* --- KẾT QUẢ CUỐI CÙNG --- */
SELECT * FROM Deduped_Data
WHERE rn = 1; -- Chỉ lấy dòng đầu tiên, loại bỏ dòng thứ 2, 3 trở đi

    """)
    with engine.connect() as conn:
        rows = conn.execute(sql).mappings().all()
    return [dict(r) for r in rows]

@idcuonbo_bp.route("/idcuonbo")
@permission_required('view_coil_id')
def idcuonbo():
    rows = get_sanluong_kho()
    trangthai_list = sorted({r["TrangThai"] for r in rows})
    nhom_list = sorted({r["Nhóm"] for r in rows if r["Nhóm"]})
    tp2_list = sorted({r["TpLoai2"] for r in rows if r["TpLoai2"] is not None})
    nha_may_list = sorted({r["NhaMay"] for r in rows if r.get("NhaMay")}) 
    return render_template("idcuonbo.html", trangthai_list=trangthai_list,nhom_list=nhom_list,
        tp2_list=tp2_list,nha_may_list=nha_may_list, )

@idcuonbo_bp.route("/idcuonbo_search")
@permission_required('view_coil_id')
def idcuonbo_search():
    keyword = request.args.get("keyword", "").strip()
    trangthai = request.args.get("trangthai", "")
    sort_col = request.args.get("sort_col", "")
    sort_dir = request.args.get("sort_dir", "asc")
    page = int(request.args.get("page", 1))
    limit = int(request.args.get("limit", 25))

    rows = get_sanluong_kho()
    nhom_list = request.args.getlist("nhom[]")
    tp2 = request.args.get("tp2", "")
    nha_may = request.args.get("nha_may", "") 

    # --------- filter ---------
    filtered = []
    keywords = [k.strip().lower() for k in keyword.split(",") if k.strip()]

    for r in rows:
        id_val = r.get("ID Cuộn Bó")
        so_mapping_val = r.get("SO Mapping")
        so_mapping_du_kien_val = r.get("SO Mapping dự kiến")
        material_val = (r.get("Material Description") or "").lower()
        lo_phoi_val = (r.get("Lô phôi") or "").lower()

        # Kiểm tra keyword (AND logic)
        if keywords:
            match_kw = all(
                (
                    (k.isdigit() and (str(id_val) == k or str(so_mapping_val) == k or str(so_mapping_du_kien_val) == k))
                    or k in material_val
                    or k in lo_phoi_val
                )
                for k in keywords
            )
        else:
            match_kw = True

        match_trangthai = (trangthai == "" or r["TrangThai"] == trangthai)
        match_nhom = (
            not nhom_list
            or (r.get("Nhóm") and r["Nhóm"].strip().lower() in [n.strip().lower() for n in nhom_list])
        )
        match_tp2 = (tp2 == "" or str(r["TpLoai2"]) == str(tp2))
        match_nha_may = (nha_may == "" or r.get("NhaMay") == nha_may)

        if match_kw and match_trangthai and match_nhom and match_tp2 and match_nha_may:
            filtered.append(r)

    rows = filtered

    # --------- parse ngày ---------
    for r in rows:
        r["_Ngày sản xuất"] = parse_date_str(r.get("Ngày sản xuất"))
        r["_NgayDuKien"] = parse_date_str(r.get("NgayDuKien"))

    # --------- sort ---------
    reverse = sort_dir == "desc"
    if sort_col == "Khối lượng":
        rows.sort(key=lambda x: x.get("Khối lượng") or 0, reverse=reverse)
    elif sort_col == "TpLoai2":
        rows.sort(key=lambda x: x.get("TpLoai2") or 0, reverse=reverse)
    elif sort_col == "Ngày sản xuất":
        rows.sort(key=lambda x: x.get("_Ngày sản xuất") or datetime.min, reverse=reverse)
    elif sort_col == "NgayDuKien":
        rows.sort(key=lambda x: x.get("_NgayDuKien") or datetime.min, reverse=reverse)

    # --------- pagination ---------
    total = len(rows)
    start = (page - 1) * limit
    end = start + limit
    page_rows = rows[start:end]

    # --------- format ngày ---------
    for r in page_rows:
        r["Ngày sản xuất"] = r["_Ngày sản xuất"]
        r["NgayDuKien"] = r["_NgayDuKien"]

    return jsonify({"rows": page_rows, "total": total})
@idcuonbo_bp.route("/idcuonbo_export")
@permission_required('view_coil_id')
def idcuonbo_export():
    # 1. Lấy tất cả các tham số lọc từ URL, giống hệt idcuonbo_search
    keyword = request.args.get("keyword", "").strip()
    trangthai = request.args.get("trangthai", "")
    nhom_list = request.args.getlist("nhom[]")
    tp2 = request.args.get("tp2", "")
    nha_may = request.args.get("nha_may", "")
    sort_col = request.args.get("sort_col", "")
    sort_dir = request.args.get("sort_dir", "asc")

    # 2. Lấy toàn bộ dữ liệu
    rows = get_sanluong_kho()

    # 3. Áp dụng logic lọc y hệt trong idcuonbo_search
    filtered = []
    keywords = [k.strip().lower() for k in keyword.split(",") if k.strip()]
    for r in rows:
        id_val = r.get("ID Cuộn Bó")
        so_mapping_val = r.get("SO Mapping")
        so_mapping_du_kien_val = r.get("SO Mapping dự kiến")
        material_val = (r.get("Material Description") or "").lower()
        lo_phoi_val = (r.get("Lô phôi") or "").lower()

        if keywords:
            match_kw = all(
                (
                    (k.isdigit() and (str(id_val) == k or str(so_mapping_val) == k or str(so_mapping_du_kien_val) == k))
                    or k in material_val
                    or k in lo_phoi_val
                )
                for k in keywords
            )
        else:
            match_kw = True

        match_trangthai = (trangthai == "" or r["TrangThai"] == trangthai)
        match_nhom = (
            not nhom_list
            or (r.get("Nhóm") and r["Nhóm"].strip().lower() in [n.strip().lower() for n in nhom_list])
        )
        match_tp2 = (tp2 == "" or str(r["TpLoai2"]) == str(tp2))
        match_nha_may = (nha_may == "" or r.get("NhaMay") == nha_may)

        if match_kw and match_trangthai and match_nhom and match_tp2 and match_nha_may:
            filtered.append(r)
    
    rows = filtered

    # 4. Áp dụng logic sắp xếp (nếu cần)
    for r in rows:
        r["_Ngày sản xuất"] = parse_date_str(r.get("Ngày sản xuất"))
        r["_NgayDuKien"] = parse_date_str(r.get("NgayDuKien"))

    reverse = sort_dir == "desc"
    if sort_col: # Áp dụng sort nếu có
        if sort_col == "Khối lượng":
            rows.sort(key=lambda x: x.get("Khối lượng") or 0, reverse=reverse)
        elif sort_col == "TpLoai2":
            rows.sort(key=lambda x: x.get("TpLoai2") or 0, reverse=reverse)
        elif sort_col == "Ngày sản xuất":
            rows.sort(key=lambda x: x.get("_Ngày sản xuất") or datetime.min, reverse=reverse)
        elif sort_col == "NgayDuKien":
            rows.sort(key=lambda x: x.get("_NgayDuKien") or datetime.min, reverse=reverse)

    # 5. Tạo file Excel bằng Pandas (BỎ QUA PHÂN TRANG)
    if not rows:
        return "Không có dữ liệu để export", 404

    df = pd.DataFrame(rows)
    
    # Format lại ngày tháng để hiển thị đẹp trong Excel
    if '_Ngày sản xuất' in df.columns:
        df['Ngày sản xuất'] = pd.to_datetime(df['_Ngày sản xuất']).dt.strftime('%d/%m/%Y')
    if '_NgayDuKien' in df.columns:
        df['NgayDuKien'] = pd.to_datetime(df['_NgayDuKien']).dt.strftime('%d/%m/%Y')
        
    # Chọn và sắp xếp lại các cột mong muốn
    columns_to_export = [
        'ID Cuộn Bó', 'Material Description', 'Nhóm', 'NhaMay', 'Khối lượng','Vị trí',"Ca","Order","Batch","Mác thép","Customer N",
        'Ngày sản xuất', 'TpLoai2', 'TrangThai', 'SO Mapping', 'NgayDuKien', 'SO Mapping dự kiến'
    ]
    # Lấy các cột tồn tại trong DataFrame để tránh lỗi
    final_columns = [col for col in columns_to_export if col in df.columns]
    df_final = df[final_columns]
    df_final = df_final.copy()
    # Đổi tên cột cho thân thiện hơn
    df_final.rename(columns={
        'TrangThai': 'Trạng thái',
        'NhaMay': 'Nhà máy'
    }, inplace=True)

    # Tạo file trong bộ nhớ
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df_final.to_excel(writer, index=False, sheet_name='IDCuonBo')
    writer.close()
    output.seek(0)

    # 6. Trả về file cho người dùng
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'idcuonbo_{datetime.now().strftime("%Y-%m-%d")}.xlsx'
    )