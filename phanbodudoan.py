import pandas as pd
import numpy as np
import pandas as pd
from datetime import datetime
from db import engine
from sqlalchemy import text
from sqlalchemy.types import NVARCHAR, Float, BigInteger

# --- KHỞI TẠO DỮ LIỆU MOCK (Thay thế bằng kết nối DB thực tế) ---
def get_rows_from_db1():
    """Lấy toàn bộ dữ liệu từ vw_so_chonhapkho (Nguồn cung chi tiết cuộn)."""
    sql = text("""
    SELECT *
    FROM vw_so_chonhapkho
    ORDER BY Material, NhaMay, NgayDuKien
    """)
    with engine.connect() as conn:
        rows = conn.execute(sql).mappings().all()
    return [dict(r) for r in rows]

def get_rows_from_db2():
    """Lấy toàn bộ dữ liệu từ vw_so_thieu (Nhu cầu thiếu)."""
    sql = text("""
  
WITH All_SO_Material_Pairs AS (
    SELECT DISTINCT
        [SO Mapping],
        [Material Description]
    FROM vw_so_kho_sumary2 -- Bảng SO chứa đầy đủ SO-Material nhất
),

SO_Request_Representative AS (
    SELECT
        [SO Mapping],
        [CW] AS Representative_CW,
        [NHÓM] AS Representative_NHOM,
        ROW_NUMBER() OVER(PARTITION BY [SO Mapping] ORDER BY [Material Description] ASC) as rn
    FROM dbo.so_request
),

so_request_complete AS (
    SELECT
        pairs.[SO Mapping],
        pairs.[Material Description],

        COALESCE(sr_direct.[CW], sr_rep.Representative_CW) AS CW,
        COALESCE(sr_direct.[NHÓM], sr_rep.Representative_NHOM) AS NHOM
    FROM All_SO_Material_Pairs pairs

    LEFT JOIN dbo.so_request sr_direct
        ON pairs.[SO Mapping] = sr_direct.[SO Mapping]
        AND pairs.[Material Description] = sr_direct.[Material Description]

    LEFT JOIN SO_Request_Representative sr_rep
        ON pairs.[SO Mapping] = sr_rep.[SO Mapping]
        AND sr_rep.rn = 1 -- ĐIỀU KIỆN QUAN TRỌNG: Chỉ lấy dòng đại diện số 1
),

data AS (
    SELECT
        lt.[TÀU/PHƯƠNG TIỆN VẬN TẢI] AS tau,
        s.[SO Mapping] AS saleO,
        s.[Material Description] AS material_desc,
        lt.[ETA_Parsed],
        ISNULL(s.[Shipped Quantity (KG)],0) AS shipped_qty,
        ISNULL(s.[Quantity (KG)],0) AS qty,
        ISNULL(s.[SL Mapping kho],0) AS Mapping_kho,
        lt.SheetMonth,
        s.NhaMay AS nhamay
    FROM vw_so_kho_sumary2 s
    JOIN dbo.lichtau lt
        ON s.[SO Mapping] = TRY_CAST(lt.[SỐ LỆNH TÁCH] AS BIGINT)
),
process AS (
    SELECT
        d.*,
        ROUND((Mapping_kho + shipped_qty) * 100.0 / NULLIF(qty,0), 2) AS process_value
    FROM data d
),
sanluongthieu AS (
    SELECT
        p.tau,
        p.saleO,
        p.material_desc,
        p.nhamay,
        p.SheetMonth,
        p.ETA_Parsed,
        (p.qty - p.shipped_qty - p.Mapping_kho) AS SanLuongThieu
    FROM process p
    WHERE (p.qty - p.shipped_qty - p.Mapping_kho) > 0 AND p.process_value < 90  AND p.tau != '' AND (
            (p.qty <= 100000 AND p.process_value < 80) -- Ngưỡng fail cho SO <= 100T
            OR 
            (p.qty > 100000  AND p.process_value < 90) -- Ngưỡng fail cho SO > 100T
        )
),filter_condition AS (
        SELECT 
            CASE 
                WHEN DAY(GETDATE()) > 10 THEN FORMAT(GETDATE(), 'MM.yyyy')
                ELSE FORMAT(DATEADD(MONTH, -1, GETDATE()), 'MM.yyyy') 
            END AS Thang1,
            FORMAT(GETDATE(), 'MM.yyyy') AS ThangHienTai
    )
-- PHẦN SELECT CUỐI CÙNG
SELECT 
    s.saleO AS [SO Mapping],
    s.material_desc AS [Material],
    s.nhamay AS [NhaMay],
    s.SanLuongThieu,
    s.SheetMonth,
    s.tau AS [TÀU/PHƯƠNG TIỆN VẬN TẢI],
    s.ETA_Parsed,
    -- JOIN với CTE "so_request_complete" đã được xử lý hoàn chỉnh
    sr_complete.CW,
    sr_complete.NHOM
FROM sanluongthieu s
CROSS JOIN filter_condition f
JOIN dbo.lichtau lt
        ON TRY_CAST(s.saleO AS BIGINT) = TRY_CAST(lt.[SỐ LỆNH TÁCH] AS BIGINT)
LEFT JOIN so_request_complete sr_complete
    ON s.saleO = sr_complete.[SO Mapping]
    AND s.material_desc = sr_complete.[Material Description]
WHERE
    ISNULL(LTRIM(RTRIM(s.tau)),'') <> '' 
    AND 
    (
        -- Chuyển đổi chuỗi 'MM.yyyy' thành 'yyyyMM' để so sánh số học hoặc chuỗi đúng
        -- Logic: Lấy 4 ký tự cuối (Năm) + 2 ký tự đầu (Tháng)
        
        (DAY(GETDATE()) > 10 
            AND (RIGHT(s.SheetMonth, 4) + LEFT(s.SheetMonth, 2)) >= (RIGHT(f.Thang1, 4) + LEFT(f.Thang1, 2)))
        OR 
        (DAY(GETDATE()) <= 10 
            AND (RIGHT(s.SheetMonth, 4) + LEFT(s.SheetMonth, 2)) >= FORMAT(DATEADD(MONTH, -1, GETDATE()), 'yyyyMM'))
    )
        ORDER BY
    CASE WHEN s.ETA_Parsed IS NULL THEN 1 ELSE 0 END,
    s.ETA_Parsed ASC,
    s.SheetMonth;          
    """)
    with engine.connect() as conn:
        rows = conn.execute(sql).mappings().all()
    return [dict(r) for r in rows]
def ExportDataSAP():
    # --- BƯỚC 1: TẢI DỮ LIỆU ---

    data1 = get_rows_from_db1()
    df_cung = pd.DataFrame(data1)
    data2 = get_rows_from_db2()
    df_thieu = pd.DataFrame(data2)
    df_thieu = df_thieu.drop(["TÀU/PHƯƠNG TIỆN VẬN TẢI", "ETA_Parsed","SheetMonth"], axis=1)
    df_thieu = df_thieu.drop_duplicates()

    # 1. Làm sạch khoảng trắng và đồng nhất chữ hoa/thường để so sánh chính xác
    df_cung['Material'] = df_cung['Material'].str.strip()
    df_cung['Nhóm'] = df_cung['Nhóm'].str.strip().str.upper() # Chuyển nhóm trong kho sang chữ HOA
    df_thieu['Material'] = df_thieu['Material'].str.strip()

    # 2. Tổng hợp chính xác các đơn hàng: Cộng dồn sản lượng thiếu thay vì xóa dòng


    
    # 3. Chuẩn hóa các cột điều kiện và kết quả
    df_cung['TongKhoiLuong'] = pd.to_numeric(df_cung['TongKhoiLuong'], errors='coerce').fillna(0)
    df_cung['NgayDuKien'] = pd.to_datetime(df_cung['NgayDuKien'], errors='coerce')
    df_thieu['SanLuongThieu'] = pd.to_numeric(df_thieu['SanLuongThieu'], errors='coerce').fillna(0)
    
    cw_split = df_thieu['CW'].fillna('').str.split('-', expand=True)
    df_thieu['CW_min'] = pd.to_numeric(cw_split[0], errors='coerce').fillna(0)
    df_thieu['CW_max'] = pd.to_numeric(cw_split[1], errors='coerce').fillna(float('inf'))
    
    # Chuyển yêu cầu nhóm sang chữ HOA để khớp với dữ liệu kho
    df_thieu['NHOM_set'] = df_thieu['NHOM'].astype(str).apply(
        lambda x: {item.strip().upper() for item in x.split(',')} if x and x.lower() not in ['nan', 'none', ''] else set()
    )

    df_ket_qua_hoan_chinh = df_thieu.copy()
    df_ket_qua_hoan_chinh['Lượng_Được_Phân_Bổ'] = 0.0
    df_ket_qua_hoan_chinh['Lượng_Còn_Thiếu'] = 0.0
    df_ket_qua_hoan_chinh['Kết_Quả_Giao_Hàng'] = ''


    # --- BƯỚC 3: SẮP XẾP KHO HÀNG ---
    df_cung['UuTien'] = np.where(df_cung['NgayDuKien'].isnull(), 0, 1)

    supply_dict = {}
    # Sắp xếp theo 2 cấp: Ưu tiên trước, sau đó đến NgayDuKien
    # Việc sắp xếp trước một lần sẽ hiệu quả hơn là sắp xếp trong vòng lặp
    df_cung_sorted = df_cung.sort_values(['UuTien', 'NgayDuKien'])
    
    for (material_key, nhamay_key), group in df_cung_sorted.groupby(['Material', 'NhaMay'], sort=False):
        # group đã được sắp xếp đúng thứ tự, chỉ cần reset_index
        supply_dict[(material_key, nhamay_key)] = group.reset_index(drop=True)
        
    so_uu_tien_list = df_thieu['SO Mapping'].drop_duplicates().tolist()
    
    supply_dict = {}
    for (material_key, nhamay_key), group in df_cung.groupby(['Material', 'NhaMay']):
        supply_dict[(material_key, nhamay_key)] = group.sort_values('NgayDuKien').reset_index(drop=True)

    
    allocation_details = []

    # --- BƯỚC 4: BẮT ĐẦU QUÁ TRÌNH PHÂN BỔ (THEO LOGIC CŨ) ---

    for so_id in so_uu_tien_list:
        df_so_materials = df_thieu[df_thieu['SO Mapping'] == so_id]


        for original_index in df_so_materials.index:
            so_row = df_thieu.loc[original_index]
            material_id = so_row['Material']
            nhamay_id = so_row['NhaMay']
            so_can_thieu = so_row['SanLuongThieu']
            supply_key = (material_id, nhamay_id)

            

            if supply_key not in supply_dict:
                df_ket_qua_hoan_chinh.loc[original_index, 'Lượng_Còn_Thiếu'] = so_can_thieu
                df_ket_qua_hoan_chinh.loc[original_index, 'Kết_Quả_Giao_Hàng'] = "Không có tồn kho và chờ nhập kho"
               
                continue

            current_supply_df = supply_dict[supply_key]
            luong_da_phan_bo = 0.0
            ngay_nhap_kho_max = pd.NaT

            # VỚI MỖI ĐƠN HÀNG, LUÔN QUÉT LẠI DANH SÁCH KHO TỪ ĐẦU
            for supply_idx in range(len(current_supply_df)):
                
                # Nếu đã đủ hàng cho đơn này, dừng quét
                if so_can_thieu <= 0.01:
                    break
                
                roll_data = current_supply_df.loc[supply_idx]
                roll_id = roll_data.get('Roll_ID', roll_data.get('ID Cuộn Bó', 'N/A'))

                # 1. Bỏ qua những cuộn đã hết sạch hàng từ các đơn trước
                if roll_data['TongKhoiLuong'] <= 0.01:
                    continue

                # 2. BÂY GIỜ MỚI BẮT ĐẦU KIỂM TRA ĐIỀU KIỆN
                is_suitable = True
                if so_row['NHOM_set']:
                    if roll_data.get('Nhóm', '') not in so_row['NHOM_set']:
                        is_suitable = False
                
                if is_suitable and (so_row['CW_min'] > 0 or so_row['CW_max'] != float('inf')):
                    khoi_luong_tan = roll_data['TongKhoiLuong'] / 1000.0
                    if not (so_row['CW_min'] <= khoi_luong_tan <= so_row['CW_max']):
                        is_suitable = False
                
                # 3. NẾU PHÙ HỢP, TIẾN HÀNH PHÂN BỔ
                if is_suitable:
                    luong_phan_bo = min(so_can_thieu, roll_data['TongKhoiLuong'])
                    if luong_phan_bo > 0:
                        
                        so_can_thieu -= luong_phan_bo
                        luong_da_phan_bo += luong_phan_bo
                        current_supply_df.loc[supply_idx, 'TongKhoiLuong'] -= luong_phan_bo
                        
                        ngay_nhap_kho_hien_tai = roll_data['NgayDuKien']
                        if pd.isna(ngay_nhap_kho_max) or ngay_nhap_kho_hien_tai > ngay_nhap_kho_max:
                            ngay_nhap_kho_max = ngay_nhap_kho_hien_tai
                        allocation_details.append({'SO_Mapping': so_id, 'Material': material_id, 'NhaMay': nhamay_id, 'Roll_ID': roll_id, 'Ngay_Du_Kien': ngay_nhap_kho_hien_tai, 'Luonng_Phan_Bo': luong_phan_bo})
            
            # Cập nhật kết quả tóm tắt cho đơn hàng này
            df_ket_qua_hoan_chinh.loc[original_index, 'Lượng_Được_Phân_Bổ'] = luong_da_phan_bo
            df_ket_qua_hoan_chinh.loc[original_index, 'Lượng_Còn_Thiếu'] = so_can_thieu
           

            if so_can_thieu <= 0.01:
                if pd.notna(ngay_nhap_kho_max): df_ket_qua_hoan_chinh.loc[original_index, 'Kết_Quả_Giao_Hàng'] = ngay_nhap_kho_max.strftime('%Y-%m-%d')
                else: df_ket_qua_hoan_chinh.loc[original_index, 'Kết_Quả_Giao_Hàng'] = f"Đang có sẵn {luong_da_phan_bo / 1000.0:,.0f}T chưa map"
            else:
                luong_da_phan_bo_tan = luong_da_phan_bo / 1000.0
                khai_bao_thieu_ban_dau_tan = (luong_da_phan_bo + so_can_thieu) / 1000.0
                if luong_da_phan_bo_tan == 0: df_ket_qua_hoan_chinh.loc[original_index, 'Kết_Quả_Giao_Hàng'] = (f"Thiếu toàn bộ ({khai_bao_thieu_ban_dau_tan:,.0f}T)")
                elif pd.notna(ngay_nhap_kho_max): df_ket_qua_hoan_chinh.loc[original_index, 'Kết_Quả_Giao_Hàng'] = (f"{luong_da_phan_bo_tan:,.0f}/{khai_bao_thieu_ban_dau_tan:,.0f}T - {ngay_nhap_kho_max.strftime('%Y-%m-%d')}")
                else: df_ket_qua_hoan_chinh.loc[original_index, 'Kết_Quả_Giao_Hàng'] = (f"Đang có {luong_da_phan_bo_tan:,.0f}/{khai_bao_thieu_ban_dau_tan:,.0f}T chưa map")

 

    # --- BƯỚC 5: LƯU KẾT QUẢ VÀO DATABASE ---
    
    df_final_output = df_ket_qua_hoan_chinh.loc[df_thieu.index].reset_index(drop=True)
    df_upload_summary = df_final_output[['SO Mapping', 'Material', 'NhaMay', 'Lượng_Được_Phân_Bổ', 'Lượng_Còn_Thiếu', 'Kết_Quả_Giao_Hàng']].copy()
    df_allocation_detail = pd.DataFrame(allocation_details)
    
  

    dtype_summary = { 'SO Mapping': BigInteger(), 'Material': NVARCHAR(length=500), 'NhaMay': NVARCHAR(length=50), 'Lượng_Được_Phân_Bổ': Float(), 'Lượng_Còn_Thiếu': Float(), 'Kết_Quả_Giao_Hàng': NVARCHAR(length=255) }
    dtype_detail = { 'SO_Mapping': BigInteger(), 'Material': NVARCHAR(length=50), 'NhaMay': NVARCHAR(length=50), 'Roll_ID': NVARCHAR(length=50), 'Ngay_Du_Kien': NVARCHAR(length=255), 'Luonng_Phan_Bo': Float() }

    try:
        with engine.begin() as connection:
            if not df_upload_summary.empty:
                df_upload_summary.to_sql('tbl_SO_Forecast_Result', connection, if_exists='replace', index=False, dtype=dtype_summary)
          
            if not df_allocation_detail.empty:
                df_allocation_detail.to_sql('tbl_SO_Allocation_Detail', connection, if_exists='replace', index=False, dtype=dtype_detail)
                
        print("✅ Lưu kết quả vào database thành công!")
    except Exception as e:
        print(f"❌ LỖI KHI GHI VÀO DATABASE: {e}")
        
    print("\n--- HOÀN THÀNH TOÀN BỘ QUÁ TRÌNH ---")