from flask import Blueprint, render_template, request, jsonify, flash, redirect, url_for
import pandas as pd
from sqlalchemy import text, types
from db import engine
from datetime import datetime, timedelta
from auth.decorator import permission_required
import unicodedata
import numpy as np
import re

# Khởi tạo Blueprint riêng cho CTD, url_prefix là /ctd để không đụng hàng
ctd_bp = Blueprint('ctd_bp', __name__, url_prefix='/ctd')

# Route truy cập trang chính: /ctd/dashboard_ctd
@ctd_bp.route('/dashboard_ctd')
@permission_required('ctd_view')
def dashboard_ctd():
    return render_template('ctd_dashboard.html')

# ==============================================================================
# PHẦN 1: TỒN KHO & XUẤT HÀNG (GIỮ NGUYÊN CODE CŨ)
# ==============================================================================
@ctd_bp.route('/upload_mapping_ctd', methods=['POST'])
@permission_required('upload_ctd')
def upload_mapping_ctd():
    if 'file' not in request.files:
        flash('Không tìm thấy file', 'danger')
        return redirect(url_for('ctd_bp.dashboard_ctd'))
    
    file = request.files['file']
    if file.filename == '':
        flash('Chưa chọn file', 'danger')
        return redirect(url_for('ctd_bp.dashboard_ctd'))

    try:
        df = pd.read_excel(file)
        df.columns = [c.strip().lower() for c in df.columns] 
        
        required_columns = ['ma_vat_tu', 'duong_kinh', 'mac_thep']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            flash(f"Lỗi: File Excel thiếu các cột bắt buộc: {', '.join(missing_columns)}", 'danger')
            return redirect(url_for('ctd_bp.dashboard_ctd'))

        df['ma_vat_tu'] = df['ma_vat_tu'].astype(str).str.strip()
        df['ma_vat_tu'] = df['ma_vat_tu'].str.replace(r'\.0$', '', regex=True)
        df = df[['ma_vat_tu', 'duong_kinh', 'mac_thep']]
        
        with engine.begin() as conn:
            dtype_mapping = {
                'ma_vat_tu': types.NVARCHAR(length=50),
                'duong_kinh': types.NVARCHAR(length=50),
                'mac_thep': types.NVARCHAR(length=100)
            }
            df.to_sql('mapping_vat_tu_ctd', conn, if_exists='replace', index=False, dtype=dtype_mapping)
            conn.execute(text("CREATE CLUSTERED INDEX IX_mapping_ctd_mvt ON mapping_vat_tu_ctd(ma_vat_tu)"))
        
        flash('Upload Mapping và Đồng bộ Vật tư CTD thành công!', 'success')
    except Exception as e:
        flash(f'Lỗi khi xử lý file: {str(e)}', 'danger')
        
    return redirect(url_for('ctd_bp.dashboard_ctd'))

@ctd_bp.route('/api/data_ctd')
def get_data_ctd():
    try:
        with engine.connect() as conn:
            check_table = conn.execute(text("SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'mapping_vat_tu_ctd'")).scalar()
            if not check_table:
                return jsonify({
                    "status": "error", 
                    "message": "Bảng Mapping chưa tồn tại! Vui lòng chọn file Excel và bấm Upload trước."
                })

            query = """
                SELECT 
                    k.ma_vat_tu, 
                    k.ten_vat_tu, 
                    (k.sl_cuoi_ky / 1000.0) as khoi_luong,
                    ISNULL(k.mrp_group_desc, N'Khác') as mrp_group_desc,
                    DATEDIFF(day, k.ngay_nhap_cuoi, GETDATE()) as tuoi_ton_kho,
                    k.snapshot_ts,
                    CASE 
                        WHEN LEFT(RTRIM(k.so_lo), 3) = 'EXP' THEN N'Xuất khẩu'
                        WHEN LEFT(RTRIM(k.so_lo), 3) = 'DOS' THEN N'Thép CLC'
                        ELSE N'Nội địa'
                    END AS thi_truong,
                    CASE 
                        WHEN RIGHT(RTRIM(k.ten_vat_tu), 2) = 'II' THEN N'Loại 2'
                        ELSE N'Loại 1'
                    END AS loai_hang,
                    ISNULL(m.duong_kinh, N'N/A') as duong_kinh,
                    ISNULL(m.mac_thep, N'N/A') as mac_thep
                FROM kho_ctd_pv k with (nolock)
                LEFT JOIN mapping_vat_tu_ctd m with (nolock) ON k.ma_vat_tu = m.ma_vat_tu
                WHERE k.nha_may = 'CTD' AND k.sl_cuoi_ky > 0
            """
            df = pd.read_sql(query, conn)
            
            query_cnk = """
                SELECT 
                    CASE 
                        WHEN s.[Material Description] LIKE N'Thép thanh%' THEN N'Thép thanh'
                        WHEN s.[Material Description] LIKE N'Thép cuộn%' THEN N'Thép cuộn'
                        ELSE N'Khác'
                    END AS nhom_vat_tu,
                    CASE 
                        WHEN s.[Tp loại 2] = 0 THEN N'Loại 1'
                        WHEN s.[Tp loại 2] = 1 THEN N'Loại 2'
                        ELSE N'Không rõ'
                    END AS loai_hang,
                    ISNULL(m.duong_kinh, N'N/A') AS duong_kinh,
                    s.[Mác thép] AS mac_thep,
                    (s.[Khối lượng] / 1000.0) AS khoi_luong_tan
                FROM sanluong s WITH (NOLOCK)
                LEFT JOIN mapping_vat_tu_ctd m WITH (NOLOCK) ON s.[Material] = m.ma_vat_tu
                WHERE s.NhaMay = 'CTD' AND s.[Khối lượng] > 0 AND s.[Đã nhập kho] = N'No'
            """
            df_cnk = pd.read_sql(query_cnk, conn)
            tong_cnk = 0
            pie_chart_cnk = {"labels": [], "values": []}
            if not df_cnk.empty:
                tong_cnk = float(df_cnk['khoi_luong_tan'].sum())
                df_cnk_grp = df_cnk.groupby('nhom_vat_tu')['khoi_luong_tan'].sum().reset_index()
                pie_chart_cnk = {
                    "labels": df_cnk_grp['nhom_vat_tu'].tolist(),
                    "values": df_cnk_grp['khoi_luong_tan'].tolist()
                }

            query_so = """
                SELECT 
                    CASE 
                        WHEN so.[Item Description] LIKE N'Thép thanh%' THEN N'Thép thanh'
                        WHEN so.[Item Description] LIKE N'Thép cuộn%' THEN N'Thép cuộn'
                        ELSE N'Khác'
                    END AS nhom_vat_tu,
                    ISNULL(m.duong_kinh, N'N/A') AS duong_kinh,
                    ISNULL(m.mac_thep, N'N/A') AS mac_thep,
                    (so.[Quantity (KG)] / 1000.0) AS kl_can,
                    (so.[Shipped Quantity (KG)] / 1000.0) AS kl_da_ship,
                    (so.[SL cần xuất] / 1000.0) AS kl_can_xuat
                FROM so WITH (NOLOCK)
                LEFT JOIN mapping_vat_tu_ctd m WITH (NOLOCK) ON so.[Material] = m.ma_vat_tu
                WHERE so.NhaMay = 'CTD' 
                  AND CAST(so.[Reqd Deliv Date] AS DATE) = CAST(GETDATE() AS DATE)
                  AND so.[Sales Document] LIKE '14%'
            """
            df_so = pd.read_sql(query_so, conn)

            query_sx_ctd = """
                SELECT 
                    [Ngày sản xuất] AS ngay_sx,
                    [Ca] AS ca_sx,
                    [Mác thép] AS mac_thep,
                    CASE 
                        WHEN PhanXuong = 1.0 THEN N'Cán 1'
                        WHEN PhanXuong = 2.0 THEN N'Cán 2'
                        WHEN PhanXuong = 3.0 THEN N'Cán 3'
                        ELSE N'Khác'
                    END AS may_can,
                    ([Khối lượng] / 1000.0) AS khoi_luong_tan
                FROM sanluong WITH (NOLOCK)
                WHERE NhaMay = 'CTD'
                AND MONTH([Ngày sản xuất]) = MONTH(GETDATE())
                AND YEAR([Ngày sản xuất]) = YEAR(GETDATE())
            """
            df_sx = pd.read_sql(query_sx_ctd, conn)            

        if df.empty:
            return jsonify({"status": "empty"})

        last_update = df['snapshot_ts'].max().strftime("%d/%m/%Y %H:%M:%S")
        tong_ton = float(df['khoi_luong'].sum())

        df_group = df.groupby('mrp_group_desc')['khoi_luong'].sum().reset_index()
        pie_chart = {"labels": df_group['mrp_group_desc'].tolist(), "values": df_group['khoi_luong'].tolist()}

        bins = [-1, 90, 180, 365, float('inf')]
        age_labels = ['Dưới 3 tháng', '3-6 tháng', '6-12 tháng', 'Trên 1 năm']
        df['nhom_tuoi'] = pd.cut(df['tuoi_ton_kho'], bins=bins, labels=age_labels)
        df_age = df.groupby(['nhom_tuoi', 'mrp_group_desc'])['khoi_luong'].sum().unstack(fill_value=0).reindex(age_labels).reset_index()
        
        age_chart = {
            "labels": age_labels,
            "datasets": [{"label": col, "data": df_age[col].tolist()} for col in df_age.columns if col != 'nhom_tuoi']
        }

        df_market = df.groupby(['mrp_group_desc', 'thi_truong'])['khoi_luong'].sum().unstack(fill_value=0).reset_index()
        market_labels = df_market['mrp_group_desc'].tolist()
        for col in ['Nội địa', 'Xuất khẩu', 'Thép CLC']:
            if col not in df_market.columns:
                df_market[col] = 0.0
                
        market_chart = {
            "labels": market_labels,
            "noidia": df_market['Nội địa'].tolist(),
            "xuatkhau": df_market['Xuất khẩu'].tolist(),
            "clc": df_market['Thép CLC'].tolist()
        }
        
        df_l1_6m = df[(df['loai_hang'] == 'Loại 1') & (df['tuoi_ton_kho'] <= 180)].copy()
        df_l1_over6m = df[(df['loai_hang'] == 'Loại 1') & (df['tuoi_ton_kho'] > 180)].copy()
        df_l2 = df[df['loai_hang'] == 'Loại 2'].copy()

        def extract_num(val):
            nums = re.findall(r'\d+\.?\d*', str(val))
            return float(nums[0]) if nums else 0.0

        def get_horiz_chart(filtered_df):
            if filtered_df.empty:
                return {"labels": [], "values": [], "groups": []}
            
            grp = filtered_df.groupby(['duong_kinh', 'mac_thep'])['khoi_luong'].sum().reset_index()
            grp['dk_num'] = grp['duong_kinh'].apply(extract_num)
            grp = grp.sort_values(['dk_num', 'mac_thep'], ascending=[True, True]).reset_index(drop=True)
            
            labels = grp['mac_thep'].tolist()
            values = grp['khoi_luong'].tolist()
            
            groups = []
            current_dk = None
            start_idx = 0
            for i, row in grp.iterrows():
                dk = row['duong_kinh']
                if dk != current_dk:
                    if current_dk is not None:
                        groups.append({"label": current_dk, "start": start_idx, "end": i - 1})
                    current_dk = dk
                    start_idx = i
            if current_dk is not None:
                groups.append({"label": current_dk, "start": start_idx, "end": len(grp) - 1})
                
            return {"labels": labels, "values": values, "groups": groups}

        chart_thanh_nd = get_horiz_chart(df_l1_6m[(df_l1_6m['mrp_group_desc'] == 'Thép thanh') & (df_l1_6m['thi_truong'] == 'Nội địa')])
        chart_cuon_nd = get_horiz_chart(df_l1_6m[(df_l1_6m['mrp_group_desc'] == 'Thép cuộn') & (df_l1_6m['thi_truong'] == 'Nội địa')])
        chart_clc = get_horiz_chart(df_l1_6m[df_l1_6m['thi_truong'] == 'Thép CLC'])

        chart_thanh_nd_l1_over6m = get_horiz_chart(df_l1_over6m[(df_l1_over6m['mrp_group_desc'] == 'Thép thanh') & (df_l1_over6m['thi_truong'] == 'Nội địa')])
        chart_cuon_nd_l1_over6m = get_horiz_chart(df_l1_over6m[(df_l1_over6m['mrp_group_desc'] == 'Thép cuộn') & (df_l1_over6m['thi_truong'] == 'Nội địa')])
        chart_clc_l1_over6m = get_horiz_chart(df_l1_over6m[df_l1_over6m['thi_truong'] == 'Thép CLC'])

        chart_thanh_nd_l2 = get_horiz_chart(df_l2[(df_l2['mrp_group_desc'] == 'Thép thanh') & (df_l2['thi_truong'] == 'Nội địa')])
        chart_cuon_nd_l2 = get_horiz_chart(df_l2[(df_l2['mrp_group_desc'] == 'Thép cuộn') & (df_l2['thi_truong'] == 'Nội địa')])
        chart_clc_l2 = get_horiz_chart(df_l2[df_l2['thi_truong'] == 'Thép CLC'])

        df_ton_l1 = df[df['loai_hang'] == 'Loại 1'].copy()
        df_ton_grp = df_ton_l1.groupby(['mrp_group_desc', 'duong_kinh', 'mac_thep'])['khoi_luong'].sum().reset_index()
        df_ton_grp.rename(columns={'mrp_group_desc': 'nhom_vat_tu', 'khoi_luong': 'ton_l1'}, inplace=True)
        
        df_cnk_l1 = df_cnk[df_cnk['loai_hang'] == 'Loại 1'].copy()
        df_cnk_grp = df_cnk_l1.groupby(['nhom_vat_tu', 'duong_kinh', 'mac_thep'])['khoi_luong_tan'].sum().reset_index()
        df_cnk_grp.rename(columns={'khoi_luong_tan': 'cnk_l1'}, inplace=True)
        
        if not df_so.empty:
            df_so_grp = df_so.groupby(['nhom_vat_tu', 'duong_kinh', 'mac_thep'])[['kl_can', 'kl_da_ship', 'kl_can_xuat']].sum().reset_index()
        else:
            df_so_grp = pd.DataFrame(columns=['nhom_vat_tu', 'duong_kinh', 'mac_thep', 'kl_can', 'kl_da_ship', 'kl_can_xuat'])
        
        df_atp = pd.merge(df_so_grp, df_ton_grp, on=['nhom_vat_tu', 'duong_kinh', 'mac_thep'], how='outer')
        df_atp = pd.merge(df_atp, df_cnk_grp, on=['nhom_vat_tu', 'duong_kinh', 'mac_thep'], how='outer')
        df_atp.fillna(0, inplace=True)
        
        df_atp['atp'] = df_atp['ton_l1'] + df_atp['cnk_l1'] - df_atp['kl_can_xuat']
        df_atp['co_don_hang'] = df_atp.apply(lambda row: 1 if (row['kl_can'] > 0 or row['kl_can_xuat'] > 0) else 0, axis=1)

        def get_multi_horiz_chart(filtered_df):
            if filtered_df.empty:
                return {"labels": [], "kl_can": [], "kl_da_ship": [], "kl_can_xuat": [], "groups": []}
            
            grp = filtered_df.sort_values(by=['duong_kinh', 'mac_thep']).reset_index(drop=True)
            grp['dk_num'] = grp['duong_kinh'].apply(extract_num)
            grp = grp.sort_values(['dk_num', 'mac_thep']).reset_index(drop=True)
            
            labels = grp['mac_thep'].tolist()
            kl_can = grp['kl_can'].tolist()
            kl_da_ship = grp['kl_da_ship'].tolist()
            kl_can_xuat = grp['kl_can_xuat'].tolist()
            
            groups = []
            current_dk, start_idx = None, 0
            for i, row in grp.iterrows():
                if row['duong_kinh'] != current_dk:
                    if current_dk is not None: groups.append({"label": current_dk, "start": start_idx, "end": i - 1})
                    current_dk, start_idx = row['duong_kinh'], i
            if current_dk is not None: groups.append({"label": current_dk, "start": start_idx, "end": len(grp) - 1})
                
            return {"labels": labels, "kl_can": kl_can, "kl_da_ship": kl_da_ship, "kl_can_xuat": kl_can_xuat, "groups": groups}

        def get_atp_horiz_chart(filtered_df):
            if filtered_df.empty: return {"labels": [], "values": [], "groups": []}
            
            grp = filtered_df.copy()
            grp['dk_num'] = grp['duong_kinh'].apply(extract_num)
            grp = grp.sort_values(['co_don_hang', 'dk_num', 'mac_thep'], ascending=[False, True, True]).reset_index(drop=True)
            
            labels = grp['mac_thep'].tolist()
            values = grp['atp'].tolist()
            
            groups = []
            current_dk, start_idx = None, 0
            for i, row in grp.iterrows():
                nhom_hien_tai = str(row['co_don_hang']) + '_' + str(row['duong_kinh'])
                if nhom_hien_tai != current_dk:
                    if current_dk is not None: 
                        groups.append({"label": current_dk.split('_')[1], "start": start_idx, "end": i - 1})
                    current_dk, start_idx = nhom_hien_tai, i
                    
            if current_dk is not None: 
                groups.append({"label": current_dk.split('_')[1], "start": start_idx, "end": len(grp) - 1})
                
            return {"labels": labels, "values": values, "groups": groups}

        chart_xuat_thanh = get_multi_horiz_chart(df_so_grp[df_so_grp['nhom_vat_tu'] == 'Thép thanh'])
        chart_xuat_cuon = get_multi_horiz_chart(df_so_grp[df_so_grp['nhom_vat_tu'] == 'Thép cuộn'])

        chart_atp_thanh = get_atp_horiz_chart(df_atp[df_atp['nhom_vat_tu'] == 'Thép thanh'])
        chart_atp_cuon = get_atp_horiz_chart(df_atp[df_atp['nhom_vat_tu'] == 'Thép cuộn'])

        tong_sx_ctd = 0
        prod_date_shift_chart = {"labels": [], "datasets": [], "tong_ngay": []}
        prod_grade_chart = {"labels": [], "values": []}
        prod_machine_chart = {"labels": [], "values": []}

        if not df_sx.empty:
            df_sx['ngay_sx'] = pd.to_datetime(df_sx['ngay_sx']).dt.strftime('%d/%m')
            tong_sx_ctd = float(df_sx['khoi_luong_tan'].sum())
            
            df_ds = df_sx.groupby(['ngay_sx', 'ca_sx'])['khoi_luong_tan'].sum().unstack(fill_value=0)
            df_ds['tong_ngay'] = df_ds.sum(axis=1)
            prod_date_shift_chart["labels"] = df_ds.index.tolist()
            prod_date_shift_chart["datasets"] = [{"label": str(shift), "data": df_ds[shift].tolist()} for shift in df_ds.columns if shift != 'tong_ngay']
            prod_date_shift_chart["tong_ngay"] = df_ds['tong_ngay'].tolist()

            df_gs = df_sx.groupby('mac_thep')['khoi_luong_tan'].sum().reset_index().sort_values('khoi_luong_tan', ascending=True)
            prod_grade_chart = {"labels": df_gs['mac_thep'].tolist(), "values": df_gs['khoi_luong_tan'].tolist()}
            
            df_mc = df_sx.groupby('may_can')['khoi_luong_tan'].sum().reset_index()
            prod_machine_chart = {"labels": df_mc['may_can'].tolist(), "values": df_mc['khoi_luong_tan'].tolist()}

        return jsonify({
            "status": "success",
            "last_update": last_update,
            "tong_ton": tong_ton,
            "pie_chart": pie_chart,
            "age_chart": age_chart,
            "market_chart": market_chart,
            "tong_cnk": tong_cnk,
            "pie_chart_cnk": pie_chart_cnk,
            "chart_thanh_nd": chart_thanh_nd,
            "chart_cuon_nd": chart_cuon_nd,
            "chart_clc": chart_clc,
            "chart_thanh_nd_l1_over6m": chart_thanh_nd_l1_over6m,
            "chart_cuon_nd_l1_over6m": chart_cuon_nd_l1_over6m,
            "chart_clc_l1_over6m": chart_clc_l1_over6m,
            "chart_thanh_nd_l2": chart_thanh_nd_l2,
            "chart_cuon_nd_l2": chart_cuon_nd_l2,
            "chart_clc_l2": chart_clc_l2,
            "chart_xuat_thanh": chart_xuat_thanh,
            "chart_xuat_cuon": chart_xuat_cuon,
            "chart_atp_thanh": chart_atp_thanh,
            "chart_atp_cuon": chart_atp_cuon,
            "tong_sx_ctd": tong_sx_ctd,
            "prod_date_shift_chart": prod_date_shift_chart,
            "prod_grade_chart": prod_grade_chart,
            "prod_machine_chart": prod_machine_chart
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

# ==============================================================================
# PHẦN 2: LỊCH TÀU 
# ==============================================================================
@ctd_bp.route('/upload_lichtau_ctd', methods=['POST'])
@permission_required('upload_ctd')
def upload_lichtau_ctd():
    if 'file' not in request.files:
        flash('Không tìm thấy file', 'danger')
        return redirect(url_for('ctd_bp.dashboard_ctd'))
    
    file = request.files['file']
    if file.filename == '':
        flash('Chưa chọn file', 'danger')
        return redirect(url_for('ctd_bp.dashboard_ctd'))

    try:
        now = datetime.now()
        current_m = now.strftime('%m')
        current_y = now.strftime('%Y')
        
        xls = pd.ExcelFile(file)
        target_sheet = None
        for sheet in xls.sheet_names:
            if current_m in sheet and current_y in sheet:
                target_sheet = sheet
                break
        
        if not target_sheet:
            target_sheet = xls.sheet_names[0]

        df = pd.read_excel(xls, sheet_name=target_sheet, skiprows=3)
        df = df.iloc[:, 1:7]
        df.columns = ['ten_tau', 'khoi_luong', 'ngay_du_kien', 'don_vi_van_tai', 'nha_may', 'cang_do']
        df = df.dropna(subset=['ten_tau'])

        def parse_vessel_dates(row):
            val = row['ngay_du_kien']
            if pd.isnull(val):
                return pd.Series([None, None])
            
            if hasattr(val, 'strftime'):
                d_str = val.strftime('%Y-%m-%d')
                return pd.Series([d_str, d_str])
            
            val_str = str(val).strip().replace(' ', '')
            
            if '-' in val_str and '/' in val_str:
                try:
                    parts = val_str.split('/')
                    days = parts[0].split('-')
                    start_d = int(days[0])
                    end_d = int(days[-1])
                    m = int(parts[1])
                    y = int(parts[2]) if len(parts) > 2 else now.year
                    
                    start_date = datetime(y, m, start_d).strftime('%Y-%m-%d')
                    end_date = datetime(y, m, end_d).strftime('%Y-%m-%d')
                    return pd.Series([start_date, end_date])
                except:
                    pass
                    
            try:
                dt = pd.to_datetime(val_str, dayfirst=True)
                if pd.notnull(dt):
                    d_str = dt.strftime('%Y-%m-%d')
                    return pd.Series([d_str, d_str])
            except:
                pass

            return pd.Series([None, None])

        df[['ngay_bat_dau', 'ngay_ket_thuc']] = df.apply(parse_vessel_dates, axis=1)
        df = df.dropna(subset=['ngay_bat_dau'])

        with engine.begin() as conn:
            dtype_mapping = {
                'ten_tau': types.NVARCHAR(length=255),
                'khoi_luong': types.FLOAT(),
                'ngay_du_kien': types.NVARCHAR(length=100),
                'ngay_bat_dau': types.DATE(),
                'ngay_ket_thuc': types.DATE(),
                'don_vi_van_tai': types.NVARCHAR(length=255),
                'nha_may': types.NVARCHAR(length=255),
                'cang_do': types.NVARCHAR(length=255)
            }
            df.to_sql('lich_tau_ctd', conn, if_exists='replace', index=False, dtype=dtype_mapping)
        
        flash(f'Upload Lịch Tàu ({target_sheet}) thành công!', 'success')
    except Exception as e:
        flash(f'Lỗi khi xử lý file Lịch tàu: {str(e)}', 'danger')
        
    return redirect(url_for('ctd_bp.dashboard_ctd'))


@ctd_bp.route('/api/lichtau_ctd')
def get_lichtau_ctd():
    try:
        with engine.connect() as conn:
            check_table = conn.execute(text("SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'lich_tau_ctd'")).scalar()
            if not check_table:
                return jsonify({"status": "empty", "message": "Chưa có dữ liệu Lịch tàu."})

            query_sql = """
                SELECT 
                    ten_tau, 
                    khoi_luong, 
                    ngay_du_kien,
                    ngay_bat_dau,
                    ngay_ket_thuc,
                    FORMAT(ngay_bat_dau, 'yyyy-MM-dd') as start_date, 
                    FORMAT(ngay_ket_thuc, 'yyyy-MM-dd') as end_date, 
                    don_vi_van_tai, 
                    nha_may, 
                    cang_do
                FROM lich_tau_ctd with (nolock)
                ORDER BY ngay_bat_dau ASC
            """
            df_tau = pd.read_sql(query_sql, conn)

        if df_tau.empty:
            return jsonify({"status": "empty"})

        df_tau['ngay_bat_dau'] = pd.to_datetime(df_tau['ngay_bat_dau'])
        df_tau['ngay_ket_thuc'] = pd.to_datetime(df_tau['ngay_ket_thuc'])

        def normalize_text(text):
            if pd.isnull(text): return ""
            return unicodedata.normalize('NFC', str(text)).strip().lower()

        df_tau['ten_tau_norm'] = df_tau['ten_tau'].apply(normalize_text)

        try:
            from db import engine_mysql 
            list_ten_tau = df_tau['ten_tau'].astype(str).str.strip().dropna().unique().tolist()
            
            if list_ten_tau:
                in_clause = ', '.join([f"N'{tau}'" for tau in list_ten_tau])
                min_date = (df_tau['ngay_bat_dau'].min() - pd.Timedelta(days=7)).strftime('%Y-%m-%d')
                max_date = (df_tau['ngay_ket_thuc'].max() + pd.Timedelta(days=7)).strftime('%Y-%m-%d')

                query_mysql = f"""
                    SELECT 
                        Transporter as ten_tau,
                        SAPCode,
                        SAPDescription,
                        Weight as khoi_luong_xuat,
                        DATE(IssueDate) as ngay_xuat
                    FROM bkmis_hpsdq.v_phieuxuathang_hrc 
                    WHERE 
                      Transporter IN ({in_clause})
                      AND IssueDate >= '{min_date}'
                      AND IssueDate <= '{max_date} 23:59:59'
                """
                
                with engine_mysql.connect() as conn_mysql:
                    df_xuat = pd.read_sql(text(query_mysql), conn_mysql)
                
                if not df_xuat.empty:
                    df_xuat['ten_tau_norm'] = df_xuat['ten_tau'].apply(normalize_text)
                    df_xuat['khoi_luong_xuat'] = (pd.to_numeric(df_xuat['khoi_luong_xuat'], errors='coerce').fillna(0)) / 1000.0                   
                    df_xuat['ngay_xuat'] = pd.to_datetime(df_xuat['ngay_xuat'])
            else:
                df_xuat = pd.DataFrame(columns=['ten_tau_norm', 'khoi_luong_xuat', 'ngay_xuat'])

        except Exception as e_mysql:
            df_xuat = pd.DataFrame(columns=['ten_tau_norm', 'khoi_luong_xuat', 'ngay_xuat'])

        def calculate_actual_export_data(row):
            if df_xuat.empty:
                return pd.Series({'da_xuat': 0.0, 'chi_tiet': []})
                
            tau_norm = row['ten_tau_norm']
            start_window = row['ngay_bat_dau'] - pd.Timedelta(days=5)
            end_window = row['ngay_ket_thuc'] + pd.Timedelta(days=5)
            
            mask = (
                (df_xuat['ten_tau_norm'] == tau_norm) & 
                (df_xuat['ngay_xuat'] >= start_window) & 
                (df_xuat['ngay_xuat'] <= end_window)
            )
            
            df_filtered = df_xuat.loc[mask]
            
            if df_filtered.empty:
                return pd.Series({'da_xuat': 0.0, 'chi_tiet': []})
                
            tong_xuat = df_filtered['khoi_luong_xuat'].sum()
            detail_grp = df_filtered.groupby(['SAPCode', 'SAPDescription'])['khoi_luong_xuat'].sum().reset_index()
            chi_tiet = detail_grp.to_dict('records')
            
            return pd.Series({'da_xuat': tong_xuat, 'chi_tiet': chi_tiet})

        res = df_tau.apply(calculate_actual_export_data, axis=1)
        df_tau['da_xuat'] = res['da_xuat']
        df_tau['chi_tiet'] = res['chi_tiet']

        df_tau = df_tau.drop(columns=['ngay_bat_dau', 'ngay_ket_thuc', 'ten_tau_norm'])
        df_tau = df_tau.replace({np.nan: None})
        
        data = df_tau.to_dict(orient='records')
        cang_do_list = [c for c in df_tau['cang_do'].unique() if c is not None]

        return jsonify({
            "status": "success",
            "data": data,
            "filters": {
                "cang_do": cang_do_list
            }
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


# ==============================================================================
# PHẦN 3: LỆNH SẢN XUẤT (LSX) & THUẬT TOÁN THÁC NƯỚC (WATERFALL)
# ==============================================================================

# Helper bóc tách ngày từ chuỗi (VD: "08h00 01/08 - 08h00 03/08")
def parse_lsx_date_range(date_str, default_year=None):
    if not default_year:
        default_year = datetime.now().year
    
    if pd.isnull(date_str):
        return None, None
    
    val_str = str(date_str).strip()
    
    # Tìm tất cả các cụm ngày/tháng dạng dd/mm hoặc dd/mm/yyyy
    matches = re.findall(r'(\d{1,2})[/.-](\d{1,2})(?:[/.-](\d{2,4}))?', val_str)
    
    try:
        if len(matches) >= 2:
            # Match 1: Ngày bắt đầu
            d1, m1, y1 = matches[0]
            year1 = int(y1) if y1 else default_year
            if year1 < 100: year1 += 2000
            start_date = datetime(year1, int(m1), int(d1)).strftime('%Y-%m-%d')
            
            # Match 2: Ngày kết thúc
            d2, m2, y2 = matches[1]
            year2 = int(y2) if y2 else default_year
            if year2 < 100: year2 += 2000
            end_date = datetime(year2, int(m2), int(d2)).strftime('%Y-%m-%d')
            
            return start_date, end_date
        elif len(matches) == 1:
            d1, m1, y1 = matches[0]
            year1 = int(y1) if y1 else default_year
            if year1 < 100: year1 += 2000
            single_date = datetime(year1, int(m1), int(d1)).strftime('%Y-%m-%d')
            return single_date, single_date
    except Exception as ex:
        print(f"Lỗi parse ngày LSX '{val_str}': {str(ex)}")
        
    return None, None


@ctd_bp.route('/upload_lsx_ctd', methods=['POST'])
@permission_required('upload_ctd')
def upload_lsx_ctd():
    uploaded_files = {
        1: request.files.get('file_can1'),
        2: request.files.get('file_can2'),
        3: request.files.get('file_can3')
    }

    has_file = any(f and f.filename != '' for f in uploaded_files.values())
    if not has_file:
        flash('Vui lòng chọn ít nhất 1 file Excel Lệnh sản xuất để upload!', 'danger')
        return redirect(url_for('ctd_bp.dashboard_ctd'))

    now = datetime.now()
    # Ưu tiên tìm sheet dạng MM.YYYY (VD: "08.2026" cho mục đích kiểm thử hiện tại)
    target_sheet_pattern = "08.2026"

    combined_dfs = []

    for may_can, file in uploaded_files.items():
        if not file or file.filename == '':
            continue
        
        try:
            xls = pd.ExcelFile(file)
            selected_sheet = None
            
            # 1. Tìm sheet có định dạng MM.YYYY (ưu tiên 08.2026)
            for sheet in xls.sheet_names:
                clean_sheet = sheet.strip()
                if target_sheet_pattern in clean_sheet:
                    selected_sheet = sheet
                    break
            
            # Nếu không thấy sheet 08.2026, thử tìm theo tháng/năm hiện tại MM.YYYY
            if not selected_sheet:
                cur_pattern = now.strftime('%m.%Y')
                for sheet in xls.sheet_names:
                    if cur_pattern in sheet.strip():
                        selected_sheet = sheet
                        break

            # Nếu vẫn không thấy, lấy sheet đầu tiên
            if not selected_sheet:
                selected_sheet = xls.sheet_names[0]

            # 2. Đọc từ Dòng 9 (skiprows=8), header=None để lấy chuẩn theo Column Index
            df_raw = pd.read_excel(xls, sheet_name=selected_sheet, skiprows=8, header=None)

            if df_raw.empty or df_raw.shape[1] < 18:
                flash(f'File Cán {may_can} ({selected_sheet}) không đúng định dạng hoặc thiếu cột!', 'warning')
                continue

            # Lấy đúng các cột theo vị trí chỉ định:
            # C (Index 2): Thời gian dự kiến sản xuất
            # D (Index 3): Kích cỡ (Đường kính)
            # E (Index 4): Mác thép
            # F (Index 5): Sản lượng /n (tấn)
            # Q (Index 16): Số Order
            # R (Index 17): Số Batch
            df_can = df_raw.iloc[:, [2, 3, 4, 5, 16, 17]].copy()
            df_can.columns = ['thoi_gian_du_kien', 'duong_kinh', 'mac_thep', 'san_luong_ke_hoach', 'so_order', 'so_batch']

            # 3. Làm sạch dữ liệu:
            # Xóa Order rỗng, NaN, hoặc toàn khoảng trắng
            df_can['so_order'] = df_can['so_order'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df_can = df_can[~df_can['so_order'].isin(['', 'nan', 'None', 'NAT', 'null', 'NULL'])].copy()

            if df_can.empty:
                continue

            df_can['so_batch'] = df_can['so_batch'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df_can['duong_kinh'] = df_can['duong_kinh'].astype(str).str.strip()
            df_can['mac_thep'] = df_can['mac_thep'].astype(str).str.strip()
            df_can['san_luong_ke_hoach'] = pd.to_numeric(df_can['san_luong_ke_hoach'], errors='coerce').fillna(0.0)

            # Gán Máy Cán (1, 2, 3 kiểu INT)
            df_can['may_can'] = int(may_can)

            # Parse ngày bắt đầu, ngày kết thúc
            year_hint = 2026 if "2026" in selected_sheet else now.year
            dates = df_can['thoi_gian_du_kien'].apply(lambda x: parse_lsx_date_range(x, default_year=year_hint))
            df_can['ngay_bat_dau'] = [d[0] for d in dates]
            df_can['ngay_ket_thuc'] = [d[1] for d in dates]

            # Loại bỏ dòng không xác định được ngày
            df_can = df_can.dropna(subset=['ngay_bat_dau']).copy()

            combined_dfs.append(df_can)

        except Exception as e:
            flash(f'Lỗi khi xử lý file Cán {may_can}: {str(e)}', 'danger')
            return redirect(url_for('ctd_bp.dashboard_ctd'))

    if not combined_dfs:
        flash('Không có dữ liệu hợp lệ nào được trích xuất từ các file Excel!', 'danger')
        return redirect(url_for('ctd_bp.dashboard_ctd'))

    df_final = pd.concat(combined_dfs, ignore_index=True)
    df_final['ngay_bat_dau'] = pd.to_datetime(df_final['ngay_bat_dau']).dt.date
    df_final['ngay_ket_thuc'] = pd.to_datetime(df_final['ngay_ket_thuc']).dt.date

    try:
        with engine.begin() as conn:
            dtype_mapping = {
                'may_can': types.INTEGER(),
                'thoi_gian_du_kien': types.NVARCHAR(length=100),
                'duong_kinh': types.NVARCHAR(length=50),
                'mac_thep': types.NVARCHAR(length=100),
                'san_luong_ke_hoach': types.FLOAT(),
                'so_order': types.NVARCHAR(length=100),
                'so_batch': types.NVARCHAR(length=100),
                'ngay_bat_dau': types.DATE(),
                'ngay_ket_thuc': types.DATE()
            }
            # Ghi đè bảng Lệnh sản xuất
            df_final.to_sql('lenh_san_xuat_ctd', conn, if_exists='replace', index=False, dtype=dtype_mapping)

        flash(f'Upload và xử lý thành công {len(df_final)} dòng Lệnh Sản Xuất!', 'success')
    except Exception as e:
        flash(f'Lỗi lưu dữ liệu LSX vào Database: {str(e)}', 'danger')

    return redirect(url_for('ctd_bp.dashboard_ctd'))


@ctd_bp.route('/api/lsx_ctd')
def get_lsx_ctd():
    try:
        with engine.connect() as conn:
            check_table = conn.execute(text("SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'lenh_san_xuat_ctd'")).scalar()
            if not check_table:
                return jsonify({"status": "empty", "message": "Chưa có dữ liệu Lệnh Sản Xuất."})

            # 1. Lấy dữ liệu Kế hoạch LSX, sắp xếp theo Order và ngày bắt đầu (FIFO)
            query_lsx = """
                SELECT 
                    may_can,
                    thoi_gian_du_kien,
                    duong_kinh,
                    mac_thep,
                    san_luong_ke_hoach,
                    so_order,
                    so_batch,
                    FORMAT(ngay_bat_dau, 'yyyy-MM-dd') as start_date,
                    FORMAT(ngay_ket_thuc, 'yyyy-MM-dd') as end_date,
                    ngay_bat_dau,
                    ngay_ket_thuc
                FROM lenh_san_xuat_ctd with (nolock)
                WHERE san_luong_ke_hoach > 0
                ORDER BY ngay_bat_dau ASC, may_can ASC
            """
            df_lsx = pd.read_sql(query_lsx, conn)

            if df_lsx.empty:
                return jsonify({"status": "empty"})

            # 2. Lấy dữ liệu Thực tế Sản lượng từ bảng sanluong (WHERE NhaMay = 'CTD')
            query_sl = """
                SELECT 
                    RTRIM(LTRIM(CAST([Order] AS NVARCHAR(100)))) AS so_order,
                    SUM([Khối lượng] / 1000.0) AS tong_da_san_xuat
                FROM sanluong WITH (NOLOCK)
                WHERE NhaMay = 'CTD' 
                  AND [Khối lượng] > 0
                  AND [Order] IS NOT NULL
                GROUP BY [Order]
            """
            df_sl = pd.read_sql(query_sl, conn)

        # Tạo Dictionary lưu tổng sản lượng thực tế theo từng Order
        actual_order_pool = {}
        if not df_sl.empty:
            for _, row in df_sl.iterrows():
                ord_key = str(row['so_order']).strip()
                actual_order_pool[ord_key] = float(row['tong_da_san_xuat'])

        # 3. THUẬT TOÁN THÁC NƯỚC (WATERFALL ALLOCATION)
        # Sắp xếp theo Order & Ngày bắt đầu để trừ dần sản lượng cho lệnh cũ trước (FIFO)
        df_lsx = df_lsx.sort_values(by=['so_order', 'ngay_bat_dau'], ascending=[True, True]).reset_index(drop=True)

        da_sx_list = []
        con_lai_list = []
        percent_list = []
        is_completed_list = []

        for _, row in df_lsx.iterrows():
            ord_key = str(row['so_order']).strip()
            ke_hoach = float(row['san_luong_ke_hoach'])
            
            # Lấy lượng tồn thực tế của Order còn lại trong pool
            available_prod = actual_order_pool.get(ord_key, 0.0)
            
            # Phân bổ thác nước
            allocated = min(ke_hoach, available_prod)
            actual_order_pool[ord_key] = max(0.0, available_prod - allocated)
            
            remaining = max(0.0, ke_hoach - allocated)
            pct = (allocated / ke_hoach * 100.0) if ke_hoach > 0 else 0.0
            
            # Điều kiện hoàn thành: Đạt >= 98%
            completed = (pct >= 98.0)

            da_sx_list.append(round(allocated, 2))
            con_lai_list.append(round(remaining, 2))
            percent_list.append(round(pct, 1))
            is_completed_list.append(completed)

        df_lsx['da_san_xuat'] = da_sx_list
        df_lsx['con_lai'] = con_lai_list
        df_lsx['percent'] = percent_list
        df_lsx['is_completed'] = is_completed_list

        # Sắp xếp lại để hiển thị Gantt: Gom nhóm theo Đường kính -> Mác thép -> Ngày
        def extract_num(val):
            nums = re.findall(r'\d+\.?\d*', str(val))
            return float(nums[0]) if nums else 0.0

        df_lsx['dk_num'] = df_lsx['duong_kinh'].apply(extract_num)
        df_lsx = df_lsx.sort_values(by=['dk_num', 'mac_thep', 'ngay_bat_dau', 'may_can'], ascending=[True, True, True, True]).reset_index(drop=True)

        # Tính toán KPI Tổng
        tong_yeu_cau = float(df_lsx['san_luong_ke_hoach'].sum())
        tong_da_sx = float(df_lsx['da_san_xuat'].sum())
        tong_con_lai = max(0.0, tong_yeu_cau - tong_da_sx)

        # Danh mục filters trả về Frontend
        filter_duong_kinh = sorted([x for x in df_lsx['duong_kinh'].unique() if x], key=extract_num)
        filter_mac_thep = sorted([x for x in df_lsx['mac_thep'].unique() if x])
        filter_may_can = [1, 2, 3]

        # Xóa các cột tạm
        df_lsx = df_lsx.drop(columns=['dk_num', 'ngay_bat_dau', 'ngay_ket_thuc'])
        df_lsx = df_lsx.replace({np.nan: None})

        return jsonify({
            "status": "success",
            "kpi": {
                "tong_yeu_cau": tong_yeu_cau,
                "tong_da_sx": tong_da_sx,
                "tong_con_lai": tong_con_lai
            },
            "filters": {
                "duong_kinh": filter_duong_kinh,
                "mac_thep": filter_mac_thep,
                "may_can": filter_may_can
            },
            "data": df_lsx.to_dict(orient='records')
        })

    except Exception as e:
        print("\n!!! [LỖI API LSX] !!!", str(e))
        return jsonify({"status": "error", "message": str(e)})