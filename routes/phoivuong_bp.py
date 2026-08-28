from flask import Blueprint, render_template, request, jsonify, flash, redirect, url_for
import pandas as pd
from sqlalchemy import text, types
from db import engine, engine_phoivuong_mysql
from datetime import datetime
from auth.decorator import permission_required
phoivuong_bp = Blueprint('phoivuong_bp', __name__, url_prefix='/phoivuong')

# ==========================================
# ROUTE MỚI: TEST KẾT NỐI MYSQL TRỰC TIẾP
# ==========================================
@phoivuong_bp.route('/test_mysql')
def test_mysql():
    try:
        with engine_phoivuong_mysql.connect() as conn:
            result = conn.execute(text("SELECT 1")).scalar()
            return f"✅ KẾT NỐI MYSQL THÀNH CÔNG! Kết quả test trả về: {result}"
    except Exception as e:
        return f"❌ LỖI KẾT NỐI MYSQL: {str(e)}"

@phoivuong_bp.route('/phoivuong_dashboard')
@permission_required('pv_view')
def phoivuong_dashboard():
    return render_template('phoivuong_dashboard.html')

@phoivuong_bp.route('/upload_mapping', methods=['POST'])
@permission_required('upload_pv')
def upload_mapping():
    if 'file' not in request.files:
        flash('Không tìm thấy file', 'danger')
        return redirect(url_for('phoivuong_bp.phoivuong_dashboard'))
    
    file = request.files['file']
    if file.filename == '':
        flash('Chưa chọn file', 'danger')
        return redirect(url_for('phoivuong_bp.phoivuong_dashboard'))

    try:
        df = pd.read_excel(file)
        df.columns = [c.strip().lower() for c in df.columns] 
        
        required_columns = ['ma_vat_tu', 'ten_khach_hang']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            flash(f"Lỗi: File Excel thiếu các cột bắt buộc: {', '.join(missing_columns)}", 'danger')
            return redirect(url_for('phoivuong_bp.phoivuong_dashboard'))
        df['ma_vat_tu'] = df['ma_vat_tu'].astype(str).str.strip()
        df['ma_vat_tu'] = df['ma_vat_tu'].str.replace(r'\.0$', '', regex=True)
        
        df['ten_khach_hang'] = df['ten_khach_hang'].astype(str).str.strip()
        df = df.drop_duplicates(subset=['ma_vat_tu'], keep='first')
        
        with engine.begin() as conn:
            dtype_mapping = {
                'ma_vat_tu': types.NVARCHAR(length=50),
                'ten_khach_hang': types.NVARCHAR(length=255)
            }
            df.to_sql('mapping_khachhang_pv', conn, if_exists='replace', index=False, dtype=dtype_mapping)
            conn.execute(text("CREATE CLUSTERED INDEX IX_mapping_mvt ON mapping_khachhang_pv(ma_vat_tu)"))
        
        flash('Upload Mapping và Đồng bộ Khách hàng thành công!', 'success')
    except Exception as e:
        flash(f'Lỗi khi xử lý file: {str(e)}', 'danger')
        
    return redirect(url_for('phoivuong_bp.phoivuong_dashboard'))

@phoivuong_bp.route('/api/data')
def get_dashboard_data():
    try:
        # ==========================================
        # 1. KẾT NỐI SQL SERVER (DỮ LIỆU TỒN KHO)
        # ==========================================
        with engine.connect() as conn:
            query = """
                SELECT 
                    k.ma_vat_tu, 
                    ISNULL(k.mac_thep, k.loai_vat_tu) as mac_thep,
                    k.sl_cuoi_ky as khoi_luong,
                    k.ngay_nhap_cuoi,
                    k.snapshot_ts,
                    ISNULL(m.ten_khach_hang, N'Phôi dân dụng') as khach_hang,
                    CASE 
                        WHEN m.ten_khach_hang IS NOT NULL THEN 'CLC' 
                        ELSE N'Phôi dân dụng' 
                    END as phan_loai_kh,
                    CASE 
                        WHEN RIGHT(RTRIM(k.so_lo), 2) = 'II' THEN N'Loại 2' 
                        ELSE N'Loại 1' 
                    END as loai_hang,
                    DATEDIFF(day, k.ngay_nhap_cuoi, GETDATE()) as tuoi_ton_kho
                FROM kho_ctd_pv k with (nolock)
                LEFT JOIN mapping_khachhang_pv m with (nolock) ON k.ma_vat_tu = m.ma_vat_tu
                WHERE k.nha_may = 'PhoiVuong' AND k.sl_cuoi_ky > 0
            """
            df = pd.read_sql(query, conn)

        if df.empty:
            return jsonify({"status": "empty"})

        last_update = df['snapshot_ts'].max().strftime("%d/%m/%Y %H:%M:%S")
        tong_ton = float(df['khoi_luong'].sum())

        df_kh = df.groupby('khach_hang')['khoi_luong'].sum().reset_index()
        pie_chart = {"labels": df_kh['khach_hang'].tolist(), "values": df_kh['khoi_luong'].tolist()}

        df_type = df.groupby(['khach_hang', 'loai_hang'])['khoi_luong'].sum().unstack(fill_value=0).reset_index()
        type_chart = {
            "labels": df_type['khach_hang'].tolist(),
            "loai1": df_type.get('Loại 1', pd.Series([0]*len(df_type))).tolist(),
            "loai2": df_type.get('Loại 2', pd.Series([0]*len(df_type))).tolist()
        }

        bins = [-1, 90, 180, 365, float('inf')]
        labels = ['Dưới 3 tháng', '3-6 tháng', '6-12 tháng', 'Trên 1 năm']
        df['nhom_tuoi'] = pd.cut(df['tuoi_ton_kho'], bins=bins, labels=labels)
        df_age = df.groupby('nhom_tuoi')['khoi_luong'].sum().reset_index()
        age_chart = {"labels": labels, "values": df_age['khoi_luong'].tolist()}

        def process_group_chart(kh_filter, is_dd=False):
            df_filtered = df[df['phan_loai_kh'] == kh_filter].copy()
            if df_filtered.empty: 
                return {"labels": [], "loai1": [], "loai2": [], "groups": []}
            
            if is_dd:
                # DÂN DỤNG: Chỉ lấy Mác thép, không có Nhóm
                df_filtered['label_col'] = df_filtered['mac_thep'].fillna('Không rõ')
                df_grp = df_filtered.groupby(['label_col', 'loai_hang'])['khoi_luong'].sum().unstack(fill_value=0).reset_index()
                for col in ['Loại 1', 'Loại 2']:
                    if col not in df_grp.columns: df_grp[col] = 0.0
                df_grp = df_grp.sort_values('label_col')
                return {
                    "labels": df_grp['label_col'].tolist(),
                    "loai1": df_grp['Loại 1'].tolist(),
                    "loai2": df_grp['Loại 2'].tolist(),
                    "groups": []
                }
            else:
                # CLC: Nhóm Rowspan theo Khách hàng -> Mác thép
                df_filtered['khach_hang'] = df_filtered['khach_hang'].fillna('Không rõ')
                df_filtered['mac_thep'] = df_filtered['mac_thep'].fillna('Không rõ')
                
                # Group theo cả 2 cấp
                df_grp = df_filtered.groupby(['khach_hang', 'mac_thep', 'loai_hang'])['khoi_luong'].sum().unstack(fill_value=0).reset_index()
                
                for col in ['Loại 1', 'Loại 2']:
                    if col not in df_grp.columns: df_grp[col] = 0.0
                
                # Sắp xếp: Khách Hàng trước, Mác thép sau
                df_grp = df_grp.sort_values(['khach_hang', 'mac_thep']).reset_index(drop=True)
                
                labels = df_grp['mac_thep'].tolist()
                loai1 = df_grp['Loại 1'].tolist()
                loai2 = df_grp['Loại 2'].tolist()
                
                # Tính toán tọa độ gộp ô (Rowspan) cho Frontend
                groups = []
                current_kh = None
                start_idx = 0
                for i, row in df_grp.iterrows():
                    kh = row['khach_hang']
                    if kh != current_kh:
                        if current_kh is not None:
                            groups.append({"label": current_kh, "start": start_idx, "end": i - 1})
                        current_kh = kh
                        start_idx = i
                if current_kh is not None:
                    groups.append({"label": current_kh, "start": start_idx, "end": len(df_grp) - 1})
                    
                return {
                    "labels": labels,
                    "loai1": loai1,
                    "loai2": loai2,
                    "groups": groups
                }

        clc_chart = process_group_chart('CLC', is_dd=False)
        dd_chart = process_group_chart('Phôi dân dụng', is_dd=True)

        # ==========================================
        # 2. KẾT NỐI MYSQL (DỮ LIỆU SẢN XUẤT) CÓ BẮT LỖI
        # ==========================================
        tong_sx = 0
        prod_date_shift_chart = {"labels": [], "datasets": []}
        prod_grade_size_chart = {"labels": [], "values": []}
        prod_machine_chart = {"labels": [], "values": []}
        mysql_error = None # Biến hứng lỗi

        try:
            with engine_phoivuong_mysql.connect() as conn_mysql:
                # BỎ hoàn toàn hàm DATE_FORMAT trong SQL để tránh lỗi ký tự %
                query_sx = """
                    SELECT 
                        ShiftName as ca_sx,
                        ProductionDate as ngay_sx,
                        MayDuc as may_duc,
                        GradeCode as mac_thep,
                        CONCAT(ProductSizeCode, 'x', CAST(Length * 1000 AS UNSIGNED)) as kich_thuoc,
                        (Weight / 1000.0) as khoi_luong_tan
                    FROM bkmis_kcshpsdq.view_dq1_nmlt_sanluongphoi
                    WHERE MONTH(ProductionDate) = MONTH(CURRENT_DATE())
                      AND YEAR(ProductionDate) = YEAR(CURRENT_DATE())
                """
                df_sx = pd.read_sql(query_sx, conn_mysql)
                
                if df_sx.empty:
                    mysql_error = "Truy vấn thành công nhưng View không có dữ liệu của Tháng hiện tại!"
                else:
                    # ĐỊNH DẠNG NGÀY THÁNG TRỰC TIẾP BẰNG PANDAS (An toàn 100%)
                    df_sx['ngay_sx'] = pd.to_datetime(df_sx['ngay_sx']).dt.strftime('%d/%m')
                    
                    tong_sx = float(df_sx['khoi_luong_tan'].sum())
                    
                    df_ds = df_sx.groupby(['ngay_sx', 'ca_sx'])['khoi_luong_tan'].sum().unstack(fill_value=0)
                    df_ds['tong_ngay'] = df_ds.sum(axis=1)
                    prod_date_shift_chart["labels"] = df_ds.index.tolist()
                    prod_date_shift_chart["datasets"] = [
                        {"label": str(shift), "data": df_ds[shift].tolist()} 
                        for shift in df_ds.columns if shift != 'tong_ngay' 
                    ]
                    prod_date_shift_chart["tong_ngay"] = df_ds['tong_ngay'].tolist()
                    df_sx['nhan_thep'] = df_sx['mac_thep'].astype(str) + ' - ' + df_sx['kich_thuoc'].astype(str)
                    df_gs = df_sx.groupby('nhan_thep')['khoi_luong_tan'].sum().reset_index().sort_values('khoi_luong_tan', ascending=True)
                    prod_grade_size_chart = {
                        "labels": df_gs['nhan_thep'].tolist(),
                        "values": df_gs['khoi_luong_tan'].tolist()
                    }
                    
                    df_mc = df_sx.groupby('may_duc')['khoi_luong_tan'].sum().reset_index()
                    prod_machine_chart = {
                        "labels": ('Máy ' + df_mc['may_duc'].astype(str)).tolist(),
                        "values": df_mc['khoi_luong_tan'].tolist()
                    }
        except Exception as e_mysql:
            mysql_error = f"Lỗi MySQL: {str(e_mysql)}"

        # ==========================================
        # 3. TRẢ VỀ JSON GỘP
        # ==========================================
        return jsonify({
            "status": "success",
            "last_update": last_update,
            "tong_ton": tong_ton,
            "pie_chart": pie_chart,
            "type_chart": type_chart,
            "age_chart": age_chart,
            "clc_chart": clc_chart,
            "dd_chart": dd_chart,
            
            # Khối sản xuất
            "tong_sx": tong_sx,
            "mysql_error": mysql_error, # Trả lỗi ra API
            "prod_date_shift_chart": prod_date_shift_chart,
            "prod_grade_size_chart": prod_grade_size_chart,
            "prod_machine_chart": prod_machine_chart
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})