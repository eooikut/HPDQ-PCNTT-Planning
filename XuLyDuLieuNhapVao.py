import pandas as pd
import re, os, math
from db import engine
import pandas as pd
from dateutil import parser
import math
# ---------------- Cấu hình kết nối SQL Server ----------------
# Thay thông tin server/user/pwd/db của bạn:


# ===== HÀM ĐỌC FILE CSV/XLSX/XLS =====
def read_file_auto(file_path, **kwargs):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        return pd.read_csv(file_path, encoding="utf-8-sig", **kwargs)
    elif ext == ".xlsx":
        return pd.read_excel(file_path, engine="openpyxl", **kwargs)
    elif ext == ".xls":
        return pd.read_excel(file_path, engine="xlrd", **kwargs)
    else:
        raise ValueError(f"Không hỗ trợ định dạng file: {ext}")

# ===== LẤY RANGE THỜI GIAN LSX =====
def get_lsx_range_from_file(file_path, sheet_name=0, row_index=5, col_index=0):
    val = read_file_auto(file_path, sheet_name=sheet_name, header=None).iloc[row_index, col_index]
    if pd.isna(val):
        return None, None
    text = str(val)
    found = re.findall(r"(\d{2}/\d{2}/\d{4})", text)
    if len(found) >= 2:
        return (pd.to_datetime(found[0], dayfirst=True, errors="coerce"),
                pd.to_datetime(found[-1], dayfirst=True, errors="coerce"))
    if len(found) == 1:
        d = pd.to_datetime(found[0], dayfirst=True, errors="coerce")
        return d, d
    return None, None

def extract_dates(val):
    if pd.isna(val):
        return None, None
    text = str(val).replace("\n", " ").strip()
    found = re.findall(r"(\d{2}/\d{2}/\d{4})", text)
    if len(found) >= 2:
        return (pd.to_datetime(found[0], dayfirst=True, errors="coerce"),
                pd.to_datetime(found[-1], dayfirst=True, errors="coerce"))
    if len(found) == 1:
        d = pd.to_datetime(found[0], dayfirst=True, errors="coerce")
        return d, d
    return None, None

# ===== XỬ LÝ FILE LSX =====
def process_lsx(file_path, sheet_name=3, skip_rows=6):
    """
    Đọc file Excel, xử lý dữ liệu và trả về final_df chuẩn cho database.
    
    Args:
        file_path (str): Đường dẫn file Excel.
        sheet_name (str|None): Tên sheet. Mặc định None.
        skip_rows (int): Số dòng bỏ qua đầu file. Mặc định 6.
    
    Returns:
        pd.DataFrame: DataFrame đã xử lý, chuẩn cho insert vào SQL Server.
    """
    
    # ---------- Đọc file ----------
    df = read_file_auto(file_path, sheet_name=sheet_name, skiprows=skip_rows)
    df.columns = [str(c).strip() for c in df.columns]

    # ---------- Tìm cột thời gian ----------
    time_col_candidates = [c for c in df.columns if "Thời gian" in c or "Time/Date" in c]
    time_col = time_col_candidates[0] if time_col_candidates else None
    if time_col:
        df[time_col] = df[time_col].ffill()
        block_days = df[[time_col]].drop_duplicates().copy()
        block_days[["Ngày bắt đầu block","Ngày kết thúc block"]] = block_days[time_col].apply(
            lambda x: pd.Series(extract_dates(x))
        )
        for c in ["Ngày bắt đầu block","Ngày kết thúc block"]:
            block_days[c] = pd.to_datetime(block_days[c], dayfirst=True, errors="coerce")
        block_days["Số ngày yêu cầu block"] = (
            (block_days["Ngày kết thúc block"] - block_days["Ngày bắt đầu block"]).dt.days + 1
        )
        df = df.merge(block_days, on=time_col, how="left")

    # ---------- Tìm cột Order ----------
    order_candidates = [c for c in df.columns if re.search(r"order", c, re.IGNORECASE) or "Số Order" in c]
    order_col = order_candidates[0] if order_candidates else None
    if not order_col and "Order" in df.columns:
        order_col = "Order"
    if not order_col:
        raise RuntimeError("Không tìm thấy cột Order.")

    # ---------- Fill text columns ----------
    numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
    cols_to_ffill = [c for c in df.columns if c not in numeric_cols]
    df[cols_to_ffill] = df[cols_to_ffill].ffill()

    # ---------- KHÁCH HÀNG ----------
    if "KHÁCH HÀNG" in df.columns:
        df["KHÁCH HÀNG"] = df["KHÁCH HÀNG"].fillna("Chưa có KHÁCH HÀNG")
    else:
        df["KHÁCH HÀNG"] = "Chưa có KHÁCH HÀNG"

    # ---------- Chuyển cột sản lượng ----------
    for col in ["Unnamed: 4","Unnamed: 5"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df = df.rename(columns={"Unnamed: 4": "Sản lượng 1A", "Unnamed: 5": "Sản lượng 1B"})

    # ---------- Tính tổng và trung bình/ngày ----------
    agg_df = df.groupby(
        [order_col, "KHÁCH HÀNG", "Ngày bắt đầu block", "Ngày kết thúc block", "Số ngày yêu cầu block"],
        as_index=False
    ).agg({"Sản lượng 1A":"sum", "Sản lượng 1B":"sum"})

    agg_df["SL yêu cầu (tấn)"] = (agg_df["Sản lượng 1A"] + agg_df["Sản lượng 1B"])
    agg_df["SL trung bình/ngày"] = agg_df["SL yêu cầu (tấn)"] / agg_df["Số ngày yêu cầu block"]
    agg_df = agg_df.rename(columns={order_col: "Order", "SL yêu cầu (tấn)": "Tổng yêu cầu"})

    # ---------- Giữ các cột chi tiết khác ----------
    detail_cols_candidates = []
    keywords = ["Kích", "mác", "phôi", "kích thước", "yêu cầu", "số lô", "khối lượng", "cuộn", "mục đích"]
    for c in df.columns:
        cname = c.lower()
        if any(k.lower() in cname for k in keywords):
            detail_cols_candidates.append(c)

    def join_unique(vals):
        vals = pd.Series(vals.dropna().astype(str).unique())
        vals = vals[vals != "nan"]
        if len(vals) == 0:
            return pd.NA
        return " | ".join(vals)

    detail_map = {col: join_unique for col in detail_cols_candidates}
    detail_map["KHÁCH HÀNG"] = join_unique

    detail_group = df.groupby(order_col, as_index=False).agg(detail_map)
    if order_col != "Order":
        detail_group = detail_group.rename(columns={order_col: "Order"})

    # ---------- Merge final_df ----------
    final_df = pd.merge(agg_df, detail_group, on="Order", how="left")
    final_df = final_df.rename(columns={"KHÁCH HÀNG_x": "KHÁCH HÀNG", "Phôi cán/Slab": "Mac thep"})

    # ---------- Drop các cột không cần thiết ----------
    cols_to_drop = ["KHÁCH HÀNG_y", "Số ngày yêu cầu block_x", "Số ngày yêu cầu block_y",
                    "Số lô/\nBatch\ntháng 9","Số lượng cuộn yêu cầu", "Số cuộn tối thiểu", "Số cuộn tối đa"]
    final_df = final_df.drop(columns=[c for c in cols_to_drop if c in final_df.columns])

    # ---------- Chuẩn hóa tên cột ----------
    final_df.columns = final_df.columns.str.replace(r"[\n/]", "_", regex=True).str.strip()

    final_df = final_df.rename(columns={
        'Số lô__Batch': 'Số_lô_Batch',
        'KL Cuộn_(Tấn)': 'KL_Cuộn_(Tấn)',
        'SL trung bình_ngày':'SL trung bình/ngày'
    })

    # ---------- Chuyển các cột số sang int/float an toàn ----------
    numeric_cols_int = ['Số_lô_Batch']
    numeric_cols_float = ['Khối lượng cuộn trung bình']

    for col in numeric_cols_int:
        if col in final_df.columns:
            final_df[col] = pd.to_numeric(final_df[col], errors='coerce').fillna(0).astype(int)

    for col in numeric_cols_float:
        if col in final_df.columns:
            final_df[col] = pd.to_numeric(final_df[col], errors='coerce').fillna(0.0).astype(float)

    return final_df

# ===== XỬ LÝ SẢN LƯỢNG THỰC TẾ =====
def process_actual(file_path, sheet_name="Data"):
    df = read_file_auto(file_path).dropna(how="all")
    df["Ngày sản xuất"] = pd.to_datetime(df["Ngày sản xuất"], errors="coerce")
    df["Khối lượng"] = pd.to_numeric(df["Khối lượng"], errors="coerce")
    df = df.dropna(subset=["Order","Ngày sản xuất"])

    df_daily = df.groupby(["Order","Ngày sản xuất"], as_index=False)["Khối lượng"].sum()
    df_daily = df_daily.rename(columns={
        "Khối lượng":"Sản lượng thực tế",
        "Ngày sản xuất":"Ngày"
    })
    total_actual = df_daily.groupby("Order", as_index=False)["Sản lượng thực tế"].sum()
    total_actual = total_actual.rename(columns={"Sản lượng thực tế":"Tổng sản lượng thực tế"})
    return df_daily, total_actual

# ===== CLASSIFY =====
##Xử lý dữ liệu file TÀU

from datetime import datetime, timedelta
import re

def filter_sheets_from_month(sheet_names):
    """
    Phiên bản SỬA LỖI: Ép buộc định dạng ngày dd.mm.yyyy để tránh hiểu sai.
    """
    now = datetime.now()
    
    # 1. Tính toán mốc thời gian (Start Date)
    if now.day > 10:
        # Nếu > ngày 10: Lấy từ ngày 1 tháng hiện tại
        start_dt = datetime(now.year, now.month, 1)
    else:
        # Nếu <= ngày 10: Lấy từ ngày 1 tháng trước
        first_day_this_month = datetime(now.year, now.month, 1)
        prev_month_date = first_day_this_month - timedelta(days=1)
        start_dt = datetime(prev_month_date.year, prev_month_date.month, 1)
    
    # In ra để kiểm tra mốc thời gian hệ thống đang hiểu
    print(f"\n⚡ [HỆ THỐNG] Hôm nay: {now.strftime('%d/%m/%Y')}")
    print(f"⚡ [MỐC LỌC] Chỉ lấy các sheet từ ngày: {start_dt.strftime('%d/%m/%Y')} trở đi")

    filtered = []
    for s in sheet_names:
        s_clean = s.strip()
        # Regex bắt: LỊCH TÀU - 10.2025
        m = re.match(r"LỊCH TÀU - (\d{2}\.\d{4})", s_clean, re.IGNORECASE)
        if m:
            month_str = m.group(1) # VD: 10.2025
            try:
                # 2. ÉP KIỂU CHÍNH XÁC: dd.mm.yyyy
                # Tạo chuỗi ngày 01.10.2025
                date_str = f"01.{month_str}"
                sheet_dt = datetime.strptime(date_str, "%d.%m.%Y")
                
                # 3. So sánh
                if sheet_dt >= start_dt:
                    filtered.append(s)
                    print(f"  ✅ Lấy: {s} (Hiểu là {sheet_dt.strftime('%d/%m/%Y')})")
                else:
                    print(f"  ❌ Bỏ : {s} (Hiểu là {sheet_dt.strftime('%d/%m/%Y')} < Mốc)")
            except Exception as e:
                print(f"  ⚠️ Lỗi parse sheet {s}: {e}")
                continue
    
    # Sắp xếp lại danh sách
    filtered.sort(key=lambda x: datetime.strptime("01." + re.search(r"(\d{2}\.\d{4})", x).group(1), "%d.%m.%Y"))
    
    return filtered



def parse_eta(eta):
    """Chuẩn hóa giá trị ETA thành datetime, lấy ngày đầu tiên hợp lệ."""
    if isinstance(eta, datetime): 
        # Nếu đã là datetime (hoặc Timestamp, vì Timestamp là một subclass của datetime)
        
        # Giả định: Pandas đã đọc D/M/Y (6/11) thành M/D/Y (Tháng 6, Ngày 11)
        # Ta thực hiện hoán đổi: (month=6, day=11) -> (month=11, day=6)
        
        original_day = eta.day    # = 11
        original_month = eta.month # = 6
        original_year = eta.year
        
        # Thử tạo datetime mới bằng cách hoán đổi Ngày và Tháng
        try:
            # datetime(Năm, Tháng Mới (11), Ngày Mới (6))
            # Nếu 11 là tháng, 6 là ngày, điều này hợp lệ.
            return datetime(original_year, original_day, original_month)
        except ValueError:
            # Nếu việc hoán đổi không hợp lệ (ví dụ: 32/12/2025 bị đọc thành 12/32/2025)
            # Thì ta chấp nhận giá trị đã được Pandas tạo ra
            return eta
    if not eta or not isinstance(eta, str):
        return None

    eta_original = eta  # lưu để debug nếu cần
    eta = eta.strip().upper()
    current_year = datetime.now().year
    iso_match = re.match(r"^\d{4}-\d{2}-\d{2}(?:\s+\d{2}:\d{2}:\d{2}(?:\.\d+)?)?$", eta_original.strip())
    if iso_match:
        try:
            return datetime.fromisoformat(eta_original.strip().split('.')[0])
        except Exception:
            pass 
    # 🔹 Bỏ ngoặc và chữ
    eta = re.sub(r"\(.*?\)", " ", eta)
    eta = re.sub(r"[^0-9./:\-\s]", " ", eta)
    eta = re.sub(r"\s+", " ", eta).strip()

    # 🧩 Các pattern đặc biệt cần xử lý trước
    patterns = [
        # 1️⃣ Dải ngày có tháng và năm: 25/10-29/10/2025
        (r"^(\d{1,2})[./](\d{1,2})-(\d{1,2})[./](\d{1,2})[./](\d{4})$", lambda g: f"{g[0]}.{g[1]}.{g[4]}"),
        # 2️⃣ Dải ngày cùng tháng, có năm: 06-08.10.2025
        (r"^(\d{1,2})-(\d{1,2})[./](\d{1,2})[./](\d{4})$", lambda g: f"{g[0]}.{g[2]}.{g[3]}"),
        # 3️⃣ Dải ngày cùng tháng, không có năm: 13.10-15.10 hoặc 03-04.09
        (r"^(\d{1,2})[./-](\d{1,2})[./-](\d{1,2})$", lambda g: f"{g[0]}.{g[1]}.{current_year}"),
        # 4️⃣ Dải ngày giao tháng: 30.09-3.10 (năm hiện tại)
        (r"^(\d{1,2})[./-](\d{1,2})[./-](\d{1,2})[./-](\d{1,2})$", lambda g: f"{g[0]}.{g[1]}.{current_year}"),
        # 5️⃣ Dải ngày kiểu 8-10/9/2025
        (r"^(\d{1,2})-(\d{1,2})/(\d{1,2})/(\d{4})$", lambda g: f"{g[0]}.{g[2]}.{g[3]}"),
        # 6️⃣ Ngày ISO / SQL: 2025-12-09 00:00:00 hoặc 2025-12-09 00:00:00.000
        (r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:\s+\d{1,2}:\d{2}:\d{2}(?:\.\d{1,3})?)?$", lambda g: f"{g[2]}.{g[1]}.{g[0]}"),
        # 7️⃣ Ngày đơn đầy đủ dd.mm.yyyy hoặc dd/mm/yyyy
        (r"^(\d{1,2})[./](\d{1,2})[./](\d{4})$", lambda g: f"{g[0]}.{g[1]}.{g[2]}"),
        # 8️⃣ Ngày đơn thiếu năm: dd.mm hoặc dd/mm
        (r"^(\d{1,2})[./-](\d{1,2})$", lambda g: f"{g[0]}.{g[1]}.{current_year}"),
    ]

    # 🔍 Tìm cụm ngày đầu tiên
    date_candidates = re.findall(r"\d{1,2}(?:[./-]\d{1,2}){1,2}(?:[./-]\d{2,4})?", eta)
    if not date_candidates:
        print(f"⚠️ Không tìm thấy ngày trong: {eta_original}")
        return None

    first = date_candidates[0]

    # 🔎 Thử match từng pattern
    for pattern, builder in patterns:
        m = re.match(pattern, first)
        if m:
            parts = builder(m.groups())
            try:
                return datetime.strptime(parts, "%d.%m.%Y")
            except ValueError:
                continue

    # Nếu vẫn chưa parse được → thử match đơn giản dd.mm.yyyy
    try:
        return datetime.strptime(first, "%d.%m.%Y")
    except Exception:
        print(f"⚠️ Không parse được ETA: {eta_original} → '{first}'")
        return None


def normalize_ship_name(name: str) -> str:
    """Chuẩn hóa tên tàu: bỏ ngoặc, ký tự phụ, đồng nhất format."""
    if pd.isna(name):
        return ""
    name = str(name).strip().upper()
    name = re.sub(r"\(.*?\)", "", name)  # bỏ phần trong ngoặc
    return name
def process_lichtau(file_path):
    all_data = []

    # 1. Định nghĩa hàm xử lý giá trị (đưa ra ngoài vòng lặp để tối ưu)
    def safe_value(val):
        if val is None or pd.isna(val) or (isinstance(val, float) and math.isnan(val)):
            return None
        if isinstance(val, pd.Timestamp):
            return val.to_pydatetime()
        return val

    with pd.ExcelFile(file_path) as xls:
        # 2. Gọi hàm filter mới (Không truyền tham số ngày nữa)
        sheets = filter_sheets_from_month(xls.sheet_names)

        # 3. [QUAN TRỌNG] Kiểm tra nếu không có sheet nào thỏa mãn -> Return ngay để tránh lỗi concat
        if not sheets:
            print(f"⚠️ Không tìm thấy sheet nào thỏa mãn điều kiện thời gian trong file {file_path}")
            return pd.DataFrame()

        required_cols = [
            "SỐ LỆNH TÁCH", "TÀU/PHƯƠNG TIỆN VẬN TẢI", "KHỐI LƯỢNG TỔNG TÀU",
            "ETA DUNG QUẤT", "ĐẠI LÝ", "ETB DUNG QUẤT", "THỜI GIAN LÀM XONG HÀNG",
            "NGÀY DK DUYỆT SO", "Cảng xếp", "CẢNG ĐẾN",
            "LỆNH XUẤT HÀNG - KẾ HOẠCH DUYỆT (SỐ LỆNH ĐẦY ĐỦ - SỐ XNĐH - KL TỔNG ĐƠN - LSD) (MỖI LỆNH 1 DÒNG)",
            "KHỐI LƯỢNG HÀNG XUẤT LÊN TÀU", "SẢN XUẤT (HRC 1/2-TÌNH TRẠNG)",
            "C.W MAX TÀU NHẬN ĐƯỢC", "GHI CHÚ", "NHỊP", "TÌNH TRẠNG",
            "SO", "TỔNG ĐÃ MAP", "ĐÃ XUẤT", "CÒN LẠI", "SheetMonth"
        ]

        for sheet in sheets:
            try:
                df = pd.read_excel(
                    xls, 
                    sheet_name=sheet, 
                    skiprows=2,
                    dtype={'ETA DUNG QUẤT': str} 
                )

                # Chuẩn hóa tên cột
                df.columns = (
                    df.columns.astype(str)
                    .str.replace(r'[\r\n]+', ' ', regex=True)
                    .str.replace(r'\s*/\s*', '/', regex=True)
                    .str.replace(r'\s+', ' ', regex=True)
                    .str.strip()
                )

                # Thêm cột SheetMonth
                month = sheet.replace("LỊCH TÀU - ", "").strip()
                df["SheetMonth"] = month
                
                # Xử lý SỐ LỆNH TÁCH
                df['SỐ LỆNH TÁCH'] = pd.to_numeric(df['SỐ LỆNH TÁCH'], errors='coerce')
                df.dropna(subset=['SỐ LỆNH TÁCH'], inplace=True)
                df['SỐ LỆNH TÁCH'] = df['SỐ LỆNH TÁCH'].astype('Int64')

                # Reindex cột
                df = df.reindex(columns=required_cols)

                # Xử lý Tên Tàu
                if 'TÀU/PHƯƠNG TIỆN VẬN TẢI' in df.columns:
                    # Lưu ý: Đảm bảo bạn đã import/định nghĩa hàm normalize_ship_name ở ngoài
                    df['TÀU/PHƯƠNG TIỆN VẬN TẢI'] = df['TÀU/PHƯƠNG TIỆN VẬN TẢI'].apply(normalize_ship_name)
                else:
                    df['TÀU/PHƯƠNG TIỆN VẬN TẢI'] = ""

                # Fill dữ liệu group theo tàu
                cols_fill = ['KHỐI LƯỢNG TỔNG TÀU', 'ETA DUNG QUẤT']
                if 'TÀU/PHƯƠNG TIỆN VẬN TẢI' in df.columns:
                    mask_has_tau = df['TÀU/PHƯƠNG TIỆN VẬN TẢI'].notna() & (df['TÀU/PHƯƠNG TIỆN VẬN TẢI'] != "")
                    if not df[mask_has_tau].empty:
                        try:
                            df.loc[mask_has_tau, cols_fill] = (
                                df[mask_has_tau]
                                .groupby(['TÀU/PHƯƠNG TIỆN VẬN TẢI'], group_keys=False)[cols_fill]
                                .transform(lambda x: x.ffill().bfill())
                            )
                        except ValueError:
                            pass

                # Convert số float
                float_cols = ["KHỐI LƯỢNG TỔNG TÀU","KHỐI LƯỢNG HÀNG XUẤT LÊN TÀU",
                              "TỔNG ĐÃ MAP","ĐÃ XUẤT","CÒN LẠI","C.W MAX TÀU NHẬN ĐƯỢC","SO"]
                for col in float_cols:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce')

                # Apply safe_value
                for col in df.columns:
                    df[col] = df[col].apply(safe_value)

                # Parse ETA
                if "ETA DUNG QUẤT" in df.columns:
                    # Lưu ý: Đảm bảo đã import/định nghĩa hàm parse_eta ở ngoài
                    df["ETA_Parsed"] = df["ETA DUNG QUẤT"].apply(parse_eta)
                else:
                    df["ETA_Parsed"] = None

                # Xóa hàng rỗng
                df = df.dropna(how='all')
                
                if not df.empty:
                    all_data.append(df)
            
            except Exception as e:
                print(f"❌ Lỗi khi đọc sheet {sheet}: {e}")
                continue

    # 4. Check lần cuối trước khi concat
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        return final_df
    else:
        return pd.DataFrame()
def process_so_details():
    file_path = "HRC2 - FILE THEO DÕI ĐƠN HÀNG.xlsx"
    df = pd.read_excel("HRC1 - FILE THEO DÕI ĐƠN HÀNG (04.04)_28.xlsx", sheet_name=1)

    # Sửa lại dòng này: thêm một cặp dấu ngoặc vuông [[...]]
    df1 = df[["SO Mapping", "CW", "NHÓM", "Material description"]]

    # In ra 5 dòng đầu tiên của DataFrame mới
    print(df1.head())

def _rename_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Tự động tìm và đổi tên các cột quan trọng về một tên chuẩn hóa.
    """
    rename_map = {
        'SO Mapping': ['so mapping', 'so_mapping', 'số lệnh tách', 'số lệnh tách'],
        'Material Description': ['material description', 'material_description', 'item description'],
        'Material description': ['material description', 'material_description', 'item description']
    }

    current_columns = {c.lower().strip(): c for c in df.columns}

    for standard_name, variations in rename_map.items():
        for var in variations:
            if var in current_columns and standard_name not in df.columns:
                df = df.rename(columns={current_columns[var]: standard_name})
                break # Đã đổi tên, chuyển sang tên chuẩn tiếp theo
    return df

def _normalize_cw(value):
    """
    Chuẩn hóa giá trị cột CW.
    - '18-24' -> '18-24'
    - 'max25', '<25', '25' -> '0-25'
    - Các giá trị khác -> ''
    """
    if pd.isna(value):
        return ""

    s_value = str(value).strip().lower()

    # 1. Ưu tiên tìm kiếm định dạng min-max ở bất kỳ đâu trong chuỗi (ví dụ: '19-23mt')
    # re.search sẽ tìm kiếm thay vì khớp từ đầu đến cuối như re.match
    range_match = re.search(r'(\d+)\s*-\s*(\d+)', s_value)
    if range_match:
        num1 = int(range_match.group(1))
        num2 = int(range_match.group(2))
        return f"{min(num1, num2)}-{max(num1, num2)}"

    # 2. Nếu không có định dạng trên, mới tìm số đơn lẻ đầu tiên trong chuỗi (ví dụ: 'max25', '<=25')
    numbers = re.findall(r'\d+', s_value)
    if numbers:
        num = int(numbers[0])
        return f"0-{num}"

    # 3. Nếu không tìm thấy bất kỳ số nào, trả về chuỗi rỗng
    return ""

def process_so_details(file_paths: list[str]):
    """
    Đọc và xử lý các file chi tiết SO,
    sau đó ghi đè vào bảng so_request trong DB.
    """
    from sqlalchemy.types import NVARCHAR, BigInteger

    all_dfs = []

    for file_path in file_paths: # Lặp qua tất cả các file được cung cấp
        try:
            # --- LOGIC MỚI: Sử dụng 'with' để đảm bảo file được đóng lại ---
            with pd.ExcelFile(file_path) as xls:
                sheet_names = xls.sheet_names
                
                target_sheet = None
                # 1. Ưu tiên tìm tên sheet chính xác (không phân biệt hoa thường)
                for name in sheet_names:
                    if name.strip().upper() in ["ĐƠN HÀNG", "ĐƠN HÀNG HRC"]:
                        target_sheet = name
                        break
                
                # 2. Nếu không tìm thấy tên, thử dùng index 1 (sheet thứ hai) làm dự phòng
                if target_sheet is None and len(sheet_names) > 1:
                    target_sheet = 1 # Dùng index
    
                if target_sheet is not None:
                    # Đọc dữ liệu từ đối tượng 'xls' đã mở
                    df = pd.read_excel(xls, sheet_name=target_sheet)
                    all_dfs.append(df)
                else:
                    # Nếu không có sheet nào phù hợp, bỏ qua file này
                    print(f"Cảnh báo: Không tìm thấy sheet 'ĐƠN HÀNG' hoặc sheet thứ 2 trong file '{file_path}'. Bỏ qua file.")
                    continue
            # --- KẾT THÚC LOGIC MỚI: File sẽ tự động được đóng khi thoát khỏi khối 'with' ---
        except Exception as e:
            print(f"Cảnh báo: Bỏ qua file '{file_path}' do lỗi: {e}")

    if not all_dfs:
        print("Không có file chi tiết SO nào được cung cấp.")
        return

    # --- Kết hợp dữ liệu và chọn cột ---
    df_combined = pd.concat(all_dfs, ignore_index=True)
    

    required_cols = ["SO Mapping", "CW", "NHÓM", "Material description"]
    # Lọc ra các cột tồn tại trong DataFrame
    existing_cols = [col for col in required_cols if col in df_combined.columns]
    df_final = df_combined[existing_cols]

    # Kiểm tra cột bắt buộc 'SO Mapping'
    if 'SO Mapping' not in df_final.columns:
        raise ValueError("Không tìm thấy cột 'SO Mapping' hoặc các biến thể của nó trong file Excel. Vui lòng kiểm tra lại tên cột.")

    # --- Chuẩn hóa kiểu dữ liệu ---
    df_final['SO Mapping'] = pd.to_numeric(df_final['SO Mapping'], errors='coerce').fillna(0).astype('Int64')
    df_final = df_final.dropna(subset=['SO Mapping']) # Bỏ các dòng không có SO Mapping

    # Áp dụng chuẩn hóa cho cột CW nếu tồn tại
    if 'CW' in df_final.columns:
        df_final['CW'] = df_final['CW'].apply(_normalize_cw)

    # Chuẩn hóa cột NHÓM: thay thế '/' bằng ','
    if 'NHÓM' in df_final.columns:
        # 1. Đảm bảo cột là kiểu string
        # 2. Loại bỏ phần trong ngoặc đơn và các khoảng trắng xung quanh nó
        # 3. Thay thế '/' bằng ','
        df_final['NHÓM'] = df_final['NHÓM'].astype(str).str.replace(r'\s*\(.*\)\s*', '', regex=True).str.replace('/', ',', regex=False)

    for col in ["CW", "NHÓM", "Material description"]:
        if col in df_final.columns:
            df_final[col] = df_final[col].astype(str).fillna('')

    # --- Ghi vào DB ---
    dtype_mapping = {
        'SO Mapping': BigInteger(),
        'CW': NVARCHAR(),
        'NHÓM': NVARCHAR(),
        'Material description': NVARCHAR()
    }
    df_final.to_sql('so_request', engine, if_exists='replace', index=False, dtype=dtype_mapping)
    print(f"Đã ghi thành công {len(df_final)} dòng vào bảng so_request.")
import numpy as np
def process_create_lsx(input_file_path):

    
    # --- BẮT ĐẦU LOGGING ---
    print("\n" + "="*50)
    print("--- [BẮT ĐẦU] Xử lý file import Đơn Hàng ---")
    print(f"Đường dẫn file: {input_file_path}")
    # 1. Định nghĩa các tên cột
    COL_KHSX = 'KHSX'
    COL_DO_DAY = 'Độ dày'              # Dùng cho cả sắp xếp (số) và hiển thị (chuỗi)
    COL_WMDD_STR = 'W\nMDĐ'          # Cột CHUỖI (vd: "123X") - Dùng để HIỂN THỊ
    COL_KHO_RONG_NUM = 'Khổ rộng'      # Cột SỐ (vd: 1230) - Dùng để SẮP XẾP
    COL_MAC_THEP = 'Mác thép'
    COL_1A = '1A'
    COL_1B = '1B\nI' 
    COL_NOTE_DAC_BIET = 'NOTE MÁC ĐẶC BIỆT\nYÊU CẦU KHÁC'
    COL_ORDER = 'Order HRC'
    COL_CW = 'CW'
    COL_MUC_DICH = 'Mục đích sử dụng'
    COL_KHACH_HANG = 'Tên KH'
    COL_DOT_SX = 'Đợt sx'
    
    # --- Bọc toàn bộ hàm trong try...except để bắt lỗi chi tiết ---
    try: 
        # 2. Đọc file input (giữ nguyên)
        print("Bước 1: Đang đọc sheet 'ĐƠN HÀNG' từ file Excel...")
        try:
            df_input = pd.read_excel(
                input_file_path, 
                sheet_name="ĐƠN HÀNG", 
                header=0,
                dtype=str 
            )
        except ValueError as e:
            if "Worksheet named 'ĐƠN HÀNG' not found" in str(e):
                raise ValueError("Lỗi: Không tìm thấy sheet có tên 'ĐƠN HÀNG' trong file Excel.")
            else:
                raise e
        print(f"✅ Đọc file thành công. Tìm thấy {len(df_input)} dòng thô.")

        # 3. Xử lý dữ liệu (Clean/Chuẩn hóa) (giữ nguyên)
        required_cols_check = [COL_DOT_SX, COL_ORDER, COL_WMDD_STR, COL_DO_DAY, COL_1A, COL_1B, COL_CW]
        for col in required_cols_check:
            if col not in df_input.columns:
                raise ValueError(f"Lỗi: Không tìm thấy cột '{col}' trong file Excel. Vui lòng kiểm tra lại tên cột.")

        print("Bước 2: Đang làm sạch và chuẩn hóa dữ liệu...")
        df_input = df_input.dropna(subset=[COL_DOT_SX])
        print(f" -> Sau khi bỏ dòng thiếu '{COL_DOT_SX}', còn lại: {len(df_input)} dòng.")

        mask_original_not_null = df_input[COL_ORDER].notna()
        mask_converted_is_null = pd.to_numeric(df_input[COL_ORDER], errors='coerce').isna()
        mask_is_bad_text = mask_original_not_null & mask_converted_is_null
        df_input = df_input[~mask_is_bad_text]
        print(f" -> Sau khi bỏ dòng có '{COL_ORDER}' là chữ, còn lại: {len(df_input)} dòng.")

        df_input['__sort_kho_rong'] = pd.to_numeric(
            df_input[COL_WMDD_STR].str.extract(r'(\d+)', expand=False), 
            errors='coerce'
        ).fillna(0)
        df_input['__sort_do_day'] = pd.to_numeric(df_input[COL_DO_DAY], errors='coerce').fillna(0)

        df_input[COL_WMDD_STR] = df_input[COL_WMDD_STR].fillna('').astype(str)
        df_input[COL_DO_DAY] = df_input[COL_DO_DAY].fillna('').astype(str)
        
        for col in [COL_KHSX, COL_MAC_THEP, COL_NOTE_DAC_BIET, COL_CW, COL_MUC_DICH, COL_KHACH_HANG]:
            if col in df_input.columns:
                df_input[col] = df_input[col].fillna('')
                
        for col in [COL_1A, COL_1B]:
            if col in df_input.columns:
                df_input[col] = pd.to_numeric(df_input[col], errors='coerce').fillna(0)

        # 4. Sắp xếp theo yêu cầu (giữ nguyên)
        print("Bước 3: Đang sắp xếp dữ liệu...")
        df_sorted = df_input.sort_values(
            by=['__sort_kho_rong', '__sort_do_day'],
            ascending=[False, False]
        )
        df_sorted = df_sorted.reset_index(drop=True)
        print("✅ Sắp xếp hoàn tất.")

        # ================================================================
        # --- [BẮT ĐẦU THAY ĐỔI] Bước 5: Tạo DataFrame kết quả (df_output) ---
        # ================================================================
        print("Bước 4: Đang tạo DataFrame kết quả và tính toán các cột...")
        df_output = pd.DataFrame()

        # === 5.1 Mapping Dữ Liệu (Phần 1: Dữ liệu thô) ===
        df_output['STT'] = np.arange(1, len(df_sorted) + 1)
        df_output['ThoiGianSX'] = df_sorted[COL_KHSX]
        df_output['KichCo'] = df_sorted[COL_DO_DAY].astype(str) + 'x' + df_sorted[COL_WMDD_STR].astype(str)
        df_output['MacThep'] = df_sorted[COL_MAC_THEP]
        df_output['SanLuong_1A'] = df_sorted[COL_1A]
        df_output['SanLuong_1B'] = df_sorted[COL_1B]
        df_output['YeuCauDacBiet'] = df_sorted[COL_NOTE_DAC_BIET]
        df_output['OrderNumber'] = pd.to_numeric(df_sorted[COL_ORDER], errors='coerce').fillna(0).astype('int64')
        df_output['KL_Cuon'] = df_sorted[COL_CW]
        df_output['MucDichSuDung'] = df_sorted[COL_MUC_DICH]
        df_output['KhachHang'] = df_sorted[COL_KHACH_HANG]
        df_output['DotSX'] = df_sorted[COL_DOT_SX]
        df_output['ID'] = None 
        df_output['CoTinh_GHC'] = np.nan
        df_output['CoTinh_GHB'] = np.nan
        df_output['CoTinh_GianDai'] = np.nan
        df_output['CoTinh_DoCung'] = np.nan
        df_output['Phoi_MacPhoi'] = np.nan
        df_output['Phoi_KichThuoc'] = np.nan
        df_output['Batch'] = np.nan

        # === 5.2 Mapping Dữ Liệu (Phần 2: Tính toán và Gắn cờ) ===
        
        # 1. Xử lý KL_Cuon (CW) - Logic này chỉ chấp nhận "num1-num2" hoặc "num"
        cw_str = df_sorted[COL_CW].astype(str).str.strip()
        
        # Trích xuất dải (vd: "18-24") -> group1=18, group2=24
        range_matches = cw_str.str.extract(r'^\s*(\d+\.?\d*)\s*-\s*(\d+\.?\d*)\s*$') # regex fullmatch
        
        # Trích xuất số đơn (vd: "18") -> group1=18
        single_matches = cw_str.str.extract(r'^\s*(\d+\.?\d*)\s*$') # regex fullmatch
        
        # 2. Tính toán giá trị max của CW
        cw_min_range = pd.to_numeric(range_matches[0], errors='coerce')
        cw_max_range = pd.to_numeric(range_matches[1], errors='coerce')
        cw_max_from_range = np.maximum(cw_min_range, cw_max_range)
        cw_max_from_single = pd.to_numeric(single_matches[0], errors='coerce')
        
        cw_max = cw_max_from_range.fillna(cw_max_from_single)
        avg_kl_cuon = cw_max - 0.5
        avg_kl_cuon = avg_kl_cuon.replace(0, np.nan) # Tránh chia cho 0

        # 3. Gắn cờ lỗi (Tên cột: `HasWarning`)
        # Lỗi = (Chuỗi CW không rỗng) VÀ (Không thể parse ra avg_kl_cuon)
        is_not_empty = cw_str.str.len() > 0
        is_parse_error = avg_kl_cuon.isna()
        df_output['HasWarning'] = (is_not_empty & is_parse_error) # Cột này là True/False

        # 4. Tính toán (An toàn với NaN)
        tong_san_luong = df_sorted[COL_1A] + df_sorted[COL_1B]
        san_luong_yeucau_float = (tong_san_luong / avg_kl_cuon).round(2)
        
        # 5. 'SanLuong_YeuCau_Cuon' (Dòng lỗi sẽ là 0)
        df_output['SanLuong_YeuCau_Cuon'] = san_luong_yeucau_float.round(0).fillna(0).astype(int)

        # 6. 'DungSai' (Dòng lỗi sẽ là "± 0")
        base_dung_sai_float = 0.1 * san_luong_yeucau_float
        adjusted_dung_sai_float = np.where(
            tong_san_luong > 2000,
            base_dung_sai_float / 2,
            base_dung_sai_float
        )
        dung_sai_int = pd.Series(adjusted_dung_sai_float).round(0).fillna(0).astype(int)
        df_output['DungSai'] = "± " + dung_sai_int.astype(str)

        # ================================================================
        # --- [KẾT THÚC THAY ĐỔI] Bước 5 ---
        # ================================================================

        print(f"✅ Xử lý hoàn tất. Trả về {len(df_output)} dòng dữ liệu sạch.")
        print("="*50 + "\n")
        return df_output
        
    except Exception as e:
        print(f"❌ LỖI BẤT NGỜ trong quá trình xử lý dữ liệu: {e}")
        import traceback
        traceback.print_exc()
        raise e