import pandas as pd
from datetime import datetime
from db import engine  # import engine SQLAlchemy đã cấu hình
# ---------- Ghi DataFrame vào SQL Server ----------
from sqlalchemy import inspect,types,text
import logging
import json
logger = logging.getLogger(__name__)
##UTILS LƯU DATAFRAME VỀ DATABASE
def save_df_to_db(df: pd.DataFrame, table_name: str, engine, batch_size=500, if_exists="append"):
    """
    Ghi DataFrame vào SQL Server an toàn, chia batch để tránh lỗi
    Có thêm debug chi tiết để phát hiện lỗi khi to_sql bị fail.
    """
    try:
        import sqlalchemy.types as types
        from sqlalchemy import inspect
        import logging

        logger = logging.getLogger(__name__)

        # === 1️⃣ Chuẩn bị kiểu dữ liệu tương ứng ===
        dtype_mapping = {}
        for col in df.columns:
            if pd.api.types.is_string_dtype(df[col]):
                dtype_mapping[col] = types.NVARCHAR(length=4000)
            elif pd.api.types.is_integer_dtype(df[col]):
                dtype_mapping[col] = types.BigInteger()
            elif pd.api.types.is_float_dtype(df[col]):
                dtype_mapping[col] = types.Float()
            elif pd.api.types.is_datetime64_any_dtype(df[col]):
                dtype_mapping[col] = types.DateTime()
            else:
                dtype_mapping[col] = types.NVARCHAR(length=4000)

        # === 2️⃣ Kiểm tra DataFrame ===
        if df.empty:
            logger.warning(f"[SKIP] No data to insert into {table_name}.")
            return

        logger.info(f"Preparing to insert into {table_name}: {len(df)} rows, {len(df.columns)} columns.")
        logger.debug(f"Columns: {list(df.columns)}")
        logger.debug(f"dtypes:\n{df.dtypes}")

        # === 3️⃣ Xử lý NULL: số → 0, chuỗi → "" ===
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                df[col] = df[col].fillna(0)
            else:
                df[col] = df[col].fillna("")

        # === 4️⃣ Ghi dữ liệu ===
        with engine.begin() as conn:
            insp = inspect(conn)
            if table_name not in insp.get_table_names():
                # Tạo bảng nếu chưa tồn tại
                logger.info(f"Table {table_name} not found — creating new table.")
                df.head(0).to_sql(table_name, conn, if_exists="replace", index=False, dtype=dtype_mapping)
                logger.info(f"Table {table_name} created successfully.")

            total_rows = len(df)
            logger.info(f"Starting insert of {total_rows} rows into {table_name} in batches of {batch_size}...")

            for i in range(0, total_rows, batch_size):
                batch_df = df.iloc[i:i+batch_size]

                try:
                    batch_df.to_sql(
                        table_name,
                        conn,
                        if_exists=if_exists,
                        index=False,
                        dtype=dtype_mapping,
                        method=None
                    )
                    logger.info(f"✅ Inserted rows {i+1}-{i+len(batch_df)} into {table_name}.")
                except Exception as e:
                    logger.error(f"❌ Error inserting batch {i+1}-{i+len(batch_df)}: {e}")
                    logger.error(f"Batch preview:\n{batch_df.head(3)}")
                    raise  # để dừng và thấy lỗi thật

            logger.info(f"✅ Finished inserting {total_rows} rows into {table_name}.")

    except Exception as e:
        logger.exception(f"🔥 save_df_to_db() failed for table {table_name}: {e}")
        print(f"⚠️ Lỗi khi ghi dữ liệu vào {table_name}: {e}")
        print(f"➡️ DataFrame shape: {df.shape}")
        print(f"➡️ Columns: {list(df.columns)}")
        print(df.head(3))
def save_lichtau(df: pd.DataFrame, table_name: str, engine):
    """
    Ghi toàn bộ DataFrame vào SQL Server (ghi đè hoàn toàn bảng).
    - Giữ nguyên thứ tự như trong Excel.
    - Ép kiểu tự động, hỗ trợ tiếng Việt (NVARCHAR).
    - Log chi tiết tiến trình và lỗi nếu có.
    """
    try:
        if df.empty:
            logger.warning(f"[⚠️] DataFrame rỗng, bỏ qua ghi vào bảng {table_name}.")
            return

        # Mapping kiểu dữ liệu SQL tương ứng với pandas dtype
        dtype_mapping = {}
        for col in df.columns:
            if pd.api.types.is_string_dtype(df[col]):
                dtype_mapping[col] = types.NVARCHAR(length=4000)
            elif pd.api.types.is_integer_dtype(df[col]):
                dtype_mapping[col] = types.BigInteger()
            elif pd.api.types.is_float_dtype(df[col]):
                dtype_mapping[col] = types.Float()
            elif pd.api.types.is_datetime64_any_dtype(df[col]):
                dtype_mapping[col] = types.DateTime()
            else:
                dtype_mapping[col] = types.NVARCHAR(length=4000)

        # Fill NaN để tránh lỗi khi ghi
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                df[col] = df[col].fillna(0)
            else:
                df[col] = df[col].fillna("")

        logger.info(f"[ℹ️] Ghi {len(df)} dòng, {len(df.columns)} cột vào bảng {table_name}...")

        # Ghi đè toàn bộ bảng (theo đúng thứ tự Excel)
        df.to_sql(
            table_name,
            engine,
            if_exists="replace",  # ghi đè toàn bộ
            index=False,
            dtype=dtype_mapping
        )

        logger.info(f"[✅] Đã ghi thành công {len(df)} dòng vào bảng {table_name}.")

    except Exception as e:
        logger.exception(f"[❌] Lỗi khi ghi DataFrame vào bảng {table_name}: {e}")
        raise
# ---------- Load dữ liệu từ DB ----------
def load_table_from_db(engine, table_name: str, lsx_id: str = None) -> pd.DataFrame:
    with engine.connect() as conn:
        if lsx_id:
            # Câu lệnh chuẩn, dùng parameter để tránh SQL Injection
            query = f"SELECT * FROM [{table_name}] WHERE lsx_id = ?"
            df = pd.read_sql(query, conn, params=(lsx_id,))
        else:
            query = f"SELECT * FROM [{table_name}]"
            df = pd.read_sql(query, conn)
    return df
# CAU HINH LAI TIME

def normalize_datetime(df: pd.DataFrame) -> pd.DataFrame:
    """Convert datetime columns to string YYYY-MM-DD HH:MM:SS"""
    df = df.copy()
    for c in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            df[c] = df[c].dt.strftime("%Y-%m-%d %H:%M:%S")
        df[c] = df[c].where(pd.notnull(df[c]), None)
    return df
# UPSERT FILE SẢN LƯỢNG TỰ ĐỘNG
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)
def sync_sap_to_production(engine):
    """Gọi Stored Procedure để đồng bộ ID từ sanluong sang coil_data"""
    try:
        with engine.begin() as conn:
            conn.execute(text("EXEC [dbo].[sp_Sync_SAP_ID_To_Production]"))
    except Exception as e:
        logger.error(f"❌ Lỗi khi chạy Sync SP: {e}")
def upsert_sanluong_from_excel(df: pd.DataFrame, table_name: str = "sanluong", nhamay: str = "HRC1", date_col_name: str = "Posting Date"):
    if df.empty:
        logger.warning(f"⚠️ [UPSERT] DataFrame rỗng cho nhà máy {nhamay}. Bỏ qua.")
        return

    logger.info(f"🚀 [START] Bắt đầu Upsert {table_name} - {nhamay}. Số dòng: {len(df)}")

    # 1. CLEAN UP TÊN CỘT (RẤT QUAN TRỌNG: Tránh lỗi 'ID Ref ' vs 'ID Ref')
    original_cols = list(df.columns)
    df.columns = [c.strip() for c in df.columns]
    if list(df.columns) != original_cols:
        logger.info("ℹ️ Đã xóa khoảng trắng thừa trong tên cột Excel.")

    # 2. TÍNH NGÀY NHỎ NHẤT ĐỂ CHẶN XÓA LỊCH SỬ
    min_date_str = "1900-01-01"
    try:
        if date_col_name in df.columns:
            temp_dates = pd.to_datetime(df[date_col_name], dayfirst=True, errors='coerce')
            min_val = temp_dates.min()
            if pd.notna(min_val):
                min_date_str = min_val.strftime("%Y-%m-%d")
                logger.info(f"📅 Dữ liệu từ ngày: {min_date_str}")
        else:
            logger.warning(f"⚠️ Không tìm thấy cột '{date_col_name}'. Upsert sẽ chạy chế độ toàn bộ.")
    except Exception as e:
        logger.error(f"⚠️ Lỗi tính ngày: {e}. Vẫn tiếp tục chạy.")

    # 3. CHUẨN BỊ DỮ LIỆU
    df = df.copy()
    df["NhaMay"] = nhamay
    snap_ts = datetime.now()
    df["snapshot_ts"] = snap_ts
    df["status"] = "active"
    
    # 4. ÉP KIỂU SỐ (PRE-PROCESSING)
    # Danh sách các cột bắt buộc phải là số nguyên (BIGINT)
    bigint_cols = ["Order"] 
    
    for col in bigint_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype("int64")

    # XỬ LÝ CÁC CỘT ID SANG CHUỖI (NVARCHAR)
    text_id_cols = ["ID Cuộn Bó", "ID cuộn bó gốc", "ID Ref"]
    
    for col in text_id_cols:
        if col in df.columns:
            # 1. Ép chuỗi, cắt khoảng trắng và đưa về in hoa
            df[col] = df[col].astype(str).str.strip()
            
            # 2. [MỚI CẦN THÊM] Xóa đuôi .0 ở cuối chuỗi do Pandas tự sinh ra
            df[col] = df[col].str.replace(r'\.0$', '', regex=True)
            
            # 3. Dùng Regex để biến mọi dạng 0, 00, 0.0 thành rỗng
            df[col] = df[col].replace(r'^0+(\.0+)?$', '', regex=True)
            df.loc[df[col].isin(['NAN', 'NONE', '']), col] = ""
            
    # Lọc bỏ các dòng không có ID Cuộn Bó (ID trống)
    df = df[df["ID Cuộn Bó"] != ""]

    staging_name = f"staging_sanluong_{nhamay}"

    # 5. DTYPE MAPPING (Cấu hình cho to_sql)
    dtype_mapping = {}
    for col in df.columns:
        if col == "snapshot_ts":                        # <--- THÊM MỚI
            dtype_mapping[col] = types.DateTime()
        elif col in bigint_cols: 
            dtype_mapping[col] = types.BigInteger() 
        elif pd.api.types.is_string_dtype(df[col]):
            dtype_mapping[col] = types.NVARCHAR(length=4000)
        elif pd.api.types.is_integer_dtype(df[col]):
            dtype_mapping[col] = types.BigInteger() 
        elif pd.api.types.is_float_dtype(df[col]):
            dtype_mapping[col] = types.Float()
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            dtype_mapping[col] = types.DateTime()
        else:
            dtype_mapping[col] = types.NVARCHAR(length=4000)

    # 6. THỰC THI SQL
    try:
        with engine.begin() as conn:
            # A. Tạo Staging
            conn.execute(text(f"DROP TABLE IF EXISTS [{staging_name}]"))
            df.to_sql(staging_name, conn, if_exists="replace", index=False, dtype=dtype_mapping)
            logger.info(f"✅ Đã tạo bảng staging: {staging_name}")

            # --- [DEBUG LOG]: Kiểm tra xem SQL đã nhận đúng kiểu BigInt chưa ---
            check_query = text(f"""
                SELECT COLUMN_NAME, DATA_TYPE 
                FROM INFORMATION_SCHEMA.COLUMNS 
                WHERE TABLE_NAME = '{staging_name}' AND COLUMN_NAME IN ('ID Cuộn Bó', 'ID Ref')
            """)
            schema_info = conn.execute(check_query).fetchall()
            logger.info(f"🔍 [CHECK SCHEMA] Cấu trúc bảng Staging: {schema_info}")
            # ------------------------------------------------------------------

            # B. Đánh dấu Removed (Logic cũ giữ nguyên)
            conn.execute(text(f"""
                UPDATE [{table_name}]
                SET status='removed', snapshot_ts=:snap
                WHERE status IN ('active','updated') 
                  AND NhaMay=:nhamay
                  AND [{date_col_name}] >= :min_date  
                  AND NOT EXISTS (
                      SELECT 1 FROM [{staging_name}] s
                      WHERE s.[ID Cuộn Bó] = [{table_name}].[ID Cuộn Bó]
                        AND s.NhaMay = :nhamay
                  )
            """), {"snap": snap_ts, "nhamay": nhamay, "min_date": min_date_str})

            # C. Lưu Removed & Xóa khỏi bảng chính
            conn.execute(text(f"""
                INSERT INTO [{table_name}_removed]
                SELECT * FROM [{table_name}] 
                WHERE status='removed' AND NhaMay=:nhamay
            """), {"nhamay": nhamay})

            conn.execute(text(f"""
                DELETE FROM [{table_name}] 
                WHERE status='removed' AND NhaMay=:nhamay
            """), {"nhamay": nhamay})

            # D. Update dữ liệu thay đổi
            cols_to_update = [c for c in df.columns if c not in ["ID Cuộn Bó", "NhaMay", "status", "snapshot_ts"]]
            if cols_to_update:
                set_clause = ", ".join([f"t.[{c}] = s.[{c}]" for c in cols_to_update])
                diff_condition = " OR ".join([f"ISNULL(t.[{c}], '') <> ISNULL(s.[{c}], '')" for c in cols_to_update])
                set_clause += ", t.status='updated', t.snapshot_ts=:snap"

                conn.execute(text(f"""
                    UPDATE t
                    SET {set_clause}
                    FROM [{table_name}] t
                    INNER JOIN [{staging_name}] s
                      ON t.[ID Cuộn Bó] = s.[ID Cuộn Bó] 
                     AND t.NhaMay = s.NhaMay
                    WHERE t.status IN ('active','updated') 
                      AND ({diff_condition})
                """), {"snap": snap_ts})

            # E. INSERT MỚI (SỬA ĐỂ TRÁNH LỖI 8114 & LỆCH CỘT)
            # Thay vì SELECT *, ta liệt kê đích danh các cột
            
            # Lấy danh sách cột có trong DataFrame (chính là cấu trúc của Staging)
            # Lưu ý: Cần bọc tên cột trong ngoặc vuông []
            common_cols = [f"[{c}]" for c in df.columns]
            cols_string = ", ".join(common_cols)
            
            insert_query = f"""
                INSERT INTO [{table_name}] ({cols_string})
                SELECT {cols_string} FROM [{staging_name}] s
                WHERE NOT EXISTS (
                    SELECT 1 FROM [{table_name}] t
                    WHERE t.[ID Cuộn Bó] = s.[ID Cuộn Bó] 
                    AND t.NhaMay = s.NhaMay
                )
            """
            conn.execute(text(insert_query))
            logger.info("✅ Insert dữ liệu mới thành công (Safe Mode).")

            # F. Dọn dẹp
            conn.execute(text(f"DROP TABLE IF EXISTS [{staging_name}]"))

        # G. Đồng bộ sang Production (Coil Data)
        sync_sap_to_production(engine)
        logger.info(f"🏁 [END] Hoàn tất quy trình Upsert cho {nhamay}.")

    except Exception as e:
        logger.exception(f"❌ [ERROR] Lỗi nghiêm trọng khi Upsert {nhamay}: {e}")
        # Không raise để chương trình không crash hoàn toàn, chỉ log lỗi
from sqlalchemy import text, types
from db import engine


# ------------------- Upsert KHO -------------------
def upsert_kho_from_excel(df: pd.DataFrame, table_name: str = "kho"):
    """
    Upsert dữ liệu KHO an toàn (chuẩn hóa BIGINT + INT)
    - ID Cuộn Bó: BIGINT (chính xác, không lỗi float)
    - Plant: INT
    - Tách staging riêng theo từng Plant tránh conflict song song
    """
    if df.empty:
        print("⚠️ Dữ liệu trống, bỏ qua upsert.")
        return

    # ==== 1️⃣ Chuẩn hóa schema ====
    KHO_SCHEMA = [
        "Plant","Material","Storage Location","Material Description",
        "ID Cuộn Bó","Vị trí","Khối lượng","Nhóm","Ca","Ngày sản xuất",
        "SO Mapping","Batch","Order"
        ,"Lô Phôi","Trạm cân",
        "Số lượng in","Nhập tay","Tp loại 2",
        "snapshot_ts","status","Mác thép","Customer N",
    ]

    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    for col in KHO_SCHEMA:
        if col not in df.columns:
            df[col] = None
    df = df[KHO_SCHEMA]

    # ==== 2️⃣ Ép kiểu dữ liệu ====
    # Các cột số nguyên (ID, Plant)
    int_cols = ["Plant"]
    bigint_cols = ["Material", "SO Mapping"]

    # Các cột float
    float_cols = [
        "Khối lượng","SO Item Ma","Batch","Order","Trạm cân","Số lượng in","Storage Location"
    ]
    if "ID Cuộn Bó" in df.columns:
        df["ID Cuộn Bó"] = df["ID Cuộn Bó"].astype(str).str.strip().str.upper()
        
        # [MỚI CẦN THÊM] Xóa đuôi .0 ở cuối
        df["ID Cuộn Bó"] = df["ID Cuộn Bó"].str.replace(r'\.0$', '', regex=True)
        
        # Lọc bỏ các hàng trống hoặc 'NAN'
        df = df[(df["ID Cuộn Bó"] != "") & (~df["ID Cuộn Bó"].isin(['NAN', 'NONE', '0']))]
    # Convert từng nhóm
    for col in int_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype("Int64")

    for col in bigint_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype("Int64")

    for col in float_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Các cột text
    nvarchar_cols = [c for c in KHO_SCHEMA if c not in int_cols + bigint_cols + float_cols + ["snapshot_ts","status"]]
    for col in nvarchar_cols:
        df[col] = df[col].astype(str).fillna("")
    snap_ts = datetime.now()
    df["snapshot_ts"] = snap_ts
    df["status"] = "active"

    # ==== 3️⃣ Mapping dtype cho SQL ====
    dtype_mapping = {}
    for c in KHO_SCHEMA:
        if c == "snapshot_ts":                          
            dtype_mapping[c] = types.DateTime()         
        elif c in bigint_cols:
            dtype_mapping[c] = types.BIGINT()
        elif c in int_cols:
            dtype_mapping[c] = types.INTEGER()
        elif c in float_cols:
            dtype_mapping[c] = types.Float()
        else:
            dtype_mapping[c] = types.NVARCHAR(length=4000)

    # ==== 4️⃣ Xử lý từng Plant riêng biệt ====
    plants = df["Plant"].dropna().unique()

    for plant in plants:
        plant_int = int(plant)
        staging_name = f"staging_kho_{plant_int}"

        df_plant = df[df["Plant"] == plant_int]

        with engine.begin() as conn:
            # ⚙️ 4.1. Ghi staging riêng cho từng plant
            df_plant.to_sql(staging_name, conn, if_exists="replace", index=False, dtype=dtype_mapping)

            # ⚙️ 4.2. Đánh dấu removed cho cuộn không còn trong staging
            conn.execute(text(f"""
                UPDATE t
                SET t.status='removed', t.snapshot_ts=:snap
                FROM [{table_name}] t
                WHERE t.status IN ('active','updated') AND t.Plant=:plant AND t.Plant=:plant
                AND NOT EXISTS (
                    SELECT 1 FROM [{staging_name}] s
                    WHERE s.[ID Cuộn Bó] = t.[ID Cuộn Bó] AND s.Plant = t.Plant
                )
            """), {"snap": snap_ts, "plant": plant_int})

            # ⚙️ 4.3. Lưu removed sang bảng _removed
            conn.execute(text(f"""
                INSERT INTO [{table_name}_removed]
                SELECT * FROM [{table_name}]
                WHERE status='removed' AND Plant=:plant
            """), {"plant": plant_int})

            # ⚙️ 4.4. Xóa record removed khỏi bảng chính
            conn.execute(text(f"""
                DELETE FROM [{table_name}]
                WHERE status='removed' AND Plant=:plant
            """), {"plant": plant_int})

            # ⚙️ 4.5. Update record đã thay đổi
            cols_to_update = [c for c in KHO_SCHEMA if c not in ["ID Cuộn Bó", "Plant", "status", "snapshot_ts"]]
            set_clause = ", ".join([f"t.[{c}] = s.[{c}]" for c in cols_to_update])
            diff_condition = " OR ".join([f"ISNULL(t.[{c}], '') <> ISNULL(s.[{c}], '')" for c in cols_to_update])

            conn.execute(text(f"""
                UPDATE t
                SET {set_clause}, t.status='updated', t.snapshot_ts=:snap
                FROM [{table_name}] t
                INNER JOIN [{staging_name}] s
                    ON s.[ID Cuộn Bó] = t.[ID Cuộn Bó] AND s.Plant = t.Plant
                WHERE {diff_condition}
            """), {"snap": snap_ts})

            # ⚙️ 4.6. Thêm record mới
            conn.execute(text(f"""
                INSERT INTO [{table_name}]
                SELECT s.* FROM [{staging_name}] s
                WHERE NOT EXISTS (
                    SELECT 1 FROM [{table_name}] t
                    WHERE t.[ID Cuộn Bó] = s.[ID Cuộn Bó] AND t.Plant = s.Plant
                )
            """))

            # ⚙️ 4.7. Dọn staging
            conn.execute(text(f"DROP TABLE IF EXISTS [{staging_name}]"))

        print(f"✅ Upsert kho cho Plant {plant_int} hoàn tất ({len(df_plant)} dòng) lúc {snap_ts}")
# ---------- Upsert SALES ORDER ----------

def upsert_so_from_excel(df: pd.DataFrame, table_name: str,date_col_name: str = "Document Date"):
    df = normalize_datetime(df)
    snap_ts = datetime.now()
    df = df.copy()
    df["snapshot_ts"] = snap_ts
    df["status1"] = "active"
    min_date_str = "1900-01-01"
    try:
        if date_col_name in df.columns:
            # Chuyển đổi cột ngày tháng an toàn, bỏ qua các giá trị lỗi
            temp_dates = pd.to_datetime(df[date_col_name], dayfirst=True, errors='coerce')
            min_val = temp_dates.min()
            
            if pd.notna(min_val):
                min_date_str = min_val.strftime("%Y-%m-%d")
                logger.info(f"📅 Dữ liệu SO bắt đầu từ ngày: {min_date_str}. Các SO cũ hơn mốc này sẽ được bảo toàn.")
            else:
                logger.warning(f"⚠️ Cột '{date_col_name}' toàn giá trị rỗng/lỗi. Nguy cơ ảnh hưởng lịch sử!")
        else:
            logger.warning(f"⚠️ Không tìm thấy cột '{date_col_name}'. Sẽ chạy chế độ quét toàn bộ (nguy cơ xóa lịch sử).")
    except Exception as e:
        logger.error(f"⚠️ Lỗi tính toán ngày SO: {e}. Vẫn tiếp tục chạy...")
    with engine.begin() as conn:
        dtype_mapping = {}
        for col in df.columns:
            if col == "snapshot_ts":                    # <--- THÊM MỚI
                dtype_mapping[col] = types.DateTime()
            elif pd.api.types.is_string_dtype(df[col]):
                dtype_mapping[col] = types.NVARCHAR(length=4000)
            elif pd.api.types.is_integer_dtype(df[col]):
                dtype_mapping[col] = types.BigInteger()
            elif pd.api.types.is_float_dtype(df[col]):
                dtype_mapping[col] = types.Float()
            elif pd.api.types.is_datetime64_any_dtype(df[col]):
                dtype_mapping[col] = types.DateTime()
            else:
                dtype_mapping[col] = types.NVARCHAR(length=4000)

        # 2️⃣ Staging table
        df.to_sql("staging_tmp", conn, if_exists="replace", index=False, dtype=dtype_mapping)

        # 3️⃣ Đánh dấu record bị loại
        conn.execute(text(f"""
            UPDATE t
            SET status1='removed', snapshot_ts=:ts
            FROM [{table_name}] t
            WHERE t.status1 IN ('active','updated')
              AND t.[{date_col_name}] >= :min_date  
              AND NOT EXISTS (
                  SELECT 1 FROM staging_tmp s
                  WHERE s.[Sales Document] = t.[Sales Document]
                    AND s.[Material] = t.[Material]
                    AND s.[Sales Document Item] = t.[Sales Document Item]
              )
        """), {"ts": snap_ts, "min_date": min_date_str})

        # 4️⃣ Chuyển sang _removed
        conn.execute(text(f"""
            INSERT INTO [{table_name}_removed]
            SELECT * FROM [{table_name}] WHERE status1='removed'
        """))

        # 5️⃣ Xóa record removed
        conn.execute(text(f"DELETE FROM [{table_name}] WHERE status1='removed'"))

        # 6️⃣ Cập nhật record trùng khóa và set trạng thái updated
        cols_to_update = [
        c for c in df.columns
        if c not in ["Sales Document", "Material", "Sales Document Item", "status1", "snapshot_ts"]
        ]
        set_clause = ", ".join([f"t.[{c}] = s.[{c}]" for c in cols_to_update])
        set_clause += ", t.status1 = 'updated', t.snapshot_ts = :ts"

        # Điều kiện khác nhau giữa staging_tmp và bảng chính
        diff_condition = " OR ".join([f"ISNULL(t.[{c}], '') <> ISNULL(s.[{c}], '')" for c in cols_to_update])

        conn.execute(text(f"""
            UPDATE t
            SET {set_clause}
            FROM [{table_name}] t
            INNER JOIN staging_tmp s
            ON s.[Sales Document]=t.[Sales Document]
            AND s.[Material]=t.[Material]
            AND s.[Sales Document Item]=t.[Sales Document Item]
            WHERE t.status1 IN ('active','updated')
            AND ({diff_condition})
        """), {"ts": snap_ts})

        # 7️⃣ Thêm mới
        cols = ", ".join([f"[{c}]" for c in df.columns])
        conn.execute(text(f"""
            INSERT INTO [{table_name}] ({cols})
            SELECT {cols} FROM staging_tmp s
            WHERE NOT EXISTS (
                SELECT 1 FROM [{table_name}] t
                WHERE t.[Sales Document]=s.[Sales Document]
                  AND t.[Material]=s.[Material]
                  AND t.[Sales Document Item]=s.[Sales Document Item]
            )
        """))

        # 8️⃣ Dọn staging
        conn.execute(text("DROP TABLE IF EXISTS staging_tmp"))
def log_activity(action: str, user_id: int = None, username: str = None, target_type: str = None, target_id=None, details: str = "", ip_address: str = None):
    """Ghi lại một hành động của người dùng vào bảng audit_log."""
    try:
        with engine.begin() as conn:
            stmt = text("""
                INSERT INTO audit_log (user_id, username, action, target_type, target_id, details, ip_address)
                VALUES (:user_id, :username, :action, :target_type, :target_id, :details, :ip_address)
            """)
            conn.execute(stmt, {
                "user_id": user_id, "username": username, "action": action,
                "target_type": target_type, "target_id": str(target_id),
                "details": details, "ip_address": ip_address
            })
    except Exception as e:
        logger.error(f"Lỗi khi ghi nhật ký hoạt động: {e}")
import pandas as pd
from sqlalchemy import text
import logging
import numpy as np
logger = logging.getLogger(__name__)

def sync_order_production_rules_via_pandas(engine):
    logger.info("🔄 [MIGRATION] Bắt đầu đồng bộ từ so_request sang Production & Sales (Chế độ linh hoạt TDC)...")
    try:
        with engine.connect() as conn:
            # 1. KÉO DỮ LIỆU TỪ so_request
            # 🌟 ĐÃ BỎ ĐIỀU KIỆN ÉP BUỘC PHẢI CÓ TDC_Code
            df_request = pd.read_sql("""
                SELECT [Order], [TDC_Code], [SO Mapping], [CW], [Target_Weight], 
                       [Material description], [KySanXuat], [is_skin_required],
                       [material_code], [thickness], [width], [alloc_thick], 
                       [gradeSteel], [purpose]
                FROM dbo.so_request WITH (NOLOCK) 
                WHERE [Order] IS NOT NULL 
            """, conn)

            # 2. Kéo Master dữ liệu luật TDC
            df_master = pd.read_sql("SELECT id as master_id, tdc_code, customer_name, grade FROM dbo.tdc_master WITH (NOLOCK)", conn)
            df_active_ver = pd.read_sql("SELECT master_id, id as tdc_version_id FROM dbo.tdc_versions WITH (NOLOCK) WHERE status = 'Active'", conn)

        if df_request.empty:
            logger.info("✅ Không có Order hợp lệ nào cần đồng bộ.")
            return

        # 3. THỰC HIỆN SO KHỚP BỘ LUẬT TDC (🌟 ĐỔI THÀNH LEFT JOIN)
        # Dùng LEFT JOIN để giữ lại các Order không có TDC hoặc gõ sai mã TDC
        df_result = pd.merge(df_request, df_master, left_on='TDC_Code', right_on='tdc_code', how='left')
        df_result = pd.merge(df_result, df_active_ver, on='master_id', how='left')

        if df_result.empty: return

        # 🌟 XỬ LÝ DỮ LIỆU BỊ NULL SAU KHI LEFT JOIN
        # Đưa các ID về kiểu Int64 của Pandas (Cho phép chứa giá trị rỗng/NaN mà không bị lỗi số thập phân)
        df_result['master_id'] = pd.to_numeric(df_result['master_id'], errors='coerce').astype('Int64')
        df_result['tdc_version_id'] = pd.to_numeric(df_result['tdc_version_id'], errors='coerce').astype('Int64')
        
        # Nếu không có TDC, mặc định khách hàng là 'CHƯA XÁC ĐỊNH'
        df_result['customer_name'] = df_result['customer_name'].fillna('Chưa xác định')

        # 4. CHUẨN HÓA DỮ LIỆU CHUNG (MTO/MTS và Khối lượng)
        df_result['production_status'] = 'MTO'

        df_result['CW'] = df_result['CW'].fillna('').astype(str).str.strip()

        split_cw = df_result['CW'].str.split('-', n=1, expand=True)
        df_result['req_min_w'] = pd.to_numeric(split_cw[0], errors='coerce').fillna(0) * 1000
        df_result['req_max_w'] = pd.to_numeric(split_cw[1], errors='coerce').fillna(0) * 1000 if split_cw.shape[1] > 1 else 0.0

        # ==============================================================================
        # TRANSACTION: BẢO VỆ DỮ LIỆU ĐA KÊNH
        # ==============================================================================
        with engine.begin() as transaction_conn:

            # ---------------------------------------------------------
            # 🌟 NHÁNH 1: ĐỒNG BỘ sales_orders (HEADER)
            # ---------------------------------------------------------
            df_sales = df_result[['SO Mapping', 'customer_name']].copy()
            df_sales.columns = ['so_number', 'customer_name']
            df_sales['so_number'] = pd.to_numeric(df_sales['so_number'], errors='coerce').astype('Int64')
            df_sales = df_sales.dropna(subset=['so_number']).drop_duplicates(subset=['so_number'])

            if not df_sales.empty:
                df_sales.to_sql("staging_sales", transaction_conn, if_exists="replace", index=False,
                                dtype={'so_number': types.BigInteger(), 'customer_name': types.NVARCHAR(255)})
                transaction_conn.execute(text("""
                    MERGE [dbo].[sales_orders] AS target
                    USING staging_sales AS source
                    ON target.so_number = source.so_number
                    WHEN NOT MATCHED THEN
                        INSERT (so_number, customer_name, order_date, created_at, status)
                        VALUES (source.so_number, source.customer_name, GETDATE(), GETDATE(), 'Pending');
                """))
                transaction_conn.execute(text("DROP TABLE IF EXISTS staging_sales"))

            # ---------------------------------------------------------
            # 🌟 NHÁNH 2: ĐỒNG BỘ so_details (PHÂN BỔ BÁN HÀNG)
            # ---------------------------------------------------------
            df_so_details = df_result.copy()
            df_so_details['so_number'] = pd.to_numeric(df_so_details['SO Mapping'], errors='coerce').astype('Int64')
            df_so_details = df_so_details.dropna(subset=['so_number', 'material_code'])
            df_so_details = df_so_details.drop_duplicates(subset=['so_number', 'material_code'], keep='last')

            # ĐÃ VÁ LỖI: Bổ sung cột 'purpose'
            df_to_so_details = df_so_details[[
                'so_number', 'material_code', 'Material description', 'gradeSteel',
                'thickness', 'alloc_thick', 'width', 'Target_Weight', 'master_id',
                'req_min_w', 'req_max_w', 'purpose'
            ]].copy()
            
            df_to_so_details.columns = [
                'so_number', 'material_code', 'description', 'grade',
                'thickness', 'alloc_thick', 'width', 'total_weight', 'tdc_id',
                'min_weight', 'max_weight', 'usage_purpose'
            ]

            if not df_to_so_details.empty:
                df_to_so_details.to_sql("staging_so_details", transaction_conn, if_exists="replace", index=False,
                                     dtype={
                                         'so_number': types.BigInteger(), 'material_code': types.NVARCHAR(100),
                                         'tdc_id': types.BigInteger(),
                                         'description': types.NVARCHAR(1000),
                                         'usage_purpose': types.NVARCHAR(1000),
                                         'grade': types.VARCHAR(100)
                                     })
                transaction_conn.execute(text("""
                    MERGE [dbo].[so_details] AS target
                    USING staging_so_details AS source
                    ON target.so_number = source.so_number AND target.material_code = source.material_code
                    WHEN MATCHED THEN
                        UPDATE SET
                            target.description = source.description,
                            target.grade = source.grade,
                            target.thickness = source.thickness,
                            target.alloc_thick = source.alloc_thick,
                            target.width = source.width,
                            target.total_weight = source.total_weight,
                            target.min_weight = source.min_weight,
                            target.max_weight = source.max_weight,
                            target.tdc_id = ISNULL(source.tdc_id, target.tdc_id),
                            
                            target.usage_purpose = source.usage_purpose
                    WHEN NOT MATCHED THEN
                        INSERT (so_number, material_code, description, grade, thickness, alloc_thick, width, total_weight, min_weight, max_weight, tdc_id, usage_purpose, status)
                        VALUES (source.so_number, source.material_code, source.description, source.grade, source.thickness, source.alloc_thick, source.width, source.total_weight, source.min_weight, source.max_weight, source.tdc_id, source.usage_purpose, 'Hidden');
                """))
                transaction_conn.execute(text("DROP TABLE IF EXISTS staging_so_details"))

            # ---------------------------------------------------------
            # 🌟 NHÁNH 3: ĐỒNG BỘ order_production_rules (XƯỞNG CÁN)
            # ---------------------------------------------------------
            df_to_sql = df_result[[
                'Order', 'tdc_version_id', 'master_id', 'req_min_w', 'req_max_w', 
                'SO Mapping', 'Target_Weight', 'Material description', 
                'KySanXuat', 'is_skin_required', 'production_status',
                'thickness', 'width', 'alloc_thick' 
            ]].copy()
            
            df_to_sql.columns = [
                'Order', 'tdc_version_id', 'master_id', 'req_min_w', 'req_max_w', 
                'SO Mapping', 'Target_Weight', 'material_desc', 
                'KySanXuat', 'is_skin_required', 'production_status',
                'req_thick', 'req_width', 'alloc_thick'
            ]
            
            df_to_sql['Order'] = df_to_sql['Order'].astype(str).str.replace(r'\.0$', '', regex=True)
            df_to_sql['SO Mapping'] = pd.to_numeric(df_to_sql['SO Mapping'], errors='coerce').astype('Int64')
            df_to_sql = df_to_sql.drop_duplicates(subset=['Order'], keep='last')

            df_to_sql.to_sql("staging_order_rules", transaction_conn, if_exists="replace", index=False, chunksize=1000, 
                             dtype={
                                 'Order': types.VARCHAR(50), 
                                 'tdc_version_id': types.BigInteger(), 
                                 'master_id': types.BigInteger(), 
                                 'SO Mapping': types.BigInteger(),
                                 'material_desc': types.NVARCHAR(1000)
                             })
            transaction_conn.execute(text("CREATE CLUSTERED INDEX IX_Stag ON staging_order_rules([Order])"))
            
            # 🌟 ĐOẠN MERGE SQL SAU ĐÂY ĐÃ ĐƯỢC KIỂM TRA KHÔNG TRÙNG LẶP BẤT KỲ CỘT NÀO
            merge_sql = text("""
               MERGE [dbo].[order_production_rules] AS target
                USING staging_order_rules AS source
                ON target.[Order] = CAST(source.[Order] AS VARCHAR(50))

                WHEN MATCHED THEN
                    UPDATE SET 
                        target.target_tdc_version_id = CASE 
                            WHEN source.tdc_version_id IS NULL THEN target.target_tdc_version_id
                            WHEN target.target_tdc_version_id IS NULL THEN source.tdc_version_id
                            WHEN ISNULL(target.is_manual_override, 0) = 0 THEN source.tdc_version_id
                            WHEN ISNULL(target.is_manual_override, 0) = 1 AND ISNULL((SELECT master_id FROM tdc_versions WHERE id = target.target_tdc_version_id), -1) = source.master_id THEN source.tdc_version_id
                            ELSE target.target_tdc_version_id 
                        END,
                        
                        target.req_min_w = CASE WHEN ISNULL(target.is_manual_override, 0) = 0 THEN source.req_min_w ELSE target.req_min_w END,
                        target.req_max_w = CASE WHEN ISNULL(target.is_manual_override, 0) = 0 THEN source.req_max_w ELSE target.req_max_w END,

                        target.is_conflict = CASE 
                            WHEN source.tdc_version_id IS NULL THEN target.is_conflict
                            WHEN target.target_tdc_version_id IS NULL THEN 0
                            WHEN ISNULL(target.is_manual_override, 0) = 0 THEN 0
                            WHEN ISNULL(target.is_manual_override, 0) = 1 THEN
                                CASE
                                    WHEN ISNULL((SELECT master_id FROM tdc_versions WHERE id = target.target_tdc_version_id), -1) <> source.master_id THEN 1
                                    WHEN ISNULL(target.req_min_w, -1) <> ISNULL(source.req_min_w, -1) OR ISNULL(target.req_max_w, -1) <> ISNULL(source.req_max_w, -1) THEN 1
                                    ELSE 0
                                END
                            ELSE target.is_conflict
                        END,

                        target.conflict_note = CASE 
                            WHEN source.tdc_version_id IS NULL THEN target.conflict_note
                            WHEN ISNULL(target.is_manual_override, 0) = 1 AND ISNULL((SELECT master_id FROM tdc_versions WHERE id = target.target_tdc_version_id), -1) <> source.master_id THEN N'⚠️ Báo động: Đơn hàng bị đổi sang mã TDC_Code khác.'
                            WHEN ISNULL(target.is_manual_override, 0) = 1 AND (ISNULL(target.req_min_w, -1) <> ISNULL(source.req_min_w, -1) OR ISNULL(target.req_max_w, -1) <> ISNULL(source.req_max_w, -1)) THEN N'⚠️ Trọng lượng Excel bị đổi lệch với bản chốt tay.'
                            ELSE NULL
                        END,

                        target.proposed_tdc_version_id = CASE 
                            WHEN source.tdc_version_id IS NULL THEN target.proposed_tdc_version_id
                            WHEN ISNULL(target.is_manual_override, 0) = 1 AND ISNULL((SELECT master_id FROM tdc_versions WHERE id = target.target_tdc_version_id), -1) <> source.master_id THEN source.tdc_version_id 
                            ELSE NULL 
                        END,
                        
                        target.proposed_min_w = CASE WHEN ISNULL(target.is_manual_override, 0) = 1 AND ISNULL(target.req_min_w, -1) <> ISNULL(source.req_min_w, -1) THEN source.req_min_w ELSE NULL END,
                        target.proposed_max_w = CASE WHEN ISNULL(target.is_manual_override, 0) = 1 AND ISNULL(target.req_max_w, -1) <> ISNULL(source.req_max_w, -1) THEN source.req_max_w ELSE NULL END,

                        target.SO_mapping = source.[SO Mapping],
                        target.total_weight = ISNULL(source.Target_Weight, 0),
                        target.material_desc = source.material_desc,
                        target.KySanXuat = source.KySanXuat,
                        target.is_skin_required = ISNULL(source.is_skin_required, 0),
                        target.production_status = source.production_status,
                        target.req_thick = source.req_thick,
                        target.req_width = source.req_width,
                        target.alloc_thick = source.alloc_thick

                WHEN NOT MATCHED THEN
                    INSERT ([Order], [target_tdc_version_id], [req_min_w], [req_max_w], is_manual_override, is_conflict,
                            [SO_mapping], [total_weight], [fulfilled_weight], [material_desc], [KySanXuat], [is_skin_required], [production_status], [req_thick], [req_width], [alloc_thick])
                    VALUES (CAST(source.[Order] AS VARCHAR(50)), source.tdc_version_id, source.req_min_w, source.req_max_w, 0, 0,
                            source.[SO Mapping], ISNULL(source.Target_Weight, 0), 0, source.material_desc, source.KySanXuat, ISNULL(source.is_skin_required, 0), source.production_status, source.req_thick, source.req_width, source.alloc_thick);
            """)
            transaction_conn.execute(merge_sql)
            transaction_conn.execute(text("DROP TABLE staging_order_rules"))
            
        logger.info(f"✅ Đã xử lý Exact Match thành công: Tất cả các nhánh dữ liệu được đồng bộ.")
    except Exception as e:
        logger.error(f"❌ Lỗi Rollback: {e}")
        raise e
import json
from sqlalchemy import text

# Định nghĩa danh sách các lỗi thuộc Giai đoạn 2 (Đợi kết quả Lab)
STAGE_2_DEFECTS = ['YieldPoint', 'Tensile', 'Elongation', 'Hardness', 'ImpactEnergy',  'C', 'Mn', 'Si', 'P', 'S', 'Cu', 'Ni', 'Cr', 'Mo', 'V', 'Ti', 'Al', 'Ca', 'B', 'Nb', 'CEV',  'O', 'N', 'H']

def evaluate_tdc_stage_1(scores_json, criteria_json, coil_weight, min_w, max_w):
    """ĐÁNH GIÁ GIAI ĐOẠN 1: Kích thước, Khối lượng và Bề mặt"""
    coil_weight = float(coil_weight) if coil_weight is not None else 0.0 # Chống lỗi NULL
    try:
        scores = json.loads(scores_json) if scores_json else {}
        criteria_list = json.loads(criteria_json) if criteria_json else []
    except Exception:
        return {'stage1_penalty': 9999, 'stage1_msg': 'Lỗi JSON', 'status': 'FAILNOCHEM'}

    penalty = 0
    failed_reasons = []
    is_thick_pass = scores.get('is_thick_pass', 1)
    if is_thick_pass == 0:
        penalty += 100
        failed_reasons.append("Chiều dày không đạt")

    # 1. Kiểm tra Khối lượng
    if min_w > 0 and coil_weight < min_w:
        penalty += 100
        failed_reasons.append(f"Khối lượng ({coil_weight}) < Min ({min_w})")
    if max_w > 0 and coil_weight > max_w:
        penalty += 100
        failed_reasons.append(f"Khối lượng ({coil_weight}) > Max ({max_w})")

    # 2. Kiểm tra Hình học / Bề mặt
    TOTAL_CRITERIA_COUNT = len(criteria_list)
    for idx, crit in enumerate(criteria_list):
        defect_key = crit['defect']
        if defect_key in STAGE_2_DEFECTS: continue # Bỏ qua cơ tính
            
        val = scores.get(defect_key, 0)
        allowed_range = crit.get('range', [])
        weight_score = TOTAL_CRITERIA_COUNT - idx 
        
        if val == 0: 
            penalty += (weight_score * 25)
            failed_reasons.append(f"{crit.get('name_vi', defect_key)}:Thiếu")
        elif allowed_range and val not in allowed_range:
            closest_limit = min(allowed_range, key=lambda x: abs(x - val))
            dist = abs(val - closest_limit)
            penalty += (weight_score * dist * 5)
            failed_reasons.append(f"{crit.get('name_vi', defect_key)}:C{val}(Lệch {dist})")

    return {
        'stage1_penalty': penalty,
        'stage1_msg': ', '.join(failed_reasons) if failed_reasons else "Đạt",
        'status': 'PASSNOCHEM' if penalty == 0 else 'FAILNOCHEM'
    }
def evaluate_tdc_stage_2(scores_json, criteria_json):
    """ĐÁNH GIÁ GIAI ĐOẠN 2: Cơ tính và Hóa học"""
    try:
        scores = json.loads(scores_json) if scores_json else {}
        criteria_list = json.loads(criteria_json) if criteria_json else []
    except Exception:
        return {'stage2_penalty': 9999, 'stage2_msg': 'Lỗi JSON'}

    penalty = 0
    failed_reasons = []
    TOTAL_CRITERIA_COUNT = len(criteria_list)
    
    for idx, crit in enumerate(criteria_list):
        defect_key = crit['defect']
        if defect_key not in STAGE_2_DEFECTS: continue # Chỉ chấm cơ tính
            
        val = scores.get(defect_key, 0)
        allowed_range = crit.get('range', [])
        weight_score = TOTAL_CRITERIA_COUNT - idx 
        
        if val == 0: 
            penalty += (weight_score * 25)
            failed_reasons.append(f"{crit.get('name_vi', defect_key)}:Thiếu")
        elif allowed_range and val not in allowed_range:
            closest_limit = min(allowed_range, key=lambda x: abs(x - val))
            dist = abs(val - closest_limit)
            penalty += (weight_score * dist * 5)
            failed_reasons.append(f"{crit.get('name_vi', defect_key)}:C{val}(Lệch {dist})")

    return {
        'stage2_penalty': penalty,
        'stage2_msg': ', '.join(failed_reasons) if failed_reasons else ""
    }
def process_etl_qc_lifecycle(engine):
    """Điều phối vòng đời QC (Đã đồng bộ chuẩn logic qc_msg từ save_manual_data)"""
    try:
        with engine.begin() as conn:
            # ==========================================
            # BLOCK 1: ĐÁNH GIÁ STAGE 1 (CUỘN MỚI TINH)
            # ==========================================
            sql_case_1 = text("""
                SELECT 
                    c.coil_id, c.weight, c.scores, v.criteria_json,
                    ISNULL(c.req_min_w, 0) as min_w, 
                    ISNULL(c.req_max_w, 0) as max_w
                FROM coil_data c WITH (NOLOCK)
                JOIN tdc_versions v WITH (NOLOCK) ON c.target_tdc_version_id = v.id
                WHERE c.qc_stage IS NULL 
                  AND c.target_tdc_version_id IS NOT NULL
            """)
            new_coils = conn.execute(sql_case_1).mappings().fetchall()
            
            if new_coils:
                update_stage1_payload = []
                for row in new_coils:
                    scores_dict = json.loads(row['scores']) if row['scores'] else {}
                    if 'is_thick_pass' not in scores_dict:
                        scores_dict['is_thick_pass'] = 1 # Mặc định luôn là Đạt
                    
                    new_scores_json = json.dumps(scores_dict)
                    res1 = evaluate_tdc_stage_1(new_scores_json, row['criteria_json'], row['weight'], row['min_w'], row['max_w'])
                    
                    # 🌟 VÁ LỖI 1: Xử lý qc_msg cho STAGE 1 (Nếu đạt thì rỗng, nếu lỗi thì lấy stage1_msg)
                    final_qc_msg = "" if res1['stage1_penalty'] == 0 else res1['stage1_msg']
                    
                    update_stage1_payload.append({
                        'cid': row['coil_id'], 
                        'pen1': res1['stage1_penalty'],
                        'msg1': res1['stage1_msg'], 
                        'status': res1['status'],
                        'qc_msg': final_qc_msg,
                        'new_scores': new_scores_json # Truyền thêm qc_msg
                    })
                
                # 🌟 VÁ LỖI SQL: Thêm qc_msg vào lệnh UPDATE
                conn.execute(text("""
                    UPDATE coil_data 
                    SET stage1_penalty = :pen1, 
                        stage1_msg = :msg1, 
                        qc_msg = :qc_msg, 
                        qc_status = :status, 
                        qc_stage = 'STAGE_1',
                        scores = :new_scores         
                    WHERE coil_id = :cid
                """), update_stage1_payload)
                logger.info(f"✅ [ETL-QC] Đã hoàn thành Stage 1 cho {len(update_stage1_payload)} cuộn mới.")

            # ==========================================
            # BLOCK 2: ĐÁNH GIÁ STAGE 2 & PHÂN BỔ (CÓ CƠ TÍNH)
            # ==========================================
            sql_case_3 = text("""
                SELECT 
                    c.coil_id, c.weight, c.scores, v.criteria_json,
                    c.stage1_penalty, c.stage1_msg,
                    r.[Order] as order_id, r.production_status, r.total_weight, r.fulfilled_weight, r.SO_mapping
                FROM coil_data c WITH (NOLOCK)
                JOIN tdc_versions v WITH (NOLOCK) ON c.target_tdc_version_id = v.id
                JOIN order_production_rules r WITH (NOLOCK) ON c.[Order] = r.[Order]
                WHERE c.qc_stage = 'STAGE_1'
                  AND JSON_VALUE(c.scores, '$.YieldPoint') != '0'
                  AND JSON_VALUE(c.scores, '$.Tensile') != '0'
                  AND JSON_VALUE(c.scores, '$.Elongation') != '0'
            """)
            ready_coils = conn.execute(sql_case_3).mappings().fetchall()
            
            if ready_coils:
                for row in ready_coils:
                    res2 = evaluate_tdc_stage_2(row['scores'], row['criteria_json'])
                    total_penalty = row['stage1_penalty'] + res2['stage2_penalty']
                    
                    cid = row['coil_id']
                    c_weight = row['weight'] or 0
                    order_id = row['order_id']
                    
                    # 🌟 VÁ LỖI 2: Logic xử lý tin nhắn ĐẠT/LỖI chuẩn như dashboard.py
                    if total_penalty == 0:
                        qc_status = 'PASS'
                        final_q_class = 'LOAI_1'
                        final_p_status = 'PRIME'
                        final_msg = "" # Xóa sạch tin nhắn nếu Pass hoàn toàn
                    else:
                        qc_status = 'FAIL'
                        final_q_class = None
                        final_p_status = None
                        # Lọc bỏ chữ "Đạt" và nối chuỗi thông minh
                        msgs = [m for m in [row['stage1_msg'], res2['stage2_msg']] if m and m != "Đạt"]
                        final_msg = " | ".join(msgs)
                        
                    # --- Logic phân bổ Room (Giữ nguyên) ---
                    final_mapped_po = None
                    if final_q_class == 'LOAI_1' and order_id:
                        check_room = conn.execute(text("""
                            SELECT fulfilled_weight, total_weight 
                            FROM order_production_rules WITH (UPDLOCK, ROWLOCK) 
                            WHERE [Order] = :oid
                        """), {"oid": order_id}).fetchone()
                        
                        if check_room:
                            curr_fulfilled = float(check_room[0] or 0)
                            total_allowed = float(check_room[1] or 0)
                            new_fulfilled = curr_fulfilled + c_weight
                            
                            conn.execute(text("UPDATE order_production_rules SET fulfilled_weight = :w WHERE [Order] = :oid"), {"w": new_fulfilled, "oid": order_id})
                            
                            if row['production_status'] == 'MTO':
                                if new_fulfilled <= total_allowed:
                                    final_mapped_po = row['SO_mapping'] if row['SO_mapping'] else '1'
                                else:
                                    final_mapped_po = '0'
                            else:
                                final_mapped_po = '0'

                    # 🌟 VÁ LỖI 3: Update đúng biến vào đúng cột (Đã thêm qc_msg)
                    conn.execute(text("""
                        UPDATE coil_data 
                        SET stage2_penalty = :pen2, 
                            stage2_msg = :s2_msg, 
                            qc_msg = :qc_msg, 
                            qc_status = :status, 
                            mapped_po = :mpo, 
                            qc_stage = 'STAGE_2',
                            quality_class = :q_class, 
                            prime_status = :p_status,
                            
                            rework_status = CASE 
                                WHEN :status <> 'PASS' THEN NULL
                                WHEN :status = 'PASS' AND ISNULL((
                                    SELECT TOP 1 is_skin_required 
                                    FROM order_production_rules 
                                    WHERE [Order] = coil_data.[Order]
                                ), 0) = 1 THEN 'SKIN_CUST'
                                WHEN :status = 'PASS' THEN 'FINAL'
                                ELSE NULL 
                            END
                            
                        WHERE coil_id = :cid
                    """), {
                        "pen2": res2['stage2_penalty'], 
                        "s2_msg": res2['stage2_msg'], 
                        "qc_msg": final_msg,          
                        "status": qc_status, 
                        "mpo": final_mapped_po, 
                        "q_class": final_q_class, 
                        "p_status": final_p_status, 
                        "cid": cid
                    })
                
                logger.info(f"✅ [ETL-QC] Đã hoàn thành Stage 2 và Phân bổ cho {len(ready_coils)} cuộn.")

    except Exception as e:
        logger.error(f"❌ [ETL-QC] Lỗi khi chạy lifecycle: {str(e)}")