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

logger = logging.getLogger(__name__)

def sync_order_production_rules_via_pandas(engine):
    logger.info("🔄 Bắt đầu Transform Order Rules bằng Pandas...")
    try:
        with engine.connect() as conn:
            # 1. KÉO DỮ LIỆU THÔ LÊN RAM (Sử dụng NOLOCK để chống nghẽn hệ thống)
            df_so = pd.read_sql("SELECT [Sales Document], [PO.], [Customer] FROM dbo.so WITH (NOLOCK)", conn)
            df_so_removed = pd.read_sql("SELECT [Sales Document], [PO.], [Customer] FROM dbo.so_removed WITH (NOLOCK)", conn)
            df_request = pd.read_sql("SELECT [SO Mapping], [CW], [Order], [Customer], [gradeSteel], [purpose] FROM dbo.so_request WITH (NOLOCK) WHERE [gradeSteel] LIKE '%SAE1006%'", conn)
            df_tdc = pd.read_sql("SELECT [id] as tdc_id, [customer_name], [grade], [usage_purpose] FROM dbo.tdc_master WITH (NOLOCK)", conn)
            df_active_ver = pd.read_sql("SELECT master_id, id as tdc_version_id FROM tdc_versions WITH (NOLOCK) WHERE status = 'Active'", conn)

        # ==========================================
        # BƯỚC 1: XỬ LÝ NHƯ VW_SO_REQUEST_WITH_CUSTOMER
        # ==========================================
        df_so_combined = pd.concat([df_so, df_so_removed]).drop_duplicates(subset=['Sales Document', 'PO.'], keep='first')
        df_request['Parsed_PO'] = df_request['Customer'].apply(lambda x: x.split('/')[-1].strip() if pd.notnull(x) and '/' in x else str(x).strip())
        
        df_request['SO Mapping'] = pd.to_numeric(df_request['SO Mapping'], errors='coerce')
        df_so_combined['Sales Document'] = pd.to_numeric(df_so_combined['Sales Document'], errors='coerce')
        
        df_merged_1 = pd.merge(df_request, df_so_combined, left_on='SO Mapping', right_on='Sales Document', how='left', suffixes=('', '_so'))
        
        df_so_combined['PO_clean'] = df_so_combined['PO.'].astype(str).str.strip().str.lower()
        df_merged_1['Parsed_PO_clean'] = df_merged_1['Parsed_PO'].astype(str).str.strip().str.lower()
        
        unmapped = df_merged_1[df_merged_1['Sales Document'].isnull()].copy()
        unmapped = unmapped.drop(columns=['Sales Document', 'PO.', 'Customer_so'])
        
        mapped_2 = pd.merge(unmapped, df_so_combined, left_on='Parsed_PO_clean', right_on='PO_clean', how='inner')
        df_final_request = pd.concat([df_merged_1[df_merged_1['Sales Document'].notnull()], mapped_2], ignore_index=True)
        df_final_request['Final_Customer'] = df_final_request['Customer_so'].combine_first(df_final_request['Customer'])

        # ==========================================
        # BƯỚC 2: XỬ LÝ NHƯ VW_TDC_PO
        # ==========================================
        def normalize_str(series, remove_spaces=False):
            s = series.astype(str).str.strip().str.lower()
            return s.str.replace(' ', '') if remove_spaces else s

        df_final_request['cust_norm'] = normalize_str(df_final_request['Final_Customer'])
        df_final_request['grade_norm'] = normalize_str(df_final_request['gradeSteel'])
        df_final_request['purp_norm'] = normalize_str(df_final_request['purpose'], remove_spaces=True)

        df_tdc['cust_norm'] = normalize_str(df_tdc['customer_name'])
        df_tdc['grade_norm'] = normalize_str(df_tdc['grade'])
        df_tdc['purp_norm'] = normalize_str(df_tdc['usage_purpose'], remove_spaces=True)

        df_result = pd.merge(df_final_request, df_tdc, on=['cust_norm', 'grade_norm', 'purp_norm'], how='inner')
        df_result = pd.merge(df_result, df_active_ver, left_on='tdc_id', right_on='master_id', how='inner')

        # ==========================================
        # BƯỚC 3: TÁCH MIN/MAX VÀ ĐẨY VÀO SQL
        # ==========================================
        if df_result.empty:
            logger.info("Không có Order nào map được.")
            return

        df_result['CW'] = df_result['CW'].fillna('').astype(str).str.strip()
        split_cw = df_result['CW'].str.split('-', n=1, expand=True)
        
        # [ĐÃ SỬA YÊU CẦU 1]: Nhân 1000 để đổi từ Tấn sang Kg
        df_result['req_min_w'] = pd.to_numeric(split_cw[0], errors='coerce').fillna(0) * 1000
        df_result['req_max_w'] = (pd.to_numeric(split_cw[1], errors='coerce').fillna(0) * 1000) if split_cw.shape[1] > 1 else 0.0

        # Lọc cột
        df_to_sql = df_result[['Order', 'tdc_version_id', 'req_min_w', 'req_max_w']].copy()
        # 1. Ép tất cả về số. Cái nào là chữ tự động biến thành NaN (Not a Number)
        df_to_sql['Order'] = pd.to_numeric(df_to_sql['Order'], errors='coerce')
        # 2. Xóa các dòng bị NaN (chính là xóa các dòng rác)
        df_to_sql = df_to_sql.dropna(subset=['Order'])
        # 3. Ép lại về kiểu số nguyên nguyên thủy (bỏ đuôi .0) rồi biến thành chuỗi chuẩn SQL
        df_to_sql['Order'] = df_to_sql['Order'].astype('Int64').astype(str)

        # Loại bỏ trùng lặp để bảo vệ Primary Key
        df_to_sql = df_to_sql.drop_duplicates(subset=['Order'])

        # ---------------------------------------------------------
        # KHU VỰC DEBUG & LOGGING DỮ LIỆU STAGING
        # ---------------------------------------------------------
        logger.info(f"🔍 [DEBUG STAGING] Tổng số Order chuẩn bị đẩy vào Database: {len(df_to_sql)}")

        # GHI VÀO DB BẰNG UPSERT
        with engine.begin() as conn:
            df_to_sql.to_sql("staging_order_rules", conn, if_exists="replace", index=False, chunksize=1000, dtype={'Order': types.VARCHAR(50)})
            
            conn.execute(text("CREATE CLUSTERED INDEX IX_Stag ON staging_order_rules([Order])"))
            
            merge_sql = text("""
                MERGE [dbo].[order_production_rules] AS target
                USING staging_order_rules AS source
                ON target.[Order] = CAST(source.[Order] AS VARCHAR(50))
                
                WHEN MATCHED AND (
                    ISNULL(target.target_tdc_version_id, -1) <> ISNULL(source.tdc_version_id, -1) OR
                    ISNULL(target.req_min_w, -1) <> ISNULL(source.req_min_w, -1) OR
                    ISNULL(target.req_max_w, -1) <> ISNULL(source.req_max_w, -1)
                ) THEN
                    UPDATE SET 
                        target.target_tdc_version_id = source.tdc_version_id, 
                        target.req_min_w = source.req_min_w, 
                        target.req_max_w = source.req_max_w
                        
                WHEN NOT MATCHED THEN
                    INSERT ([Order], [target_tdc_version_id], [req_min_w], [req_max_w])
                    VALUES (CAST(source.[Order] AS VARCHAR(50)), source.tdc_version_id, source.req_min_w, source.req_max_w);
            """)
            conn.execute(merge_sql)
            conn.execute(text("DROP TABLE staging_order_rules"))
            
        logger.info(f"✅ Đã xử lý và đồng bộ bằng Pandas: {len(df_to_sql)} Orders. Chớp mắt!")

    except Exception as e:
        logger.error(f"❌ Lỗi: {e}")
def evaluate_coil_against_tdc(scores_json, criteria_json, coil_weight, min_w, max_w):
    """
    Hàm tính điểm chuẩn QC. So sánh điểm lỗi thực tế với giới hạn của TDC Version.
    Trả về Dictionary chứa các chỉ số để lưu xuống Database.
    """
    try:
        scores = json.loads(scores_json) if scores_json else {}
        criteria_list = json.loads(criteria_json) if criteria_json else []
    except Exception:
        return {'penalty': 9999, 'match_pct': 0, 'match_ratio': '0/0', 'match_type': 'JSON_ERROR', 'failed_msg': 'Lỗi đọc dữ liệu JSON'}

    TOTAL_CRITERIA_COUNT = len(criteria_list)
    total_penalty = 0
    failed_reasons = [] 
    match_type = 'PERFECT'
    met_count = 0

    # 1. Kiểm tra Khối lượng
    if min_w > 0 and coil_weight < min_w:
        total_penalty += 100
        failed_reasons.append(f"Khối lượng ({coil_weight}) < Min ({min_w})")
        match_type = 'WEIGHT_FAIL'
    if max_w > 0 and coil_weight > max_w:
        total_penalty += 100
        failed_reasons.append(f"Khối lượng ({coil_weight}) > Max ({max_w})")
        match_type = 'WEIGHT_FAIL'

    # 2. Kiểm tra Tiêu chí Bề mặt / Cơ tính / Kích thước
    for idx, crit in enumerate(criteria_list):
        defect_key = crit['defect']
        defect_name = crit.get('name_vi', defect_key) 
        val = scores.get(defect_key, 0)
        allowed_range = crit.get('range', [])
        
        # Trọng số: Lỗi ở đầu list (quan trọng) bị phạt nặng hơn
        weight_score = TOTAL_CRITERIA_COUNT - idx 
        
        if val == 0: 
            penalty = weight_score * 25
            total_penalty += penalty
            failed_reasons.append(f"{defect_name}:Thiếu")
            if match_type == 'PERFECT': match_type = 'MISSING_DATA'
            
        elif allowed_range and val not in allowed_range:
            closest_limit = min(allowed_range, key=lambda x: abs(x - val))
            dist = abs(val - closest_limit)
            penalty = weight_score * dist * 5
            total_penalty += penalty
            failed_reasons.append(f"{defect_name}:C{val}(Lệch {dist})")
            if match_type in ('PERFECT', 'MISSING_DATA'): match_type = 'PROP_MISMATCH'
            
        else:
            met_count += 1

    match_pct = round((met_count / TOTAL_CRITERIA_COUNT) * 100, 1) if TOTAL_CRITERIA_COUNT > 0 else 0

    return {
        'penalty': total_penalty,
        'match_type': match_type,
        'match_pct': match_pct,
        'match_ratio': f"{met_count}/{TOTAL_CRITERIA_COUNT}",
        'failed_msg': ', '.join(failed_reasons)
    }

def run_auto_qc_batch(engine):
    """
    Background Job: Quét các cuộn PENDING và WAIT_LAB, kiểm tra cơ tính Lab, chấm điểm và cập nhật trạng thái.
    """
    logger.info("🔍 [AUTO-QC] Bắt đầu quét các cuộn chờ đánh giá...")
    
    try:
        with engine.begin() as conn:
            # 1. LẤY DỮ LIỆU ĐÃ TỐI ƯU
            query = text("""
                SELECT 
                    c.coil_id, 
                    ISNULL(c.weight, 0) as weight, 
                    c.scores, 
                    c.raw_data,
                    v.criteria_json, 
                    ISNULL(c.req_min_w, 0) as min_w, 
                    ISNULL(c.req_max_w, 0) as max_w,
                    c.qc_status -- Lấy thêm trạng thái hiện tại để so sánh
                FROM coil_data c
                JOIN tdc_versions v ON c.target_tdc_version_id = v.id 
                
                -- [TỐI ƯU 1]: Quét cả cuộn PENDING lẫn WAIT_LAB
                WHERE c.qc_status IN ('PENDING', 'WAIT_LAB')
                  AND c.target_tdc_version_id IS NOT NULL
            """)
            
            pending_coils = conn.execute(query).mappings().fetchall()
            
            if not pending_coils:
                logger.info("✅ [AUTO-QC] Không có cuộn nào cần xử lý.")
                return

            update_data = []
            
            # 2. XỬ LÝ LOGIC
            for coil in pending_coils:
                # [TỐI ƯU 2]: Sửa lại tên biến cho chuẩn để không bị NameError
                scores_str = coil['scores'] or "{}"
                raw_str = coil['raw_data'] or "{}"
                current_status = coil['qc_status']
                
                try:
                    scores_dict = json.loads(scores_str)
                    raw_dict = json.loads(raw_str)
                except Exception:
                    scores_dict = {}
                    raw_dict = {}
                
                # Rút giá trị số học ra
                yp_score = scores_dict.get('YieldPoint', 0)
                ts_score = scores_dict.get('Tensile', 0)
                yp_raw = raw_dict.get('YieldPoint', 0)
                
                has_lab_data = (yp_score > 0) or (ts_score > 0) or (yp_raw > 0)
                
                if not has_lab_data:
                    # [TỐI ƯU 3 - SMART UPDATE]: 
                    # Nếu cuộn đang là WAIT_LAB rồi và vẫn chưa có điểm -> Bỏ qua luôn!
                    # Cứu SQL Server khỏi việc bị SPAM hàng ngàn lệnh UPDATE vô nghĩa.
                    if current_status == 'WAIT_LAB':
                        continue 
                        
                    # Chỉ cập nhật đối với cuộn PENDING lần đầu tiên được đưa về WAIT_LAB
                    update_data.append({
                        'c_id': coil['coil_id'],
                        'st': 'WAIT_LAB', 
                        'pen': 0, 
                        'pct': 0, 
                        'typ': 'WAITING', 
                        'msg': 'Chờ kết quả cơ tính/hóa học'
                    })
                else:
                    # Đã có đủ dữ liệu -> Chấm điểm
                    res = evaluate_coil_against_tdc(
                        scores_str, 
                        coil['criteria_json'], 
                        coil['weight'], 
                        coil['min_w'], 
                        coil['max_w']
                    )
                    
                    qc_status = 'PASS' if res['penalty'] == 0 else 'FAIL'
                    
                    update_data.append({
                        'c_id': coil['coil_id'],
                        'st': qc_status,
                        'pen': res['penalty'],
                        'pct': res['match_pct'],
                        'typ': res['match_type'],
                        'msg': res['failed_msg'][:4000]
                    })

            # 3. GHI XUỐNG DB
            if update_data:
                update_sql = text("""
                    UPDATE coil_data 
                    SET qc_status = :st, 
                        qc_penalty = :pen, 
                        qc_match_pct = :pct, 
                        qc_match_type = :typ, 
                        qc_msg = :msg
                    WHERE coil_id = :c_id
                """)
                conn.execute(update_sql, update_data)
                logger.info(f"✅ [AUTO-QC] Đã chấm điểm và cập nhật {len(update_data)} cuộn.")
            else:
                logger.info(f"✅ [AUTO-QC] Đã quét {len(pending_coils)} cuộn nhưng tất cả đều đang ở trạng thái chờ Lab, không cần thao tác DB.")
                
    except Exception as e:
        logger.error(f"❌ [AUTO-QC] Lỗi hệ thống: {str(e)}")