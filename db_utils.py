import pandas as pd
from datetime import datetime
from db import engine  # import engine SQLAlchemy đã cấu hình
# ---------- Ghi DataFrame vào SQL Server ----------
from sqlalchemy import inspect,types,text

import logging

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

# ------------------- Helper -------------------

# ------------------- Upsert SANLUONG -------------------
from datetime import datetime
from sqlalchemy import text, types
from db import engine
import pandas as pd

def upsert_sanluong_from_excel(df: pd.DataFrame, table_name: str = "sanluong", nhamay: str = "HRC1"):
    if df.empty:
        return

    df = df.copy()
    df["NhaMay"] = nhamay
    snap_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df["snapshot_ts"] = snap_ts
    df["status"] = "active"
    if "ID Cuộn Bó" in df.columns:
        df["ID Cuộn Bó"] = pd.to_numeric(df["ID Cuộn Bó"], errors="coerce").fillna(0).astype("Int64")
    if "Order" in df.columns:
        df["Order"] = pd.to_numeric(df["Order"], errors="coerce").fillna(0).astype("Int64")
    # Tạo staging riêng cho từng nhà máy
    staging_name = f"staging_sanluong_{nhamay}"

    # dtype mapping
    dtype_mapping = {}
    for col in df.columns:
        # Ép kiểu BIGINT cho ID Cuộn Bó và Order
        if col in ["ID Cuộn Bó", "Order"]: 
            dtype_mapping[col] = types.BigInteger()
        elif pd.api.types.is_string_dtype(df[col]):
            dtype_mapping[col] = types.NVARCHAR(length=4000)
        elif pd.api.types.is_integer_dtype(df[col]):
            # Các cột Integer khác (nếu có)
            dtype_mapping[col] = types.BigInteger() 
        elif pd.api.types.is_float_dtype(df[col]):
            dtype_mapping[col] = types.Float()
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            dtype_mapping[col] = types.DateTime()
        else:
            dtype_mapping[col] = types.NVARCHAR(length=4000)

    with engine.begin() as conn:
        # 1️⃣ Ghi staging tạm (riêng cho nhà máy)
        df.to_sql(staging_name, conn, if_exists="replace", index=False, dtype=dtype_mapping)

        # 2️⃣ Đánh dấu removed trong cùng nhà máy
        conn.execute(text(f"""
            UPDATE [{table_name}]
            SET status='removed', snapshot_ts=:snap
            WHERE status IN ('active','updated') AND NhaMay=:nhamay
              AND NOT EXISTS (
                  SELECT 1 FROM [{staging_name}] s
                  WHERE s.[ID Cuộn Bó] = [{table_name}].[ID Cuộn Bó]
                    AND s.NhaMay = :nhamay
              )
        """), {"snap": snap_ts, "nhamay": nhamay})

        # 3️⃣ Chuyển removed sang bảng _removed
        conn.execute(text(f"""
            INSERT INTO [{table_name}_removed]
            SELECT * FROM [{table_name}] 
            WHERE status='removed' AND NhaMay=:nhamay
        """), {"nhamay": nhamay})

        # 4️⃣ Xóa các dòng removed khỏi bảng chính
        conn.execute(text(f"""
            DELETE FROM [{table_name}] 
            WHERE status='removed' AND NhaMay=:nhamay
        """), {"nhamay": nhamay})

        # 5️⃣ Update dữ liệu khác biệt giữa staging và bảng chính
        cols_to_update = [c for c in df.columns if c not in ["ID Cuộn Bó", "NhaMay", "status", "snapshot_ts"]]
        if cols_to_update:
            set_clause = ", ".join([f"t.[{c}] = s.[{c}]" for c in cols_to_update])
            diff_condition = " OR ".join([f"ISNULL(t.[{c}], '') <> ISNULL(s.[{c}], '')" for c in cols_to_update])
            # thêm status và snapshot_ts
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

        # 6️⃣ Insert mới từ staging
        conn.execute(text(f"""
            INSERT INTO [{table_name}]
            SELECT s.* 
            FROM [{staging_name}] s
            WHERE NOT EXISTS (
                SELECT 1 FROM [{table_name}] t
                WHERE t.[ID Cuộn Bó] = s.[ID Cuộn Bó] 
                  AND t.NhaMay = s.NhaMay
            )
        """))

        # 7️⃣ Dọn staging
        conn.execute(text(f"DROP TABLE IF EXISTS [{staging_name}]"))

import pandas as pd
from datetime import datetime
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
    bigint_cols = ["ID Cuộn Bó","Material", "SO Mapping"]

    # Các cột float
    float_cols = [
        "Khối lượng","SO Item Ma","Batch","Order","Trạm cân","Số lượng in","Storage Location"
    ]

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

    # Lọc bỏ hàng không có ID
    df = df[df["ID Cuộn Bó"].notna() & (df["ID Cuộn Bó"] > 0)]

    snap_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df["snapshot_ts"] = snap_ts
    df["status"] = "active"

    # ==== 3️⃣ Mapping dtype cho SQL ====
    dtype_mapping = {}
    for c in KHO_SCHEMA:
        if c in bigint_cols:
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


def upsert_so_from_excel(df: pd.DataFrame, table_name: str):
    df = normalize_datetime(df)
    snap_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df = df.copy()
    df["snapshot_ts"] = snap_ts
    df["status1"] = "active"

    with engine.begin() as conn:
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

        # 2️⃣ Staging table
        df.to_sql("staging_tmp", conn, if_exists="replace", index=False, dtype=dtype_mapping)

        # 3️⃣ Đánh dấu record bị loại
        conn.execute(text(f"""
            UPDATE t
            SET status1='removed', snapshot_ts=:ts
            FROM [{table_name}] t
            WHERE NOT EXISTS (
                SELECT 1 FROM staging_tmp s
                WHERE s.[Sales Document]=t.[Sales Document]
                  AND s.[Material]=t.[Material]
                  AND s.[Sales Document Item]=t.[Sales Document Item]
            ) AND t.status1 IN ('active','updated')
        """), {"ts": snap_ts})

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