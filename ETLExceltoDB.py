import os
import time
import logging
from datetime import datetime,timedelta
from apscheduler.schedulers.background import BackgroundScheduler
from db_utils import upsert_kho_from_excel, upsert_sanluong_from_excel, upsert_so_from_excel, sync_order_production_rules_via_pandas,process_etl_qc_lifecycle
from sqlalchemy import text
from db import engine
from phanbodudoan import ExportDataSAP
# ---------- Cấu hình ----------
LOCAL_FOLDER = "data_auto_update"
FACTORIES = ["nm1", "nm2"]
FILES = {
    "kho": "kho_{factory}.xlsx",
    "sanluong": "sanluong_{factory}.xlsx",
    "so": "so.xlsx"  # file SO chung cho tất cả
}

# Map factory -> NhaMay cho sanluong
FACTORY_MAP = {"nm1": "HRC1", "nm2": "HRC2"}

# ---------- Logger ----------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------- Lấy file path ----------
def get_file_path(file_name):
    path = os.path.join(LOCAL_FOLDER, file_name)
    if os.path.exists(path):
        return path
    return None

# ---------- UTILITY: Chờ file được giải phóng (Mã mới quan trọng) ----------
def wait_for_file_release(file_path, timeout=120):
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            # Thử mở file ở chế độ ghi ('a') để kiểm tra quyền truy cập độc quyền.
            # Nếu thành công, file không bị khóa, và ta đóng nó lại ngay lập tức.
            with open(file_path, 'a') as f:
                pass
            logger.info(f"File {file_path} đã được giải phóng.")
            return True
        except PermissionError:
            # File vẫn đang bị khóa bởi tiến trình Export SAP hoặc Excel.exe.
            logger.warning(f"File {file_path} đang bị khóa. Đang chờ 5 giây...")
            time.sleep(5)
        except Exception as e:
             # Xử lý các lỗi khác như FileNotFound, IO Error, ...
            logger.error(f"Lỗi không xác định khi kiểm tra file {file_path}: {e}")
            return False

    # Nếu hết thời gian timeout
    logger.error(f"!!! Timeout 300s: File {file_path} vẫn bị khóa. Bỏ qua lần cập nhật này.")
    return False


# ---------- Update sanluong ----------
def update_sanluong(factory):
    from XuLyDuLieuNhapVao import read_file_auto,process_actual
    try:
        file_name = FILES["sanluong"].format(factory=factory)
        path = get_file_path(file_name)
        if not path:
            logger.warning(f"File {file_name} không tồn tại")
            return
        

        if not wait_for_file_release(path):
            return

        nhamay = FACTORY_MAP.get(factory, factory.upper())
        df_sl = read_file_auto(path)
        df_sl = df_sl.dropna(subset=["ID Cuộn Bó"])
        df_sl["NhaMay"] = nhamay
        df_sl = df_sl.drop_duplicates(subset=['ID Cuộn Bó', 'NhaMay'])
        # Truyền NhaMay vào hàm upsert
        upsert_sanluong_from_excel(df_sl, "sanluong", nhamay=nhamay, date_col_name="Ngày sản xuất")
        logger.info(f"[{datetime.now()}] Cập nhật sanluong nhà máy {factory} ({nhamay}) xong")
    except Exception as e:
        logger.error(f"Error update sanluong ({factory}): {e}")

# ---------- Update kho ----------
def update_kho(factory):
    from XuLyDuLieuNhapVao import read_file_auto
    try:
        file_name = FILES["kho"].format(factory=factory)
        path = get_file_path(file_name)
        if not path:
            logger.warning(f"File {file_name} không tồn tại")
            return

        if not wait_for_file_release(path):
            return
        df_kho = read_file_auto(path)
        df_kho = df_kho.dropna(subset=["ID Cuộn Bó"])
        upsert_kho_from_excel(df_kho, "kho")
        logger.info(f"[{datetime.now()}] Cập nhật kho nhà máy {factory} xong")
    except Exception as e:
        logger.error(f"Error update kho ({factory}): {e}")

# ---------- Update SO chung ----------
def update_so():
    from XuLyDuLieuNhapVao import read_file_auto
    try:
        path = get_file_path(FILES["so"])
        if not path:
            logger.warning("File SO không tồn tại")
            return

        if not wait_for_file_release(path):
            return
            
        df_so = read_file_auto(path, skiprows=2)
        df_so = df_so.loc[:, ~df_so.columns.str.contains('^Unnamed')]
        upsert_so_from_excel(df_so, "so", date_col_name="Document Date")
        logger.info(f"[{datetime.now()}] Cập nhật SO xong")
    except Exception as e:
        logger.error(f"Error update so: {e}")

def job_update_so():
    """Chạy cập nhật SO trước."""
    logger.info(f"[{datetime.now()}] --- BẮT ĐẦU JOB: update_so() ---")
    update_so()
    logger.info(f"[{datetime.now()}] --- KẾT THÚC JOB: update_so() ---")
def run_post_etl_stored_procedures():
    """Chạy các SP đồng bộ nội bộ trong Database sau khi ETL đã nạp xong dữ liệu mới"""
    logger.info(f"[{datetime.now()}] --- BẮT ĐẦU CHẠY STORED PROCEDURES ĐỒNG BỘ ---")
    try:
        # --- NHỊP 0: Chạy Pandas ---
        sync_order_production_rules_via_pandas(engine)
        
        # --- NHỊP 1: Chạy SP Đồng bộ (CÓ CƠ CHẾ CHỐNG DEADLOCK - RETRY 3 LẦN) ---
        max_retries = 3
        for attempt in range(max_retries):
            try:
                with engine.begin() as conn:
                    logger.info(f"Đang chạy SP: sp_Sync_SAP_ID_To_Production... (Lần thử {attempt + 1})")
                    # Thêm SET DEADLOCK_PRIORITY LOW để nhường đường nếu đụng độ
                    conn.execute(text("SET DEADLOCK_PRIORITY LOW; EXEC [dbo].[sp_Sync_SAP_ID_To_Production]"))
                    
                    logger.info("Đang chạy SP: sp_Sync_SanLuong_Order_TieuChuan...")
                    conn.execute(text("EXEC [dbo].[sp_Sync_SanLuong_Order_TieuChuan]"))
                
                # Nếu chạy đến đây không có lỗi -> Break (thoát vòng lặp retry)
                break 
                
            except Exception as db_err:
                if '1205' in str(db_err): # 1205 là mã lỗi Deadlock của SQL Server
                    logger.warning(f"⚠️ Bị Deadlock ở Nhịp 1 (Lần {attempt + 1}). Chờ 5 giây để thử lại...")
                    time.sleep(5)
                    if attempt == max_retries - 1:
                        logger.error("❌ Đã thử lại 3 lần nhưng vẫn bị Deadlock!")
                        raise db_err # Thất bại hoàn toàn thì ném lỗi ra ngoài
                else:
                    raise db_err # Lỗi khác không phải Deadlock thì ném ra ngoài luôn
                    
        # --- NHỊP 2: Chạy Python Auto-QC chấm điểm (Chạy ĐỘC LẬP bên ngoài khối with) ---
        # Lúc này Store vừa dập xong, các cuộn đã thành 'PENDING', Python lập tức vào xử lý.
        logger.info("Đang chạy Python Auto-QC Batch...")

        # --- NHỊP 3: Đồng bộ Kho (CŨNG BỌC CHỐNG DEADLOCK LUÔN CHO AN TOÀN) ---
        for attempt in range(max_retries):
            try:
                with engine.begin() as conn:
                    logger.info(f"Đang chạy SP: sp_SyncKhoToCoilData... (Lần thử {attempt + 1})")
                    conn.execute(text("SET DEADLOCK_PRIORITY LOW; EXEC [dbo].[sp_SyncKhoToCoilData]")) 
                break # Thành công thì thoát
            except Exception as db_err:
                if '1205' in str(db_err):
                    logger.warning(f"⚠️ Bị Deadlock ở Kho (Lần {attempt + 1}). Chờ 5 giây để thử lại...")
                    time.sleep(5)
                    if attempt == max_retries - 1:
                        raise db_err
                else:
                    raise db_err

    except Exception as e:
        logger.error(f"LỖI NGHIÊM TRỌNG KHI CHẠY SP ĐỒNG BỘ: {e}")
def job_update_factory():
    """Chạy cập nhật sanluong, kho và ExportDataSAP sau SO 5 phút."""
    logger.info(f"[{datetime.now()}] --- BẮT ĐẦU JOB: update_factory() ---")
    for factory in FACTORIES:
        update_sanluong(factory)
        update_kho(factory)
    run_post_etl_stored_procedures()
    process_etl_qc_lifecycle(engine)
    ExportDataSAP()
    logger.info(f"[{datetime.now()}] --- KẾT THÚC JOB: update_factory() ---")

def start_scheduler():
    os.makedirs(LOCAL_FOLDER, exist_ok=True)
    scheduler = BackgroundScheduler()

    now = datetime.now()
    scheduler.add_job(
        job_update_so,
        'interval',
        minutes=20,
        next_run_time=now
    )

    # Job 2: update_factory chạy mỗi 10 phút, nhưng trễ hơn 5 phút
    scheduler.add_job(
        job_update_factory,
        'interval',
        minutes=20,
        next_run_time=now + timedelta(minutes=5)
    )
    scheduler.start()
    logger.info("Scheduler đã khởi chạy:")
    logger.info(" - job_update_so() chạy mỗi 10 phút")
    logger.info(" - job_update_factory() chạy sau đó 5 phút")

if __name__ == "__main__":
    start_scheduler()
    while True:
        time.sleep(60)