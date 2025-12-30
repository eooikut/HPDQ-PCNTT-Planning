import os
import time
import logging
from datetime import datetime,timedelta
from apscheduler.schedulers.background import BackgroundScheduler
from db_utils import upsert_kho_from_excel, upsert_sanluong_from_excel, upsert_so_from_excel
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
        upsert_sanluong_from_excel(df_sl, "sanluong", nhamay=nhamay)
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
        upsert_so_from_excel(df_so, "so")
        logger.info(f"[{datetime.now()}] Cập nhật SO xong")
    except Exception as e:
        logger.error(f"Error update so: {e}")

def job_update_so():
    """Chạy cập nhật SO trước."""
    logger.info(f"[{datetime.now()}] --- BẮT ĐẦU JOB: update_so() ---")
    update_so()
    logger.info(f"[{datetime.now()}] --- KẾT THÚC JOB: update_so() ---")

def job_update_factory():
    """Chạy cập nhật sanluong, kho và ExportDataSAP sau SO 5 phút."""
    logger.info(f"[{datetime.now()}] --- BẮT ĐẦU JOB: update_factory() ---")
    for factory in FACTORIES:
        update_sanluong(factory)
        update_kho(factory)
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
        next_run_time=now+timedelta(minutes=10)
    )
    scheduler.start()
    logger.info("Scheduler đã khởi chạy:")
    logger.info(" - job_update_so() chạy mỗi 10 phút")
    logger.info(" - job_update_factory() chạy sau đó 5 phút")

if __name__ == "__main__":
    start_scheduler()
    while True:
        time.sleep(60)