import win32com.client
import time
import os
import sys
import subprocess
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
import win32gui
import win32con
# ===============================
# 📌 CẤU HÌNH ĐĂNG NHẬP (QUAN TRỌNG: CẦN SỬA)
# ===============================
# Tên kết nối phải GIỐNG Y HỆT trong SAP Logon Pad của bạn (Ví dụ: "1. PRD - Production")
SAP_CONNECTION_NAME = "PRD_UPGRADE" 
SAP_USERNAME = "zit06"
SAP_PASSWORD = "Vtkn2k167@1"
# Đường dẫn tới file chạy SAP (Thường mặc định như dưới, nếu khác hãy sửa lại)
SAP_LOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

# ===============================
# 📌 CẤU HÌNH CHUNG
# ===============================
CUSTOM_DIR = r"C:\Users\Administrator\Desktop\ProjectPKH\data_auto_update"
LOG_PATH = os.path.join(CUSTOM_DIR, "master_export_log.txt")
def clear_ghost_icons():
    """Dùng Windows API quét và xóa icon ma dựa trên X-Quang thực tế."""
    try:
        toolbars = []

        # 1. BẮT KHAY HIỂN THỊ (Visible Tray)
        hwnd_tray = win32gui.FindWindow("Shell_TrayWnd", None)
        if hwnd_tray:
            hwnd_notify = win32gui.FindWindowEx(hwnd_tray, 0, "TrayNotifyWnd", None)
            if hwnd_notify:
                # Bắt Toolbar nằm trong SysPager (Như ảnh X-Quang)
                hwnd_pager = win32gui.FindWindowEx(hwnd_notify, 0, "SysPager", None)
                if hwnd_pager:
                    tb1 = win32gui.FindWindowEx(hwnd_pager, 0, "ToolbarWindow32", None)
                    if tb1: toolbars.append(tb1)
                
                # Bắt Toolbar nằm ngoài (Phòng hờ)
                tb2 = win32gui.FindWindowEx(hwnd_notify, 0, "ToolbarWindow32", None)
                if tb2: toolbars.append(tb2)

        # 2. BẮT KHAY ẨN (Overflow Tray - Ổ chứa icon ma của bạn)
        hwnd_overflow = win32gui.FindWindow("NotifyIconOverflowWindow", None)
        if hwnd_overflow:
            tb3 = win32gui.FindWindowEx(hwnd_overflow, 0, "ToolbarWindow32", None)
            if tb3: toolbars.append(tb3)

        # 3. TIẾN HÀNH QUÉT CHUỘT ẢO
        for tb in toolbars:
            rect = win32gui.GetClientRect(tb)
            w = rect[2]
            h = rect[3]
            
            # 🔥 MẤU CHỐT Ở ĐÂY: Nếu khay bị ẩn (kích thước 0x0), ép quét vùng 400x400
            if w == 0 or h == 0:
                w = 400
                h = 400

            # Quét từng pixel với bước nhảy 5px
            for x in range(0, w, 5):
                for y in range(0, h, 5):
                    win32gui.SendMessage(tb, win32con.WM_MOUSEMOVE, 0, (y << 16) | x)

        log_message("🔄 Đã quét dọn triệt để Ghost Icons (bao gồm cả khay ẩn).", "INFO")
    except Exception as e:
        log_message(f"⚠️ Lỗi clear ghost icon: {e}", "WARN")
# --- Hàm ghi log tập trung ---
def log_message(message, level="INFO"):
    """Ghi thông báo ra console và file log."""
    time_stamp = datetime.now().strftime('%H:%M:%S')
    full_message = f"[{time_stamp}] {level.upper()[:1]}️ {message}"
    print(full_message)
    if level in ["ERROR", "WARN", "CRITICAL", "SUCCESS"]:
        with open(LOG_PATH, "a", encoding="utf-8") as log:
             log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {'✅' if level == 'SUCCESS' else '🛑'} {message}\n")
def kill_all_sap_processes():
    processes = ["sapgui.exe", "saplogon.exe", "saplgpad.exe"]
    for process in processes:
        try:
            subprocess.run(
                ['taskkill', '/f', '/im', process, '/t'], 
                capture_output=True, 
                check=False
            )
        except Exception as e:
            pass
    time.sleep(2)
    clear_ghost_icons()
    log_message("✅ Đã dọn dẹp xong toàn bộ tiến trình SAP.", level="SUCCESS")
def graceful_exit_sap(session):
    """
    Thử đóng SAP một cách 'êm ái' bằng lệnh /nex để tránh lỗi Ghost Icon.
    Nếu thất bại (do treo, mất kết nối) thì mới dùng đến biện pháp mạnh (Kill).
    """
    try:
        if session is None:
            log_message("⚠️ Không có phiên SAP để đóng (Session is None).", level="WARN")
            return

        log_message("🚪 Đang gửi lệnh đóng SAP (/nex)...", level="INFO")
        
        # /nex là T-code đặc biệt để thoát ngay lập tức không cần hỏi Yes/No
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
        session.findById("wnd[0]").sendVKey(0) # Enter
        
        time.sleep(3) # Chờ 3 giây để cửa sổ SAP tự biến mất
        session = None
        log_message("✅ Đã đóng SAP an toàn (Graceful Exit).", level="SUCCESS")

    except Exception as e:
        log_message(f"⚠️ Không thể đóng êm (Lỗi: {e}). Đang chuyển sang chế độ Kill...", level="WARN")
        # Nếu gửi lệnh thất bại (ví dụ SAP đang bị treo), buộc phải Kill
        kill_all_sap_processes()
# --- HÀM CHỜ THÔNG MINH ---
def wait_for_sap_ready(session, max_wait_minutes=15):
    log_message(f"⏳ Bắt đầu chờ SAP xử lý (Tối đa {max_wait_minutes} phút)...", level="INFO")
    start_time = time.time()
    while True:
        elapsed_time = (time.time() - start_time) / 60
        if elapsed_time > max_wait_minutes:
            raise Exception("TIMEOUT: SAP chạy quá lâu.")

        try:
            if session.Busy:
                time.sleep(1)
                continue
            # Xử lý Popup cảnh báo dữ liệu lớn
            if session.Children.Count > 1:
                try:
                    session.findById("wnd[1]").sendVKey(0) # Enter
                    log_message("⚠️ Đã tự động đóng Popup cảnh báo.", level="WARN")
                    time.sleep(1)
                except: pass

            status_text = session.findById("wnd[0]/sbar").Text
            break 
        except:
            time.sleep(1)
            pass
        time.sleep(1)
    log_message(f"✅ SAP phản hồi sau {round(elapsed_time * 60, 0)} giây.", level="SUCCESS")


# --- CÁC ID SAP CHUNG ---
MULTI_SELECT_TABLE_PATH = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
MULTI_SELECT_INPUT_BASE = "ctxtRSCSEL_255-SLOW_I"
PLANT_INPUT_ID_ZPP04A = "wnd[0]/usr/ctxtS_WERKS-LOW"
STORAGE_LOC_BUTTON_ID = "wnd[0]/usr/btn%_S_LGORT_%_APP_%-VALU_PUSH"
DATE_FROM_ID_ZBC04B = "wnd[0]/usr/ctxtS_NGAYSX-LOW"
DATE_TO_ID_ZBC04B = "wnd[0]/usr/ctxtS_NGAYSX-HIGH"
PLANT_ID_ZBC04B = "wnd[0]/usr/ctxtS_WERKS-LOW" 
PRODUCT_GROUP_ID = "wnd[0]/usr/ctxtS_PX-LOW"
PRODUCT_GROUP_BUTTON_ID = "wnd[0]/usr/btn%_S_PX_%_APP_%-VALU_PUSH"
L1_CHECKBOX_ID = "wnd[0]/usr/chkP_L1"
L2_CHECKBOX_ID = "wnd[0]/usr/chkP_L2"
DATE_FROM_ID_ZSD04A = "wnd[0]/usr/ctxtS_VDATU-LOW"
DATE_TO_ID_ZSD04A = "wnd[0]/usr/ctxtS_VDATU-HIGH" 
ORDER_TYPE_BUTTON_ID = "wnd[0]/usr/btn%_S_AUART_%_APP_%-VALU_PUSH" 

# ===============================
# 📝 CẤU HÌNH TÁC VỤ
# ===============================
TASK_CONFIGS = [
    # 1. ZSD04A
    {
        "name": "ZSD04A_ALL", 
        "tcode": "ZSD04A",
        "output_filename": "so.xlsx",
        "menu_export_path": "wnd[0]/mbar/menu[0]/menu[3]/menu[0]",
        "group": "SLOW",
        "params": {
            "DATE_FROM": "{ZSD04A_FROM}",
            "DATE_TO": "{ZSD04A_TO}",
            "ORDER_TYPES_LIST": ["ZOR5", "ZOR6", "ZOR8", "ZOR7", "ZORI", "ZORZ", "ZORY","ZORL"], 
        }
    },
    # 2. ZPP04A - HRC2
    {
        "name": "ZPP04A_HRC2",
        "tcode": "ZPP04A",
        "output_filename": "kho_nm2.xlsx",
        "menu_export_path": "wnd[0]/mbar/menu[0]/menu[3]/menu[0]",
        "group": "FAST",
        "params": { "PLANT_VALUE": "1600", "STORAGE_LOCATIONS_LIST": ["1505", "1506" , "1522" ] }
    },
    # 3. ZPP04 - HRC1
    {
        "name": "ZPP04_HRC1",
        "tcode": "ZPP04",
        "output_filename": "kho_nm1.xlsx",
        "menu_export_path": "wnd[0]/mbar/menu[0]/menu[3]/menu[0]",
        "group": "FAST",
        "params": { "PLANT_VALUE": "1000", "STORAGE_LOCATIONS_LIST": ["1519", "1522"] }
    },
    # 4. ZBC04B - HRC1
    {
        "name": "ZBC04B_HRC1",
        "tcode": "ZBC04B",
        "output_filename": "sanluong_nm1.xlsx",
        "menu_export_path": "wnd[0]/mbar/menu[0]/menu[1]/menu[0]",
        "group": "FAST",
        "params": { "DATE_FROM": "{ZBC04B_FROM}", "DATE_TO": "{ZBC04B_TO}", "PLANT_VALUE": "1000", "PRODUCT_GROUP_LIST": ["7", "7H"], "UNCHECK_L1_L2": True }
    },
    # 5. ZBC04B - HRC2
    {
        "name": "ZBC04B_HRC2",
        "tcode": "ZBC04B",
        "output_filename": "sanluong_nm2.xlsx",
        "menu_export_path": "wnd[0]/mbar/menu[0]/menu[1]/menu[0]",
        "group": "FAST",
        "params": { "DATE_FROM": "{ZBC04B_FROM}", "DATE_TO": "{ZBC04B_TO}", "PLANT_VALUE": "1600", "PRODUCT_GROUP_LIST": ["8", "7H"], "UNCHECK_L1_L2": True }
    },
]

# ===============================
# CÁC HÀM HỖ TRỢ
# ===============================
def calculate_dynamic_dates():
    today = date.today()
    tomorrow = today + timedelta(days=1)
    today_sap = today.strftime("%d.%m.%Y")
    tomorrow_sap = tomorrow.strftime("%d.%m.%Y")
    # ZSD04A
    six_months_ago = today - relativedelta(months=4)
    start_zsd04a = six_months_ago.replace(day=1).strftime("%d.%m.%Y")
    
    # ZBC04B
    first_day_prev_month = (today - relativedelta(months=2)).replace(day=1)
    start_zbc04b = first_day_prev_month.strftime("%d.%m.%Y")
    
    return {"ZSD04A_FROM": start_zsd04a, "ZSD04A_TO": tomorrow_sap, "ZBC04B_FROM": start_zbc04b, "ZBC04B_TO": tomorrow_sap}

def force_close_process(process_name):
    """Tắt cưỡng bức tiến trình."""
    try:
        subprocess.run(['taskkill', '/f', '/im', process_name], capture_output=True, check=False)
    except: pass

# ===============================
# 🔐 HÀM LOGIN & CONNECT 
# ===============================
def sap_login_and_connect():
    """
    Kết nối thông minh: 
    1. Thử dùng session đang mở.
    2. Nếu chưa login (đang ở màn hình User/Pass), tiến hành login.
    3. Nếu không thấy SAP, mới khởi động mới hoàn toàn.
    """
    session = None
    try:
        # Bước 1: Kiểm tra xem có tiến trình SAP nào đang chạy không
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        
        if application.Children.Count > 0:
            connection = application.Children(0)
            if connection.Children.Count > 0:
                session = connection.Children(0)
                
                # Bước 2: Kiểm tra trạng thái Login thực tế
                try:
                    # Thử tìm ô nhập Username. Nếu thấy -> Chưa login.
                    session.findById("wnd[0]/usr/txtRSYST-BNAME")
                    log_message("⚠️ SAP đang mở nhưng chưa login. Tiến hành đăng nhập...", level="WARN")
                except:
                    # Nếu không tìm thấy ô nhập User -> Đã vào màn hình chính -> Dùng luôn
                    log_message("✅ Đã kết nối vào Session đang hoạt động.", level="SUCCESS")
                    return session
        else:
            log_message("🔍 Không tìm thấy kết nối SAP hiện có.", level="INFO")
    except Exception:
        # Lỗi ở đây thường là do chưa mở SAP Logon
        pass

    # Bước 3: Nếu không tận dụng được session cũ, dọn dẹp để mở mới
    log_message("🚀 Đang khởi động và đăng nhập SAP mới từ đầu...", level="INFO")
    kill_all_sap_processes() # Quét sạch Ghost Icons cũ trước khi bắt đầu
    
    try:
        # Mở SAP Logon
        subprocess.Popen(SAP_LOGON_PATH)
        time.sleep(5) 

        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        
        log_message(f"Kết nối tới Server: {SAP_CONNECTION_NAME}...", level="INFO")
        connection = application.OpenConnection(SAP_CONNECTION_NAME, True)
        time.sleep(3)
        session = connection.Children(0)

        # Điền thông tin đăng nhập
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = SAP_USERNAME
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = SAP_PASSWORD
        session.findById("wnd[0]").sendVKey(0) # Enter
        
        time.sleep(3)
        log_message("✅ Đăng nhập SAP thành công!", level="SUCCESS")
        return session

    except Exception as e:
        log_message(f"🛑 Đăng nhập thất bại: {e}", level="CRITICAL")
        # Nếu lỗi nặng, dọn dẹp sạch sẽ để không để lại icon ma
        kill_all_sap_processes()
        sys.exit(1)


# 🚀 HÀM ĐIỀN THAM SỐ & CHẠY (Bản Fix lỗi ZPP04)
def run_tcode_and_fill_selections(session, config, dummy_wait=0):
    tcode = config['tcode']
    params = config['params']
    log_message(f"Bắt đầu chạy {tcode} ({config['name']})...")
    
    try:
        session.StartTransaction(tcode)
        time.sleep(2)

        # ---------------------------
        # 1. ĐIỀN THAM SỐ
        # ---------------------------
        if tcode in ["ZPP04A", "ZPP04"]:
            session.findById(PLANT_INPUT_ID_ZPP04A).text = params.get("PLANT_VALUE")
            if params.get("STORAGE_LOCATIONS_LIST"):
                session.findById(STORAGE_LOC_BUTTON_ID).press()
                time.sleep(1)
                for i, loc in enumerate(params["STORAGE_LOCATIONS_LIST"]):
                    session.findById(f"{MULTI_SELECT_TABLE_PATH}/{MULTI_SELECT_INPUT_BASE}[0,{i}]").text = loc
                session.findById("wnd[1]/tbar[0]/btn[8]").press()
                time.sleep(1)
        
        elif tcode == "ZBC04B":
            session.findById(DATE_FROM_ID_ZBC04B).text = params["DATE_FROM"]
            session.findById(DATE_TO_ID_ZBC04B).text = params["DATE_TO"]
            session.findById(PLANT_ID_ZBC04B).text = params["PLANT_VALUE"]
            if params.get("PRODUCT_GROUP_LIST"):
                session.findById(PRODUCT_GROUP_BUTTON_ID).press()
                time.sleep(1)
                for i, px in enumerate(params["PRODUCT_GROUP_LIST"]):
                    # Điền vào bảng chọn nhiều
                    session.findById(f"{MULTI_SELECT_TABLE_PATH}/{MULTI_SELECT_INPUT_BASE}[0,{i}]").text = px
                session.findById("wnd[1]/tbar[0]/btn[8]").press() # F8 (Copy)
                time.sleep(1)
            else:
                # Logic cũ (giữ lại để tương thích với NM2 nếu chỉ chạy 1 giá trị)
                session.findById(PRODUCT_GROUP_ID).text = params.get("PRODUCT_GROUP_VALUE", "")
            if params.get("UNCHECK_L1_L2"):
                session.findById(L1_CHECKBOX_ID).selected = False
                session.findById(L2_CHECKBOX_ID).selected = False

        elif tcode == "ZSD04A":
            session.findById(DATE_FROM_ID_ZSD04A).text = params["DATE_FROM"]
            session.findById(DATE_TO_ID_ZSD04A).text = params["DATE_TO"]
            if params.get("ORDER_TYPES_LIST"):
                session.findById(ORDER_TYPE_BUTTON_ID).press()
                time.sleep(1)
                for i, otype in enumerate(params["ORDER_TYPES_LIST"]):
                    try: session.findById(f"{MULTI_SELECT_TABLE_PATH}/{MULTI_SELECT_INPUT_BASE}[0,{i}]").text = otype
                    except: break
                session.findById("wnd[1]/tbar[0]/btn[8]").press()
                time.sleep(1)

        # ---------------------------
        # 2. THỰC THI (F8)
        # ---------------------------
        log_message("Đang nhấn F8...")
        session.findById("wnd[0]").sendVKey(8)
        
        # ---------------------------
        # 3. XỬ LÝ ZPP04 & CHECK NO DATA
        # ---------------------------
        
        # Nếu là ZPP04: Chờ Popup -> Đóng -> Back -> CHỜ TIẾP
        if tcode == "ZPP04":
            log_message("⚠️ Đang xử lý ZPP04 (Đóng Popup -> Back)...", level="WARN")
            wait_for_sap_ready(session, max_wait_minutes=5)
            
            # Đóng Popup (wnd[1])
            try:
                if session.Children.Count > 1:
                    session.findById("wnd[1]").close()
                    time.sleep(1)
            except: pass

            # Bấm Back (btn[3])
            try:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
                log_message("-> Đã bấm nút Back. Đang chờ màn hình kết quả hiện ra...", level="INFO")
                
                # 🔥 QUAN TRỌNG: Chờ 3 giây để màn hình chuyển từ Log về Grid kết quả
                time.sleep(3) 
                
                # Check lại lần nữa xem SAP có bận không
                wait_for_sap_ready(session, max_wait_minutes=2)
            except: pass

        else:
            # Các T-code khác
            wait_for_sap_ready(session, max_wait_minutes=20)
        
        # 4. KIỂM TRA "NO DATA" (Tránh lỗi Export failed)
        try:
            status = session.findById("wnd[0]/sbar").Text
            # Nếu status bar báo không có dữ liệu -> Ném lỗi để dừng Export
            if "No data" in status or "Không tìm thấy" in status or "No list generated" in status: 
                raise Exception("NO_DATA_FOUND")
            
            # Kiểm tra xem có đang ở màn hình nhập liệu không (Nếu F8 xong mà vẫn ở màn hình cũ nghĩa là lỗi/ko có data)
            # Dấu hiệu nhận biết màn hình nhập liệu: Có nút Execute (btn[8])
            # (Logic này mang tính tương đối, dùng để chặn lỗi)
            try:
                session.findById("wnd[0]/tbar[1]/btn[8]") # Thử tìm nút Execute
                # Nếu tìm thấy nút Execute -> Nghĩa là vẫn đang ở màn hình nhập liệu -> Không Export được
                log_message("Cảnh báo: Vẫn đang ở màn hình nhập liệu (có thể do ko có data).", level="WARN")
                raise Exception("NO_DATA_FOUND")
            except:
                pass # Nếu ko thấy nút Execute -> Nghĩa là đã sang màn hình kết quả -> OK

        except Exception as e:
            if str(e) == "NO_DATA_FOUND": raise e

    except Exception as e:
        # Nếu là lỗi No Data thì ném ra ngoài để Main Loop xử lý nhẹ nhàng
        if str(e) == "NO_DATA_FOUND": raise e
        
        log_message(f"Lỗi khi chạy T-Code: {e}", level="ERROR")
        raise

# 📤 EXPORT & SAVE
def export_data_to_excel(session, output_filename, custom_dir, menu_export_path):
    log_message("Đang Export...")
    try: session.findById(menu_export_path).select()
    except Exception as e: raise Exception(f"Menu Export Failed: {e}")

    time.sleep(2)
    try:
        # Xử lý cửa sổ Save As
        if session.Children.Count > 1: # Kiểm tra có wnd[1]
             # Nút Unconverted
            try: session.findById("wnd[1]/tbar[0]/btn[20]").press()
            except: pass
            
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = output_filename
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = custom_dir
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            log_message("Đang ghi file...", level="INFO")
            time.sleep(5)
            
            if os.path.exists(os.path.join(custom_dir, output_filename)):
                log_message(f"Export thành công: {output_filename}", level="SUCCESS")
            else:
                log_message("Cảnh báo: Không thấy file sau khi Save!", level="WARN")
    except Exception as e:
        raise Exception(f"Lỗi Save Window: {e}")

# ===============================
# MAIN
# ===============================
def main_sequence():
    # 1. TÍNH TOÁN NGÀY THÁNG
    date_map = calculate_dynamic_dates()
    for config in TASK_CONFIGS:
        for k, v in config['params'].items():
            if isinstance(v, str) and v.startswith("{"): 
                config['params'][k] = date_map.get(v.strip("{}"), v)

    # 2. CHUẨN BỊ MÔI TRƯỜNG (Sử dụng try...finally để đảm bảo dọn dẹp)
    try:
        if not os.path.isdir(CUSTOM_DIR): 
            os.makedirs(CUSTOM_DIR)
        
        # Đóng toàn bộ các tác vụ cũ để tránh xung đột
        force_close_process("excel.exe") 
        kill_all_sap_processes() # Hàm này cần bổ sung "sapgui.exe" như tôi đã gợi ý

        # 3. KHỞI TẠO KẾT NỐI MỚI (Start Fresh)
        sap_session = sap_login_and_connect()
        
        for config in TASK_CONFIGS:
            print("\n" + "="*60)
            try:
                path = os.path.join(CUSTOM_DIR, config['output_filename'])
                if os.path.exists(path): 
                    os.remove(path)

                # Thực thi T-Code
                run_tcode_and_fill_selections(sap_session, config)
                
                # Export dữ liệu
                export_data_to_excel(sap_session, config['output_filename'], CUSTOM_DIR, config['menu_export_path'])
                
                # Tắt Excel ngay nếu là ZSD04A (để giải phóng file)
                if config['tcode'] == "ZSD04A": 
                    force_close_process("excel.exe")

            except Exception as e:
                if "NO_DATA_FOUND" in str(e): 
                    log_message(f"Bỏ qua {config['name']} (Không có dữ liệu).", level="WARN")
                else: 
                    log_message(f"❌ Lỗi tại Task {config['name']}: {e}", level="ERROR")
            
            # Quay về màn hình chính sau mỗi task bằng lệnh /n
            try:
                sap_session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                sap_session.findById("wnd[0]").sendVKey(0)
                time.sleep(1)
            except: 
                pass

        log_message("🏁 Đã hoàn thành danh sách tác vụ.", level="SUCCESS")

    except Exception as fatal_e:
        log_message(f"🛑 Lỗi nghiêm trọng trong quá trình chạy: {fatal_e}", level="CRITICAL")
    finally:
        # 4. DỌN DẸP CUỐI CÙNG (LUÔN CHẠY)
        # Thử đóng SAP êm ái trước, nếu không được sẽ Kill sạch
        try:
            if 'sap_session' in locals():
                graceful_exit_sap(sap_session) # Sử dụng lệnh /nex
        except:
            pass
            
        log_message("🧹 Đang thực hiện dọn dẹp tài nguyên hệ thống...", level="INFO")
        kill_all_sap_processes() 
        force_close_process("excel.exe")
        log_message("✅ Hệ thống đã sạch sẽ. Sẵn sàng cho lần chạy sau.", level="SUCCESS")

if __name__ == "__main__":
    try:
        main_sequence()
    except Exception as e:
        log_message(f"🛑 Lỗi Fatal (Dừng chương trình): {e}", level="CRITICAL")
        # Nếu crash giữa chừng, cũng nên kill SAP để lần sau chạy không bị lỗi
        kill_all_sap_processes()