# Sửa lại trong file app.py của bạn
import os
from cheroot.wsgi import Server
from your_app_module import app # Thay 'your_app_module' bằng tên file chứa 'create_app()'

# ---------- BẮT ĐẦU CẤU HÌNH TỐI ƯU ----------

# 1. Tính toán số luồng dựa trên CPU, nhưng có điều chỉnh
# Server có 8 vCPU. Công thức (2*cores)+1 = 17.
# Tuy nhiên, để chừa tài nguyên cho RPA, chúng ta sẽ dùng một con số thấp hơn nhưng vẫn hiệu quả.
worker_threads = 12

# 2. Lấy logger từ ứng dụng của bạn (nếu có)
# Giả sử bạn có một logger đã cấu hình
try:
    from your_app_module import logger
    logger.info(f"🚀 Starting Cheroot server with {worker_threads} worker threads...")
except ImportError:
    print(f"--> Starting Cheroot server with {worker_threads} worker threads...")


# 3. Khởi tạo server Cheroot với cấu hình tối ưu
server = Server(
    ('0.0.0.0', 5001),                   # Host và port
    app,                                 # Ứng dụng Flask
    numthreads=worker_threads,           # Số luồng xử lý chính
    max_request_body_size=52428800,      # Giới hạn kích thước upload là 50MB
    request_queue_size=2048,             # Tăng hàng đợi kết nối để xử lý traffic đột biến
    timeout=60                           # Đóng các kết nối treo sau 60 giây
)

# ---------- KẾT THÚC CẤU HÌNH ----------

if __name__ == "__main__":
    # Các tác vụ khởi động của bạn (nếu có)
    # start_scheduler()
    
    try:
        server.start()
    except KeyboardInterrupt:
        print("\n--> Server is shutting down.")
        server.stop()