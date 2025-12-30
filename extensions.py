from flask_caching import Cache
import os
# Khởi tạo một đối tượng cache rỗng
# Chúng ta sẽ liên kết nó với app của bạn ở Bước 3
cache_dir = os.path.join(os.getcwd(), 'flask_cache_data')
cache = Cache(config={
    'CACHE_TYPE': 'FileSystemCache',   # <--- Đổi loại cache
    'CACHE_DIR': cache_dir,            # <--- Nơi lưu file
    'CACHE_DEFAULT_TIMEOUT': 86400,      # 5 phút
    'CACHE_THRESHOLD': 500             # Lưu tối đa 500 file
})