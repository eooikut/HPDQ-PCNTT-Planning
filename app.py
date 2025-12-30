from flask import Flask, redirect, url_for, session,jsonify
import os
import logging
from datetime import datetime
from functools import wraps
from flask_wtf.csrf import CSRFProtect
# Import scheduler và engine SQLAlchemy
from ETLExceltoDB import start_scheduler
from db import engine
from storage_utils import load_metadata
from cheroot.wsgi import Server
from extensions import cache
from flask_compress import Compress
# ---------- Logger ----------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------- Hàm lấy ngày tạo file SO ----------
def get_sl_kho_created_at():
    """
    Lấy ngày tạo file Sales Order (SO) trong thư mục uploads.
    Trả về datetime hoặc None nếu chưa có file.
    """
    so_path = os.path.join("data_auto_update", "sanluong_nm2.xlsx")  # đổi tên nếu khác
    if os.path.exists(so_path):
        return datetime.fromtimestamp(os.stat(so_path).st_mtime)
    return None
def get_so_created_at():
    """
    Lấy ngày tạo file Sales Order (SO) trong thư mục uploads.
    Trả về datetime hoặc None nếu chưa có file.
    """
    so_path = os.path.join("data_auto_update", "so.xlsx")  # đổi tên nếu khác
    if os.path.exists(so_path):
        return datetime.fromtimestamp(os.stat(so_path).st_mtime)
    return None

# ---------- Decorator login_required ----------
def login_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('auth.login'))
        return f(*args, **kwargs)
    return decorated

# ---------- Flask App ----------
def create_app():
    app = Flask(__name__)
    cache.init_app(app)
    Compress(app)
    # Config
    app.config["UPLOAD_FOLDER"] = "uploads"
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    # Lấy secret key từ biến môi trường để bảo mật hơn trên production
    app.secret_key = os.getenv("SECRET_KEY", "default-secret-key-for-dev-environment")

    # Khởi tạo CSRF Protection - BẮT BUỘC BẬT TRÊN PRODUCTION
    CSRFProtect(app)

    # ---------- Đăng ký blueprint ----------
    from auth.routes import auth_bp
    from routes.upload import upload_bp
    from routes.lsx import lsx_bp
    from routes.Order import report_bp
    from routes.khachhang import so_bp
    from routes.lichtau import tau_bp
    from routes.reportlsx import reportlsx_bp
    from routes.idcuonbo import idcuonbo_bp
    from routes.users import user_bp
    from routes.dashboard import dashboard_bp
    from routes.dashboardso import dashboard_so_bp
    from routes.tools import tools_bp
    from routes.kho2d import kho2d_bp
    app.register_blueprint(tools_bp)
    app.register_blueprint(auth_bp)
    app.register_blueprint(upload_bp)
    app.register_blueprint(lsx_bp)
    app.register_blueprint(report_bp)
    app.register_blueprint(so_bp)
    app.register_blueprint(tau_bp)
    app.register_blueprint(reportlsx_bp)
    app.register_blueprint(idcuonbo_bp)
    app.register_blueprint(user_bp)
    app.register_blueprint(dashboard_bp)
    app.register_blueprint(dashboard_so_bp)
    app.register_blueprint(kho2d_bp)
    # ---------- Biến toàn cục cho template ----------
    @app.context_processor
    def inject_file_times():
        return {
            "so_created_at": get_so_created_at(),
            "sl_kho_created_at": get_sl_kho_created_at()
        }

    # ---------- API cho AJAX ----------
    @app.route('/api/file-times')
    def api_file_times():
        metadata = load_metadata()
        lichtau_time = None
        lichtau_name = None
        donhangchitiet_time = None 
        so_time = get_so_created_at()
        sl_time = get_sl_kho_created_at()

        def format_time(t):
            return t.strftime("%Y-%m-%d %H:%M:%S") if t else None
        
        for m in reversed(metadata):  # reversed để lấy bản mới nhất
            if m.get("type") == "lichtau" and not lichtau_time:
                lichtau_time = m.get("uploaded_at")
            elif m.get("type") == "donhangchitiet" and not donhangchitiet_time:
                donhangchitiet_time = m.get("uploaded_at")
        # Nếu file bị xóa hoặc chưa có → báo updating
        return jsonify({
            "status": "ok",
            "so_created_at": format_time(so_time),
            "sl_kho_created_at": format_time(sl_time),
            "so_status": "ok" if so_time else "updating",
            "sl_status": "ok" if sl_time else "updating",
            "lichtau_time": lichtau_time,
            "lichtau_name": "LỊCH TÀU",
            "donhangchitiet_time": donhangchitiet_time,
            "donhangchitiet_name": "ĐƠN HÀNG"
        })

    # ---------- Route home redirect login ----------
    @app.route('/')
    def home():
        return redirect(url_for('auth.login'))

    return app


# ---------- Main ----------
app = create_app()
if __name__ == "__main__":
    # Khởi động các tác vụ nền nếu cần
    start_scheduler() 
    app.run(host="0.0.0.0", port=5002, debug=True)