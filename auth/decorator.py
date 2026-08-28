from functools import wraps
from flask import session, redirect, url_for, flash, current_app,render_template

from db import engine 
from sqlalchemy import text

def login_required(f):
    """
    Đảm bảo người dùng đã đăng nhập và phiên hợp lệ.
    """
    @wraps(f)
    def decorated_function(*args, **kwargs):
        # Kiểm tra session tồn tại ở trình duyệt
        if 'user_id' not in session:
            flash("Bạn cần đăng nhập để truy cập trang này.", "warning")
            return redirect(url_for('auth.login'))
            
        # KIỂM TRA ĐỐI CHIẾU VỚI DATABASE
        with engine.begin() as conn:
            query = text("SELECT session_version, status FROM users WHERE id = :id")
            user = conn.execute(query, {"id": session['user_id']}).mappings().fetchone()
            
            # Nếu user không tồn tại, bị khóa, hoặc version ở trình duyệt lệch với DB
            if not user or user['status'] != 1 or user['session_version'] != session.get('session_version'):
                session.clear() # Hủy toàn bộ session hiện tại
                flash("Phiên đăng nhập đã hết hạn, quyền bị thay đổi hoặc tài khoản bị khóa. Vui lòng đăng nhập lại.", "danger")
                return redirect(url_for('auth.login'))
                
        return f(*args, **kwargs)
    return decorated_function
def admin_required(f):
    """
    Đảm bảo người dùng là admin.
    """
    @wraps(f)
    @login_required
    def decorated_function(*args, **kwargs):
        if session.get('role') != 'admin':
            flash("Bạn không có quyền truy cập chức năng này.", "danger")
            return redirect(url_for('tau_bp.lichtau')) # Hoặc trang chủ
        return f(*args, **kwargs)
    return decorated_function

def permission_required(permission):
    """
    Đảm bảo người dùng có một quyền cụ thể.
    """
    def decorator(f):
        @wraps(f)
        @login_required
        def decorated_function(*args, **kwargs):
            if 'permissions' in session and (permission in session['permissions'] or session.get('role') == 'admin'):
                return f(*args, **kwargs)
            flash("Bạn không có quyền truy cập chức năng này.", "danger")
            return render_template('403.html'), 403
        return decorated_function
    return decorator