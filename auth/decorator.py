from functools import wraps
from flask import session, redirect, url_for, flash, current_app,render_template

def login_required(f):
    """
    Đảm bảo người dùng đã đăng nhập.
    """
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash("Bạn cần đăng nhập để truy cập trang này.", "warning")
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