from flask import Blueprint, render_template, request, redirect, url_for, flash, session
from werkzeug.security import check_password_hash
from db import engine
from db_utils import log_activity
from sqlalchemy import text

auth_bp = Blueprint('auth', __name__, template_folder='templates')

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()

        if not username or not password:
            flash('Vui lòng nhập đầy đủ tên đăng nhập và mật khẩu', 'danger')
            return redirect(url_for('auth.login'))

        # Sử dụng một khối `with` duy nhất cho tất cả các thao tác DB
        with engine.begin() as conn:
            query = text("SELECT id, username, password_hash, users.role, status FROM users WHERE username = :username")
            result = conn.execute(query, {"username": username}).mappings().fetchone()

            if result and check_password_hash(result['password_hash'], password):
                # Kiểm tra tài khoản có bị vô hiệu hóa không
                if result.get('status') != 1:
                    flash('Tài khoản của bạn đã bị vô hiệu hóa. Vui lòng liên hệ quản trị viên.', 'danger')
                    return redirect(url_for('auth.login'))
    
                # Lưu session
                session['user_id'] = result['id']
                session['username'] = result['username']
                session['role'] = result['role'].strip()
    
                # Lấy danh sách quyền của user và lưu vào session
                permissions_result = conn.execute(text("SELECT permission_name FROM user_permissions WHERE user_id = :uid"), {"uid": result['id']}).fetchall()
                session['permissions'] = [p[0] for p in permissions_result]
    
                # Cập nhật last_login
                update_query = text("UPDATE users SET last_login = GETDATE() WHERE id = :id")
                conn.execute(update_query, {"id": result['id']})
    
                # Ghi nhật ký đăng nhập thành công
                log_activity('login_success', user_id=result['id'], username=result['username'], ip_address=request.remote_addr)
                
                # --- LOGIC CHUYỂN HƯỚNG THÔNG MINH ---
                # 1. Admin luôn được chuyển đến trang mặc định
                if session['role'] == 'admin':
                    flash('Đăng nhập với quyền Admin thành công', 'success')
                    return redirect(url_for('dashboard_bp.dashboard'))

                # 2. User sẽ được chuyển đến trang đầu tiên họ có quyền
                # Danh sách ưu tiên các quyền và route tương ứng
                permission_routes = {
                    'view_ship_schedule': 'dashboard_bp.dashboard',
                    'view_lsx_report': 'reportlsx.lsx_all',
                    'view_order': 'report.order',
                    'view_customer': 'so.so_all',
                    'view_coil_id': 'idcuonbo_bp.idcuonbo',
                    'manage_lsx': 'lsx.danhsach_lsx',
                    'manage_users': 'user.list_users',
                    'upload_files': 'upload.upload_files',
                }

                for perm, route in permission_routes.items():
                    if perm in session['permissions']:
                        flash('Đăng nhập thành công', 'success')
                        return redirect(url_for(route))

                # 3. Nếu user không có quyền nào, hiển thị trang thông báo
                return render_template('no_permissions.html')

        # Nếu result không tồn tại hoặc sai mật khẩu
        log_activity('login_failure', username=username, ip_address=request.remote_addr, details="Sai tên đăng nhập hoặc mật khẩu")
        flash('Sai tên đăng nhập hoặc mật khẩu', 'danger')
        return redirect(url_for('auth.login'))

    return render_template('login.html')


@auth_bp.route('/logout')
def logout():
    user_id = session.get('user_id')
    username = session.get('username')
    # Ghi nhật ký đăng xuất
    log_activity('logout', user_id=user_id, username=username, ip_address=request.remote_addr)
    session.pop('user_id', None)
    session.pop('username', None)
    session.pop('role', None)
    flash('Đã đăng xuất', 'success')
    return redirect(url_for('auth.login'))
