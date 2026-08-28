from flask import Blueprint, render_template, request, redirect, url_for, flash, session
from sqlalchemy import text
from werkzeug.security import generate_password_hash
import pandas as pd
from datetime import datetime
from db import engine
from db_utils import log_activity, save_df_to_db
from auth.decorator import login_required, permission_required

user_bp = Blueprint('user', __name__, template_folder='templates')


# --- Danh sách user ---
@user_bp.route('/users')
@permission_required('manage_users')
def list_users():
    with engine.connect() as conn:
        # Chỉ hiển thị những user đang hoạt động (status = 1)
        users = conn.execute(text("""
            SELECT id, username, full_name, role, created_at, last_login, status FROM users WHERE status = 1
        """)).mappings().all()
    return render_template('users.html', users=users)


# --- Tạo user ---
@user_bp.route('/users/create', methods=['GET', 'POST'])
@permission_required('manage_users')
def create_user():
    # Lấy danh sách tất cả các quyền để hiển thị trên form
    with engine.connect() as conn:
        all_permissions_result = conn.execute(text("SELECT name, description FROM permissions ORDER BY name")).mappings().all()
        all_permissions = [dict(p) for p in all_permissions_result]

    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        full_name = request.form['full_name']
        role = request.form['role']

        current_time = datetime.now()
        password_hash = generate_password_hash(password)
        with engine.begin() as conn:
            # Mặc định tài khoản mới tạo là active (status=1)
            # Sử dụng OUTPUT INSERTED.id để lấy ID vừa tạo một cách an toàn trong một câu lệnh
            result = conn.execute(text("""
                INSERT INTO users (username, password_hash, role, full_name, created_at, status)
                OUTPUT INSERTED.id
                VALUES (:u, :p, :r, :fn, :ca, 1)
            """), {"u": username, "p": password_hash, "r": role, "fn": full_name, "ca": current_time})
            new_user_id = result.scalar()

            # Gán quyền chi tiết nếu role là 'user'
            if role == 'user':
                assigned_permissions = request.form.getlist('permissions')
                if assigned_permissions:
                    new_perms_data = [{"user_id": new_user_id, "permission_name": p} for p in assigned_permissions]
                    df_perms = pd.DataFrame(new_perms_data)
                    df_perms.to_sql('user_permissions', conn, if_exists='append', index=False)

        flash("Tạo tài khoản thành công!", "success")
        log_activity('create_user', user_id=session.get('user_id'), username=session.get('username'), target_type='user', target_id=new_user_id, details=f"Tạo tài khoản mới: {username}", ip_address=request.remote_addr)
        return redirect(url_for('user.list_users'))

    return render_template('register.html', all_permissions=all_permissions)


# --- Sửa quyền user ---
@user_bp.route('/users/edit/<int:id>', methods=['GET', 'POST'])
@permission_required('manage_users')
def edit_user(id):
    # Lấy danh sách tất cả các quyền để hiển thị trên form
    all_permissions = []
    with engine.begin() as conn:
        # Ngăn admin tự hạ quyền hoặc tự vô hiệu hóa tài khoản của chính mình
        if id == session.get('user_id'):
            role = request.form.get('role')
            status = request.form.get('status')
            if (role and role != 'admin') or (status and status != '1'):
                flash("Bạn không thể tự hạ quyền hoặc vô hiệu hóa tài khoản của chính mình.", "danger")
                return redirect(url_for('user.list_users'))

        all_permissions_result = conn.execute(text("SELECT name, description FROM permissions ORDER BY name")).mappings().all()
        all_permissions = [dict(p) for p in all_permissions_result]

        if request.method == 'POST':
            role = request.form['role']
            full_name = request.form['full_name']
            status = request.form.get('status', 1) # Mặc định là 1 nếu không có
            
            # Cập nhật thông tin cơ bản
            conn.execute(text("UPDATE users SET role=:r, full_name=:fn, status=:st WHERE id=:i"), 
                         {"r": role, "fn": full_name, "st": status, "i": id})

            # Cập nhật quyền chi tiết
            # Xóa hết quyền cũ
            conn.execute(text("DELETE FROM user_permissions WHERE user_id = :uid"), {"uid": id})

            # Nếu role là 'user', thêm quyền mới. Nếu là 'admin', không thêm quyền nào.
            if role == 'user':
                assigned_permissions = request.form.getlist('permissions')
                if assigned_permissions:
                    new_perms_data = [{"user_id": id, "permission_name": p} for p in assigned_permissions]
                    df_perms = pd.DataFrame(new_perms_data)
                    df_perms.to_sql('user_permissions', conn, if_exists='append', index=False)

            flash("Cập nhật thông tin người dùng thành công!", "success")
            log_activity('edit_user', user_id=session.get('user_id'), username=session.get('username'), target_type='user', target_id=id, details=f"Cập nhật tài khoản ID {id}", ip_address=request.remote_addr)
            return redirect(url_for('user.list_users'))

        user = conn.execute(text("SELECT id, username, full_name, role, status FROM users WHERE id=:i"), {"i": id}).mappings().fetchone()
        user_permissions_result = conn.execute(text("SELECT permission_name FROM user_permissions WHERE user_id = :uid"), {"uid": id}).fetchall()
        user_permissions = [p[0] for p in user_permissions_result]
    return render_template('edit_user.html', user=user, all_permissions=all_permissions, user_permissions=user_permissions)


# --- Đặt lại mật khẩu (bởi Admin) ---
@user_bp.route('/users/reset_password/<int:id>', methods=['GET', 'POST'])
@permission_required('manage_users')
def reset_password(id):
    # Ngăn admin tự reset mật khẩu của chính mình qua chức năng này.
    if id == session.get('user_id'):
        flash("Bạn không thể tự đặt lại mật khẩu của chính mình qua chức năng này.", "warning")
        return redirect(url_for('user.list_users'))

    with engine.begin() as conn:
        user = conn.execute(text("SELECT id, username, full_name FROM users WHERE id = :id"), {"id": id}).mappings().fetchone()
        if not user:
            flash("Không tìm thấy người dùng.", "danger")
            return redirect(url_for('user.list_users'))

    if request.method == 'POST':
        new_password = request.form.get('new_password')
        confirm_password = request.form.get('confirm_password')

        if not new_password or new_password != confirm_password:
            flash("Mật khẩu mới không khớp hoặc không được để trống.", "danger")
            return render_template('reset_password.html', user=user)

        password_hash = generate_password_hash(new_password)
        with engine.begin() as conn:
            conn.execute(text("UPDATE users SET password_hash = :p WHERE id = :id"), {"p": password_hash, "id": id})
        
        flash(f"Đã đặt lại mật khẩu cho tài khoản '{user['username']}' thành công!", "success")
        log_activity('reset_password', user_id=session.get('user_id'), username=session.get('username'), target_type='user', target_id=id, details=f"Đặt lại mật khẩu cho user ID {id} ({user['username']})", ip_address=request.remote_addr)
        return redirect(url_for('user.list_users'))

    return render_template('reset_password.html', user=user)

# --- Xóa user ---
@user_bp.route('/users/delete/<int:id>', methods=['POST'])
@permission_required('manage_users')
def delete_user(id):
    # Ngăn admin tự xóa tài khoản của chính mình
    if id == session.get('user_id'):
        flash("Bạn không thể tự xóa tài khoản của chính mình.", "danger")
        return redirect(url_for('user.list_users'))

    with engine.begin() as conn:
        # Lấy username để ghi log trước khi xóa
        user_to_delete = conn.execute(text("SELECT username FROM users WHERE id=:i"), {"i": id}).scalar()
        # Thực hiện xóa mềm: cập nhật status = 0
        conn.execute(text("UPDATE users SET status = 0 WHERE id=:i"), {"i": id})

    flash(f"Đã xóa tài khoản '{user_to_delete}' thành công!", "success")
    log_activity('delete_user', user_id=session.get('user_id'), username=session.get('username'), target_type='user', target_id=id, details=f"Xóa tài khoản ID {id} ({user_to_delete})", ip_address=request.remote_addr)
    return redirect(url_for('user.list_users'))