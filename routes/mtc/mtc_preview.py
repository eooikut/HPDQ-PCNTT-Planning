from flask import Blueprint, render_template, request

# Khởi tạo Blueprint
mtc_preview_bp = Blueprint('mtc_preview_bp', __name__)

@mtc_preview_bp.route('/mtc_preview')
def preview_label():
    # Lấy tham số 'type' từ URL (mặc định là 'sni' nếu không nhập)
    cert_type = request.args.get('type', 'sni').lower()

    # Xác định class CSS tương ứng với 4 cases
    if cert_type == 'sni':
        cert_class = 'cert-sni'
    elif cert_type == 'ce':
        cert_class = 'cert-ce'
    elif cert_type == 'ms-right':
        cert_class = 'cert-ms-right'
    elif cert_type == 'ms-left':
        cert_class = 'cert-ms-left'
    else:
        cert_class = 'cert-sni' # Fallback an toàn

    # Render file HTML và truyền biến cert_class sang giao diện
    return render_template('mtc/mtc_preview.html', cert_class=cert_class)