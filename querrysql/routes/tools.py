from flask import Blueprint, render_template, request, flash
import pandas as pd
from caculator import run_calculation_tool # Import hàm logic ở trên

# Giả sử bạn có blueprint tên là 'tools_bp' hoặc tạo mới
tools_bp = Blueprint('tools', __name__, url_prefix='/tools')

@tools_bp.route('/production-calc', methods=['GET', 'POST'])
def production_calc():
    results = None
    warnings = None
    
    # 1. THIẾT LẬP GIÁ TRỊ MẶC ĐỊNH
    selected_grade = 'LOW_CARBON' 
    selected_factory = 'HRC1' # Mặc định là HRC1

    if request.method == 'POST':
        try:
            # 2. LẤY DỮ LIỆU TỪ FORM
            selected_grade = request.form.get('grade_type', 'LOW_CARBON')
            selected_factory = request.form.get('factory_type', 'HRC1') # Lấy nhà máy user chọn
            
            df = pd.DataFrame()
            
            # Xử lý Upload File
            if 'file' in request.files and request.files['file'].filename != '':
                df = pd.read_excel(request.files['file'])
            
            # Xử lý Nhập tay (Giữ nguyên logic cũ)
            elif 'manual_width' in request.form:
                widths = request.form.getlist('manual_width')
                thicks = request.form.getlist('manual_thick')
                masses = request.form.getlist('manual_mass')
                data = {
                    'Khổ rộng': [float(w) for w in widths if w],
                    'Chiều dày': [float(t) for t in thicks if t],
                    'Khối lượng': [float(m) for m in masses if m]
                }
                if len(data['Khổ rộng']) > 0:
                    df = pd.DataFrame(data)

            # 3. GỌI HÀM TÍNH TOÁN (TRUYỀN THÊM factory_code)
            if not df.empty:
                # [QUAN TRỌNG] Truyền selected_factory vào hàm
                calc_response = run_calculation_tool(df, selected_grade, factory_code=selected_factory)
                
                if isinstance(calc_response, dict) and 'error' in calc_response:
                    flash(calc_response['error'], 'danger')
                else:
                    results = calc_response.get('data', [])
                    warnings = calc_response.get('warnings', [])
                    
                    if warnings:
                        flash(f'Đã tính xong cho {selected_factory}. Có {len(warnings)} dòng cảnh báo.', 'warning')
                    else:
                        flash(f'Tính toán thành công: {selected_factory} - {selected_grade}', 'success')
            else:
                flash('Chưa có dữ liệu đầu vào.', 'warning')

        except Exception as e:
            flash(f'Lỗi hệ thống: {str(e)}', 'danger')

    # 4. TRẢ VỀ TEMPLATE KÈM BIẾN selected_factory ĐỂ GIỮ TRẠNG THÁI
    return render_template('production_calc.html', 
                           results=results, 
                           warnings=warnings, 
                           selected_grade=selected_grade,
                           selected_factory=selected_factory) # Nhớ truyền biến này