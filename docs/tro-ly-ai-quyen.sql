-- Thêm 3 quyền của trợ lý AI để trang Quản lý người dùng cấp được (docs/TRO-LY-AI.md, bước 3).
-- Chạy trên CSDL của web Planning. Chạy lại nhiều lần không sao.
-- Nếu bảng permissions có thêm cột bắt buộc ngoài name, description thì thêm giá trị cho cột đó.

IF NOT EXISTS (SELECT 1 FROM permissions WHERE name = N'ai_hoi_dap')
    INSERT INTO permissions (name, description) VALUES (N'ai_hoi_dap', N'Trợ lý AI: hỏi đáp số liệu');

IF NOT EXISTS (SELECT 1 FROM permissions WHERE name = N'xem_bao_cao')
    INSERT INTO permissions (name, description) VALUES (N'xem_bao_cao', N'Trợ lý AI: xem, tải báo cáo ngày, tuần');

IF NOT EXISTS (SELECT 1 FROM permissions WHERE name = N'chot_bao_cao')
    INSERT INTO permissions (name, description) VALUES (N'chot_bao_cao', N'Trợ lý AI: chốt bản báo cáo');

-- Cấp nhanh cho một tài khoản (thay ten_dang_nhap). Có thể cấp bằng trang Quản lý người dùng thay cho đoạn này.
-- Người được cấp phải đăng xuất rồi đăng nhập lại thì quyền mới có hiệu lực.
/*
DECLARE @uid INT = (SELECT id FROM users WHERE username = N'ten_dang_nhap');
INSERT INTO user_permissions (user_id, permission_name)
SELECT @uid, p.name FROM permissions p
WHERE p.name IN (N'ai_hoi_dap', N'xem_bao_cao')
  AND NOT EXISTS (SELECT 1 FROM user_permissions up WHERE up.user_id = @uid AND up.permission_name = p.name);
*/
