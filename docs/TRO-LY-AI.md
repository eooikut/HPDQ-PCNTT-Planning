# Trợ lý AI trên web Planning

Nhánh `feat/tro-ly-ai`. Tài liệu này dành cho người giữ repo web Planning và người vận hành máy chủ web: tính năng gồm gì, cần cài gì để chạy, kiểm tra thế nào, lỗi thường gặp.

## 1. Tính năng

| | Ở đâu | Quyền |
|---|---|---|
| **Hỏi đáp số liệu** bằng lời thường: đơn hàng, tàu, tồn kho, sản xuất, dừng sản xuất, bán hàng, chất lượng. Trả lời ngắn kèm bảng số gốc, liên kết tới trang đã lọc, câu hỏi gợi ý. Nút **Trang này** hỏi đúng bộ lọc đang chọn. | Nút nổi góc dưới bên phải, mọi trang | `ai_hoi_dap` |
| **Báo cáo ngày, tuần**: số dựng sẵn (không dùng AI), tải Excel | Menu Công cụ > Báo cáo ngày, tuần (`/bao-cao`) | `xem_bao_cao` |
| **Chốt báo cáo**: lưu bản đã chốt | Nút Chốt trên trang báo cáo | `chot_bao_cao` |

Trợ lý chỉ báo số và phép so bằng số, không nhận xét, không đề xuất. Số năng lượng (điện, than, khí, hơi) chưa có.

## 2. Cách chạy

```
Trình duyệt ──(1) GET /ai/token──► Web Planning (Flask)     kiểm đăng nhập + quyền, ký token 5 phút
     │
     └──(2) gọi thẳng, kèm token──► Dịch vụ trợ lý (WEB_AI_API)  ──► MCP Server HPDQ (lấy số)
                                                                 ──► LLM nội bộ (viết câu trả lời)
```

- Câu hỏi **không đi qua Flask**: web không giữ luồng waitress trong lúc AI trả lời (5-10 giây mỗi câu).
- Dịch vụ trợ lý do nhóm AI P.CNTT vận hành (module `@trolyweb` của MCP Server HPDQ). Phía web chỉ cần **hai giá trị cấu hình** và **ba quyền**.
- Thiếu cấu hình thì trợ lý tự ẩn, web chạy như cũ.

Tệp thay đổi trong nhánh:

| Tệp | Nội dung |
|---|---|
| `routes/ai_bp.py` (mới) | `/ai/token`, `/bao-cao`, đọc cấu hình từ `.env.ai` |
| `app.py` | đăng ký `ai_bp` (2 hàng) |
| `templates/base.html` | nạp CSS, mục menu Báo cáo, include partial (chỉ khi có quyền) |
| `templates/partials/ai_panel.html`, `templates/ai/*` (mới) | nạp script khung chat, trang báo cáo |
| `static/js/ai-chung.js`, `static/js/ai-panel.js`, `static/css/ai-panel.css` (mới) | khung chat, trang báo cáo. Không dùng thư viện ngoài |
| `.env.ai.example`, `docs/tro-ly-ai-quyen.sql`, `docs/TRO-LY-AI.md` (mới) | cấu hình mẫu, script thêm quyền, tài liệu này |
| `.gitignore` | bỏ qua `.env.ai` |

Không đổi bảng dữ liệu nghiệp vụ, không đổi trang cũ. Chỉ thêm 3 hàng vào bảng `permissions`.

## 3. Điều kiện trước khi bật

Nhóm AI P.CNTT chuẩn bị (mục 6), rồi gửi cho người vận hành web:

1. **Địa chỉ trợ lý** (`WEB_AI_API`), dạng `https://.../web`. Phải là **HTTPS** (trang Planning chạy HTTPS, trình duyệt chặn gọi HTTP) và máy người dùng trong LAN mở được.
2. **Khoá ký token** (`WEB_AI_SECRET`), gửi qua kênh riêng, không gửi qua nhóm chat, email chung hay commit.

## 4. Cài trên máy chạy web

1. **Lấy mã**: merge nhánh `feat/tro-ly-ai` vào nhánh đang chạy.

2. **Cấu hình**: chép `.env.ai.example` thành `.env.ai` ở thư mục gốc (cùng chỗ với `app.py`), điền hai giá trị ở mục 3.
   > **Không ghi vào `.env`.** Repo này công khai và `.env` đang được commit, ai cũng đọc được khoá. `.env.ai` đã có trong `.gitignore`.
   >
   > Cách khác: đặt `WEB_AI_SECRET`, `WEB_AI_API` thành biến môi trường của máy hoặc dịch vụ Windows; biến môi trường được ưu tiên hơn tệp.

3. **Thêm quyền**: chạy `docs/tro-ly-ai-quyen.sql` trên CSDL của web (thêm `ai_hoi_dap`, `xem_bao_cao`, `chot_bao_cao` vào bảng `permissions`; chạy lại nhiều lần không sao).

4. **Cấp quyền**: trang Quản lý người dùng > sửa người dùng > tích các quyền trên. Tài khoản admin có sẵn mọi quyền. Người được cấp **đăng xuất rồi đăng nhập lại** (quyền nạp vào phiên lúc đăng nhập).

5. **Khởi động lại web** (waitress lưu đệm template, phải khởi động lại mới thấy `base.html` mới).

6. **Đồng giờ máy**: token có hạn 5 phút, tính theo giờ máy. Máy web và máy trợ lý lệch nhau quá 5 phút thì mọi câu hỏi bị từ chối. Bật đồng bộ giờ (NTP) ở cả hai máy.

## 5. Kiểm tra sau khi bật

Đăng nhập bằng tài khoản có quyền:

| # | Làm | Đúng khi |
|---|---|---|
| 1 | Mở `https://planning.hoaphatdungquat.vn/ai/token` | Trả JSON có `token`, `api` đúng địa chỉ trợ lý |
| 2 | Mở trang bất kỳ | Có nút tròn góc dưới bên phải; bấm mở khung "Trợ lý Planning" có lời chào và câu gợi ý |
| 3 | Gõ "xin chào" | Trả lời ngay (dưới 1 giây) |
| 4 | Gõ "Tồn kho thép cuộn hiện tại?" | Thấy các bước chạy (Lấy số, Viết câu trả lời), sau 5-10 giây có câu trả lời, bảng số, mốc "Số liệu tính đến" |
| 5 | Menu Công cụ > Báo cáo ngày, tuần | Có 4 mục kèm bảng; bấm tải Excel được tệp |
| 6 | Đăng nhập tài khoản **không** có quyền | Không có nút trợ lý, không có mục Báo cáo; mở `/bao-cao` ra trang 403 |

Bước 1 hoặc 2 sai: xem mục 7. Bước 3 tới 5 sai: thường do dịch vụ trợ lý, báo nhóm AI P.CNTT kèm ảnh màn hình và giờ hỏi.

**Tắt nhanh** khi cần: xoá (hoặc để trống) `WEB_AI_API` trong `.env.ai`, khởi động lại web. Trợ lý ẩn, mọi trang khác chạy như cũ.

## 6. Dịch vụ trợ lý (nhóm AI P.CNTT)

Mã nguồn: `hpdq-mcp-nest`, module `src/@trolyweb`, route `/web/*`. Chạy cùng tiến trình MCP Server ở chế độ HTTP (mặc định bật, tắt bằng `WEB_AI=0`), hoặc chạy riêng `node dist/trolyweb.main.js` (`TROLY_WEB_PORT`, mặc định 8940).

| Biến | Bắt buộc | Ý nghĩa |
|---|---|---|
| `WEB_AI_SECRET` | có | Khoá ký token, giống hệt bên web. Tạo một lần, ví dụ `openssl rand -hex 32`. Trống thì mọi lời gọi `/web/*` bị từ chối |
| `WEB_AI_ORIGINS` | | Nguồn được gọi (CORS), cách nhau bằng dấu phẩy. Mặc định đã có `https://planning.hoaphatdungquat.vn` |
| `LLM_BASE_URL` | | Mặc định `https://10.192.220.12/v1` (vLLM trên view360-llm) |
| `LLM_API_KEY` | có | Khoá vLLM |
| `LLM_MODEL` | | Để trống thì lấy model đầu tiên của `/models` |
| `LLM_CA_FILE` | nên có | Chứng chỉ Caddy tự ký của view360-llm. Chỉ dùng `LLM_TLS_INSECURE=1` khi thử |
| `WEB_AI_DONG_THOI` | | Số câu gọi LLM cùng lúc, mặc định 3 (vLLM dùng chung với View360, Dify) |
| `WEB_AI_HET_GIO_MS` | | Hết giờ chờ LLM, mặc định 45000 |
| `TROLY_MCP_URL`, `TROLY_MCP_TOKEN` | | MCP lấy số, mặc định chính máy chủ này |

Yêu cầu mạng:

- Trình duyệt người dùng tới `WEB_AI_API` qua **HTTPS, chứng chỉ hợp lệ**. Ở gateway, mở `/web/*` **chỉ cho dải LAN**.
- Máy trợ lý tới được LLM `10.192.220.12:443` (đã kiểm từ máy MCP .31, .32).
- Nhật ký câu hỏi và phản hồi đúng/sai: `troly-web.jsonl`, `troly-web-phan-hoi.jsonl` trong `MCP_DATA_DIR`; bản báo cáo đã chốt: `MCP_DATA_DIR/bao-cao`.

## 7. Lỗi thường gặp

| Người dùng thấy | Nguyên nhân | Xử lý |
|---|---|---|
| Không có nút trợ lý | Chưa có `.env.ai` (hoặc thiếu một giá trị), chưa cấp quyền `ai_hoi_dap`, chưa đăng nhập lại sau khi cấp, chưa khởi động lại web | Kiểm lần lượt mục 4 |
| `/ai/token` trả "Trợ lý AI chưa được cấu hình." | Thiếu `WEB_AI_SECRET` hoặc `WEB_AI_API` | Mục 4 bước 2, rồi khởi động lại web |
| "Anh/chị chưa có quyền dùng trợ lý AI." | Tài khoản chưa có `ai_hoi_dap` hoặc `xem_bao_cao` | Mục 4 bước 4 |
| "Phiên đăng nhập đã hết. Anh/chị tải lại trang để đăng nhập." | Phiên web hết hạn | Tải lại trang, đăng nhập |
| "Phiên trợ lý hết hạn hoặc không hợp lệ. Anh/chị tải lại trang." ở **mọi** câu hỏi | Khoá hai bên khác nhau, hoặc giờ hai máy lệch quá 5 phút | So `WEB_AI_SECRET` hai bên; bật NTP |
| "Mất kết nối tới trợ lý." | Máy người dùng không tới được `WEB_AI_API`: sai địa chỉ, dùng HTTP, chứng chỉ không hợp lệ, gateway chặn, hoặc nguồn web chưa nằm trong `WEB_AI_ORIGINS` | Mở `WEB_AI_API` + `/suc-khoe` từ máy người dùng (đúng là trả 401, không phải lỗi mạng); xem tab Console của trình duyệt (F12) |
| "Trợ lý đang trả lời câu trước của anh/chị." | Mỗi người một câu cùng lúc | Chờ câu trước xong |
| "Trợ lý AI chưa phản hồi. Anh/chị thử lại sau ít phút..." hoặc ghi chú "chưa có câu dẫn; số trong bảng là số gốc" | LLM quá tải hoặc không chạy | Báo nhóm AI P.CNTT. Bảng số vẫn đúng |
| "Chưa lấy được số liệu cho câu hỏi này." hoặc "...: chưa lấy được số" | MCP Server không trả được số | Báo nhóm AI P.CNTT |
| Trả lời chậm hơn 20 giây | Nhiều người hỏi cùng lúc (khung chat hiện "chờ lượt") hoặc công cụ lấy số chậm | Bình thường 2-10 giây. Thường xuyên chậm thì báo nhóm AI P.CNTT |

## 8. Bảo mật

- Token chỉ có hạn 5 phút, ký bằng HMAC-SHA256, chứa tên đăng nhập, vai trò và 3 quyền AI. Khoá không nằm trong mã, không gửi xuống trình duyệt.
- Dịch vụ trợ lý chỉ đọc số (không ghi CSDL Planning), chỉ nhận nguồn trong `WEB_AI_ORIGINS`, mỗi người một câu cùng lúc.
- Chữ từ máy chủ đều được thoát HTML trước khi hiển thị.
- Tải Excel đưa token vào địa chỉ tải (`?t=`), hết hạn sau 5 phút.
