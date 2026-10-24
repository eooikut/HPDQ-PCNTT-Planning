# Trợ lý AI trên web Planning

Nhánh `feat/tro-ly-ai`. Tài liệu cho người quản trị web Planning (máy 10.192.212.24): tính năng gồm gì, cài và cấu hình thế nào, kiểm tra ra sao, lỗi thường gặp.

## 1. Tính năng

| | Ở đâu | Quyền |
|---|---|---|
| **Hỏi đáp số liệu** bằng lời thường: đơn hàng, tàu, tồn kho, sản xuất, dừng sản xuất, bán hàng, chất lượng. Trả lời ngắn kèm bảng số gốc, liên kết tới trang đã lọc, câu hỏi gợi ý. Nút **Trang này** hỏi đúng bộ lọc đang chọn. | Nút nổi góc dưới bên phải, mọi trang | `ai_hoi_dap` |
| **Báo cáo ngày, tuần**: số dựng sẵn (không dùng AI), tải Excel | Menu Công cụ > Báo cáo ngày, tuần (`/bao-cao`) | `xem_bao_cao` |
| **Chốt báo cáo**: lưu bản đã chốt | Nút Chốt trên trang báo cáo | `chot_bao_cao` |

Trợ lý chỉ báo số và phép so bằng số, không nhận xét, không đề xuất. Số năng lượng (điện, than, khí, hơi) chưa có.

## 2. Cách chạy

```
Trình duyệt ──► planning.hoaphatdungquat.vn (nginx trên .24)
                  ├─ /ai/token, mọi trang  → Flask: kiểm đăng nhập + quyền, ký token 5 phút
                  └─ /troly-api/*          → https://mcp.hoaphatdungquat.vn/web/* (máy chủ MCP .32)
                                                 ├─ lấy số: MCP Server HPDQ
                                                 └─ viết câu trả lời: LLM nội bộ view360 (10.192.220.12)
```

- Trình duyệt chỉ nói chuyện với `planning.hoaphatdungquat.vn`, cùng nguồn với trang: không CORS, không cần mở mạng cho máy người dùng, chỉ dùng chứng chỉ của planning.
- Câu hỏi đi qua nginx thẳng sang máy chủ trợ lý, **không qua Flask**: web không giữ luồng waitress trong lúc AI trả lời (2-10 giây mỗi câu).
- Máy chủ trợ lý chỉ nhận yêu cầu đến từ IP 10.192.212.24. Dịch vụ trợ lý do nhóm AI P.CNTT vận hành; phía web chỉ cần **khối nginx**, **hai giá trị cấu hình** và **ba quyền**.
- Thiếu cấu hình thì trợ lý tự ẩn, web chạy như cũ.

Tệp thay đổi trong nhánh:

| Tệp | Nội dung |
|---|---|
| `routes/ai_bp.py` (mới) | `/ai/token`, `/bao-cao`, đọc cấu hình từ `.env.ai` |
| `app.py` | đăng ký `ai_bp` (2 hàng) |
| `templates/base.html` | nạp CSS, mục menu Báo cáo, include partial (chỉ khi có quyền) |
| `templates/partials/ai_panel.html`, `templates/ai/*` (mới) | nạp script khung chat, trang báo cáo |
| `static/js/ai-chung.js`, `static/js/ai-panel.js`, `static/css/ai-panel.css` (mới) | khung chat, trang báo cáo. Không dùng thư viện ngoài |
| `docs/nginx-tro-ly-ai.conf` (mới) | khối nginx `/troly-api/` |
| `.env.ai.example`, `docs/tro-ly-ai-quyen.sql`, `docs/TRO-LY-AI.md` (mới) | cấu hình mẫu, script thêm quyền, tài liệu này |
| `.gitignore` | bỏ qua `.env.ai` |

Không đổi bảng dữ liệu nghiệp vụ, không đổi trang cũ. Chỉ thêm 3 hàng vào bảng `permissions`.

## 3. Trước khi cài

Nhóm AI P.CNTT báo đã xong hai việc, kèm **khoá `WEB_AI_SECRET`** (gửi qua kênh riêng):

1. Dịch vụ trợ lý đã chạy trên máy chủ MCP .32, đường `/web/` đã mở cho IP 10.192.212.24.
2. Mạng từ 10.192.212.24 tới 10.192.212.32 cổng 443 đã thông.

Kiểm tra từ máy .24 (PowerShell):

```powershell
curl.exe -s -o NUL -w "%{http_code}`n" --max-time 10 https://mcp.hoaphatdungquat.vn/web/suc-khoe
```

Ra `401` là đúng (đã tới trợ lý, chỉ thiếu token). Ra `403`: máy chủ trợ lý chưa nhận IP của .24. Ra `000` hoặc treo: mạng chưa thông.

## 4. Cài trên máy web

1. **Lấy mã**: merge nhánh `feat/tro-ly-ai` vào nhánh đang chạy.

2. **Cấu hình**: chép `.env.ai.example` thành `.env.ai` ở thư mục gốc (cùng chỗ với `app.py`):
   ```
   WEB_AI_SECRET=<khoá nhận từ nhóm AI P.CNTT>
   WEB_AI_API=/troly-api
   ```
   > **Không ghi vào `.env`.** Repo này công khai và `.env` đang được commit, ai cũng đọc được khoá. `.env.ai` đã có trong `.gitignore`. Không dán khoá vào issue, pull request hay nhóm chat.
   >
   > Cách khác: đặt hai biến này thành biến môi trường của máy hoặc dịch vụ Windows; biến môi trường được ưu tiên hơn tệp.

3. **Nginx**: chép khối trong `docs/nginx-tro-ly-ai.conf` vào khối `server` của `planning.hoaphatdungquat.vn` (cùng chỗ với `location /` đang chuyển vào Flask), rồi:
   ```
   nginx -t
   nginx -s reload
   ```

4. **Thêm quyền**: chạy `docs/tro-ly-ai-quyen.sql` trên CSDL của web (thêm `ai_hoi_dap`, `xem_bao_cao`, `chot_bao_cao` vào bảng `permissions`; chạy lại nhiều lần không sao).

5. **Cấp quyền**: trang Quản lý người dùng > sửa người dùng > tích các quyền trên. Tài khoản admin có sẵn mọi quyền. Người được cấp **đăng xuất rồi đăng nhập lại** (quyền nạp vào phiên lúc đăng nhập).

6. **Khởi động lại web** (waitress lưu đệm template, phải khởi động lại mới thấy `base.html` mới).

7. **Đồng giờ máy**: token có hạn 5 phút, tính theo giờ máy. Máy web và máy trợ lý lệch nhau quá 5 phút thì mọi câu hỏi bị từ chối. Bật đồng bộ giờ (NTP).

## 5. Kiểm tra sau khi bật

Đăng nhập bằng tài khoản có quyền:

| # | Làm | Đúng khi |
|---|---|---|
| 1 | Mở `https://planning.hoaphatdungquat.vn/troly-api/suc-khoe` | Trả JSON "Phiên trợ lý hết hạn hoặc không hợp lệ..." (mã 401): nginx đã chuyển đúng. Ra trang 404 của web: khối nginx chưa có hiệu lực |
| 2 | Mở `https://planning.hoaphatdungquat.vn/ai/token` | Trả JSON có `token` và `"api": "/troly-api"` |
| 3 | Mở trang bất kỳ | Có nút tròn góc dưới bên phải; bấm mở khung "Trợ lý Planning" có lời chào và câu gợi ý |
| 4 | Gõ "xin chào" | Trả lời ngay (dưới 1 giây) |
| 5 | Gõ "Tồn kho thép cuộn hiện tại?" | Thấy các bước chạy (Lấy số, Viết câu trả lời) và chữ hiện dần; sau 5-10 giây có câu trả lời, bảng số, mốc "Số liệu tính đến" |
| 6 | Menu Công cụ > Báo cáo ngày, tuần | Có 4 mục kèm bảng; bấm tải Excel được tệp |
| 7 | Đăng nhập tài khoản **không** có quyền | Không có nút trợ lý, không có mục Báo cáo; mở `/bao-cao` ra trang 403 |

Ở bước 5, nếu chữ không hiện dần mà hiện cả đoạn một lần ở cuối: thiếu `proxy_buffering off` trong khối nginx.

**Tắt nhanh** khi cần: để trống `WEB_AI_API` trong `.env.ai`, khởi động lại web. Trợ lý ẩn, mọi trang khác chạy như cũ.

## 6. Lỗi thường gặp

| Người dùng thấy | Nguyên nhân | Xử lý |
|---|---|---|
| Không có nút trợ lý | Chưa có `.env.ai` (hoặc thiếu một giá trị), chưa cấp quyền `ai_hoi_dap`, chưa đăng nhập lại sau khi cấp, chưa khởi động lại web | Kiểm lần lượt mục 4 |
| "Anh/chị chưa có quyền dùng trợ lý AI." | Tài khoản chưa có `ai_hoi_dap` hoặc `xem_bao_cao` | Mục 4 bước 5 |
| "Phiên đăng nhập đã hết. Anh/chị tải lại trang để đăng nhập." | Phiên web hết hạn | Tải lại trang, đăng nhập |
| "Phiên trợ lý hết hạn hoặc không hợp lệ. Anh/chị tải lại trang." ở **mọi** câu hỏi | Khoá hai bên khác nhau, hoặc giờ hai máy lệch quá 5 phút | So `WEB_AI_SECRET` với nhóm AI P.CNTT; bật NTP |
| "Trợ lý AI chưa được cấu hình xong trên máy chủ web (mã 404)" | Khối nginx chưa có hiệu lực, hoặc sai đường `WEB_AI_API` | Mục 4 bước 2, 3 |
| "...(mã 403)" | Máy chủ trợ lý chưa nhận IP của .24 (ví dụ nginx đi ra bằng IP khác) | Báo nhóm AI P.CNTT kèm IP đi ra của máy web |
| "Trợ lý AI tạm thời chưa kết nối được (mã 502, 503, 504)" | Máy .24 không tới được .32, hoặc dịch vụ trợ lý đang khởi động lại | Kiểm lệnh ở mục 3; kéo dài thì báo nhóm AI P.CNTT |
| "Trợ lý đang trả lời câu trước của anh/chị." | Mỗi người một câu cùng lúc | Chờ câu trước xong |
| "Trợ lý AI chưa phản hồi. Anh/chị thử lại sau ít phút..." hoặc ghi chú "chưa có câu dẫn; số trong bảng là số gốc" | LLM quá tải hoặc không chạy | Báo nhóm AI P.CNTT. Bảng số vẫn đúng |
| "Chưa lấy được số liệu cho câu hỏi này." hoặc "...: chưa lấy được số" | MCP Server không trả được số | Báo nhóm AI P.CNTT |
| Trả lời chậm hơn 20 giây | Nhiều người hỏi cùng lúc (khung chat hiện "chờ lượt") hoặc công cụ lấy số chậm | Bình thường 2-10 giây. Thường xuyên chậm thì báo nhóm AI P.CNTT |

Khi báo lỗi, gửi kèm ảnh màn hình, giờ hỏi và tên đăng nhập.

## 7. Bảo mật

- Token có hạn 5 phút, ký bằng HMAC-SHA256, chứa tên đăng nhập, vai trò và 3 quyền AI. Khoá không nằm trong mã, không gửi xuống trình duyệt.
- Máy chủ trợ lý chỉ nhận yêu cầu từ máy web .24, chỉ đọc số (không ghi CSDL Planning), mỗi người một câu cùng lúc. Khoá LLM chỉ nằm ở máy chủ trợ lý.
- Chữ từ máy chủ đều được thoát HTML trước khi hiển thị.
- Tải Excel đưa token vào địa chỉ tải (`?t=`), hết hạn sau 5 phút.
