"""
Trợ lý AI trên web Planning: hỏi đáp bằng lời và báo cáo ngày, tuần.

Phần AI chạy ở máy chủ MCP (module @trolyweb), dùng LLM nội bộ. Web chỉ làm hai việc:
  1. /ai/token: kiểm phiên đăng nhập + quyền, ký token ngắn hạn (5 phút) để trình duyệt gọi trợ lý.
     Câu hỏi đi qua nginx thẳng sang trợ lý (/troly-api/), không qua Flask, nên không giữ luồng waitress.
  2. /bao-cao: trang xem, chốt, tải báo cáo ngày, tuần.

Cấu hình: tệp .env.ai ở thư mục gốc (chép từ .env.ai.example) hoặc biến môi trường của máy chạy web.
KHÔNG ghi vào .env: repo này công khai và .env đang được commit.
  WEB_AI_SECRET  khoá ký token, giống WEB_AI_SECRET của dịch vụ trợ lý
  WEB_AI_API     đường gọi trợ lý, mặc định /troly-api (nginx của web chuyển sang máy chủ trợ lý, cùng nguồn)
Thiếu một trong hai thì trợ lý tự ẩn, web chạy như cũ. Hướng dẫn cài: docs/TRO-LY-AI.md.

Quyền (bảng permissions + user_permissions, thêm bằng docs/tro-ly-ai-quyen.sql; admin có tất cả):
  ai_hoi_dap     dùng khung hỏi đáp
  xem_bao_cao    xem, tải báo cáo
  chot_bao_cao   chốt bản báo cáo
"""
import base64
import hashlib
import hmac
import json
import os
import threading
import time

from dotenv import load_dotenv
from flask import Blueprint, jsonify, render_template, request, session
from sqlalchemy import bindparam, text

from auth.decorator import login_required, permission_required
from db import engine
from db_utils import log_activity

# Biến môi trường của máy (nếu có) được ưu tiên hơn tệp.
load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env.ai"))

ai_bp = Blueprint("ai_bp", __name__)

QUYEN_AI = ("ai_hoi_dap", "xem_bao_cao", "chot_bao_cao")
MO_TA_QUYEN = {
    "ai_hoi_dap": "Trợ lý AI: hỏi đáp số liệu",
    "xem_bao_cao": "Trợ lý AI: xem, tải báo cáo ngày, tuần",
    "chot_bao_cao": "Trợ lý AI: chốt bản báo cáo",
}

# Quyền AI đọc từ CSDL, nhớ 60 giây mỗi người: cấp hay thu hồi ở trang phân quyền có hiệu lực trong 1 phút, người dùng
# không phải đăng nhập lại (quyền trong phiên chỉ nạp lúc đăng nhập).
_BO_DEM_GIAY = 60
_bo_dem_quyen = {}
_khoa_bo_dem = threading.Lock()


def _cau_hinh():
    return os.getenv("WEB_AI_SECRET", "").strip(), os.getenv("WEB_AI_API", "").strip().rstrip("/")


def _quyen_ai_cua(uid):
    """Các quyền AI của một người dùng (bộ đệm 60 giây). CSDL lỗi thì dùng quyền trong phiên."""
    if not uid:
        return set()
    bay_gio = time.time()
    with _khoa_bo_dem:
        nho = _bo_dem_quyen.get(uid)
    if nho and nho[0] > bay_gio:
        return nho[1]
    try:
        with engine.connect() as conn:
            hang = conn.execute(text("SELECT permission_name FROM user_permissions WHERE user_id = :u"), {"u": uid}).fetchall()
        quyen = {h[0] for h in hang if h[0] in QUYEN_AI}
    except Exception:
        quyen = {p for p in session.get("permissions", []) if p in QUYEN_AI}
    with _khoa_bo_dem:
        _bo_dem_quyen[uid] = (bay_gio + _BO_DEM_GIAY, quyen)
    return quyen


def _co_quyen(ten):
    return session.get("role") == "admin" or ten in _quyen_ai_cua(session.get("user_id"))


def _dam_bao_quyen(conn):
    """Thêm 3 quyền AI vào bảng permissions nếu chưa có, để trang Quản lý người dùng cũng thấy (thay cho chạy SQL tay)."""
    co = {h[0] for h in conn.execute(text("SELECT name FROM permissions WHERE name IN :ds").bindparams(bindparam("ds", expanding=True)), {"ds": list(QUYEN_AI)}).fetchall()}
    for ten in QUYEN_AI:
        if ten not in co:
            conn.execute(text("INSERT INTO permissions (name, description) VALUES (:n, :d)"), {"n": ten, "d": MO_TA_QUYEN[ten]})


def _b64(b):
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode()


def ky_token(khoa, giay=300):
    """Token = base64url(JSON) + "." + base64url(HMAC-SHA256). Cùng cách với src/@trolyweb/xac-thuc.ts."""
    now = int(time.time())
    nd = {
        "sub": session.get("username"),
        "vai_tro": session.get("role"),
        "quyen": sorted(_quyen_ai_cua(session.get("user_id"))),
        "aud": "troly-web",
        "iat": now,
        "exp": now + giay,
    }
    phan = _b64(json.dumps(nd, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))
    chu_ky = _b64(hmac.new(khoa.encode("utf-8"), phan.encode("ascii"), hashlib.sha256).digest())
    return f"{phan}.{chu_ky}", nd["exp"]


@ai_bp.app_context_processor
def _bien_template():
    khoa, api = _cau_hinh()
    bat = bool(khoa and api and session.get("user_id"))
    return {
        "ai_hoi_dap": bat and _co_quyen("ai_hoi_dap"),
        "ai_bao_cao": bat and _co_quyen("xem_bao_cao"),
        "ai_chot_bao_cao": bat and _co_quyen("chot_bao_cao"),
    }


@ai_bp.route("/ai/token")
@login_required
def token():
    khoa, api = _cau_hinh()
    if not (khoa and api):
        return jsonify({"loi": "Trợ lý AI chưa được cấu hình."}), 503
    if not any(_co_quyen(q) for q in ("ai_hoi_dap", "xem_bao_cao")):
        return jsonify({"loi": "Anh/chị chưa có quyền dùng trợ lý AI."}), 403
    t, het_han = ky_token(khoa)
    return jsonify({"token": t, "het_han": het_han, "api": api})


@ai_bp.route("/bao-cao")
@login_required
def bao_cao():
    khoa, api = _cau_hinh()
    if not (khoa and api) or not _co_quyen("xem_bao_cao"):
        return render_template("403.html"), 403
    return render_template("ai/bao_cao.html")


@ai_bp.route("/ai/phan-quyen")
@permission_required("manage_users")
def phan_quyen():
    """Trang phân quyền trợ lý AI: người dùng đang hoạt động và 3 quyền AI của từng người."""
    khoa, api = _cau_hinh()
    with engine.begin() as conn:
        _dam_bao_quyen(conn)
        nguoi = conn.execute(text("SELECT id, username, full_name, role, last_login FROM users WHERE status = 1 ORDER BY username")).mappings().all()
        gan = conn.execute(
            text("SELECT user_id, permission_name FROM user_permissions WHERE permission_name IN :ds").bindparams(bindparam("ds", expanding=True)),
            {"ds": list(QUYEN_AI)},
        ).fetchall()
    theo_nguoi = {}
    for uid, ten in gan:
        theo_nguoi.setdefault(uid, set()).add(ten)
    ds = [
        {
            "id": n["id"],
            "username": n["username"],
            "full_name": n["full_name"] or "",
            "admin": (n["role"] or "").strip() == "admin",
            "last_login": n["last_login"].strftime("%d/%m/%Y %H:%M") if n["last_login"] else "",
            "quyen": sorted(theo_nguoi.get(n["id"], set())),
        }
        for n in nguoi
    ]
    cau_hinh = {"khoa": bool(khoa), "api": api}
    return render_template("ai/phan_quyen.html", nguoi_dung=ds, quyen_ai=QUYEN_AI, mo_ta=MO_TA_QUYEN, cau_hinh=cau_hinh)


@ai_bp.route("/ai/phan-quyen", methods=["POST"])
@permission_required("manage_users")
def luu_phan_quyen():
    """Bật, tắt quyền AI cho một hoặc nhiều người: {"user_ids": [..], "quyen": {"ai_hoi_dap": true, ...}}."""
    body = request.get_json(silent=True) or {}
    try:
        ids = sorted({int(x) for x in body.get("user_ids", [])})[:500]
    except (TypeError, ValueError):
        return jsonify({"loi": "Danh sách người dùng không hợp lệ."}), 400
    dat = {q: bool(v) for q, v in (body.get("quyen") or {}).items() if q in QUYEN_AI}
    if not ids or not dat:
        return jsonify({"loi": "Chưa chọn người dùng hoặc quyền."}), 400
    with engine.begin() as conn:
        _dam_bao_quyen(conn)
        # Chỉ người dùng đang hoạt động, không phải admin (admin đã có mọi quyền).
        hop_le = [
            h[0]
            for h in conn.execute(
                text("SELECT id FROM users WHERE id IN :ids AND status = 1 AND LTRIM(RTRIM(role)) <> 'admin'").bindparams(bindparam("ids", expanding=True)),
                {"ids": ids},
            ).fetchall()
        ]
        for uid in hop_le:
            for ten, bat in dat.items():
                conn.execute(text("DELETE FROM user_permissions WHERE user_id = :u AND permission_name = :q"), {"u": uid, "q": ten})
                if bat:
                    conn.execute(text("INSERT INTO user_permissions (user_id, permission_name) VALUES (:u, :q)"), {"u": uid, "q": ten})
    with _khoa_bo_dem:
        for uid in hop_le:
            _bo_dem_quyen.pop(uid, None)
    thay = ", ".join(f"{'cấp' if bat else 'thu hồi'} {ten}" for ten, bat in dat.items())
    try:
        log_activity(
            "ai_phan_quyen",
            user_id=session.get("user_id"),
            username=session.get("username"),
            target_type="user",
            target_id=hop_le[0] if len(hop_le) == 1 else None,
            details=f"Phân quyền trợ lý AI: {thay} cho {len(hop_le)} tài khoản (ID {', '.join(map(str, hop_le))})"[:1000],
            ip_address=request.remote_addr,
        )
    except Exception:
        pass
    return jsonify({"ok": True, "so_nguoi": len(hop_le), "user_ids": hop_le})
