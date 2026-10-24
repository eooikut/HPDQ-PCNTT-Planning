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
import time

from dotenv import load_dotenv
from flask import Blueprint, jsonify, render_template, session

from auth.decorator import login_required

# Biến môi trường của máy (nếu có) được ưu tiên hơn tệp.
load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env.ai"))

ai_bp = Blueprint("ai_bp", __name__)

QUYEN_AI = ("ai_hoi_dap", "xem_bao_cao", "chot_bao_cao")


def _cau_hinh():
    return os.getenv("WEB_AI_SECRET", "").strip(), os.getenv("WEB_AI_API", "").strip().rstrip("/")


def _co_quyen(ten):
    return session.get("role") == "admin" or ten in session.get("permissions", [])


def _b64(b):
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode()


def ky_token(khoa, giay=300):
    """Token = base64url(JSON) + "." + base64url(HMAC-SHA256). Cùng cách với src/@trolyweb/xac-thuc.ts."""
    now = int(time.time())
    nd = {
        "sub": session.get("username"),
        "vai_tro": session.get("role"),
        "quyen": [q for q in session.get("permissions", []) if q in QUYEN_AI],
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
