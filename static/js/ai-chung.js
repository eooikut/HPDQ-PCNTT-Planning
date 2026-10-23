/*
 * Dùng chung cho khung hỏi đáp và trang báo cáo AI: lấy token từ /ai/token, gọi trợ lý, hiển thị markdown của bảng.
 * Không phụ thuộc thư viện ngoài. Mọi chữ từ máy chủ đều được thoát HTML trước khi hiển thị.
 */
(function () {
  "use strict";
  var cache = null;

  function layToken(lamMoi) {
    if (!lamMoi && cache && cache.het_han * 1000 - Date.now() > 30000) return Promise.resolve(cache);
    return fetch((window.HPDQ_AI && window.HPDQ_AI.token) || "/ai/token", { credentials: "same-origin", headers: { Accept: "application/json" } })
      .then(function (r) {
        var ct = r.headers.get("content-type") || "";
        if (ct.indexOf("application/json") < 0) throw new Error("Phiên đăng nhập đã hết. Anh/chị tải lại trang để đăng nhập.");
        return r.json().then(function (j) {
          if (!r.ok) throw new Error(j.loi || "Không lấy được quyền dùng trợ lý.");
          cache = j;
          return j;
        });
      });
  }

  /** Gọi trợ lý (GET/POST JSON); hết hạn token thì lấy token mới và thử lại một lần. */
  function goi(duong, tuyChon, daThu) {
    tuyChon = tuyChon || {};
    return layToken(daThu).then(function (tk) {
      var h = Object.assign({ Authorization: "Bearer " + tk.token }, tuyChon.headers || {});
      return fetch(tk.api + duong, Object.assign({}, tuyChon, { headers: h })).then(function (r) {
        if (r.status === 401 && !daThu) return goi(duong, tuyChon, true);
        return r;
      });
    });
  }

  function thoat(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function dong(s) {
    return thoat(s).replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
  }

  function tachO(h) {
    return h.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map(function (o) { return o.trim(); });
  }

  /**
   * Markdown -> HTML: bảng, danh sách (- * hoặc 1.), tiêu đề (#), chữ đậm, đoạn văn. Đủ cho câu trả lời của AI và bảng
   * dựng sẵn; mọi chữ đều được thoát HTML trước.
   */
  function mdSangHtml(md) {
    var hang = String(md || "")
      // AI hay viết danh sách trên một dòng: "... như sau: - A. - B." hoặc "1. X 2. Y" -> tách thành từng dòng.
      .replace(/([.:;])\s+(-|\*|\d{1,2}\.)\s+(?=\S)/g, "$1\n$2 ")
      .split(/\r?\n/);
    var ra = [];
    var doan = [];
    var ds = null;
    function dayDoan() {
      if (doan.length) ra.push("<p>" + doan.map(dong).join("<br>") + "</p>");
      doan = [];
    }
    function dayDs() {
      if (ds) ra.push("<" + ds.loai + ">" + ds.muc.map(function (m) { return "<li>" + dong(m) + "</li>"; }).join("") + "</" + ds.loai + ">");
      ds = null;
    }
    for (var i = 0; i < hang.length; i++) {
      var h = hang[i];
      var li = h.match(/^\s*(?:([-*•])|(\d{1,2})[.)])\s+(.*)$/);
      if (li && !(h.trim().charAt(0) === "|")) {
        dayDoan();
        var loai = li[1] ? "ul" : "ol";
        if (ds && ds.loai !== loai) dayDs();
        if (!ds) ds = { loai: loai, muc: [] };
        ds.muc.push(li[3]);
        continue;
      }
      if (ds && h.trim()) dayDs();
      var td = h.match(/^\s*#{1,6}\s+(.*)$/);
      if (td) {
        dayDoan();
        ra.push('<p class="ai-td">' + dong(td[1].replace(/\*\*/g, "")) + "</p>");
        continue;
      }
      if (h.trim().charAt(0) === "|" && hang[i + 1] && /^\s*\|?\s*:?-{3,}/.test(hang[i + 1])) {
        dayDoan();
        var dau = tachO(h);
        var phai = tachO(hang[i + 1]).map(function (o) { return /:$/.test(o); });
        var html = '<div class="ai-bang-cuon"><table class="ai-bang"><thead><tr>';
        dau.forEach(function (d, k) { html += '<th class="' + (phai[k] ? "so" : "") + '">' + dong(d) + "</th>"; });
        html += "</tr></thead><tbody>";
        i += 2;
        while (i < hang.length && hang[i].trim().charAt(0) === "|") {
          var o = tachO(hang[i]);
          var dam = o.some(function (x) { return x.indexOf("**") === 0; });
          html += "<tr" + (dam ? ' class="tong"' : "") + ">";
          o.forEach(function (x, k) { html += '<td class="' + (phai[k] ? "so" : "") + '">' + dong(x) + "</td>"; });
          html += "</tr>";
          i++;
        }
        i--;
        ra.push(html + "</tbody></table></div>");
      } else if (h.trim()) {
        doan.push(h);
      } else {
        dayDoan();
        dayDs();
      }
    }
    dayDoan();
    dayDs();
    return ra.join("");
  }

  window.HPDQ_AI_CHUNG = { layToken: layToken, goi: goi, thoat: thoat, mdSangHtml: mdSangHtml };
})();
