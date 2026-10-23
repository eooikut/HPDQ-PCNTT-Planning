/*
 * Trợ lý AI trên web Planning: nút nổi + khung chat nổi góc dưới phải (cùng phong cách chatbot View360).
 * Khung tự dựng phần tử, chỉ hiện khi lấy được token (/ai/token). Câu trả lời: vài câu do AI viết (đã kiểm số),
 * bảng số do máy chủ dựng sẵn, chip mở trang đã lọc, câu hỏi gợi ý tiếp. "Trang này" đọc bộ lọc đang chọn trên trang.
 * Hội thoại giữ trong sessionStorage nên chuyển trang (vd bấm "Mở Lịch tàu") không mất.
 */
(function () {
  "use strict";
  var C = window.HPDQ_AI_CHUNG;
  if (!C || window.__hpdqAiPanel) return;
  window.__hpdqAiPanel = true;

  /* Trang hỗ trợ "Trang này": mã trang + ô lọc (id phần tử) trên trang. Trang tự khai window.HPDQ_TRANG thì dùng của trang. */
  var TRANG_WEB = {
    "/lichtau": { ma: "lichtau", ten: "Lịch tàu", loc: { tau: "#filter-tau", nhamay: "#filter-nhamay", cangden: "#filter-cangden", sheetmonth: "#filter-sheetmonth" } },
    "/tauscheduler": { ma: "tauscheduler", ten: "Kế hoạch tàu", loc: {} },
    "/dashboard": { ma: "dashboard", ten: "Lịch tàu", loc: {} },
    "/so_all": { ma: "so_all", ten: "Tiến độ SO theo khách", loc: { customer: "#filter-customer", factory: "#filter-factory" } },
    "/idcuonbo": { ma: "idcuonbo", ten: "Danh sách cuộn", loc: { nha_may: "#filter-nhamay", trangthai: "#filter-trangthai", tu_ngay: "#tu-ngay", den_ngay: "#den-ngay", keyword: "#keyword" } },
    "/dashboard-inventory": { ma: "dashboard-inventory", ten: "Tồn kho và SO", loc: {} },
  };
  var NHAN_LOC = { tau: "tàu", nhamay: "nhà máy", nha_may: "nhà máy", cangden: "cảng đến", sheetmonth: "tháng", customer: "khách", factory: "nhà máy", trangthai: "trạng thái", tu_ngay: "từ", den_ngay: "đến", keyword: "từ khoá" };
  var KHOA_HT = "hpdq_ai_hoi_thoai";
  var KHOA_MO = "hpdq_ai_mo";
  var BANG_GON = 8; // bảng dài hơn thì thu gọn, có nút "Xem thêm"

  function trangHienTai() {
    if (window.HPDQ_TRANG) {
      var t = window.HPDQ_TRANG;
      return { ma: t.ma, ten: t.ten, loc: typeof t.loc === "function" ? t.loc() : t.loc || {} };
    }
    var p = window.location.pathname.replace(/\/+$/, "") || "/";
    var cfg = TRANG_WEB[p] || TRANG_WEB[Object.keys(TRANG_WEB).find(function (k) { return p.indexOf(k + "/") === 0; })];
    if (!cfg) return null;
    var loc = {};
    Object.keys(cfg.loc).forEach(function (k) {
      var el = document.querySelector(cfg.loc[k]);
      var v = el && el.value != null ? String(el.value).trim() : "";
      if (v) loc[k] = v;
    });
    return { ma: cfg.ma, ten: cfg.ten, loc: loc };
  }

  function tao(tag, lop, html) {
    var e = document.createElement(tag);
    if (lop) e.className = lop;
    if (html != null) e.innerHTML = html;
    return e;
  }

  function doc(khoa) {
    try { return sessionStorage.getItem(khoa); } catch (e) { return null; }
  }
  function ghi(khoa, gt) {
    try { sessionStorage.setItem(khoa, gt); } catch (e) { /* bỏ qua: chế độ riêng tư hoặc đầy */ }
  }

  /* ===== Dựng phần tử ===== */
  var fab = tao("button", "ai-fab", '<i class="fas fa-robot"></i><span class="ai-fab__nhan">AI</span>');
  fab.type = "button";
  fab.title = "Trợ lý Planning (AI)";
  fab.setAttribute("aria-label", "Mở trợ lý AI");
  fab.setAttribute("data-ai-mo", "");

  var khung = tao("aside", "ai-khung");
  khung.id = "ai-khung";
  khung.setAttribute("aria-hidden", "true");
  khung.setAttribute("aria-label", "Trợ lý AI");
  khung.innerHTML =
    '<span class="ai-goc ai-goc--tr"></span><span class="ai-goc ai-goc--bl"></span>' +
    '<div class="ai-dau">' +
    '<div class="ai-bieu-tuong"><i class="fas fa-robot"></i></div>' +
    '<div class="ai-tieu-de"><div class="ai-tieu-de__hang"><span>Trợ lý Planning</span></div>' +
    '<div class="ai-phu" title="Bản thử nghiệm, AI chạy trên máy chủ nội bộ"><i class="fas fa-shield-alt"></i> AI nội bộ</div></div>' +
    '<button type="button" class="ai-nut-dau" id="ai-ve-trang" hidden title="Hỏi về trang đang xem, đúng bộ lọc đang chọn"><i class="fas fa-crosshairs"></i><span> Trang này</span></button>' +
    '<button type="button" class="ai-nut-dau" id="ai-moi" title="Cuộc hỏi mới" aria-label="Cuộc hỏi mới"><i class="fas fa-plus"></i></button>' +
    '<button type="button" class="ai-nut-dau" id="ai-phong-to" title="Phóng to"><i class="fas fa-expand-alt"></i></button>' +
    '<button type="button" class="ai-nut-dau" id="ai-dong" title="Đóng (Esc)" aria-label="Đóng"><i class="fas fa-times"></i></button>' +
    "</div>" +
    '<div class="ai-ngu-canh" id="ai-ngu-canh" hidden></div>' +
    '<div class="ai-tin-ds" id="ai-tin" aria-live="polite"></div>' +
    '<div class="ai-nhap-khung">' +
    '<textarea id="ai-nhap" class="ai-nhap" rows="1" maxlength="1000" autocomplete="off" placeholder="Hỏi về đơn hàng, tàu, tồn kho, sản xuất..."></textarea>' +
    '<button type="button" id="ai-gui" class="ai-gui" title="Gửi (Enter)" aria-label="Gửi"><i class="fas fa-paper-plane"></i></button>' +
    "</div>" +
    '<div class="ai-luu-y">Enter để gửi, Shift+Enter xuống dòng. Bảng là số gốc của hệ thống; câu chữ do AI viết đã đối chiếu với số gốc.</div>';

  var tin, nhap, nutGui, nutTrang, nutPhongTo, nguCanh;
  var lichSu = []; // 3 lượt gần nhất gửi kèm câu hỏi để AI hiểu câu hỏi tiếp
  var hoiThoai = []; // toàn bộ để vẽ lại khi chuyển trang
  var dangHoi = false;
  var daChao = false;

  function cuonXuong() {
    tin.scrollTop = tin.scrollHeight;
  }

  /* Web Planning không khai meta viewport: điện thoại vẽ trang rộng ~980px rồi thu nhỏ. Khi đó phóng khung theo tỉ lệ
     thu nhỏ để chữ đọc được và khung phủ đủ màn hình. */
  function canhDienThoai() {
    var nho = Math.min(screen.width, screen.height) <= 768;
    var z = window.innerWidth / screen.width;
    var dt = nho && (z >= 1.15 || window.innerWidth <= 768);
    document.body.classList.toggle("ai-dien-thoai", dt);
    if (dt && z >= 1.15) {
      khung.style.zoom = z;
      fab.style.zoom = z;
      khung.style.width = screen.width + "px";
      khung.style.height = window.innerHeight / z + "px";
    } else {
      khung.style.zoom = fab.style.zoom = khung.style.width = khung.style.height = "";
    }
  }

  function hienNguCanh() {
    var t = trangHienTai();
    nutTrang.hidden = !t;
    if (!t) {
      nguCanh.hidden = true;
      return t;
    }
    var loc = Object.keys(t.loc).map(function (k) { return (NHAN_LOC[k] || k) + ": " + C.thoat(t.loc[k]); });
    nguCanh.innerHTML = '<i class="fas fa-map-marker-alt"></i> Đang xem <strong>' + C.thoat(t.ten) + "</strong>" + (loc.length ? " · " + loc.join(" · ") : "");
    nguCanh.hidden = false;
    return t;
  }

  function mo() {
    canhDienThoai();
    document.body.classList.add("ai-dang-mo");
    khung.setAttribute("aria-hidden", "false");
    var t = hienNguCanh();
    if (!daChao) chao(t);
    ghi(KHOA_MO, "1");
    setTimeout(function () { nhap.focus(); cuonXuong(); }, 150);
  }

  function dong() {
    document.body.classList.remove("ai-dang-mo");
    khung.setAttribute("aria-hidden", "true");
    ghi(KHOA_MO, "0");
  }

  /* ===== Lưu, vẽ lại hội thoại ===== */
  function luu() {
    hoiThoai = hoiThoai.slice(-16);
    ghi(KHOA_HT, JSON.stringify({ ht: hoiThoai, ls: lichSu }));
  }

  function cuocMoi() {
    if (dangHoi) return;
    hoiThoai = [];
    lichSu = [];
    luu();
    tin.innerHTML = "";
    daChao = false;
    chao(hienNguCanh());
    nhap.focus();
  }

  /* ===== Tin nhắn ===== */
  function bongBot(lop) {
    var b = tao("div", "ai-tin ai-tro-ly" + (lop ? " " + lop : ""));
    tin.appendChild(b);
    cuonXuong();
    return b;
  }

  function bongNguoi(text) {
    var n = tao("div", "ai-tin ai-nguoi");
    n.textContent = text;
    tin.appendChild(n);
    return n;
  }

  function xoaGoiY() {
    var cu = document.getElementById("ai-goi-y");
    if (cu) cu.remove();
  }

  function hienGoiY(ds) {
    xoaGoiY();
    if (!ds || !ds.length) return;
    var w = tao("div", "ai-goi-y-ds");
    w.id = "ai-goi-y";
    ds.forEach(function (g) {
      var b = tao("button", "ai-chip");
      b.type = "button";
      b.textContent = g.hoi;
      b.addEventListener("click", function () { hoi({ cau_hoi: g.hoi, goi: g.goi }); });
      w.appendChild(b);
    });
    tin.appendChild(w);
    cuonXuong();
  }

  function chao(t) {
    daChao = true;
    var gio = new Date().getHours();
    var buoi = gio < 11 ? "Chào buổi sáng" : gio < 13 ? "Chào buổi trưa" : gio < 18 ? "Chào buổi chiều" : "Chào buổi tối";
    var b = bongBot("ai-chao");
    b.innerHTML =
      "<p><strong>" + buoi + " anh/chị.</strong> Tôi là trợ lý số liệu của web Planning: đơn hàng, tàu, tồn kho, sản xuất, dừng sản xuất, bán hàng, chất lượng.</p>" +
      (t ? "<p>Anh/chị đang xem <strong>" + C.thoat(t.ten) + "</strong>. Bấm <strong>Trang này</strong> để hỏi đúng bộ lọc đang chọn.</p>" : "") +
      '<p class="ai-mo-nhat">Tôi chỉ báo số và phép so bằng số, không nhận xét. Anh/chị hỏi bằng lời thường, hoặc chọn một câu dưới đây.</p>';
    var q = new URLSearchParams(Object.assign({ trang: t ? t.ma : "" }, t ? t.loc : {}));
    C.goi("/goi-y?" + q.toString())
      .then(function (r) { return r.json(); })
      .then(function (j) { if (!dangHoi) hienGoiY(j.goi_y || []); })
      .catch(function () { /* không có gợi ý vẫn hỏi được */ });
  }

  /* Bảng dài: hiện BANG_GON hàng đầu, nút "Xem thêm N mục". */
  function thuGonBang(khoi) {
    khoi.querySelectorAll("table.ai-bang").forEach(function (bg) {
      var hang = bg.querySelectorAll("tbody tr");
      if (hang.length <= BANG_GON + 2) return;
      bg.classList.add("ai-bang-gon");
      var nut = tao("button", "ai-xem-them", '<i class="fas fa-chevron-down"></i> Xem thêm ' + (hang.length - BANG_GON) + " mục");
      nut.type = "button";
      nut.addEventListener("click", function () {
        var gon = bg.classList.toggle("ai-bang-gon");
        nut.innerHTML = gon ? '<i class="fas fa-chevron-down"></i> Xem thêm ' + (hang.length - BANG_GON) + " mục" : '<i class="fas fa-chevron-up"></i> Thu gọn';
      });
      bg.parentNode.insertAdjacentElement("afterend", nut);
    });
  }

  /* Chữ thuần của câu trả lời (để chép): câu chữ + bảng dạng tab. */
  function chuThuan(kq) {
    var p = [String(kq.van_ban || "").replace(/\*\*/g, "")];
    (kq.bang || []).forEach(function (b) {
      p.push(
        b.md
          .split(/\n/)
          .filter(function (h) { return !/^\s*\|?\s*:?-{3,}/.test(h); })
          .map(function (h) { return h.replace(/^\s*\|/, "").replace(/\|\s*$/, "").split("|").map(function (o) { return o.replace(/\*\*/g, "").trim(); }).join("\t"); })
          .join("\n"),
      );
    });
    if (kq.thoi_diem) p.push("Số liệu tính đến " + kq.thoi_diem);
    return p.join("\n\n");
  }

  function veTraLoi(o, kq, cauHoi, coPhanHoi) {
    o.bong.classList.remove("ai-dang-chay");
    o.van.innerHTML = C.mdSangHtml(kq.van_ban);
    if (kq.ghi_chu) o.bong.appendChild(tao("div", "ai-ghi-chu", '<i class="fas fa-info-circle"></i> ' + C.thoat(kq.ghi_chu)));
    (kq.bang || []).forEach(function (b) {
      var k = tao("div", "ai-khoi-bang", C.mdSangHtml(b.md));
      o.bong.appendChild(k);
      thuGonBang(k);
    });
    if (kq.loi && kq.loi.length) o.bong.appendChild(tao("div", "ai-ghi-chu", C.thoat(kq.loi.join("; "))));
    (kq.lien_ket || []).forEach(function (l) {
      if (!/^\/[^/]/.test(l.url)) return;
      var a = tao("a", "ai-lien-ket", '<i class="fas fa-location-arrow"></i> ' + C.thoat(l.nhan));
      a.href = l.url;
      o.bong.appendChild(a);
    });
    var chan = tao("div", "ai-chan");
    var moc = tao("span", "ai-moc", kq.thoi_diem ? '<i class="far fa-clock"></i> Số liệu tính đến ' + C.thoat(kq.thoi_diem) : "");
    chan.appendChild(moc);
    var nut = tao("span", "ai-phan-hoi");
    var chep = tao("button", "", '<i class="far fa-copy"></i>');
    chep.type = "button";
    chep.title = "Chép câu trả lời";
    chep.addEventListener("click", function () {
      var xong = function () { chep.innerHTML = '<i class="fas fa-check"></i>'; setTimeout(function () { chep.innerHTML = '<i class="far fa-copy"></i>'; }, 1500); };
      if (navigator.clipboard) navigator.clipboard.writeText(chuThuan(kq)).then(xong).catch(function () {});
    });
    nut.appendChild(chep);
    if (coPhanHoi) {
      [["dung", "fa-thumbs-up", "Đúng"], ["sai", "fa-thumbs-down", "Chưa đúng"]].forEach(function (x) {
        var b = tao("button", "", '<i class="far ' + x[1] + '"></i>');
        b.type = "button";
        b.title = x[2];
        b.addEventListener("click", function () {
          C.goi("/phan-hoi", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ danh_gia: x[0], cau_hoi: cauHoi, van_ban: kq.van_ban }) });
          nut.querySelectorAll("button").forEach(function (y) { if (y !== chep) y.remove(); });
          nut.appendChild(tao("span", "ai-moc", "Đã ghi nhận, cảm ơn anh/chị."));
        });
        nut.appendChild(b);
      });
    }
    chan.appendChild(nut);
    o.bong.appendChild(chan);
    o.moc = moc;
    return o;
  }

  /* Khi đang chạy: ẩn phần "GỢI Ý TIẾP" AI viết ở cuối (máy chủ tách ra thành nút gợi ý). */
  function catGoiY(van) {
    var i = van.search(/\n?\s*\**\s*gợi\s*(?:ý\s*)?tiếp\s*\**\s*:/i);
    return i >= 0 ? van.slice(0, i) : van;
  }

  function hoi(body) {
    if (dangHoi) return;
    var cauHoi = body.cau_hoi || "";
    if (!cauHoi && !body.ve_trang) return;
    dangHoi = true;
    nutGui.disabled = true;
    xoaGoiY();
    var t = hienNguCanh();
    if (t) {
      body.trang = t.ma;
      body.loc = t.loc;
    }
    body.lich_su = lichSu;
    var nguoi = bongNguoi(cauHoi || "Trang " + (t ? t.ten : "này") + " đang thế nào?");
    // Bóng trả lời có ngay: ba chấm, rồi các bước (đã xong có dấu tích và số giây, đang chạy có đồng hồ), chữ hiện dần bên dưới.
    var bong = bongBot("ai-dang-chay");
    var buoc = tao("div", "ai-buoc");
    var cho = tao("span", "ai-cham", "<i></i><i></i><i></i>");
    var van0 = tao("div", "ai-van");
    bong.appendChild(buoc);
    bong.appendChild(cho);
    bong.appendChild(van0);
    var o = { bong: bong, van: van0 };
    var dongChu = "";
    var xong = false;
    var kqCuoi = null;
    var dangChay = [];
    var dongHo = setInterval(function () {
      dangChay.forEach(function (m) { m.giay.textContent = giay(Date.now() - m.bd); });
    }, 200);

    function giay(ms) { return (ms / 1000).toFixed(1).replace(".", ",") + " s"; }
    function xongBuoc() {
      dangChay.forEach(function (m) {
        m.el.classList.remove("dang");
        m.el.classList.add("xong");
        var ms = Date.now() - m.bd;
        m.giay.textContent = ms < 100 ? "" : giay(ms);
      });
      dangChay = [];
    }
    function themBuoc(chu, songSong) {
      if (!songSong) xongBuoc();
      cho.remove();
      var el = tao("div", "ai-buoc-muc dang", '<i class="ai-buoc-dau"></i><span class="ai-buoc-chu"></span><em></em>');
      el.querySelector(".ai-buoc-chu").textContent = chu;
      buoc.appendChild(el);
      dangChay.push({ el: el, bd: Date.now(), giay: el.querySelector("em"), chu: el.querySelector(".ai-buoc-chu") });
      cuonXuong();
    }
    function bat() {
      cho.remove();
      return o;
    }
    function ketThucBuoc() {
      clearInterval(dongHo);
      buoc.remove();
      cho.remove();
    }

    function xuLy(e, d) {
      if (e === "cau_hoi" && body.ve_trang) nguoi.textContent = d.cau_hoi;
      else if (e === "trang_thai") {
        var chu = d.text.replace(/\.\.\.$/, "").replace(/^Đang /, "");
        chu = chu.charAt(0).toUpperCase() + chu.slice(1);
        // Chờ lượt: ghi vào bước đang chạy, không thêm bước mới.
        if (d.buoc === "cho" && dangChay.length) dangChay[dangChay.length - 1].chu.textContent += " (chờ lượt " + d.text.replace(/\D+/g, "") + ")";
        else themBuoc(chu);
      } else if (e === "cong_cu") {
        // Chữ model viết trước khi gọi công cụ ("để em kiểm tra...") bỏ đi; nhiều công cụ chạy song song.
        dongChu = "";
        o.van.innerHTML = "";
        var songSong = dangChay.length && /^Lấy số/.test(dangChay[0].chu.textContent);
        themBuoc("Lấy số: " + d.nhan, songSong);
      } else if (e === "chu") {
        dongChu += d.t;
        // Câu trả lời thẳng (không cần lấy số) bắt đầu hiện ngay trong bước hiểu câu hỏi.
        if (dangChay.length === 1 && dangChay[0].chu.textContent === "Hiểu câu hỏi") dangChay[0].chu.textContent = "Trả lời";
        // Hiện có định dạng ngay khi đang chạy (danh sách, chữ đậm), bỏ phần gợi ý ở cuối.
        bat().van.innerHTML = C.mdSangHtml(catGoiY(dongChu)) + '<span class="ai-con-tro"></span>';
        cuonXuong();
      } else if (e === "ket_qua") {
        kqCuoi = d;
        ketThucBuoc();
        veTraLoi(bat(), d, nguoi.textContent, true);
        cuonXuong();
        if (d.goi_y && d.goi_y.length) hienGoiY(d.goi_y.map(function (h) { return { hoi: h }; }));
        lichSu.push({ hoi: nguoi.textContent, tra_loi: d.van_ban });
        lichSu = lichSu.slice(-3);
        hoiThoai.push({ vai: "nguoi", text: nguoi.textContent }, { vai: "tro_ly", kq: d, hoi: nguoi.textContent });
        luu();
        xong = true;
      } else if (e === "xong") {
        if (o && o.moc && d.ms && kqCuoi) o.moc.innerHTML += (o.moc.innerHTML ? " · " : "") + "trả lời trong " + (d.ms / 1000).toFixed(1).replace(".", ",") + " giây";
      } else if (e === "loi") {
        ketThucBuoc();
        var b = bat();
        b.bong.classList.remove("ai-dang-chay");
        b.bong.classList.add("ai-loi");
        b.van.innerHTML = '<i class="fas fa-exclamation-triangle"></i> ' + C.thoat(d.text);
        xong = true;
      }
    }

    C.goi("/hoi-dap", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })
      .then(function (r) {
        if (!r.ok) return r.json().then(function (j) { throw new Error(j.loi || "Trợ lý chưa trả lời được (" + r.status + ")."); });
        var docSse = r.body.getReader();
        var giai = new TextDecoder();
        var dem = "";
        function doc1() {
          return docSse.read().then(function (p) {
            if (p.done) return;
            dem += giai.decode(p.value, { stream: true });
            var i;
            while ((i = dem.indexOf("\n\n")) >= 0) {
              var k = dem.slice(0, i);
              dem = dem.slice(i + 2);
              var e = (k.match(/^event: (.*)$/m) || [])[1];
              var d = (k.match(/^data: (.*)$/m) || [])[1];
              if (e && d) xuLy(e, JSON.parse(d));
            }
            return doc1();
          });
        }
        return doc1();
      })
      .catch(function (err) { xuLy("loi", { text: err.message || "Mất kết nối tới trợ lý." }); })
      .then(function () {
        if (!xong) xuLy("loi", { text: "Mất kết nối tới trợ lý giữa chừng. Anh/chị hỏi lại." });
        ketThucBuoc();
        dangHoi = false;
        nutGui.disabled = false;
        nhap.focus();
        cuonXuong();
      });
  }

  function gui() {
    var v = nhap.value.trim();
    if (!v || dangHoi) return;
    nhap.value = "";
    tuGian();
    hoi({ cau_hoi: v });
  }

  function tuGian() {
    nhap.style.height = "auto";
    nhap.style.height = Math.min(nhap.scrollHeight, 120) + "px";
  }

  function veLai() {
    var luuCu = null;
    try { luuCu = JSON.parse(doc(KHOA_HT) || "null"); } catch (e) { luuCu = null; }
    if (!luuCu || !luuCu.ht || !luuCu.ht.length) return false;
    hoiThoai = luuCu.ht;
    lichSu = luuCu.ls || [];
    hoiThoai.forEach(function (m) {
      if (m.vai === "nguoi") bongNguoi(m.text);
      else if (m.vai === "tro_ly" && m.kq) {
        var bong = bongBot();
        var van = tao("div", "ai-van");
        bong.appendChild(van);
        veTraLoi({ bong: bong, van: van }, m.kq, m.hoi, false);
      }
    });
    var cuoi = hoiThoai[hoiThoai.length - 1];
    if (cuoi && cuoi.kq && cuoi.kq.goi_y && cuoi.kq.goi_y.length) hienGoiY(cuoi.kq.goi_y.map(function (h) { return { hoi: h }; }));
    daChao = true;
    return true;
  }

  function gan() {
    document.body.appendChild(fab);
    document.body.appendChild(khung);
    tin = document.getElementById("ai-tin");
    nhap = document.getElementById("ai-nhap");
    nutGui = document.getElementById("ai-gui");
    nutTrang = document.getElementById("ai-ve-trang");
    nutPhongTo = document.getElementById("ai-phong-to");
    nguCanh = document.getElementById("ai-ngu-canh");
    fab.addEventListener("click", mo);
    document.getElementById("ai-dong").addEventListener("click", dong);
    document.getElementById("ai-moi").addEventListener("click", cuocMoi);
    nutPhongTo.addEventListener("click", function () {
      var lon = khung.classList.toggle("lon");
      nutPhongTo.innerHTML = lon ? '<i class="fas fa-compress-alt"></i>' : '<i class="fas fa-expand-alt"></i>';
      nutPhongTo.title = lon ? "Thu nhỏ" : "Phóng to";
    });
    nutGui.addEventListener("click", gui);
    nutTrang.addEventListener("click", function () { hoi({ ve_trang: true, cau_hoi: "" }); });
    nhap.addEventListener("input", tuGian);
    nhap.addEventListener("keydown", function (e) {
      if (e.key === "Enter" && !e.shiftKey && !e.isComposing) {
        e.preventDefault();
        gui();
      }
    });
    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && document.body.classList.contains("ai-dang-mo")) dong();
    });
    canhDienThoai();
    window.addEventListener("resize", canhDienThoai);
    veLai();
    if (doc(KHOA_MO) === "1") setTimeout(mo, 200);
  }

  /* Chỉ gắn trợ lý khi lấy được token (đã đăng nhập, có quyền, trợ lý đã cấu hình). */
  function batDau() {
    C.layToken().then(gan).catch(function () { /* không có quyền hoặc chưa đăng nhập: không hiện trợ lý */ });
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", batDau);
  else batDau();
})();
