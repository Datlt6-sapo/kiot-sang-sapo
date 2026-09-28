/*
 * Lõi chuyển đổi file (KiotViet hoặc file khác) -> file mẫu nhập Sapo.
 * Chạy được cả trong trình duyệt lẫn Node (để test).
 * Quy tắc map giữ đúng như 2 skill kiotviet-sapo-split và kiotviet-khach-hang-cong-no.
 *
 * createConverter({ XLSX, ExcelJS, templates, dmhcCsv })
 *   templates: { donLe, quyDoi, khachHang, congNo } -> Uint8Array/ArrayBuffer của file mẫu Sapo
 *
 * Luồng: analyze(buf) -> đọc file, tìm dòng tiêu đề, đoán loại file và ghép cột
 *        run(input, type, mapping) -> tạo file Sapo theo cách ghép cột đã chốt
 */
(function (root) {
  "use strict";

  // ---------- Tiện ích chung ----------
  function isEmpty(v) {
    return v === null || v === undefined || (typeof v === "number" && isNaN(v)) || String(v).trim() === "";
  }
  function val(v) {
    return isEmpty(v) ? null : v;
  }
  function str(v) {
    if (isEmpty(v)) return null;
    if (typeof v === "number" && Number.isInteger(v)) return v.toFixed(0);
    return String(v).trim();
  }
  function num(v) {
    if (isEmpty(v)) return null;
    if (typeof v === "number") return v;
    let s = String(v).replace(/\s|đ|₫|vnd|%/gi, "").trim();
    // Số kiểu Việt Nam "1.250.000" hoặc "1.250.000,5"
    if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
    const n = Number(s);
    return isNaN(n) ? null : n;
  }
  function round2(n) {
    return Math.round(n * 100) / 100;
  }
  function stripMarks(s) {
    return s.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D");
  }
  function norm(s) {
    if (isEmpty(s)) return "";
    return stripMarks(String(s).normalize("NFC")).replace(/\s+/g, " ").trim().toLowerCase();
  }
  function nospace(s) {
    return s.replace(/[^a-z0-9]/g, "");
  }
  // Khoá so khớp tên cột: bỏ dấu, chữ thường, bỏ mọi ký tự không phải chữ/số
  const key = (s) => nospace(norm(s));
  function slugify(text) {
    if (isEmpty(text)) return null;
    const s = stripMarks(String(text).trim()).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/-+/g, "-").replace(/^-|-$/g, "");
    return s || null;
  }

  // =====================================================================
  // DANH SÁCH CỘT SAPO CẦN ĐIỀN + các tên gọi thường gặp ở phần mềm khác
  // (tên đầu tiên trong aliases là tên cột của KiotViet)
  // =====================================================================
  const SCHEMAS = {
    product: {
      label: "Danh sách hàng hóa",
      fields: [
        { id: "code", label: "Mã SKU", required: true, aliases: ["Mã hàng", "Mã SP", "Mã sản phẩm", "Mã hàng hóa", "SKU", "Mã SKU", "Mã"] },
        { id: "name", label: "Tên sản phẩm", required: true, aliases: ["Tên hàng", "Tên sản phẩm", "Tên SP", "Tên hàng hóa", "Tên"] },
        { id: "qty", label: "Quy đổi (1 = đơn lẻ)", hint: "Không có cột này thì mọi dòng là sản phẩm đơn lẻ", aliases: ["Quy đổi", "Tỷ lệ quy đổi", "Hệ số quy đổi", "Giá trị quy đổi", "Số lượng quy đổi"] },
        { id: "baseRef", label: "SKU gốc của hàng quy đổi", aliases: ["Mã ĐVT Cơ bản", "Mã hàng cơ bản", "SKU gốc", "SKU tham chiếu", "Mã hàng gốc", "Mã SP gốc"] },
        { id: "relatedRef", label: "Mã hàng liên quan", hint: "Dùng khi SKU gốc trống", aliases: ["Mã HH Liên quan", "Mã hàng liên quan"] },
        { id: "unit", label: "Đơn vị tính", aliases: ["ĐVT", "Đơn vị tính", "Đơn vị"] },
        { id: "priceAfterTax", label: "Giá bán sau thuế", hint: "Ưu tiên dùng nếu có", aliases: ["Giá bán sau thuế", "Giá sau thuế"] },
        { id: "price", label: "Giá bán", aliases: ["Giá bán", "Giá bán lẻ", "Đơn giá", "Giá lẻ", "Giá"] },
        { id: "priceBeforeTax", label: "Giá bán trước thuế", aliases: ["Giá bán trước thuế", "Giá trước thuế"] },
        { id: "cost", label: "Giá vốn", aliases: ["Giá vốn", "Giá nhập", "Giá mua", "Giá vốn trung bình"] },
        { id: "stock", label: "Tồn kho", aliases: ["Tồn kho", "Số lượng tồn", "Tồn", "Tồn cuối kỳ", "SL tồn", "Số lượng"] },
        { id: "barcode", label: "Barcode", aliases: ["Mã vạch", "Barcode"] },
        { id: "brand", label: "Nhãn hiệu", aliases: ["Thương hiệu", "Nhãn hiệu", "Hãng"] },
        { id: "category", label: "Loại sản phẩm", aliases: ["Nhóm hàng(3 Cấp)", "Nhóm hàng (3 Cấp)", "Nhóm hàng", "Loại sản phẩm", "Danh mục", "Nhóm sản phẩm", "Nhóm SP", "Nhóm hàng hóa", "Ngành hàng"] },
        { id: "vat", label: "Thuế VAT (%)", hint: "Có giá trị thì Áp dụng thuế = Có", aliases: ["VAT (%)", "VAT(%)", "VAT", "Thuế suất", "Thuế GTGT", "% VAT"] },
        { id: "description", label: "Mô tả sản phẩm", aliases: ["Mô tả", "Mô tả sản phẩm"] },
        { id: "image", label: "Ảnh đại diện", hint: "Lấy link đầu tiên", aliases: ["Hình ảnh (url1,url2...)", "Hình ảnh", "Ảnh", "Link ảnh", "Ảnh đại diện"] },
        { id: "weight", label: "Khối lượng", aliases: ["Trọng lượng", "Khối lượng"] },
      ],
    },
    customer: {
      label: "Danh sách khách hàng",
      fields: [
        { id: "ten", label: "Tên khách hàng", required: true, aliases: ["Tên khách hàng", "Tên KH", "Họ tên", "Họ và tên", "Khách hàng", "Tên"] },
        { id: "sdt", label: "Số điện thoại", required: true, hint: "Dòng trống SĐT được gán 0900000001…", aliases: ["Điện thoại", "SĐT", "Số điện thoại", "Di động", "Điện thoại di động", "Số ĐT", "SĐT di động", "ĐT", "Phone", "Mobile"] },
        { id: "email", label: "Email", aliases: ["Email", "E-mail", "Thư điện tử"] },
        { id: "diaChi", label: "Địa chỉ", aliases: ["Địa chỉ", "Địa chỉ chi tiết", "Số nhà, đường"] },
        { id: "khuVuc", label: "Khu vực (Tỉnh - Huyện)", hint: "Dạng “Tỉnh - Huyện” như KiotViet", aliases: ["Khu vực giao hàng", "Khu vực"] },
        { id: "tinh", label: "Tỉnh/Thành phố", aliases: ["Tỉnh/Thành phố", "Tỉnh/Thành", "Tỉnh thành", "Tỉnh", "Thành phố"] },
        { id: "huyen", label: "Quận/Huyện", aliases: ["Quận/Huyện", "Quận huyện", "Huyện", "Quận"] },
        { id: "xa", label: "Phường/Xã", aliases: ["Phường/Xã", "Phường xã", "Xã", "Phường"] },
        { id: "gioiTinh", label: "Giới tính", aliases: ["Giới tính"] },
        { id: "ghiChu", label: "Ghi chú", hint: "Cắt tối đa 255 ký tự", aliases: ["Ghi chú", "Note"] },
        { id: "no", label: "Công nợ", hint: "Tạo file công nợ, bỏ dòng = 0", aliases: ["Nợ cần thu hiện tại", "Nợ hiện tại", "Công nợ", "Dư nợ", "Nợ phải thu", "Công nợ hiện tại", "Nợ"] },
        { id: "tinhHd", label: "Tỉnh (xuất hóa đơn)", hint: "Dùng khi không có Khu vực/Tỉnh", aliases: ["Tỉnh/Thành phố (xuất HĐ)"] },
      ],
    },
  };

  // Ghép cột: khớp chính xác trước (theo khoá bỏ dấu), sau đó mới đoán gần đúng.
  // Đoán gần đúng chỉ khi tên cột chứa tên gọi và chỉ dư ít ký tự (vd "Giá bán lẻ" ~ "Giá bán"),
  // để không ghép nhầm kiểu "Tỉnh/Thành phố (xuất HĐ)" vào "Tỉnh/Thành phố".
  function autoMap(headers, schema) {
    const used = new Set();
    const mapping = {}, status = {};
    const hk = headers.map((h) => ({ h, k: key(h) }));
    for (const f of schema.fields) {
      for (const a of f.aliases) {
        const hit = hk.find((x) => !used.has(x.h) && x.k === key(a));
        if (hit) { mapping[f.id] = hit.h; status[f.id] = "exact"; used.add(hit.h); break; }
      }
    }
    for (const f of schema.fields) {
      if (mapping[f.id]) continue;
      let best = null;
      for (const a of f.aliases) {
        const ak = key(a);
        if (ak.length < 4) continue;
        for (const x of hk) {
          if (used.has(x.h) || !x.k.includes(ak)) continue;
          const extra = x.k.length - ak.length;
          if (extra > 5) continue;
          if (!best || extra < best.extra) best = { h: x.h, extra };
        }
      }
      if (best) { mapping[f.id] = best.h; status[f.id] = "guess"; used.add(best.h); }
    }
    const missingRequired = schema.fields.filter((f) => f.required && !mapping[f.id]).map((f) => f.id);
    const exact = Object.values(status).filter((s) => s === "exact").length;
    return { mapping, status, missingRequired, exact };
  }

  function createConverter(opts) {
    const XLSX = opts.XLSX;
    const ExcelJS = opts.ExcelJS;
    const T = opts.templates;
    let dmhc = null; // nạp lười khi cần

    // ---------- Đọc file ----------
    const allAliasKeys = new Set();
    for (const s of Object.values(SCHEMAS)) for (const f of s.fields) for (const a of f.aliases) allAliasKeys.add(key(a));

    // Số ô trong 1 dòng trùng tên cột quen thuộc (dùng để tìm dòng tiêu đề)
    function headerScore(row) {
      return (row || []).filter((c) => !isEmpty(c) && allAliasKeys.has(key(c))).length;
    }
    function findHeaderRow(grid, limit) {
      let headerIdx = 0, bestScore = -1;
      for (let i = 0; i < Math.min(limit || 20, grid.length); i++) {
        const score = headerScore(grid[i]);
        if (score > bestScore) { bestScore = score; headerIdx = i; }
      }
      return headerIdx;
    }

    function gridFromWorkbook(buf) {
      const wb = XLSX.read(buf, { type: "array" });
      // Chọn sheet có nhiều dữ liệu nhất (file từ phần mềm khác hay có sheet hướng dẫn ở đầu)
      let grid = [];
      for (const n of wb.SheetNames) {
        const g = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: null, raw: true, blankrows: false });
        if (g.length > grid.length) grid = g;
      }
      return grid;
    }

    // grid: mảng các dòng (mảng ô). headerIdx bỏ trống thì tự tìm trong 20 dòng đầu.
    function inputFromGrid(grid, headerIdx) {
      if (headerIdx === undefined || headerIdx === null) headerIdx = findHeaderRow(grid);
      const raw = grid[headerIdx] || [];
      const width = Math.max(raw.length, ...grid.slice(headerIdx, headerIdx + 50).map((r) => r.length));
      const seen = {};
      const headers = [];
      for (let c = 0; c < width; c++) {
        let h = isEmpty(raw[c]) ? `Cột ${XLSX.utils.encode_col(c)}` : String(raw[c]).trim();
        if (seen[h]) h = `${h} (${++seen[h]})`; else seen[h] = 1;
        headers.push(h);
      }
      const rows = [];
      for (const arr of grid.slice(headerIdx + 1)) {
        if (!arr.some((v) => !isEmpty(v))) continue;
        const o = {};
        headers.forEach((h, c) => (o[h] = arr[c] === undefined ? null : arr[c]));
        rows.push(o);
      }
      // Bỏ cột tự đặt tên mà không có dữ liệu
      const nonEmpty = headers.filter((h) => !/^Cột [A-Z]+$/.test(h) || rows.some((r) => !isEmpty(r[h])));
      return { headers: nonEmpty, rows, headerRow: headerIdx + 1 };
    }

    function readInput(buf) {
      return inputFromGrid(gridFromWorkbook(buf));
    }

    // Chữ ký file: dùng để nhớ cách ghép cột cho các file cùng kiểu
    function signature(headers) {
      return headers.map(key).sort().join("|");
    }

    // src: nội dung file Excel (ArrayBuffer/Uint8Array) hoặc { grid, headerIdx } đã đọc từ PDF/ảnh
    function analyze(src) {
      const input = src && src.grid ? inputFromGrid(src.grid, src.headerIdx) : readInput(src);
      const cands = Object.keys(SCHEMAS).map((t) => ({ type: t, ...autoMap(input.headers, SCHEMAS[t]) }));
      cands.sort((a, b) => (a.missingRequired.length - b.missingRequired.length) || (b.exact - a.exact));
      const best = cands[0];
      const recognized = best.missingRequired.length === 0 && best.exact > 0;
      const guessed = Object.values(best.status).includes("guess");
      // Hàng hóa không có cột Quy đổi (file phần mềm khác, bảng giá...) -> cho người dùng xem lại trước
      const noQty = best.type === "product" && !best.mapping.qty;
      return {
        input,
        type: best.type,
        recognized,
        mapping: best.mapping,
        status: best.status,
        needsReview: !recognized || guessed || noQty,
        signature: signature(input.headers),
        candidates: Object.fromEntries(cands.map((c) => [c.type, { mapping: c.mapping, status: c.status }])),
      };
    }

    async function loadTemplate(bytes) {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(bytes);
      return wb;
    }

    function clearRows(ws, fromRow, lastCol) {
      const last = Math.max(ws.rowCount, fromRow);
      for (let r = fromRow; r <= last; r++) {
        const row = ws.getRow(r);
        for (let c = 1; c <= lastCol; c++) row.getCell(c).value = null;
      }
    }

    const FONT = { name: "Arial", size: 11 };
    function put(ws, r, c, value, isText) {
      if (value === null || value === undefined) return;
      const cell = ws.getRow(r).getCell(c);
      cell.value = value;
      cell.font = FONT;
      if (isText) cell.numFmt = "@";
    }

    async function toBuffer(wb) {
      return await wb.xlsx.writeBuffer();
    }

    function checkRequired(type, m) {
      const miss = SCHEMAS[type].fields.filter((f) => f.required && !m[f.id]).map((f) => f.label);
      if (miss.length) throw new Error("Chưa chọn cột cho: " + miss.join(", ") + ".");
    }

    // =====================================================================
    // SẢN PHẨM  (đơn lẻ + quy đổi)
    // =====================================================================
    async function convertProducts(input, m) {
      checkRequired("product", m);
      // Giá: ưu tiên Giá bán sau thuế -> Giá bán -> Giá bán trước thuế
      const priceCol = m.priceAfterTax || m.price || m.priceBeforeTax || null;
      const g = (r, k) => (m[k] ? val(r[m[k]]) : null);
      // SKU tham chiếu: SKU gốc, nếu trống thì Mã hàng liên quan
      const refCols = [m.baseRef, m.relatedRef].filter(Boolean);
      const baseRef = (r) => {
        for (const c of refCols) if (!isEmpty(r[c])) return str(r[c]);
        return null;
      };

      const donLe = [], quyDoi = [];
      for (const r of input.rows) {
        const q = m.qty ? num(r[m.qty]) : null;
        (q === null || q === 1 ? donLe : quyDoi).push(r);
      }

      const warnings = [];
      if (!m.qty) warnings.push("Không có cột Quy đổi: mọi dòng được coi là sản phẩm đơn lẻ.");
      if (!refCols.length && quyDoi.length) warnings.push("Không có cột SKU gốc: cột “SKU tham chiếu” trong file quy đổi để trống.");
      if (!priceCol) warnings.push("Không có cột giá bán: cột “Giá” để trống.");

      // ---- File 1: đơn lẻ ----
      const wb1 = await loadTemplate(T.donLe);
      const ws1 = wb1.getWorksheet("Mẫu file nhập");
      clearRows(ws1, 2, 36);
      donLe.forEach((r, idx) => {
        const i = idx + 2;
        put(ws1, i, 2, g(r, "name"));
        put(ws1, i, 3, g(r, "description"));
        put(ws1, i, 4, g(r, "brand"));
        put(ws1, i, 5, g(r, "category"));
        put(ws1, i, 8, "Có");
        put(ws1, i, 9, "Có");
        if (m.vat) put(ws1, i, 16, isEmpty(r[m.vat]) ? "Không" : "Có");
        put(ws1, i, 17, str(r[m.code]), true);
        put(ws1, i, 18, str(g(r, "barcode")), true);
        put(ws1, i, 19, g(r, "unit"));
        const img = g(r, "image");
        if (img) put(ws1, i, 20, String(img).split(",")[0].trim() || null);
        put(ws1, i, 25, "Sapo");
        const w = num(g(r, "weight"));
        if (w) put(ws1, i, 28, w);
        put(ws1, i, 31, "Không");
        const p = priceCol ? num(r[priceCol]) : null;
        if (p !== null) put(ws1, i, 32, round2(p));
        const cost = num(g(r, "cost"));
        if (cost !== null) put(ws1, i, 34, round2(cost));
        const st = num(g(r, "stock"));
        if (st !== null) put(ws1, i, 35, st);
      });

      const files = [{ kind: "don_le", label: "Sản phẩm đơn lẻ", name: "San_pham_don_le", buffer: await toBuffer(wb1), rows: donLe.length }];

      // ---- File 2: quy đổi ----
      const orphans = [];
      if (quyDoi.length) {
        const baseName = new Map(donLe.map((r) => [str(r[m.code]), r[m.name]]));
        const wb2 = await loadTemplate(T.quyDoi);
        const ws2 = wb2.getWorksheet("V2 Mẫu file nhập sản phẩm quy đ");
        clearRows(ws2, 2, 18);
        const seenAlias = new Set();
        quyDoi.forEach((r, idx) => {
          const i = idx + 2;
          const ref = baseRef(r);
          const bn = ref ? baseName.get(ref) : undefined;
          if (!bn) orphans.push({ sku: str(r[m.code]), name: g(r, "name"), ref: ref || "(trống)" });
          const alias = bn ? slugify(bn) : null;
          put(ws2, i, 1, alias);
          if (!(alias && seenAlias.has(alias))) {
            put(ws2, i, 2, bn || g(r, "name"));
            if (alias) seenAlias.add(alias);
          }
          put(ws2, i, 3, str(r[m.code]), true);
          put(ws2, i, 4, ref, true);
          put(ws2, i, 7, g(r, "unit"));
          const q = num(r[m.qty]);
          if (q !== null) put(ws2, i, 8, q);
          const p = priceCol ? num(r[priceCol]) : null;
          if (p !== null) put(ws2, i, 10, round2(p));
          put(ws2, i, 17, "Có");
        });
        files.push({ kind: "quy_doi", label: "Sản phẩm quy đổi", name: "San_pham_quy_doi", buffer: await toBuffer(wb2), rows: quyDoi.length });
      }

      const checks = [
        { ok: donLe.length + quyDoi.length === input.rows.length, text: `Đơn lẻ ${donLe.length} + quy đổi ${quyDoi.length} = ${input.rows.length} dòng gốc` },
      ];
      if (quyDoi.length) checks.push({ ok: orphans.length === 0, text: orphans.length ? `${orphans.length} sản phẩm quy đổi không tìm thấy sản phẩm gốc` : "Mọi SKU tham chiếu đều khớp 1 sản phẩm đơn lẻ" });

      return {
        type: "product",
        files,
        stats: { total: input.rows.length, donLe: donLe.length, quyDoi: quyDoi.length, priceCol },
        checks,
        warnings,
        issues: orphans.length ? [{ title: "Sản phẩm quy đổi chưa khớp sản phẩm gốc", columns: ["Mã SKU", "Tên hàng", "SKU tham chiếu"], rows: orphans.map((o) => [o.sku, o.name, o.ref]) }] : [],
      };
    }

    // =====================================================================
    // KHÁCH HÀNG + CÔNG NỢ
    // =====================================================================
    const PRE_TINH = ["tinh", "thanh pho", "tp"];
    const PRE_HUYEN = ["huyen", "quan", "thi xa", "thanh pho"];
    const PRE_XA = ["xa", "phuong", "thi tran"];
    function stripPref(s, prefixes) {
      for (const p of [...prefixes].sort((a, b) => b.length - a.length)) if (s.startsWith(p + " ")) return s.slice(p.length + 1);
      return s;
    }
    // Khoá so khớp tỉnh: bỏ dấu, bỏ tiền tố, bỏ khoảng trắng/gạch -> "Bà Rịa - Vũng Tàu" = "Bà Rịa-Vũng Tàu", "Thành phố Hồ Chí Minh" = "TP Hồ Chí Minh"
    const tinhKey = (s) => nospace(stripPref(norm(s), PRE_TINH));

    function getDmhc() {
      if (dmhc) return dmhc;
      const lines = String(opts.dmhcCsv).replace(/^﻿/, "").split(/\r?\n/).slice(1);
      const byTinh = new Map(); // tên tỉnh chuẩn -> [{xa, huyen, ...}]
      const tinhIndex = new Map(); // khoá tỉnh -> tên tỉnh chuẩn
      for (const line of lines) {
        const [xa, huyen, tinh] = line.split(",").map((s) => (s || "").trim());
        if (!xa || !huyen || !tinh) continue;
        const tn = tinhKey(tinh);
        if (!tinhIndex.has(tn)) tinhIndex.set(tn, tinh);
        if (!byTinh.has(tinh)) byTinh.set(tinh, []);
        const hn = stripPref(norm(huyen), PRE_HUYEN);
        const xn = stripPref(norm(xa), PRE_XA);
        byTinh.get(tinh).push({ xa, huyen, hn, hns: nospace(hn), xn, xns: nospace(xn) });
      }
      dmhc = { byTinh, tinhIndex };
      return dmhc;
    }

    // Chỉ trả về giá trị khi khớp CHÍNH XÁC với DMHC (sau khi bỏ dấu/khoảng trắng/tiền tố).
    // src có dạng "Tỉnh - Huyện" hoặc chỉ "Tỉnh".
    function matchAddress(src, xaRaw) {
      const D = getDmhc();
      // Tên tỉnh có thể chứa " - " (Bà Rịa - Vũng Tàu) nên thử cả chuỗi trước,
      // rồi lần lượt từng vị trí " - " cho tới khi phần bên trái khớp 1 tỉnh.
      let tinh = null, rawHuyen = null;
      if (src) {
        tinh = D.tinhIndex.get(tinhKey(src)) || null;
        let k = src.indexOf(" - ");
        while (!tinh && k >= 0) {
          const t = D.tinhIndex.get(tinhKey(src.slice(0, k)));
          if (t) { tinh = t; rawHuyen = src.slice(k + 3); }
          k = src.indexOf(" - ", k + 3);
        }
      }
      if (!tinh) return [null, null, null];
      const sub = D.byTinh.get(tinh);

      let huyen = null;
      const hn = rawHuyen ? stripPref(norm(rawHuyen), PRE_HUYEN) : null;
      if (hn) {
        const hit = sub.find((x) => x.hn === hn) || sub.find((x) => x.hns === nospace(hn));
        if (hit) huyen = hit.huyen;
      }
      let xa = null;
      const xn = !isEmpty(xaRaw) ? stripPref(norm(xaRaw), PRE_XA) : null;
      if (xn) {
        const scope = huyen ? sub.filter((x) => x.huyen === huyen) : sub;
        const hit = scope.find((x) => x.xn === xn) || scope.find((x) => x.xns === nospace(xn));
        if (hit) { xa = hit.xa; if (!huyen) huyen = hit.huyen; }
      }
      return [tinh, huyen, xa];
    }

    function cleanPhone(p) {
      if (isEmpty(p)) return null;
      let s = String(typeof p === "number" ? p.toFixed(0) : p).replace(/\D/g, "");
      // Excel đọc SĐT dạng số sẽ mất số 0 đầu (vd 948435767) -> thêm lại
      if (typeof p === "number" && s.length === 9) s = "0" + s;
      return s || null;
    }

    const GENDER = { nam: "male", male: "male", nu: "female", female: "female" };

    async function convertCustomers(input, m) {
      checkRequired("customer", m);
      const g = (r, k) => (m[k] ? val(r[m[k]]) : null);

      // SĐT dùng chung cho cả 2 file; thiếu SĐT -> 0900000001, 0900000002... theo thứ tự dòng
      let seq = 0;
      const phones = input.rows.map((r) => {
        const p = cleanPhone(r[m.sdt]);
        if (p) return p;
        seq += 1;
        return "09" + String(seq).padStart(8, "0");
      });

      const wb = await loadTemplate(T.khachHang);
      const ws = wb.worksheets[0];
      clearRows(ws, 2, 24);
      const unmatched = [];
      let notesCut = 0;

      input.rows.forEach((r, idx) => {
        const i = idx + 2;
        const ten = g(r, "ten");
        const xaRaw = g(r, "xa");
        // Nguồn tỉnh/huyện: Khu vực giao hàng -> cột Tỉnh + Huyện riêng -> Tỉnh xuất HĐ
        const khuVuc = g(r, "khuVuc"), tinhCol = g(r, "tinh"), huyenCol = g(r, "huyen"), tinhHd = g(r, "tinhHd");
        let src = null;
        if (khuVuc) src = String(khuVuc).trim();
        else if (tinhCol) src = String(tinhCol).trim() + (huyenCol ? " - " + String(huyenCol).trim() : "");
        else if (tinhHd) src = String(tinhHd).trim();
        const [tinh, huyen, xa] = matchAddress(src, xaRaw);

        const miss = [];
        if (!tinh && src) miss.push("Tỉnh");
        if (tinh && !huyen && src && src.includes(" - ")) miss.push("Huyện");
        if (tinh && !xa && xaRaw) miss.push("Xã");
        if (miss.length) unmatched.push([ten, [src, xaRaw].filter(Boolean).join(" / "), miss.join(", ")]);

        const addr = [g(r, "diaChi"), xaRaw].filter((x) => !isEmpty(x)).map((x) => String(x).trim()).join(", ") || null;
        let note = g(r, "ghiChu");
        note = note === null ? null : String(note).trim();
        if (note && note.length > 255) { note = note.slice(0, 255); notesCut++; }
        const gt = g(r, "gioiTinh");
        const gender = gt === null ? null : GENDER[key(gt)] || null;
        const email = g(r, "email");

        put(ws, i, 2, ten);
        put(ws, i, 4, email === null ? null : String(email).trim());
        put(ws, i, 5, phones[idx], true);
        put(ws, i, 7, addr);
        put(ws, i, 10, tinh);
        put(ws, i, 12, huyen);
        put(ws, i, 14, xa);
        put(ws, i, 16, "Vietnam");
        put(ws, i, 17, "VN");
        put(ws, i, 19, phones[idx], true);
        put(ws, i, 22, note);
        put(ws, i, 23, gender);
      });

      const files = [{ kind: "khach_hang", label: "Khách hàng", name: "Khach_hang_Sapo", buffer: await toBuffer(wb), rows: input.rows.length }];
      const checks = [{ ok: true, text: `${input.rows.length} khách hàng = ${input.rows.length} dòng gốc` }];
      const warnings = [];

      let debtRows = 0, zeroSkipped = 0, debtTotal = 0;
      if (m.no) {
        const wb2 = await loadTemplate(T.congNo);
        const ws2 = wb2.worksheets[0];
        clearRows(ws2, 2, 3);
        let i = 2;
        input.rows.forEach((r, idx) => {
          const n = num(r[m.no]) || 0;
          if (n === 0) { zeroSkipped++; return; }
          put(ws2, i, 1, phones[idx], true);
          put(ws2, i, 3, n);
          debtTotal += n;
          i++;
        });
        debtRows = i - 2;
        files.push({ kind: "cong_no", label: "Công nợ", name: "Cong_no_khach_hang_Sapo", buffer: await toBuffer(wb2), rows: debtRows });
        checks.push({ ok: true, text: `Công nợ: ${debtRows} dòng, đã bỏ ${zeroSkipped} dòng nợ = 0` });
      } else {
        warnings.push("Không có cột công nợ nên không tạo file công nợ.");
      }
      checks.push({ ok: unmatched.length === 0, text: unmatched.length ? `${unmatched.length} khách chưa khớp đủ Tỉnh/Huyện/Xã theo danh mục Sapo (để trống, không đoán)` : "Toàn bộ địa chỉ khớp danh mục hành chính Sapo" });
      if (notesCut) warnings.push(`${notesCut} ghi chú dài hơn 255 ký tự đã được cắt bớt.`);
      if (seq) warnings.push(`${seq} khách không có SĐT được gán số tạm 0900000001 → ${"09" + String(seq).padStart(8, "0")}.`);

      return {
        type: "customer",
        files,
        stats: { total: input.rows.length, fakePhones: seq, debtRows, zeroSkipped, debtTotal, unmatched: unmatched.length },
        checks,
        warnings,
        issues: unmatched.length ? [{ title: "Địa chỉ chưa khớp danh mục Sapo", columns: ["Khách hàng", "Khu vực / Phường xã (file gốc)", "Thiếu"], rows: unmatched }] : [],
      };
    }

    async function run(input, type, mapping) {
      if (!input.rows.length) throw new Error("File không có dòng dữ liệu nào bên dưới dòng tiêu đề.");
      if (type === "product") return convertProducts(input, mapping);
      if (type === "customer") return convertCustomers(input, mapping);
      throw new Error("Loại file không hợp lệ.");
    }

    // Dùng cho test: tự nhận + chuyển luôn, báo lỗi nếu không nhận ra
    async function convert(buf) {
      const a = analyze(buf);
      if (!a.recognized) throw new Error("Không nhận ra loại file. Cần chọn cột bằng tay.");
      return run(a.input, a.type, a.mapping);
    }

    return { analyze, run, convert, readInput, inputFromGrid, findHeaderRow, headerScore, signature, slugify, matchAddress, SCHEMAS };
  }

  const api = { createConverter, SCHEMAS };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.KiotConverter = api;
})(typeof window !== "undefined" ? window : globalThis);
