/*
 * Lõi chuyển đổi KiotViet -> Sapo. Chạy được cả trong trình duyệt lẫn Node (để test).
 * Quy tắc map giữ đúng như 2 skill kiotviet-sapo-split và kiotviet-khach-hang-cong-no.
 *
 * createConverter({ XLSX, ExcelJS, templates, dmhcCsv })
 *   templates: { donLe, quyDoi, khachHang, congNo } -> Uint8Array/ArrayBuffer của file mẫu Sapo
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
    const n = typeof v === "number" ? v : Number(String(v).replace(/,/g, "").trim());
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
  function slugify(text) {
    if (isEmpty(text)) return null;
    const s = stripMarks(String(text).trim()).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/-+/g, "-").replace(/^-|-$/g, "");
    return s || null;
  }

  // Dò tên cột: khớp không phân biệt hoa thường, bỏ khoảng trắng 2 đầu.
  function findCol(headers, candidates) {
    const map = {};
    headers.forEach((h) => (map[String(h).trim().toLowerCase()] = h));
    for (const c of candidates) {
      const hit = map[c.trim().toLowerCase()];
      if (hit !== undefined) return hit;
    }
    return null;
  }

  function createConverter(opts) {
    const XLSX = opts.XLSX;
    const ExcelJS = opts.ExcelJS;
    const T = opts.templates;
    let dmhc = null; // nạp lười khi cần

    // ---------- Đọc file KiotViet ----------
    function readInput(buf) {
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
      const clean = rows.map((r) => {
        const o = {};
        for (const k in r) o[String(k).trim()] = r[k];
        return o;
      });
      const headerRow = (XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })[0] || []).map((h) => String(h == null ? "" : h).trim());
      // Bỏ các dòng trống hoàn toàn
      const data = clean.filter((r) => Object.values(r).some((v) => !isEmpty(v)));
      return { headers: headerRow.filter((h) => h), rows: data };
    }

    function detect(headers) {
      if (findCol(headers, ["Tên khách hàng"])) return "customer";
      if (findCol(headers, ["Mã hàng"]) && findCol(headers, ["Tên hàng"])) return "product";
      return null;
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

    // =====================================================================
    // SẢN PHẨM  (đơn lẻ + quy đổi)
    // =====================================================================
    async function convertProducts(input) {
      const h = input.headers;
      const cols = {
        code: findCol(h, ["Mã hàng"]),
        name: findCol(h, ["Tên hàng"]),
        brand: findCol(h, ["Thương hiệu"]),
        category: findCol(h, ["Nhóm hàng(3 Cấp)", "Nhóm hàng (3 Cấp)"]),
        barcode: findCol(h, ["Mã vạch"]),
        priceAfterTax: findCol(h, ["Giá bán sau thuế"]),
        pricePlain: findCol(h, ["Giá bán"]),
        priceBeforeTax: findCol(h, ["Giá bán trước thuế"]),
        cost: findCol(h, ["Giá vốn"]),
        stock: findCol(h, ["Tồn kho"]),
        unit: findCol(h, ["ĐVT"]),
        baseRefs: ["Mã ĐVT Cơ bản", "Mã HH Liên quan"].map((c) => findCol(h, [c])).filter(Boolean),
        qty: findCol(h, ["Quy đổi"]),
        vat: findCol(h, ["VAT (%)", "VAT(%)"]),
        description: findCol(h, ["Mô tả"]),
        image: findCol(h, ["Hình ảnh (url1,url2...)", "Hình ảnh"]),
        weight: findCol(h, ["Trọng lượng"]),
      };
      const missing = [["code", "Mã hàng"], ["name", "Tên hàng"], ["qty", "Quy đổi"]].filter(([k]) => !cols[k]).map(([, n]) => n);
      if (missing.length) throw new Error("File thiếu cột bắt buộc: " + missing.join(", ") + ". Hãy xuất lại file “Danh sách hàng hóa” từ KiotViet.");

      // Giá: ưu tiên Giá bán sau thuế -> Giá bán -> Giá bán trước thuế
      const priceCol = cols.priceAfterTax || cols.pricePlain || cols.priceBeforeTax;
      const g = (r, k) => (cols[k] ? val(r[cols[k]]) : null);
      // SKU tham chiếu: lấy giá trị đầu tiên có dữ liệu trong "Mã ĐVT Cơ bản" / "Mã HH Liên quan"
      const baseRef = (r) => {
        for (const c of cols.baseRefs) if (!isEmpty(r[c])) return str(r[c]);
        return null;
      };

      const donLe = [], quyDoi = [];
      for (const r of input.rows) {
        const q = num(r[cols.qty]);
        (q === null || q === 1 ? donLe : quyDoi).push(r);
      }

      const warnings = [];
      if (!cols.baseRefs.length && quyDoi.length) warnings.push("Không có cột “Mã ĐVT Cơ bản” / “Mã HH Liên quan”: cột “SKU tham chiếu” trong file quy đổi để trống.");

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
        if (cols.vat) put(ws1, i, 16, isEmpty(r[cols.vat]) ? "Không" : "Có");
        put(ws1, i, 17, str(r[cols.code]), true);
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
        const baseName = new Map(donLe.map((r) => [str(r[cols.code]), r[cols.name]]));
        const wb2 = await loadTemplate(T.quyDoi);
        const ws2 = wb2.getWorksheet("V2 Mẫu file nhập sản phẩm quy đ");
        clearRows(ws2, 2, 18);
        const seenAlias = new Set();
        quyDoi.forEach((r, idx) => {
          const i = idx + 2;
          const ref = baseRef(r);
          const bn = ref ? baseName.get(ref) : undefined;
          if (!bn) orphans.push({ sku: str(r[cols.code]), name: g(r, "name"), ref: ref || "(trống)" });
          const alias = bn ? slugify(bn) : null;
          put(ws2, i, 1, alias);
          if (!(alias && seenAlias.has(alias))) {
            put(ws2, i, 2, bn || g(r, "name"));
            if (alias) seenAlias.add(alias);
          }
          put(ws2, i, 3, str(r[cols.code]), true);
          put(ws2, i, 4, ref, true);
          put(ws2, i, 7, g(r, "unit"));
          const q = num(r[cols.qty]);
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
    // Khoá so khớp tỉnh: bỏ dấu, bỏ tiền tố, bỏ khoảng trắng/gạch -> "Bà Rịa - Vũng Tàu" = "Bà Rịa-Vũng Tàu", "Thành phố Hồ Chí Minh" = "TP Hồ Chí Minh"
    const tinhKey = (s) => nospace(stripPref(norm(s), PRE_TINH));
    const PRE_HUYEN = ["huyen", "quan", "thi xa", "thanh pho"];
    const PRE_XA = ["xa", "phuong", "thi tran"];
    function stripPref(s, prefixes) {
      for (const p of [...prefixes].sort((a, b) => b.length - a.length)) if (s.startsWith(p + " ")) return s.slice(p.length + 1);
      return s;
    }

    function getDmhc() {
      if (dmhc) return dmhc;
      const lines = String(opts.dmhcCsv).replace(/^﻿/, "").split(/\r?\n/).slice(1);
      const byTinh = new Map(); // tên tỉnh chuẩn -> [{xa, huyen, ...}]
      const tinhIndex = new Map(); // tỉnh chuẩn hoá -> tên tỉnh chuẩn
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
    function matchAddress(khuVuc, tinhHd, xaRaw) {
      const D = getDmhc();
      // "Khu vực giao hàng" dạng "Tỉnh - Huyện"; tên tỉnh có thể chứa " - " (Bà Rịa - Vũng Tàu)
      // nên thử cả chuỗi trước, rồi lần lượt từng vị trí " - " cho tới khi phần bên trái khớp 1 tỉnh.
      let tinh = null, rawHuyen = null;
      const src = !isEmpty(khuVuc) ? String(khuVuc).trim() : !isEmpty(tinhHd) ? String(tinhHd).trim() : null;
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

    async function convertCustomers(input) {
      const h = input.headers;
      const cols = {
        ten: findCol(h, ["Tên khách hàng"]),
        sdt: findCol(h, ["Điện thoại", "SĐT", "Số điện thoại"]),
        diaChi: findCol(h, ["Địa chỉ"]),
        khuVuc: findCol(h, ["Khu vực giao hàng"]),
        xa: findCol(h, ["Phường/Xã", "Phường xã"]),
        gioiTinh: findCol(h, ["Giới tính"]),
        email: findCol(h, ["Email"]),
        ghiChu: findCol(h, ["Ghi chú"]),
        no: findCol(h, ["Nợ cần thu hiện tại", "Nợ hiện tại"]),
        tinhHd: findCol(h, ["Tỉnh/Thành phố (xuất HĐ)"]),
      };
      const missing = [["ten", "Tên khách hàng"], ["sdt", "Điện thoại"]].filter(([k]) => !cols[k]).map(([, n]) => n);
      if (missing.length) throw new Error("File thiếu cột bắt buộc: " + missing.join(", ") + ". Hãy xuất lại file “Danh sách khách hàng” từ KiotViet.");
      const g = (r, k) => (cols[k] ? val(r[cols[k]]) : null);

      // SĐT dùng chung cho cả 2 file; thiếu SĐT -> 0900000001, 0900000002... theo thứ tự dòng
      let seq = 0;
      const phones = input.rows.map((r) => {
        const p = cleanPhone(r[cols.sdt]);
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
        const khuVuc = g(r, "khuVuc");
        const tinhHd = g(r, "tinhHd");
        const [tinh, huyen, xa] = matchAddress(khuVuc, tinhHd, xaRaw);

        const miss = [];
        if (!tinh && (khuVuc || tinhHd)) miss.push("Tỉnh");
        if (tinh && !huyen && khuVuc && String(khuVuc).includes(" - ")) miss.push("Huyện");
        if (tinh && !xa && xaRaw) miss.push("Xã");
        if (miss.length) unmatched.push([ten, [khuVuc || tinhHd, xaRaw].filter(Boolean).join(" / "), miss.join(", ")]);

        const addr = [g(r, "diaChi"), xaRaw].filter((x) => !isEmpty(x)).map((x) => String(x).trim()).join(", ") || null;
        let note = g(r, "ghiChu");
        note = note === null ? null : String(note).trim();
        if (note && note.length > 255) { note = note.slice(0, 255); notesCut++; }
        const gt = g(r, "gioiTinh");
        const gender = gt === null ? null : ({ "Nam": "male", "Nữ": "female" }[String(gt).trim()] || null);
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
      if (cols.no) {
        const wb2 = await loadTemplate(T.congNo);
        const ws2 = wb2.worksheets[0];
        clearRows(ws2, 2, 3);
        let i = 2;
        input.rows.forEach((r, idx) => {
          const n = num(r[cols.no]) || 0;
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
        warnings.push("File không có cột “Nợ cần thu hiện tại” nên không tạo file công nợ.");
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
        issues: unmatched.length ? [{ title: "Địa chỉ chưa khớp danh mục Sapo", columns: ["Khách hàng", "Khu vực / Phường xã (KiotViet)", "Thiếu"], rows: unmatched }] : [],
      };
    }

    async function convert(buf) {
      const input = readInput(buf);
      if (!input.rows.length) throw new Error("File không có dòng dữ liệu nào.");
      const type = detect(input.headers);
      if (type === "product") return convertProducts(input);
      if (type === "customer") return convertCustomers(input);
      throw new Error("Không nhận ra loại file. Cần file “Danh sách hàng hóa” (có cột Mã hàng, Tên hàng) hoặc “Danh sách khách hàng” (có cột Tên khách hàng) xuất từ KiotViet.");
    }

    return { convert, readInput, detect, slugify, matchAddress };
  }

  if (typeof module !== "undefined" && module.exports) module.exports = { createConverter };
  else root.KiotConverter = { createConverter };
})(typeof window !== "undefined" ? window : globalThis);
