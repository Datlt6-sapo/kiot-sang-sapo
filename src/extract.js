/*
 * Đọc bảng từ PDF và ảnh chụp, trả về lưới ô (grid) để đi tiếp như file Excel.
 *  - PDF có chữ thật: lấy chữ + toạ độ bằng pdf.js
 *  - PDF scan / ảnh chụp: nhận dạng chữ tiếng Việt bằng Tesseract (chạy trong trình duyệt)
 * Thư viện chỉ tải về khi cần (lần đầu dùng PDF/ảnh), nên cần mạng ở lần đó.
 */
(function (root) {
  "use strict";

  const PDFJS = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js";
  const PDFJS_WORKER = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
  const TESSERACT = "https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js";

  const loaded = {};
  function loadScript(url) {
    if (!loaded[url]) {
      loaded[url] = new Promise((resolve, reject) => {
        const s = document.createElement("script");
        s.src = url;
        s.onload = resolve;
        s.onerror = () => { delete loaded[url]; reject(new Error("Không tải được thư viện đọc PDF/ảnh. Kiểm tra kết nối mạng rồi thử lại.")); };
        document.head.appendChild(s);
      });
    }
    return loaded[url];
  }

  // Nối chữ trong ô: phần trước kết thúc bằng "-" (mã hàng bị xuống dòng "NE50-" + "MHO-35-36") thì nối liền
  const joinText = (a, b) => (a ? a + (/-$/.test(a) ? "" : " ") + b : b);

  const median = (arr) => {
    if (!arr.length) return 0;
    const s = [...arr].sort((a, b) => a - b);
    return s[Math.floor(s.length / 2)];
  };

  const UNITS = ["Đôi", "Cái", "Chiếc", "Hộp", "Gói", "Chai", "Lon", "Thùng", "Túi", "Kg", "Bộ", "Lốc", "Bao", "Cây", "Hũ", "Lọ", "Tuýp", "Vỉ", "Viên", "Ống", "Can", "Két", "Cuộn", "Tờ", "Quyển", "Cuốn", "Mét", "Lít", "Bịch", "Bình", "Chục", "Cặp", "Tấm", "Sợi", "Bó", "Khay", "Hộc", "Xấp"];

  // =====================================================================
  // Dựng bảng từ các "từ" có toạ độ: {text, x0, x1, y0, y1, conf}
  // y đã cộng dồn theo trang nên nhiều trang nối liền nhau.
  // vlines: toạ độ x các đường kẻ dọc của bảng (nếu có kẻ ô) -> ranh giới cột
  // hlines: toạ độ y các đường kẻ ngang (nếu có) -> ranh giới dòng, để ô chữ xuống 2-3 dòng vẫn là 1 ô
  // =====================================================================
  function buildTable(words, headerScore, vlines, hlines) {
    words = words.filter((w) => w.text && w.text.trim());
    if (!words.length) return { grid: [], conf: [] };
    const h = median(words.map((w) => w.y1 - w.y0)) || 10;
    const ycOf = (w) => (w.y0 + w.y1) / 2;

    // 1) Gom từ thành dòng theo toạ độ dọc
    const groupLines = (ws) => {
      const out = [];
      for (const w of [...ws].sort((a, b) => ycOf(a) - ycOf(b))) {
        const cur = out[out.length - 1];
        if (cur && Math.abs(ycOf(w) - cur.yc) < h * 0.55) {
          cur.words.push(w);
          cur.yc = cur.words.reduce((s, x) => s + ycOf(x), 0) / cur.words.length;
        } else out.push({ yc: ycOf(w), words: [w] });
      }
      for (const l of out) l.words.sort((a, b) => a.x0 - b.x0);
      return out;
    };
    // Mỗi "dòng bảng" có thể gồm nhiều dòng chữ (sub). Có đường kẻ ngang thì 1 dòng bảng = 1 dải giữa 2 đường kẻ.
    let lines;
    const hs = (hlines || []).slice().sort((a, b) => a - b).filter((y, i, arr) => i === 0 || y - arr[i - 1] > h * 0.6);
    if (hs.length >= 3) {
      const bands = [];
      for (let i = 0; i < hs.length - 1; i++) if (hs[i + 1] - hs[i] > h * 0.8) bands.push([hs[i], hs[i + 1]]);
      const inBand = bands.map(() => []), outside = [];
      for (const w of words) {
        const y = ycOf(w);
        const b = bands.findIndex(([a, z]) => y > a && y < z);
        (b >= 0 ? inBand[b] : outside).push(w);
      }
      // Dải quá cao (nhiều dòng hàng không kẻ ngăn giữa) thì không coi là 1 dòng bảng: chia theo dòng chữ như thường
      const tall = bands.map(([a, z]) => z - a > h * 7);
      lines = [
        ...groupLines(outside).map((l) => ({ yc: l.yc, sub: [l.words] })),
        ...bands.flatMap(([a, z], i) => tall[i]
          ? groupLines(inBand[i]).map((l) => ({ yc: l.yc, sub: [l.words] }))
          : inBand[i].length ? [{ yc: (a + z) / 2, sub: groupLines(inBand[i]).map((l) => l.words) }] : []),
      ].sort((a, b) => a.yc - b.yc);
    } else {
      lines = groupLines(words).map((l) => ({ yc: l.yc, sub: [l.words] }));
    }
    for (const l of lines) l.words = l.sub.flat();
    const banded = hs.length >= 3;

    // Ô trong 1 dòng chữ: nối các từ đứng sát nhau
    const cellsOfSub = (ws) => {
      const cells = [];
      for (const w of ws) {
        const c = cells[cells.length - 1];
        if (c && w.x0 - c.x1 < h * 0.9) { c.text += " " + w.text; c.x1 = Math.max(c.x1, w.x1); }
        else cells.push({ text: w.text, x0: w.x0, x1: w.x1, conf: w.conf });
      }
      return cells;
    };
    const cellsOf = (l) => l.sub.flatMap(cellsOfSub);
    // Chữ đủ tin cậy (PDF luôn tin cậy): chữ rác do vết bẩn/đường kẻ không được làm hỏng việc chia cột
    const solid = (w) => w.conf === null || w.conf === undefined || w.conf >= 60;

    // 2) Tìm dòng tiêu đề: nhiều ô trùng tên cột quen thuộc nhất; không có thì lấy dòng nhiều ô nhất
    let hi = 0, best = -1;
    const scan = Math.min(lines.length, 40);
    for (let i = 0; i < scan; i++) {
      const sc = headerScore(cellsOf(lines[i]).map((c) => c.text));
      if (sc > best) { best = sc; hi = i; }
    }
    if (best <= 0) {
      best = -1;
      for (let i = 0; i < scan; i++) {
        const n = cellsOf(lines[i]).length;
        if (n > best) { best = n; hi = i; }
      }
    }
    // Tiêu đề 2 tầng ("Giá" / "nhập" trên 2 dòng sát nhau) -> gộp dòng ngay dưới nếu nó không có số
    const headLines = [lines[hi]];
    const nx = lines[hi + 1];
    if (!banded && nx && nx.yc - lines[hi].yc < h * 1.3 && !nx.words.some((w) => /\d/.test(w.text))) headLines.push(nx);
    const bodyLines = lines.slice(hi + headLines.length);
    const headerWords = headLines.flatMap((l) => l.words);

    // 3) Ranh giới cột
    const allX0 = Math.min(...[...headerWords, ...bodyLines.flatMap((l) => l.words)].map((w) => w.x0));
    const allX1 = Math.max(...[...headerWords, ...bodyLines.flatMap((l) => l.words)].map((w) => w.x1));
    // a) Đường kẻ dọc nằm trong vùng chữ (bảng có thể chỉ kẻ vài cột)
    // Chỉ lấy đường kẻ nằm giữa chữ đầu và chữ cuối của dòng tiêu đề (bỏ viền ngoài bảng)
    const hx0 = Math.min(...headerWords.map((w) => w.x0)), hx1 = Math.max(...headerWords.map((w) => w.x1));
    const lineBounds = (vlines || []).filter((x) => x > hx0 + h * 0.3 && x < hx1 - h * 0.3);
    let bounds = [];
    // b) Dải dọc trống xuyên suốt các dòng (cho các cột không có đường kẻ).
    //    Dòng thân bảng tính theo từng chữ; dòng tiêu đề tính theo cả ô (để "Giá nhập" không bị tách đôi).
    {
      const body = bodyLines.filter((l) => cellsOf(l).length >= 2);
      const bin = Math.max(1, h / 4);
      const n = Math.ceil((allX1 - allX0) / bin) + 1;
      const cov = new Array(n).fill(0);
      const noise = Math.max(0, Math.floor(body.length * 0.08));
      const markLine = (spans, weight) => {
        const mark = new Uint8Array(n);
        for (const c of spans) for (let b = Math.floor((c.x0 - allX0) / bin); b <= Math.floor((c.x1 - allX0) / bin); b++) mark[b] = 1;
        for (let b = 0; b < n; b++) cov[b] += mark[b] * weight;
      };
      for (const l of headLines) markLine(l.sub.flatMap((ws) => cellsOfSub(ws.filter(solid))), noise + 1);
      for (const l of body) markLine(l.words.filter(solid), 1);
      let startGap = -1;
      for (let b = 0; b <= n; b++) {
        const empty = b < n && cov[b] <= noise;
        if (empty && startGap < 0) startGap = b;
        if (!empty && startGap >= 0) {
          if ((b - startGap) * bin >= h * 0.35 && startGap > 0) bounds.push(allX0 + ((startGap + b) / 2) * bin);
          startGap = -1;
        }
      }
    }
    // Bảng kẻ đủ cột (>= 3 đường kẻ dọc bên trong) -> tin đường kẻ. Kẻ thiếu -> gộp thêm ranh giới theo khoảng trống.
    if (lineBounds.length >= 3) bounds = lineBounds.slice().sort((a, b) => a - b);
    else bounds = [...lineBounds, ...bounds.filter((b) => !lineBounds.some((x) => Math.abs(x - b) < h * 1.5))].sort((a, b) => a - b);
    const colOf = (w) => {
      const xc = (w.x0 + w.x1) / 2;
      let c = 0;
      while (c < bounds.length && xc > bounds[c]) c++;
      return c;
    };
    let ncol = bounds.length + 1;

    // 4) Xếp chữ vào ô
    // Trong 1 ô: xếp chữ theo dòng (dựa trên mép trên của chữ) rồi từ trái sang phải,
    // để ô xuống 2-3 dòng đọc đúng thứ tự
    // Chữ cùng dòng = khoảng dọc chồng lên nhau đủ nhiều (chữ có dấu cao như "Dầu" vẫn cùng dòng với "1")
    const orderCell = (ws) => {
      const rowsOf = [];
      for (const w of [...ws].sort((a, b) => ycOf(a) - ycOf(b))) {
        const r = rowsOf.find((r) => Math.min(r.y1, w.y1) - Math.max(r.y0, w.y0) > 0.4 * Math.min(r.y1 - r.y0, w.y1 - w.y0));
        if (r) { r.ws.push(w); r.y0 = Math.min(r.y0, w.y0); r.y1 = Math.max(r.y1, w.y1); }
        else rowsOf.push({ y0: w.y0, y1: w.y1, ws: [w] });
      }
      return rowsOf.sort((a, b) => a.y0 + a.y1 - b.y0 - b.y1).flatMap((r) => r.ws.sort((a, b) => a.x0 - b.x0));
    };
    const place = (ws) => {
      const row = Array(ncol).fill(null), rc = Array(ncol).fill(null);
      const per = Array.from({ length: ncol }, () => []);
      for (const w of ws) per[colOf(w)].push(w);
      per.forEach((cw, c) => {
        for (const w of orderCell(cw)) {
          row[c] = joinText(row[c], w.text);
          if (w.conf !== undefined && w.conf !== null) rc[c] = rc[c] === null ? w.conf : Math.min(rc[c], w.conf);
        }
      });
      return { row, rc };
    };
    const headerLineWords = headLines.map((l) => l.words);
    const head = Array(ncol).fill(null);
    for (const ws of headerLineWords) {
      const p = place(ws).row;
      p.forEach((v, c) => { if (v) head[c] = head[c] ? head[c] + " " + v : v; });
    }
    const headerText = cellsOf(lines[hi]).map((c) => c.text).join("|");
    let rows = [];
    for (const l of bodyLines) {
      if (cellsOf(l).map((c) => c.text).join("|") === headerText) continue; // tiêu đề lặp lại mỗi trang
      const { row, rc } = place(l.words);
      rows.push({ row, rc, yc: l.yc });
    }

    // 5) Gộp các cột thừa: 2 cột kề nhau mà không dòng nào (kể cả tiêu đề) có chữ ở cả hai
    //    (vd số canh phải tách khỏi tiêu đề canh trái của chính cột đó)
    for (let c = 0; c < ncol - 1; ) {
      const all = [head, ...rows.map((r) => r.row)];
      const clash = all.some((r) => r[c] && r[c + 1]);
      if (!clash) {
        const merge = (arr, isConf) => {
          const a = arr[c], b = arr[c + 1];
          arr[c] = isConf ? (a === null ? b : b === null ? a : Math.min(a, b)) : a || b;
          arr.splice(c + 1, 1);
        };
        merge(head);
        for (const r of rows) { merge(r.row); merge(r.rc, true); }
        ncol--;
      } else c++;
    }

    // 6) Chữ xuống dòng trong ô (tên hàng dài chia 2-3 dòng, mã hàng nằm ở dòng giữa):
    //    dòng "mảnh" = ít ô có chữ. Nếu các ô đó của dòng ngay dưới đang trống thì ghép xuống;
    //    nếu dòng mảnh chỉ có 1 ô và dòng ngay trên có chữ ở ô đó thì nối vào dòng trên.
    const filledOf = (r) => r.row.map((v, i) => (v ? i : -1)).filter((i) => i >= 0);
    const joinConf = (a, b) => (a === null ? b : b === null ? a : Math.min(a, b));
    const isFrag = (r) => { const f = filledOf(r).length; return f > 0 && f <= Math.max(1, Math.floor(ncol / 3)); };
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!isFrag(r)) continue;
      const f = filledOf(r);
      const next = rows[i + 1], prev = rows[i - 1];
      if (next && !isFrag(next) && next.yc - r.yc < h * 1.9 && f.every((c) => !next.row[c])) {
        for (const c of f) { next.row[c] = r.row[c]; next.rc[c] = joinConf(next.rc[c], r.rc[c]); }
        rows.splice(i, 1); i--;
      } else if (prev && !isFrag(prev) && f.length === 1 && prev.row[f[0]] && r.yc - prev.yc < h * 1.9) {
        const c = f[0];
        prev.row[c] = joinText(prev.row[c], r.row[c]);
        prev.rc[c] = joinConf(prev.rc[c], r.rc[c]);
        rows.splice(i, 1); i--;
      }
    }

    // 7a) Đơn vị tính đọc sai dấu ("Đối", "Đồi" -> "Đôi"): so theo chữ bỏ dấu với danh sách đơn vị thường gặp
    const ukey = (t) => String(t).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[đĐ]/g, "d").toLowerCase().replace(/[^a-z0-9]/g, "");
    const unitCol = head.findIndex((t) => t && ["dvt", "donvi", "donvitinh"].includes(ukey(t)));
    if (unitCol >= 0) {
      for (const r of rows) {
        const v = r.row[unitCol];
        if (!v || r.rc[unitCol] === null) continue;
        const std = UNITS.find((u) => ukey(u) === ukey(v));
        if (std && std !== v) r.row[unitCol] = std;
      }
    }

    // 7) Số tiền đọc từ ảnh: ký hiệu "đ" hay bị đọc thành #, 4, s, 2, : ở cuối ("61,000đ" -> "61.000#").
    //    Số có dạng nhóm 3 chữ số mà dư 1 ký tự cuối -> bỏ ký tự đó và hạ độ tin cậy để người dùng soát.
    for (const r of rows) {
      r.row.forEach((v, c) => {
        if (!v || r.rc[c] === null) return;
        const m = String(v).match(/^\D{0,2}?(\d{1,3}(?:[.,]\d{3})+)\s?(?:[^\d\s]|\d)[^\d\s]?$/);
        if (m && m[1] !== v) { r.row[c] = m[1]; r.rc[c] = Math.min(r.rc[c], 65); }
        const plain = String(v).match(/^(\d{1,3}(?:[.,]\d{3})+)[đdĐ₫]$/);
        if (plain) r.row[c] = plain[1];
      });
    }

    return { grid: [head, ...rows.map((r) => r.row)], conf: [head.map(() => null), ...rows.map((r) => r.rc)] };
  }

  // =====================================================================
  // PDF
  // =====================================================================
  async function loadPdf(buf) {
    await loadScript(PDFJS);
    const lib = root.pdfjsLib;
    lib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER;
    return lib.getDocument({ data: buf }).promise;
  }

  async function pdfTextWords(pdf, onProgress) {
    const words = [];
    let offset = 0;
    for (let p = 1; p <= pdf.numPages; p++) {
      onProgress && onProgress(`Đang đọc trang ${p}/${pdf.numPages}`);
      const page = await pdf.getPage(p);
      const vp = page.getViewport({ scale: 1 });
      const tc = await page.getTextContent();
      for (const it of tc.items) {
        if (!it.str || !it.str.trim()) continue;
        const [a, b, c, d, e, f] = it.transform;
        const hgt = it.height || Math.hypot(c, d) || 10;
        const y1 = offset + (vp.height - f);
        const y0 = y1 - hgt;
        // 1 đoạn chữ có thể chứa nhiều cột cách nhau bằng nhiều khoảng trắng -> tách ra, chia toạ độ theo số ký tự
        const parts = it.str.split(/(\s{2,})/);
        const total = it.str.length || 1;
        let pos = 0;
        for (const part of parts) {
          const x0 = e + (it.width * pos) / total;
          pos += part.length;
          if (!part.trim()) continue;
          const x1 = e + (it.width * pos) / total;
          words.push({ text: part.trim(), x0, x1, y0, y1, conf: null });
        }
      }
      offset += vp.height + 20;
    }
    return words;
  }

  // =====================================================================
  // OCR (ảnh chụp, PDF scan)
  // =====================================================================
  let workerP = null;
  let progressCb = null;
  async function getWorker() {
    if (!workerP) {
      workerP = (async () => {
        await loadScript(TESSERACT);
        const w = await root.Tesseract.createWorker("vie", 1, {
          logger: (m) => {
            if (!progressCb) return;
            if (m.status === "recognizing text") progressCb(`Đang nhận dạng chữ ${Math.round(m.progress * 100)}%`);
            else if (/load|initializ/i.test(m.status)) progressCb("Đang tải bộ nhận dạng tiếng Việt (chỉ lần đầu)…");
          },
        });
        // Chế độ đọc (PSM) đặt theo từng lượt trong ocrCanvas. Mặc định của tesseract.js (PSM 6) đọc bảng rất sai.
        await w.setParameters({ preserve_interword_spaces: "1" });
        return w;
      })().catch((e) => { workerP = null; throw e; });
    }
    return workerP;
  }

  // Chuẩn bị ảnh trước khi nhận dạng:
  //  1) chỉ phóng to ảnh thật nhỏ (phóng to ảnh mờ làm kết quả tệ hơn)
  //  2) xoá đường kẻ bảng dài, vì đường kẻ làm bộ nhận dạng đọc ra ký tự rác
  //  3) xoay thẳng ảnh chụp bị nghiêng để các ô cùng hàng nằm trên 1 dòng
  function toCanvas(img, w, h) {
    const scale = w < 1000 ? Math.min(2, 1600 / Math.max(w, 1)) : 1;
    const cv = document.createElement("canvas");
    cv.width = Math.round(w * scale);
    cv.height = Math.round(h * scale);
    const ctx = cv.getContext("2d", { willReadFrequently: true });
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, cv.width, cv.height);
    ctx.drawImage(img, 0, 0, cv.width, cv.height);
    const { angle, vx, hy } = preprocess(ctx, cv.width, cv.height);
    // Toạ độ đường kẻ sau khi xoay thẳng (xoay quanh tâm ảnh)
    const rad = (-angle * Math.PI) / 180, cx = cv.width / 2, cy = cv.height / 2;
    const vlines = vx.map(([x, y]) => cx + (x - cx) * Math.cos(rad) - (y - cy) * Math.sin(rad));
    const hlines = hy.map(([x, y]) => cy + (x - cx) * Math.sin(rad) + (y - cy) * Math.cos(rad));
    if (Math.abs(angle) < 0.15) return { cv, vlines, hlines };
    const out = document.createElement("canvas");
    out.width = cv.width; out.height = cv.height;
    const o = out.getContext("2d");
    o.fillStyle = "#fff"; o.fillRect(0, 0, out.width, out.height);
    o.translate(out.width / 2, out.height / 2);
    o.rotate((-angle * Math.PI) / 180);
    o.drawImage(cv, -cv.width / 2, -cv.height / 2);
    return { cv: out, vlines, hlines };
  }

  // Trả về góc nghiêng (độ) + các đường kẻ dọc [x, yGiữa], và ghi ảnh xám đã xoá đường kẻ vào ctx
  function preprocess(ctx, W, H) {
    const im = ctx.getImageData(0, 0, W, H);
    const d = im.data;
    const g = new Uint8Array(W * H);
    const hist = new Array(256).fill(0);
    for (let i = 0, j = 0; j < g.length; i += 4, j++) {
      const v = (d[i] * 299 + d[i + 1] * 587 + d[i + 2] * 114) / 1000 | 0;
      g[j] = v; hist[v]++;
    }
    // Ngưỡng Otsu để tách nét chữ/đường kẻ khỏi nền (chỉ dùng để phân tích, ảnh gửi đi vẫn là ảnh xám)
    let sum = 0; for (let t = 0; t < 256; t++) sum += t * hist[t];
    let sB = 0, wB = 0, best = 0, thr = 128;
    for (let t = 0; t < 256; t++) {
      wB += hist[t]; if (!wB) continue;
      const wF = g.length - wB; if (!wF) break;
      sB += t * hist[t];
      const mB = sB / wB, mF = (sum - sB) / wF;
      const between = wB * wF * (mB - mF) * (mB - mF);
      if (between > best) { best = between; thr = t; }
    }
    const dark = new Uint8Array(W * H);
    for (let j = 0; j < g.length; j++) dark[j] = g[j] <= thr ? 1 : 0;
    // Đường kẻ bảng thường là màu xám nhạt, sáng hơn ngưỡng Otsu -> dò đường kẻ bằng ngưỡng riêng:
    // chỉ cần tối hơn màu nền (màu sáng phổ biến nhất) một chút
    let bgv = 255, bgc = -1;
    for (let t = 128; t < 256; t++) if (hist[t] > bgc) { bgc = hist[t]; bgv = t; }
    const lineThr = Math.max(thr, bgv - 30);
    const ld = new Uint8Array(W * H);
    for (let j = 0; j < g.length; j++) ld[j] = g[j] < lineThr ? 1 : 0;

    // Góc nghiêng: thử -4°..4°, góc nào làm các dòng chữ "dồn" thành vạch ngang rõ nhất thì chọn
    const pts = [];
    for (let y = 0; y < H; y += 2) for (let x = 0; x < W; x += 2) if (dark[y * W + x]) pts.push(x, y);
    let angle = 0, bestScore = -1;
    if (pts.length > 200) {
      for (let a = -4; a <= 4.001; a += 0.2) {
        const t = Math.tan((a * Math.PI) / 180);
        const bins = new Map();
        for (let k = 0; k < pts.length; k += 2) {
          const b = Math.round((pts[k + 1] - pts[k] * t) / 2);
          bins.set(b, (bins.get(b) || 0) + 1);
        }
        let sc = 0; for (const v of bins.values()) sc += v * v;
        if (sc > bestScore) { bestScore = sc; angle = a; }
      }
    }

    // Đường kẻ: vệt tối liền rất dài. Cho phép lệch ±1 điểm ảnh để bắt cả đường hơi nghiêng.
    const erase = new Uint8Array(W * H);
    const minH = Math.max(60, W * 0.06), minV = Math.max(60, H * 0.05);
    const hrows = []; // [y, độ dài, x giữa] của các vệt ngang dài
    for (let y = 1; y < H - 1; y++) {
      let run = 0, total = 0, xsum = 0;
      for (let x = 0; x <= W; x++) {
        const on = x < W && (ld[y * W + x] || ld[(y - 1) * W + x] || ld[(y + 1) * W + x]);
        if (on) { run++; continue; }
        if (run >= minH) { for (let k = x - run; k < x; k++) erase[y * W + k] = 1; total += run; xsum += (x - run / 2) * run; }
        run = 0;
      }
      if (total) hrows.push([y, total, xsum / total]);
    }
    // Đường kẻ ngang: gom các hàng điểm ảnh sát nhau, giữ đường dài >= 50% đường dài nhất
    const hg = [];
    for (const v of hrows) {
      const gr = hg[hg.length - 1];
      if (gr && v[0] - gr.ye <= 3) { gr.ye = v[0]; gr.items.push(v); } else hg.push({ ye: v[0], items: [v] });
    }
    const hl = hg.map((gr) => {
      const top = gr.items.reduce((a, b) => (b[1] > a[1] ? b : a));
      return { y: gr.items.reduce((s, v) => s + v[0], 0) / gr.items.length, len: top[1], x: top[2] };
    });
    const maxH = Math.max(0, ...hl.map((v) => v.len));
    const hy = hl.filter((v) => v.len >= maxH * 0.5).map((v) => [v.x, v.y]);
    const vcols = []; // [x, độ dài, y giữa] của các vệt dọc dài
    for (let x = 1; x < W - 1; x++) {
      let run = 0, total = 0, ysum = 0;
      for (let y = 0; y <= H; y++) {
        const on = y < H && (ld[y * W + x] || ld[y * W + x - 1] || ld[y * W + x + 1]);
        if (on) { run++; continue; }
        if (run >= minV) { for (let k = y - run; k < y; k++) erase[k * W + x] = 1; total += run; ysum += (y - run / 2) * run; }
        run = 0;
      }
      if (total) vcols.push([x, total, ysum / total]);
    }
    // Gom các cột điểm ảnh sát nhau thành 1 đường kẻ; giữ đường đủ dài (>= 35% đường dài nhất)
    const groups = [];
    for (const v of vcols) {
      const gr = groups[groups.length - 1];
      if (gr && v[0] - gr.xe <= 3) { gr.xe = v[0]; gr.items.push(v); } else groups.push({ xe: v[0], items: [v] });
    }
    const vl = groups.map((gr) => {
      const top = gr.items.reduce((a, b) => (b[1] > a[1] ? b : a));
      return { x: gr.items.reduce((s, v) => s + v[0], 0) / gr.items.length, len: top[1], y: top[2] };
    });
    const maxLen = Math.max(0, ...vl.map((v) => v.len));
    const vx = vl.filter((v) => v.len >= maxLen * 0.35).map((v) => [v.x, v.y]);
    // Tô màu nền lên chỗ đường kẻ đã xoá
    for (let j = 0, i = 0; j < g.length; j++, i += 4) {
      const v = erase[j] && ld[j] ? bgv : g[j];
      d[i] = d[i + 1] = d[i + 2] = v; d[i + 3] = 255;
    }
    ctx.putImageData(im, 0, 0);
    return { angle, vx, hy };
  }

  // Đọc 2 lượt rồi gộp: PSM 11 (mẩu chữ rời, hợp với bảng) + PSM 3 (theo khối, bắt được chữ lượt 1 bỏ sót)
  async function ocrCanvas(cv, yOffset, label, onProgress) {
    const worker = await getWorker();
    // Bỏ "chữ" chỉ toàn dấu câu (thường là vết đường kẻ, bụi ảnh)
    const toWords = (data) => (data.words || []).filter((w) => w.text && /[\p{L}\p{N}]/u.test(w.text)).map((w) => ({
      text: w.text.trim(),
      x0: w.bbox.x0, x1: w.bbox.x1,
      y0: w.bbox.y0 + yOffset, y1: w.bbox.y1 + yOffset,
      conf: Math.round(w.confidence),
    }));
    const run = async (psm, step) => {
      progressCb = (msg) => onProgress && onProgress(`${label ? label + " · " : ""}Lượt ${step}/2 · ${msg}`);
      await worker.setParameters({ tessedit_pageseg_mode: psm });
      const { data } = await worker.recognize(cv);
      return toWords(data);
    };
    const a = await run("11", 1);
    const b = await run("3", 2);
    progressCb = null;
    const overlap = (p, q) => {
      const ix = Math.min(p.x1, q.x1) - Math.max(p.x0, q.x0), iy = Math.min(p.y1, q.y1) - Math.max(p.y0, q.y0);
      if (ix <= 0 || iy <= 0) return 0;
      return (ix * iy) / Math.min((p.x1 - p.x0) * (p.y1 - p.y0), (q.x1 - q.x0) * (q.y1 - q.y0));
    };
    // Chỉ thêm chữ của lượt 2 ở chỗ lượt 1 chưa có gì, và chữ đó phải đủ tin cậy
    for (const w of b) if (w.conf >= 50 && !a.some((x) => overlap(x, w) > 0.3)) a.push(w);
    return a;
  }

  async function loadImage(file) {
    try {
      return await createImageBitmap(file);
    } catch (e) {
      const url = URL.createObjectURL(file);
      try {
        return await new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = rej; im.src = url; });
      } catch (e2) {
        throw new Error(/heic|heif/i.test(file.name + file.type)
          ? "Trình duyệt này không mở được ảnh HEIC của iPhone. Hãy chụp màn hình ảnh đó, hoặc đổi sang JPG rồi thả lại."
          : "Không mở được ảnh này. Hãy dùng ảnh JPG hoặc PNG.");
      } finally { URL.revokeObjectURL(url); }
    }
  }

  // =====================================================================
  // API
  // =====================================================================
  async function fromPdf(file, headerScore, onProgress) {
    const pdf = await loadPdf(new Uint8Array(await file.arrayBuffer()));
    const words = await pdfTextWords(pdf, onProgress);
    const chars = words.reduce((s, w) => s + w.text.length, 0);
    if (chars >= 30) return { method: "pdf", pages: pdf.numPages, ...buildTable(words, headerScore) };

    // PDF scan (không có chữ): vẽ từng trang ra ảnh rồi nhận dạng
    const all = [];
    let offset = 0, vl = [], hl = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const vp = page.getViewport({ scale: 2 });
      const raw = document.createElement("canvas");
      raw.width = vp.width; raw.height = vp.height;
      await page.render({ canvasContext: raw.getContext("2d"), viewport: vp }).promise;
      const { cv, vlines, hlines } = toCanvas(raw, raw.width, raw.height);
      if (p === 1) vl = vlines; // các trang cùng mẫu bảng -> dùng đường kẻ dọc của trang đầu
      hl.push(...hlines.map((y) => y + offset));
      all.push(...(await ocrCanvas(cv, offset, `Trang ${p}/${pdf.numPages}`, onProgress)));
      offset += cv.height + 40;
    }
    return { method: "ocr", pages: pdf.numPages, ...buildTable(all, headerScore, vl, hl) };
  }

  async function fromImage(file, headerScore, onProgress) {
    onProgress && onProgress("Đang mở ảnh…");
    const img = await loadImage(file);
    const { cv, vlines, hlines } = toCanvas(img, img.width, img.height);
    const words = await ocrCanvas(cv, 0, "", onProgress);
    return { method: "ocr", pages: 1, ...buildTable(words, headerScore, vlines, hlines) };
  }

  const isPdf = (f) => /\.pdf$/i.test(f.name) || f.type === "application/pdf";
  const isImage = (f) => /^image\//.test(f.type) || /\.(jpe?g|png|webp|bmp|gif|heic|heif)$/i.test(f.name);

  root.KiotExtract = { fromPdf, fromImage, isPdf, isImage, buildTable, _toCanvas: toCanvas };
})(typeof window !== "undefined" ? window : globalThis);
