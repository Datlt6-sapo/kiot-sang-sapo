// Đóng gói ứng dụng thành 1 file HTML duy nhất (thư viện + file mẫu Sapo nhúng sẵn, chạy offline).
// Chạy: node build.js
const fs = require("fs");
const path = require("path");
const r = (p) => fs.readFileSync(path.join(__dirname, p));
const inline = (s) => s.replace(/<\/script/gi, "<\\/script");

const assets = {
  donLe: r("assets/mau_san_pham_don_le.xlsx").toString("base64"),
  quyDoi: r("assets/mau_san_pham_quy_doi.xlsx").toString("base64"),
  khachHang: r("assets/mau_khach_hang_sapo.xlsx").toString("base64"),
  congNo: r("assets/mau_cong_no_sapo.xlsx").toString("base64"),
  dmhc: r("assets/dmhc_63tinh.csv").toString("base64"),
};
const libs = inline(r("vendor/xlsx.full.min.js").toString()) + "\n" + inline(r("vendor/exceljs.min.js").toString());

const html = r("src/app.html").toString()
  .replace("/*__LIBS__*/", () => libs)
  .replace("/*__CONVERT__*/", () => inline(r("src/convert.js").toString()))
  .replace("/*__ASSETS__*/", () => JSON.stringify(assets));

// "Kiot sang Sapo.html" để mở trên máy, "index.html" cho GitHub Pages
for (const name of ["Kiot sang Sapo.html", "index.html"]) {
  const out = path.join(__dirname, name);
  fs.writeFileSync(out, html);
  console.log("Đã tạo", out, (html.length / 1024 / 1024).toFixed(2) + " MB");
}
