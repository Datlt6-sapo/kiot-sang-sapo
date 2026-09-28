# Kiot sang Sapo

Mở `Kiot sang Sapo.html` bằng Chrome/Edge, thả file KiotViet vào, file Sapo tự tải về. Chạy offline, dữ liệu không rời máy.

| File KiotViet | File Sapo tạo ra |
|---|---|
| Danh sách hàng hóa | `San_pham_don_le`, `San_pham_quy_doi` |
| Danh sách khách hàng | `Khach_hang_Sapo`, `Cong_no_khach_hang_Sapo` |

## Cấu trúc

- `src/convert.js`: toàn bộ quy tắc chuyển đổi (dò cột, tách quy đổi, đối chiếu DMHC, SĐT tạm, công nợ).
- `src/app.html`: giao diện.
- `assets/`: file mẫu nhập của Sapo và danh mục hành chính `dmhc_63tinh.csv`.
- `vendor/`: SheetJS (đọc file) và ExcelJS (ghi vào mẫu, giữ định dạng).

Sửa xong thì chạy `node build.js` để tạo lại file HTML.

Cập nhật mẫu Sapo hoặc DMHC: thay file trong `assets/` (giữ nguyên tên) rồi build lại.
