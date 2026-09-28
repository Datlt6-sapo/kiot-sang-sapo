# Kiot sang Sapo

Dùng online: https://datlt6-sapo.github.io/kiot-sang-sapo/ (hoặc mở `Kiot sang Sapo.html` trên máy). Thả file KiotViet hoặc file Excel từ phần mềm khác vào, file Sapo tự tải về. Dữ liệu không rời máy.

File lạ: ứng dụng tự tìm dòng tiêu đề và đoán cột. Cột nào chưa chắc thì hiện màn hình ghép cột để bạn chọn, và nhớ cách ghép cho các file cùng kiểu (lưu trong trình duyệt).

| File KiotViet | File Sapo tạo ra |
|---|---|
| Danh sách hàng hóa | `San_pham_don_le`, `San_pham_quy_doi` |
| Danh sách khách hàng | `Khach_hang_Sapo`, `Cong_no_khach_hang_Sapo` |

## Cấu trúc

- `src/convert.js`: toàn bộ quy tắc chuyển đổi. Danh sách cột Sapo và các tên gọi cột quen thuộc nằm trong `SCHEMAS`, thêm tên gọi mới vào `aliases`.
- `src/app.html`: giao diện.
- `assets/`: file mẫu nhập của Sapo và danh mục hành chính `dmhc_63tinh.csv`.
- `vendor/`: SheetJS (đọc file) và ExcelJS (ghi vào mẫu, giữ định dạng).

Sửa xong thì chạy `node build.js` để tạo lại file HTML.

Cập nhật mẫu Sapo hoặc DMHC: thay file trong `assets/` (giữ nguyên tên) rồi build lại.
