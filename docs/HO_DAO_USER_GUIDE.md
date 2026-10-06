# Hướng Dẫn Nhanh Chức Năng Hố Đào

## Luồng Khuyến Nghị

1. Mở form `Cài đặt` trong group `Hố đào`.
2. Chọn `Loại hố đào`: 3m hoặc 5m.
3. Nhập min/max cho `d1`, `r1`, `d2`, `r2`, `H` hoặc `V`.
4. Chọn `Table Data` bằng nút `Chọn`.
5. Nhập cột nguồn theo dạng `A;B` cho 3m hoặc `A;B;C;D` cho cả 3m và 5m.
6. Map các cột trong DataGrid:
   - `Tên dữ liệu`: tên hiển thị trên hàng header kết quả.
   - `Cột trong Data`: cột lấy dữ liệu nguồn hoặc `key_STT` để đánh số thứ tự.
   - `Cột trong đào đắp`: cột đổ kết quả sang sheet đích. Bỏ trống nếu không muốn ghi dữ liệu đó.
   - `Định dạng`: double-click để chọn format Excel.
7. Bấm `Lưu và chạy`, đọc preview, xác nhận ghi Excel.

## Các Key Mặc Định

```text
TH3; V3; D1_3; R1_3; D2_3; R2_3; H_3
TH5; V5; D1_5; R1_5; D2_5; R2_5; H_5
```

Nếu chỉ chạy 3m, các key 5m có thể để nguyên hoặc bỏ trống cột đổ kết quả.

## Key Đặc Biệt

`key_STT` trong cột `Cột trong Data` dùng để đánh số thứ tự các dòng tổng. Dòng này cần có `Cột trong đào đắp` để biết vị trí ghi số thứ tự.

## Sheet Và Link Ngược

- `New sheet`: tạo sheet kết quả mới.
- `Sheet hiện có`: ghi vào sheet đã chọn.
- `Link ngược`: link tổng số tín hiệu và tổng V từ sheet kết quả về lại sheet nguồn.

## Hỗ Trợ Khi Lỗi

Trong form hố đào, bấm `Hỗ trợ`. App sẽ tạo một thư mục chẩn đoán gồm:

- `app-info.txt`
- `daodat-settings.json`
- `runtime.log` nếu có

Gửi thư mục này cho người phát triển để kiểm tra lỗi trên máy người dùng.
