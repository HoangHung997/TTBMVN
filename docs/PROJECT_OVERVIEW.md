# Tổng quan dự án

## Mục tiêu

TTBMVN Excel Tools đưa các bước tính toán rà phá bom mìn, vật nổ vào Excel nhưng giữ dữ liệu,
công thức và kết quả có thể kiểm tra trực tiếp trên workbook. Ứng dụng không thay Excel; nó bổ
sung Ribbon, form modeless và các engine có kiểm thử để người dùng vừa thao tác bảng tính vừa
điều khiển quy trình nghiệp vụ.

## Hai miền chức năng

### Xử lý hố đào

Form hố đào sinh bộ `d1`, `r1`, `d2`, `r2`, `H` cho hố 3 m hoặc 5 m sao cho thể tích nằm trong
khoảng yêu cầu. Người dùng có thể:

- đặt cận Min/Max cho kích thước, H hoặc V;
- chọn số luồng tính toán theo logical processor;
- đọc bảng nguồn, ánh xạ các cột chuẩn và cột phụ;
- đổ kết quả sang sheet mới hoặc sheet hiện có;
- ghi tiêu đề, định dạng, công thức, link ngược tổng tín hiệu/tổng thể tích;
- lưu, import và export cấu hình không gắn với vị trí workbook cụ thể.

Engine tính toán chạy ngoài Excel COM; kết quả được ghi theo batch trong một ngữ cảnh khôi phục
`ScreenUpdating`, `Calculation`, `EnableEvents` và `DisplayAlerts` kể cả khi có lỗi.

### Dự toán rà phá bom mìn

Luồng nghiệp vụ chính:

1. Gắn vai trò cho các sheet thay vì phụ thuộc tên sheet.
2. Chọn gói quy định theo ngày hiệu lực và pin package vào workbook.
3. Lập hồ sơ giá VL-NC-M theo đối tượng hưởng lương và nguồn giá.
4. Chọn định mức, điều kiện/hệ số và sinh đơn giá cạn/nước theo đúng ngữ cảnh công tác.
5. Áp đơn giá vào Phụ lục dự toán, giữ công thức khối lượng Excel khi cần.
6. Tạo Tổng hợp kinh phí theo cấu hình giám sát, VAT và các khoản chi phí.
7. Ghi dấu vết căn cứ, kiểm tra sai lệch và hỗ trợ migration/rollback khi đổi package.

## Kiến trúc mã nguồn

`ExcelAddIn1.Core` chứa mô hình, bộ đọc package và engine tính toán thuần .NET. Module này không
tham chiếu Excel Interop và là nơi ưu tiên đặt mọi quy tắc có thể kiểm thử độc lập.

`ExcelAddIn1` là lớp tích hợp VSTO: Ribbon, WinForms, đọc/ghi workbook, metadata sheet, transaction
Excel và chẩn đoán runtime. Các class `Workbook*Service` chuyển dữ liệu giữa workbook và Core.

`ExcelAddIn1.Tests` là test runner không phụ thuộc framework test ngoài. Mỗi test trả về PASS/FAIL
và exit code khác 0 khi lỗi, phù hợp chạy trong script phát hành.

`ExcelAddIn1.LicenseTool` và `ExcelAddIn1.RegulationTool` là công cụ của nhà phát hành. Chúng không
được đóng vào gói cài đặt người dùng cuối.

## Dữ liệu bền vững

- Vai trò sheet lưu trong `Worksheet.CustomProperties`, nên đổi tên sheet không làm mất mapping.
- Hồ sơ dự án, PriceProfile, package pin và audit lưu trong workbook bằng payload có schema.
- Cấu hình hố đào lưu ở Local AppData và có import/export chọn lọc.
- Package quy định có manifest, checksum, nguồn pháp lý và phiên bản dữ liệu.
- Gói cập nhật offline dùng chữ ký bất đối xứng; client chỉ giữ public key.
- License offline từ phiên bản mã nguồn này dùng ECDSA P-256 và gắn với Machine ID.

## Nguyên tắc không được phá vỡ

- Không đọc/ghi Excel COM từ worker thread.
- Không dựa vào tên sheet để xác định vai trò nghiệp vụ sau khi đã gắn mapping.
- Không bỏ qua dòng tính lỗi; phải preview, chặn ghi và trả lỗi có ngữ cảnh.
- Không sửa một phần workbook khi transaction thất bại.
- Không ghi đè tay người dùng mà không có preview hoặc lý do override.
- Không đưa private key, PFX, license người dùng, log khách hàng hoặc package hỗ trợ vào Git.

Chi tiết class, schema và quyết định thiết kế nằm trong
[du-toan/ARCHITECTURE.md](du-toan/ARCHITECTURE.md).
