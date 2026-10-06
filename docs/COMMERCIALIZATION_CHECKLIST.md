# Commercialization Checklist

## Bản Sẵn Sàng Bán

- Product name: `TTBMVN Excel Tools`.
- Publisher name: `TTBMVN`.
- Release build: `Release|x64`.
- Certificate ký code thật.
- Publish location cố định.
- File hướng dẫn cài đặt, gỡ cài đặt và update.
- File Excel mẫu cho hố đào.
- Quy trình nhận log/support package từ người dùng.

## Kiểm Thử Bắt Buộc

- Windows 10/11.
- Office 2016/2019/2021/Microsoft 365.
- Office 32-bit và 64-bit nếu có khách dùng cả hai.
- Excel dùng dấu thập phân `,`.
- Excel dùng dấu thập phân `.`.
- Workbook có merged cells.
- Workbook có protected sheet.
- Workbook nhiều sheet, đổi tên/xóa/thêm sheet khi form đang mở.

## License

Code đã có nền license offline trong `LicenseManager`.

Trước khi bán cần quyết định:

- Có cho dùng thử không.
- Thời hạn dùng thử.
- Có khóa chức năng khi hết hạn không.
- Quy trình tạo license key cho từng `MachineId`.
- Chính sách đổi máy/cấp lại key.

## Support

Khi người dùng báo lỗi:

1. Yêu cầu bấm `Hỗ trợ` trong form hố đào.
2. Nhận thư mục support được tạo.
3. Kiểm tra `runtime.log`, `app-info.txt`, `daodat-settings.json`.
4. Nếu lỗi liên quan workbook, yêu cầu file mẫu đã xóa dữ liệu nhạy cảm.
