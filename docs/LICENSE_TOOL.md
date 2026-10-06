# License Tool

`ExcelAddIn1.LicenseTool` là công cụ nội bộ của nhà phát hành để sinh key offline theo từng
Machine ID. Công cụ không được phân phối trong bộ cài người dùng cuối.

## Mô hình bảo mật

- Key mới có tiền tố `TTB26` và dùng chữ ký ECDSA P-256.
- Add-in chỉ chứa public key để xác minh.
- LicenseTool cần private key nằm ngoài repository để ký.
- Ngày hết hạn và Machine ID nằm trong payload đã ký nhưng được mã hóa Base32, không hiển thị
  trực tiếp dưới dạng ngày dễ sửa.
- Sửa một ký tự, dùng sai máy hoặc thay payload đều làm chữ ký không hợp lệ.

Cơ chế HMAC cũ đã bị loại vì secret đặt trong client cho phép người có mã nguồn tự sinh key. Key
`TTB25` và key legacy phải được cấp lại sau khi triển khai phiên bản dùng cơ chế mới.

## Vị trí private key

Đường dẫn mặc định trên máy phát hành:

```text
%LOCALAPPDATA%\TTBMVN Excel Tools\Publisher\license-signing-key-v2.bin
```

Có thể đặt private key trong kho bí mật khác rồi cấu hình đường dẫn tuyệt đối:

```powershell
$env:TTBMVN_LICENSE_SIGNING_KEY = "D:\Secure\license-signing-key-v2.bin"
```

Private key hiện hành không nằm trong repository. Chỉ chủ phát hành được phép nhận bản sao. Khi
chuyển máy, phải truyền qua kênh mã hóa và kiểm tra quyền truy cập tệp trước khi chạy tool.

Không tự tạo private key mới để thay file bị thiếu. Public key nhúng trong add-in chỉ xác minh đúng
key tương ứng; thay private key yêu cầu phát hành lại add-in và cấp lại toàn bộ license.

## Build

```powershell
./scripts/build-release.ps1 -Platform x64 -Configuration Release
```

Executable sau build:

```text
ExcelAddIn1.LicenseTool\bin\Release\ExcelAddIn1.LicenseTool.exe
```

## Cấp key

1. Người dùng mở form `Hỗ trợ` trong add-in và copy `Machine ID`.
2. Người phát hành mở LicenseTool trên máy có private key.
3. Dán Machine ID, chọn ngày hết hạn và bấm `Tạo key`.
4. Tool ký key rồi tự kiểm tra bằng public key nhúng trong Core.
5. Gửi key cho đúng người dùng; không gửi private key.
6. Người dùng nhập key tại phần `Kích hoạt offline` của form Hỗ trợ.

## Kiểm tra sự cố

- `Không tìm thấy private key phát hành`: kiểm tra đường dẫn mặc định hoặc biến môi trường.
- `Key sinh ra không qua được bước tự kiểm tra`: private key không khớp public key của bản build.
- `License key không hợp lệ cho máy này`: kiểm tra đúng Machine ID, không bị thiếu ký tự và đang
  dùng bản add-in hỗ trợ tiền tố `TTB26`.
- License hết hạn: cấp key mới với cùng Machine ID và ngày hết hạn mới.

Việc sao lưu, xoay và xử lý rò rỉ khóa được mô tả trong `SECURITY.md`.
