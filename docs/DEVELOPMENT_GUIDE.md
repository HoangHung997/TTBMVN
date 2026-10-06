# Hướng dẫn phát triển

## Môi trường yêu cầu

- Windows 11 khuyến nghị.
- Microsoft Excel desktop x64; ma trận hiện đã xác minh trên Office 365 x64.
- Visual Studio có workload Office/SharePoint development và MSBuild.
- .NET Framework 4.8.1 Developer Pack.
- PowerShell 7 hoặc Windows PowerShell 5.1.
- NuGet restore cho `ExcelAddIn1/packages.config`.

Host Office x86, Office cũ và Windows 10 chưa có bằng chứng acceptance đầy đủ; xem
`docs/du-toan/COMPATIBILITY.md` trước khi công bố hỗ trợ.

## Chuẩn bị sau khi clone

1. Mở `ExcelAddIn1.sln` và restore NuGet packages.
2. Đảm bảo `packages/ncalc.1.3.8/lib/NCalc.dll` đã được khôi phục.
3. Chọn `Release|x64` cho bản dùng với Excel x64.
4. Không đưa certificate, private key hoặc file cấu hình người dùng vào workspace.

## Build chuẩn

```powershell
./scripts/build-release.ps1 -Platform x64 -Configuration Release
```

Script tự tìm MSBuild bằng `vswhere`, tạo hoặc lấy chứng thư ký manifest dành cho kênh phát
triển và chạy test Core. Build trực tiếp `.sln` mà không truyền thumbprint có thể lỗi ở target
VSTO `ManageCertificateStore`; đó không phải lỗi biên dịch C#.

## Chạy kiểm thử

```powershell
./ExcelAddIn1.Tests/bin/Release/ExcelAddIn1.Tests.exe
```

Bộ Core hiện có 71 test. Với thay đổi liên quan workbook, phải chạy thêm đúng nhóm script trong
`docs/du-toan/TESTING.md` và kiểm tra trên bản sao workbook chuẩn. Không mở và ghi trực tiếp lên
file baseline gốc.

Các gate quan trọng:

- build không warning/error;
- Core test đạt toàn bộ;
- test rollback và khôi phục trạng thái Excel đạt;
- workbook checkpoint mở lại được, checksum và ô tổng đúng theo baseline;
- log/support package không chứa đường dẫn, tên sheet, Machine ID hoặc dữ liệu khách hàng dạng rõ.

## Quy trình sửa chức năng

1. Đọc `docs/PROJECT_STATUS.md` và task liên quan trong `docs/du-toan/TASKS.md`.
2. Chốt hành vi với QA nếu mục đó còn nằm trong danh sách `LR-*`.
3. Đặt quy tắc tính toán ở Core; lớp VSTO chỉ điều phối và ghi Excel.
4. Bổ sung test nhỏ nhất đủ bảo vệ rủi ro thay đổi.
5. Chạy test Core, test workbook liên quan và kiểm tra thủ công UI khi cần.
6. Cập nhật `PROGRESS.md`, `HANDOFF.md` và `WORKLOG.md` nếu thay đổi trạng thái task hoặc baseline.

## Debug VSTO và cài đặt

- Không chạy đồng thời bản debug Visual Studio và bản ClickOnce đã cài cùng identity.
- Trước khi kiểm tra installer, đọc `docs/du-toan/INSTALLATION.md` và `docs/VSTO_DEPLOYMENT.md`.
- Script cài đặt phát hành xử lý đăng ký VSTO cũ; không dùng `git clean` hay xóa thư mục làm việc
  để xử lý lỗi đăng ký.
- Form dự toán là modeless. Mọi dialog chọn vùng Excel phải giữ state form và không tự reload dữ
  liệu khi form active lại.

## Khóa phát hành

LicenseTool cần private key ECDSA ở ngoài repository. Đường dẫn mặc định:

```text
%LOCALAPPDATA%\TTBMVN Excel Tools\Publisher\license-signing-key-v2.bin
```

Có thể trỏ sang kho bí mật khác bằng biến môi trường `TTBMVN_LICENSE_SIGNING_KEY`. Quy trình đầy
đủ ở `docs/LICENSE_TOOL.md`. Không tạo khóa mới nếu chưa có kế hoạch phát hành lại public key và
cấp lại toàn bộ license.
