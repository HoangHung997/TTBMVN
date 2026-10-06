# TTBMVN Excel Tools

TTBMVN Excel Tools là add-in Microsoft Excel phục vụ tính toán và lập dự toán rà phá bom mìn,
vật nổ. Dự án gồm công cụ xử lý hố đào 3 m/5 m, không gian làm việc dự toán, dữ liệu quy định,
kiểm tra workbook, cập nhật offline, chẩn đoán và kích hoạt theo máy.

Trạng thái hiện tại: **1.0.4 PILOT x64**. Bộ Core có **71/71 kiểm thử đạt**; các bước
`DT-101` đến `DT-701` đã hoàn thành. Trước khi phát hành thương mại vẫn phải chốt các câu hỏi
logic P0/P1 trong đặc tả QA.

## Bắt đầu đọc

1. [Chỉ mục tài liệu](docs/INDEX.md): chọn đúng tài liệu theo vai trò và công việc.
2. [Tổng quan dự án](docs/PROJECT_OVERVIEW.md): phạm vi, luồng nghiệp vụ và cấu trúc mã nguồn.
3. [Trạng thái và việc tiếp theo](docs/PROJECT_STATUS.md): phần đã xong, rủi ro còn mở và điểm tiếp tục.
4. [Hướng dẫn phát triển](docs/DEVELOPMENT_GUIDE.md): chuẩn bị máy, build, test và debug Excel.
5. [Đặc tả chức năng cho QA](docs/QA_FUNCTIONAL_SPEC_1.0.4.md): hành vi cần kiểm thử và các quyết định logic cần chốt.
6. [Hồ sơ điều phối module dự toán](docs/du-toan/README.md): task, checkpoint, baseline và nhật ký chi tiết.

Người tiếp tục công việc phải đọc ít nhất `PROJECT_STATUS.md`, `QA_FUNCTIONAL_SPEC_1.0.4.md`,
`docs/du-toan/PROGRESS.md` và `docs/du-toan/HANDOFF.md` trước khi sửa logic nghiệp vụ.

## Chức năng chính

- Sinh dữ liệu hố đào 3 m/5 m theo ràng buộc hình học, chạy song song và ghi Excel theo batch.
- Xuất sang sheet mới hoặc sheet hiện có, ánh xạ dữ liệu nguồn, định dạng và link ngược tổng hợp.
- Gắn vai trò sheet bền vững khi người dùng đổi tên sheet.
- Quản lý hồ sơ dự toán, gói quy định 2021/2025, định mức, giá VL-NC-M và đơn giá cạn/nước.
- Lập phụ lục dự toán, tổng hợp kinh phí, truy vết căn cứ và kiểm tra sai lệch workbook.
- Cập nhật dữ liệu quy định bằng gói offline có chữ ký số và rollback.
- Tạo gói chẩn đoán đã lọc dữ liệu nhạy cảm; kích hoạt offline gắn với Machine ID.

## Cấu trúc repository

| Thư mục | Nội dung |
|---|---|
| `ExcelAddIn1/` | VSTO add-in, Ribbon, WinForms và lớp tích hợp Excel |
| `ExcelAddIn1.Core/` | Mô hình và engine nghiệp vụ không phụ thuộc Excel |
| `ExcelAddIn1.Tests/` | Bộ kiểm thử Core dạng executable |
| `ExcelAddIn1.LicenseTool/` | Công cụ nội bộ sinh key kích hoạt offline |
| `ExcelAddIn1.RegulationTool/` | Công cụ nội bộ tạo gói cập nhật quy định |
| `data/regulations/` | Nguồn pháp lý, dữ liệu trích xuất và package đã kiểm chứng |
| `Dutoanmau/` | Workbook chuẩn, biến thể và checkpoint kiểm thử |
| `scripts/` | Build, test, kiểm tra phát hành và tự động hóa Excel |
| `packaging/` | Script cài đặt, gỡ cài đặt và xác minh gói phát hành |
| `docs/` | Đặc tả, kiến trúc, QA, vận hành và hồ sơ bàn giao |

Các thư mục `bin`, `obj`, `packages`, `artifacts`, `tmp` và `.vs` là đầu ra cục bộ, không được
đưa vào Git. Gói cài đặt phát hành được tạo lại theo tài liệu release.

## Build và test nhanh

Yêu cầu Windows, Microsoft Excel desktop x64, Visual Studio có workload Office/VSTO và
.NET Framework 4.8.1 Developer Pack.

```powershell
./scripts/build-release.ps1 -Platform x64 -Configuration Release
./ExcelAddIn1.Tests/bin/Release/ExcelAddIn1.Tests.exe
```

Quy trình đầy đủ và các kiểm thử có điều khiển Excel nằm trong
[docs/DEVELOPMENT_GUIDE.md](docs/DEVELOPMENT_GUIDE.md) và
[docs/du-toan/TESTING.md](docs/du-toan/TESTING.md).

## Bảo mật và quyền sử dụng

Repository không chứa private key phát hành, chứng thư ký riêng hay dữ liệu kích hoạt của người
dùng. Không commit các tệp bị loại bởi nhóm `Signing material` trong `.gitignore`.

Mã nguồn được công khai để cộng tác và kiểm thử; repository hiện chưa cấp giấy phép tái sử dụng,
phân phối hoặc thương mại hóa. Xem [SECURITY.md](SECURITY.md) trước khi báo cáo lỗ hổng hoặc xử
lý khóa phát hành.
