# Kiểm chứng khi tiếp quản trên Windows

## 2026-10-06 — Build và test Core

Nguồn tiếp quản: `origin/main` tại commit `37d60dd133d2`.
Đã đọc HANDOVER, PROGRESS, UI-CONTRACT, ARCHITECTURE, MIGRATION-PLAN và README.

Môi trường: Windows, Visual Studio 18 Enterprise / MSBuild 18.10.1,
.NET Framework 4.8.1, cấu hình Release x64, ký manifest bằng chứng thư PILOT.

Lệnh kiểm chứng:

```powershell
./scripts/build-release.ps1 -Platform x64 -Configuration Release
```

Build ban đầu phát hiện ba lỗi biên dịch:

| Lỗi | Nguyên nhân | Sửa chữa |
|---|---|---|
| CS0234 | Core dùng System.Xml.XmlException nhưng thiếu assembly reference | Thêm reference System.Xml |
| CS0104 | CustomTaskPane trùng giữa Office.Core và Office.Tools | Alias rõ Microsoft.Office.Tools.CustomTaskPane |
| CS0102 | RateSheetRow có property và factory cùng tên NormCode | Đổi factory thành NormCodeRow và cập nhật nơi gọi |

Kết quả sau sửa:

- Build solution Release x64 thành công; không có warning/error trong output MSBuild.
- Cả 80 test trong ExcelAddIn1.Tests chạy PASS, gồm các test V2 state, fingerprint,
  resource plan, price projection, rate plan, cost links, validation, compatibility và settings.
- Các sửa chữa chỉ giải quyết tên kiểu/member và assembly reference; không thay công thức,
  định mức, mapping, identity hoặc hành vi nghiệp vụ.

## Kiểm thử Excel bổ sung 2026-10-06

- V2-801: preview không ghi workbook, cancel, pin đúng checksum, backup trước ghi,
  rollback khi cố tình gây lỗi, save/close/reopen giữ package: PASS.
- Migration nhận 26 công tác của mẫu HoaLuNamDinh; sheet Gia DT TC không có CodeName.
  Đã sửa resolver không được coi hai CodeName rỗng là cùng worksheet.
- Fixture riêng: ba công tác, hai biến thể forest-1/forest-2 tạo hai block DG Cạn;
  ba link dự toán, khối lượng `=1+1` giữ nguyên, tổng NC Excel 20.900.000: PASS.
- ScanReadOnly không đổi serialized V2 state: PASS.
- V2-901: PDF một/hai sheet, rename theo CodeName, chặn `=1/0`, bảo toàn sheet count
  và active sheet nguồn: PASS. PDF hai sheet có ba trang; đã render kiểm tra bảng THKP.
- 81 test Core PASS, build Release x64 không warning/error.
- Chạy lại sau sửa metadata ẩn: script exit 0, Print Area trống/cột kỹ thuật ẩn
  bị chặn và lỗi xuất giữ nguyên hash PDF cũ. Artifact cuối:
  `tmp/V2-runtime-f836e8dd8fa346a996938ef888c88a11/`.
- Excel thật: Ribbon MyTools, task pane Tổng quan, Gói pháp lý, Báo cáo mở được;
  task pane không khóa worksheet. Sửa màu chữ kế thừa từ theme tối của Office.
- Xử lý xung đột subscription bản cài 1.0.4 / dev 1.0.0; gỡ đúng subscription cũ,
  cài dev manifest và bật lại riêng ExcelAddIn1 trong Disabled Items. Không thay security.
  Mở lại Excel đã tự xuất hiện MyTools; dev manifest hiện dùng `|vstolocal`.

Lệnh runtime (Windows PowerShell STA, phiên Excel test riêng):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File scripts/test-estimate-v2.ps1
```

File gốc `Y:\Mau DuToan\Du toan RPBM HoaLuNamDinh_Ver1.xlsx` không bị sửa;
test làm trên bản sao `tmp/V2-runtime-<guid>/`. Script kiểm tra SHA256 nguồn.
`HoaLuNamDinh-V2-user-test.xlsx` là bản chuyển đổi trước fixture, để người dùng thử UI.

### Giới hạn của mẫu chuẩn

Kiểm tra XML file gốc trước đối chiếu: 285 cached error cells và 91 công thức chứa
`#REF!`. Đây là lỗi có sẵn ở đầu vào, không được quy thành kết quả V2 đạt/không đạt.
UI validation trên bản migration phát hiện 144 lỗi công thức và chặn xuất; không bypass.
Con số 33.107.729.000 trong PDF là tổng có sẵn được giữ khi xuất, không phải bằng chứng
engine V2 tái tạo toàn bộ dự toán đúng bằng tổng đó. Chưa tự suy đoán sửa tham chiếu mất.

## Phạm vi chưa kiểm chứng đầy đủ

Build thành công và test Core PASS không phải bằng chứng Excel/VSTO end-to-end PASS.
Chưa được coi là full runtime PASS:

- toàn bộ thao tác gắn định mức/giá/settings qua UI, save/close/reopen mọi thiết lập;
- đối chiếu giao diện với ảnh gốc tại DPI 100/125/150%;
- đối chiếu toàn bộ chuỗi giá -> đơn giá -> dự toán -> THKP trên mẫu không lỗi;
- DG Nước/Biển, hai đối tượng lương, sort/copy/delete WorkItem và Print Preview UI.

Đường dẫn `/mnt/data/chuan_UI` chưa truy cập được trên host Windows này.
Repo chỉ chứa ảnh mẫu sheet trong `docs/du-toan-v2/images/`, chưa chứa bộ ảnh UI 01–08.
Không dùng các ảnh mẫu sheet thay cho ảnh UI chuẩn. Theo HANDOVER, ảnh chuẩn 09–10
cũng chưa có tại thời điểm bàn giao; các màn hình này cần bám UI-CONTRACT hoặc mockup
cùng ngôn ngữ giao diện đã chốt.

V2-801/V2-901 đã triển khai. Các milestone khác đã có kiểm thử mục tiêu nêu trên nhưng
không tự chuyển toàn bộ roadmap thành full runtime PASS. Tiếp theo là regression còn lại.
