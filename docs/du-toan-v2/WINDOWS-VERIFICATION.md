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

## Phạm vi chưa kiểm chứng

Build thành công và test Core PASS không phải bằng chứng Excel/VSTO end-to-end PASS.
Chưa thực hiện trong lượt tiếp quản này:

- cài bản build và chạy task pane trong Excel thật;
- kiểm thử writer/Custom XML/migration/settings trên workbook rồi save/close/reopen;
- đối chiếu giao diện với ảnh gốc tại DPI 100/125/150%;
- kiểm thử chuỗi giá -> đơn giá -> dự toán -> THKP trong Excel.

Đường dẫn `/mnt/data/chuan_UI` chưa truy cập được trên host Windows này.
Repo chỉ chứa ảnh mẫu sheet trong `docs/du-toan-v2/images/`, chưa chứa bộ ảnh UI 01–08.
Không dùng các ảnh mẫu sheet thay cho ảnh UI chuẩn. Theo HANDOVER, ảnh chuẩn 09–10
cũng chưa có tại thời điểm bàn giao; các màn hình này cần bám UI-CONTRACT hoặc mockup
cùng ngôn ngữ giao diện đã chốt.

Các milestone V2-001 đến V2-701 giữ trạng thái implementation / runtime pending.
Task phát triển kế tiếp: V2-801, sau đó V2-901; checklist runtime tiếp tục theo HANDOVER.
