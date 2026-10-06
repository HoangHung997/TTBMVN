# TIẾN ĐỘ TRIỂN KHAI DỰ TOÁN RPBM V2

> Tài liệu này là nguồn trạng thái chính cho quá trình tiếp quản và phát triển module Dự toán V2.
> Mỗi task phải ghi rõ: mục tiêu, thay đổi đã làm, file liên quan, trạng thái kiểm tra, việc tiếp theo.
> Người tiếp quản sau chỉ cần đọc file này cùng README, ARCHITECTURE và MIGRATION-PLAN là biết dự án đang ở đâu.

## Trạng thái tổng quan

| Mã | Task | Trạng thái |
|---|---|---|
| V2-001 | Bỏ gate khởi động Dự toán | DONE - implementation |
| V2-101 | WorkItemId + lưu binding định mức bền vững | NEXT |
| V2-201 | Tổng hợp và sinh VL-NC-M | TODO |
| V2-301 | Sinh DG Cạn / DG Nước / DG Biển | TODO |
| V2-401 | Link Gia DT TC + THKP-TC | TODO |
| V2-501 | Validation / phục hồi / phát hiện lỗi | TODO |
| V2-601 | Tương thích file cũ và migration | TODO |

---

## V2-001 — Bỏ gate khởi động Dự toán

**Trạng thái:** DONE - implementation  
**Commit code:** `63c54d886e2d3b463044a98a6328d4c04c8fe0c9`

### Mục tiêu

Sửa hành vi cũ: bấm **Dự toán** nhưng phải mở `FrmProjectSetup`, ánh xạ đủ role/sheet và resolve package pháp lý trước thì mới vào được form Dự toán.

Hành vi V2 cần đạt:

- Bấm **Dự toán** là mở workspace ngay.
- Không bắt chọn / map `THKP-TC` trước.
- Không bắt đủ 7 worksheet role trước khi vào module.
- Thiếu package pháp lý không được chặn toàn bộ module.
- Mapping sheet và package phải được kiểm tra **theo từng chức năng cần dùng**.

### Phân tích luồng cũ

Điểm gate nằm trong:

`ExcelAddIn1/Ribbon1.cs -> btnDutoan_Click`

Luồng cũ:

```text
Bấm Dự toán
  -> kiểm tra workbook
  -> mở FrmProjectSetup (modal)
  -> ProjectSetupPlanner.CreatePlan
  -> yêu cầu package + worksheet role hợp lệ
  -> Commit setup
  -> mới new Dutoan(...)
```

Do `FrmProjectSetup` đứng trước constructor `Dutoan`, mọi lỗi hoặc missing dependency ở setup đều làm module không mở.

### Thay đổi đã thực hiện

Đã bỏ `FrmProjectSetup` khỏi **startup gate** của `btnDutoan_Click`.

Luồng mới:

```text
Bấm Dự toán
  -> kiểm tra workbook
  -> nếu form của workbook đã mở: activate
  -> new Dutoan(...)
  -> Show modeless
```

Không thay đổi logic package/role hiện có bên trong các chức năng cũ. Các control phụ thuộc package vẫn có thể báo lỗi khi người dùng thực sự mở chức năng tương ứng; phần này sẽ được refactor dần ở các task sau.

### File đã thay đổi

- `ExcelAddIn1/Ribbon1.cs`

### Thành phần cũ vẫn được giữ lại

- `FrmProjectSetup` chưa xóa.
- `ProjectSetupPlanner` chưa xóa.
- Worksheet role service chưa xóa.
- Package resolver/store chưa xóa.

Lý do: các thành phần này còn giá trị cho migration/settings/compatibility và sẽ được chuyển thành **on-demand settings**, không còn là cửa chặn module.

### Inventory role hiện tại

| Role | Hướng V2 |
|---|---|
| ResourcePrices | output/user sheet: VL-NC-M |
| UnitRateLand | output/user sheet: DG Can |
| UnitRateWater | output/user sheet: DG Nuoc |
| EstimateAppendix | output/user sheet: Gia DT TC |
| CostSummary | output/user sheet: THKP-TC |
| NormLookupView | bỏ yêu cầu sheet visible; chuyển sang UI/task pane |
| CostRuleView | bỏ yêu cầu sheet visible; chuyển sang UI/task pane |

V2 không yêu cầu tất cả role tồn tại khi mở module.

### Kiểm tra đã thực hiện

- Static review xác nhận `btnDutoan_Click` không còn gọi `FrmProjectSetup.ShowDialog(...)`.
- Constructor `Dutoan(workbook, coordinator)` không tự resolve package trong lúc khởi tạo.
- View mặc định `ShowGeneralInformation()` hiện không phụ thuộc RegulationPackage resolver.

### Kiểm tra runtime còn phải chạy trên máy Excel dev

Không tuyên bố runtime PASS trong tài liệu này cho đến khi chạy build/live test.

Checklist khi chạy local:

- Build x64 Release.
- Mở workbook không có role mapping -> bấm Dự toán -> form phải mở.
- Mở workbook không pin package -> bấm Dự toán -> form phải mở.
- Workbook có package cũ/missing -> form vẫn mở.
- Workbook đóng -> form modeless phải tự đóng như test hiện có.
- Không phát sinh setup modal trước form Dự toán.

Nếu runtime phát hiện regression, mở lại V2-001 với trạng thái FIXING.

---

## V2-101 — WorkItemId + binding định mức bền vững

**Trạng thái:** NEXT

### Mục tiêu

Đưa “dòng công tác trong Excel” thành đối tượng trung tâm, không nhận diện theo RowIndex.

Mỗi công tác có:

- `WorkItemId` GUID bền vững.
- `NormCode` / package/version binding.
- Metadata nằm trong Custom XML Part.
- Cột kỹ thuật ẩn làm anchor trên sheet.
- Xóa ô định mức hiển thị không làm mất binding.
- Chèn/xóa/sort dòng không làm lệch mapping.
- Copy/paste dòng phải phát hiện duplicate WorkItemId.

### Thiết kế đã chốt

Nguồn sự thật:

```text
Custom XML Part
    WorkItemId -> NormBinding
```

Anchor trên sheet:

```text
Gia DT TC:
M  __TTB_ID
N  __TTB_NORM
O  __TTB_KIND
P  __TTB_HASH
```

Tên/cột cụ thể sẽ được xác nhận khi coding để không phá vùng in mẫu Hoa Lư - Nam Định.

### Các bước dự kiến

1. Tạo model Core cho WorkItem / NormBinding.
2. Tạo serializer + workbook Custom XML store.
3. Tạo hidden-column anchor service.
4. Tạo reconcile service khi mở workbook.
5. Xử lý insert/delete/sort/copy.
6. Viết test Core cho serialize/reconcile.
7. Kết nối vào UI Gắn định mức.

---

## Quy tắc cập nhật file tiến độ

Khi hoàn thành một task:

1. Đổi trạng thái task thành `DONE`.
2. Ghi commit SHA.
3. Ghi file đã thay đổi.
4. Ghi acceptance criteria đã đạt.
5. Ghi test đã chạy và kết quả thật.
6. Chuyển task kế tiếp thành `NEXT`.
7. Không ghi “PASS” nếu chưa thực sự chạy test tương ứng.
