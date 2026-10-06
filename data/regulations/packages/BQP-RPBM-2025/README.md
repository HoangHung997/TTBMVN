# BQP-RPBM-2025

Gói snapshot đã resolve, áp dụng từ 28/10/2025. Nguồn gồm TT101/2025/TT-BQP và
94-98/VBHN-BQP; quy tắc chuyển tiếp tại Điều 6 TT101 vẫn giữ package 2021 cho hồ sơ
đã được phê duyệt trước ngày cắt chuyển.

## Cấu trúc

- `source/`: định nghĩa package và 248 record TSV có source locator.
- `bundle/`: sáu module runtime đã niêm phong checksum.
- `audit/change-reasons.tsv`: một lý do và locator cho từng record thay đổi.
- `audit/record-diff.txt`: diff cấp field giữa package 2021 và 2025.
- `audit/DT-302-review.md`: biên bản đối chiếu và kiểm thử.
- `audit/DT-303-machine-rate-review.md`: phương pháp, rounding, hai đối tượng và sai khác bảng in.

## Tái tạo và kiểm tra

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\scripts\test-bqp-2025-package.ps1 -RegenerateSource
```

Builder không ghi đè một bundle cùng ID/version có checksum khác. Bản hiện hành là `2.0.1`;
mọi sửa dữ liệu tiếp theo phải tăng `dataVersion`.

## Checksum

- Package: `E2537C08E7156414585BEFC446E74DD51E79AB8D67A3774B1F5250E7EECCDA90`.
- TechnicalProcess: `C79C190CC5EADD7AF3056800FB9654F61BA5602DF20A09AE10C97A050BA6C444`.
- Norm: `E1B9D047976A147BF527ECC4684A1FD60FFB945D0B09486C118EE76C236FE4ED`.
- CostRule: `9F8472DA0C8F3604E70DC672060DABF5289267238386481ACF68C8ECAD7DA39C`.
- MachineRate: `90118DFE452FC4F6ABAF6D0D0A465791415D81AFC389876DE6A632BE65E2C6A0`.
- Geography: `42C7AF952546BFF7F25147100CF44069A608A8F8F8CC3F08436AB78CF2663C8D`.
- Compliance: `81A5769101B245EE65218CC179AB0CA56840EFC2CECAF90DF164A88C6A245D07`.
- Record: 248, gồm TechnicalProcess 60, Norm 42, CostRule 37, MachineRate 33,
  Geography 69 và Compliance 7.
- Diff 2021 -> 2025: 50 added, 0 removed, 198 changed.
