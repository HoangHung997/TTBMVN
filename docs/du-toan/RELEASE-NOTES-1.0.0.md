# TTBMVN Excel Tools 1.0.0 PILOT

Ngay dong goi: 2026-08-03. Nen tang: Excel x64, .NET Framework 4.8.1.

## Chuc nang chinh

- Ho dao 3 m/5 m voi rang buoc H/V, preview, batch write, sheet moi/hien co, link nguoc,
  mapping du lieu phu, format Excel va so worker tu dong.
- Project profile va 7 worksheet role ben vung khi sheet doi ten.
- Goi phap ly BQP 2021 va BQP 2025 co version/checksum/source locator; workbook pin package.
- Luong VL-NC-M, DG Can/DG Nuoc, Gia DT TC va THKP-TC; rule engine, tra cuu va PriceProfile.
- Audit ket qua, validation, migration preview/backup/rollback.
- Update Center offline voi `.ttbupdate` ky RSA-SHA256.
- License offline theo Machine ID va form Ho tro tao goi chan doan da sanitize.

## Nghiem thu

- 68/68 Core test PASS.
- DT-605 release acceptance 14/14 PASS trong 1044,4 giay.
- Migration 2021 -> 2025 qua 6 fault phase va restore checkpoint.
- Workbook chuan co `THKP-TC!E27=1201557000`; 92 formula error legacy da ghi nhan va
  khong tang sau luong app.

## Gioi han pilot

- Chi host Windows 11 + Microsoft 365 Excel x64 da chay live.
- Channel PILOT dung self-signed certificate; can import public certificate sau khi doi chieu
  thumbprint. Chua du dieu kien phat hanh thuong mai cong khai/SmartScreen reputation.
- Update online dang tat; cap nhat phap ly dung file offline.
- Workbook legacy co 4034 defined name hong duoc gom thanh mot warning ke thua.
