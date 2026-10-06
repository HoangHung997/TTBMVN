# Ke hoach tiep quan va chuyen huong module Du toan

## 1. Muc tieu

Khong xoa module cu va lam lai tu dau. Tiep quan theo huong:

1. Bao toan nen tang da test.
2. Tach luong V2 khoi luong legacy.
3. Prototype tren workbook thuc te moi nhat.
4. Chuyen tung writer sang formula/link.
5. Chi loai legacy khi V2 dat regression va nguoi dung chap nhan.

## 2. Thanh phan nen giu

| Thanh phan hien tai | Huong V2 |
|---|---|
| ExcelAddIn1.Core | Giu, tiep tuc dat business model/validation o day |
| RegulationPackage 2021/2025 | Giu |
| Source locator/checksum/version | Giu |
| Package store + offline update | Giu |
| Effective-date/transition resolver | Giu |
| Worksheet role CustomProperties | Giu, nhung role V2 co the map on-demand |
| ProjectProfile | Giu phan pin package/schema; giam metadata khong can thiet cho luong V2 |
| ExcelWriteContext + batch write | Giu |
| Runtime diagnostics/support package | Giu |
| Validation engine | Giu va mo rong cho formula/link |
| MachineRate/Norm catalogs | Giu du lieu, dieu chinh cach render ra workbook |

## 3. Thanh phan can refactor

### Estimate Workspace

Hien tai module da co workspace/modeless va luong du toan kha day du, nhung V2 phai dua "dong cong tac trong Excel" thanh trung tam.

Workspace moi nen tap trung vao:

- danh sach cong tac da quet;
- trang thai da/chua gan dinh muc;
- tra cuu dinh muc;
- batch bind;
- dong bo bang output;
- validation.

### Worksheet role gate

Hien tai dac ta yeu cau 7 role bat buoc truoc khi vao luong.

V2 de xuat:

- `ResourcePrices`, `UnitRateLand`, `UnitRateWater`, `EstimateAppendix`, `CostSummary` chi can khi thuc su sinh/doi chieu.
- `NormLookupView`, `CostRuleView` khong bat buoc phai co sheet in; co the la UI cua add-in.
- Mapping theo role van ben vung khi doi ten sheet.

### PriceProfile

PriceProfile co the tiep tuc luu metadata/version/audit, nhung gia dang ap dung cho du toan phai co vung nhin thay/kiem tra duoc trong workbook va cac ket qua phai link toi do.

### UnitRate/Estimate writers

Writer cu neu ghi snapshot so can duoc chuyen sang FormulaBuilder.

## 4. Thanh phan khong nen mang sang nhu logic chinh

- Hard-code ten sheet.
- Hard-code vi tri dong/cot cua tung mau.
- Hard-code he so trong code khi he so co the bieu dien bang tham so/can cu.
- Tu dong suy luan so khu vuc va bat wizard du an theo khu vuc.
- Tu dong gan dinh muc thay nguoi lap.
- Render ket qua tinh san bang C# thanh so chet.

## 5. Lo trinh de xuat

### V2-001 - Baseline huong moi

- Chon `Du toan RPBM HoaLuNamDinh_Ver1.xlsx` lam workbook tham chieu chinh.
- Lap inventory 5 sheet in.
- Danh dau input/formula/manual region.
- Ghi checksum va tao working copy rieng.

### V2-101 - WorkItem + Norm Binding

- Model `EstimateWorkItem`.
- Dang ky Table/vung cong tac.
- WorkItemId ben vung.
- Form tra cuu + gan/doi/bo dinh muc.
- Batch bind nhieu dong.
- Save/reopen + insert/delete/sort test.

### V2-201 - Resource aggregation + VL-NC-M

- Gom resource unique tu dinh muc da gan.
- Sinh/cap nhat danh sach VL/NC/M.
- Sinh cong thuc gia NC va gia ca may.
- Khong xoa/sua nham cac gia nguoi dung da nhap.

### V2-301 - DG Can / DG Nuoc / DG Bien

- Sinh block theo NormCode unique.
- Don gia link VL-NC-M.
- Thanh tien = hao phi * don gia.
- Rebuild/incremental update an toan.
- Bao toan format in.

### V2-401 - Gia DT TC + THKP-TC

- Link tung cong tac toi block don gia.
- Thanh tien theo khoi luong.
- THKP link tu Gia DT TC.
- Cac ty le/quy tac link toi vung tham so/can cu.

### V2-501 - Validation

- Cong tac chua gan dinh muc.
- Dinh muc/package khong ton tai.
- Thieu gia VL/NC/M.
- Cong thuc bi ghi de.
- `#REF!`, `#VALUE!`, `#N/A`.
- Block thua/khong con duoc dung.
- Package workbook khac package dang duoc app de xuat.

### V2-601 - Tuong thich va migration

- Thu 1/2/3 khu vuc.
- Thu VT/DN.
- Thu workbook doi ten sheet.
- Mo lai tren may khong co add-in: cong thuc van doc/in duoc.
- Neu can, viet converter cho workbook legacy sau khi V2 on dinh.

## 6. Gate nghiem thu V2 co ban

Một luong V2 chi duoc coi la dat khi:

- Nguoi dung lap xong khoi luong va gan dinh muc ma khong can sua code/template.
- Moi don gia duoc sinh mot lan theo ma unique.
- Tat ca ket qua tinh trong 5 sheet chinh la cong thuc/link.
- Thay gia dau/NC/VL lam Excel tu cap nhat den THKP.
- 1, 2 va 3 khu vuc cho cung ket qua theo cac dinh muc da gan.
- Insert/delete/sort dong khong mat mapping.
- Save/reopen khong mat metadata.
- Khong co `#REF!` do add-in tao.
- Workbook van xem, audit va in duoc neu add-in khong duoc cai.
