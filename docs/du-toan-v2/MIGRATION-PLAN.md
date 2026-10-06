# Ke hoach tiep quan va chuyen huong module Du toan

## 1. Muc tieu

Khong xoa module cu va lam lai tu dau. Tiep quan theo huong:

1. Bao toan nen tang da test.
2. Tach luong V2 khoi luong legacy.
3. Prototype tren workbook thuc te moi nhat.
4. Chuyen tung writer sang formula/link.
5. Don gian hoa man hinh khoi dong va loai cac gate khong can thiet.
6. Chi loai legacy khi V2 dat regression va nguoi dung chap nhan.

## 2. Cac van de cua luong hien tai phai sua trong V2

Logic tinh toan hien tai **khong sai nhieu**; diem chinh can sua la trai nghiem su dung va cach to chuc workbook.

### 2.1. Qua nhieu tab trung gian/rac

Ban hien tai co xu huong tao/yeu cau nhieu sheet phuc vu lookup, cost rule, workspace va mapping. Ket qua la workbook co nhieu tab khong phai ho so can in.

V2 phai theo nguyen tac:

- Chi cac sheet ho so thuc te moi ton tai tren thanh tab: `THKP-TC`, `Gia DT TC`, `DG Can`, `DG Nuoc`, `VL-NC-M` va `DG Bien` neu co.
- Lookup dinh muc, CostRule, validation, cau hinh va metadata uu tien nam trong UI, database/package hoac sheet VeryHidden.
- Khong tao sheet chi de dap ung mot role ky thuat neu nguoi dung khong can xem/in sheet do.
- Khong tao `DG Nuoc`/`DG Bien` neu du an khong co dinh muc thuoc nhom do.

### 2.2. Khong duoc bat quet/chon THKP truoc khi mo Du toan

Ban hien tai yeu cau validate/mapping role, quet va chon `THKP` truoc khi cho mo form Du toan. Day la UX khong phu hop.

V2 phai cho phep:

1. Bam **Du toan** la mo duoc workspace ngay.
2. Workbook chua co sheet du toan van mo duoc.
3. Chua map `THKP-TC` van tra cuu/gan dinh muc duoc.
4. Chi khi nguoi dung thuc hien thao tac can mot output cu the (vi du Sinh THKP) thi moi yeu cau map/tao sheet do.
5. Mapping sheet theo role la **on-demand**, khong la gate toan module.

Noi cach khac: `THKP-TC` la mot output nghiep vu, khong phai "chiec khoa" de vao module Du toan.

### 2.3. Thieu goi phap ly khong duoc lam module Du toan loi/khong mo

Ban hien tai co the chan mo Du toan neu workbook chua pin/cai package phap ly. V2 phai degrade gracefully.

Hanh vi mong muon:

- Luon mo duoc workspace Du toan.
- Neu chua co package phap ly: hien trang thai **Chua co du lieu dinh muc/phap ly** va nut `Cai/Chon goi du lieu`.
- Van cho phep xem workbook, dang ky bang cong tac, map sheet va cac thao tac khong can NormCatalog.
- Chuc nang `Tra/Gan dinh muc` bi vo hieu hoa co ly do ro rang cho den khi co package.
- Chuc nang sinh don gia/kiem tra phap ly chi chay khi du package can thiet.
- Khong throw loi khoi dong form chi vi package missing/corrupt.
- Neu package workbook pin bi thieu, phai thong bao ro identity dang thieu va cho phep nguoi dung cai lai/chon huong xu ly; khong tu dong chuyen sang package "latest".

Muc tieu la **missing dependency chi chan dung hanh dong phu thuoc vao dependency do**, khong chan ca module.

## 3. Thanh phan nen giu

| Thanh phan hien tai | Huong V2 |
|---|---|
| ExcelAddIn1.Core | Giu, tiep tuc dat business model/validation o day |
| RegulationPackage 2021/2025 | Giu |
| Source locator/checksum/version | Giu |
| Package store + offline update | Giu |
| Effective-date/transition resolver | Giu |
| Worksheet role CustomProperties | Giu, nhung role V2 map on-demand |
| ProjectProfile | Giu phan pin package/schema; giam metadata khong can thiet cho luong V2 |
| ExcelWriteContext + batch write | Giu |
| Runtime diagnostics/support package | Giu |
| Validation engine | Giu va mo rong cho formula/link |
| MachineRate/Norm catalogs | Giu du lieu, dieu chinh cach render ra workbook |

## 4. Thanh phan can refactor

### Estimate Workspace

Hien tai module da co workspace/modeless va luong du toan kha day du, nhung V2 phai dua "dong cong tac trong Excel" thanh trung tam.

Workspace moi nen tap trung vao:

- danh sach cong tac da quet;
- trang thai da/chua gan dinh muc;
- tra cuu dinh muc;
- batch bind;
- dong bo bang output;
- validation;
- trang thai dependency (package, role sheet) theo kieu canh bao/on-demand thay vi gate khoi dong.

### Worksheet role gate

Hien tai dac ta yeu cau 7 role bat buoc truoc khi vao luong. V2 **bo gate bat buoc nay**.

V2 de xuat:

- Form Du toan mo khong phu thuoc da map du role hay chua.
- `ResourcePrices`, `UnitRateLand`, `UnitRateWater`, `EstimateAppendix`, `CostSummary` chi map/tao khi chuc nang tuong ung can dung.
- `NormLookupView`, `CostRuleView` khong bat buoc co sheet; mac dinh la UI cua add-in.
- Mapping theo role van ben vung khi doi ten sheet.
- Neu sheet da ton tai, cho nguoi dung chon/map; neu chua co, add-in co the tao theo mau khi nguoi dung yeu cau sinh output.
- Khong hard-code ten tab de nhan dien role.

### Package gate

Resolver/package store van giu de dam bao dung phap ly, nhung phai tach:

- **Module availability**: luon mo duoc Du toan.
- **Feature availability**: tung chuc nang tu kiem tra dependency cua no.
- Missing package la state co the xu ly, khong la exception chan form.

### PriceProfile

PriceProfile co the tiep tuc luu metadata/version/audit, nhung gia dang ap dung cho du toan phai co vung nhin thay/kiem tra duoc trong workbook va cac ket qua phai link toi do.

### UnitRate/Estimate writers

Writer cu neu ghi snapshot so can duoc chuyen sang FormulaBuilder.

## 5. Thanh phan khong nen mang sang nhu logic chinh

- Hard-code ten sheet.
- Hard-code vi tri dong/cot cua tung mau.
- Hard-code he so trong code khi he so co the bieu dien bang tham so/can cu.
- Bat buoc map du 7 role moi cho mo module.
- Bat buoc chon/quyet dinh `THKP-TC` truoc khi nguoi dung co nhu cau sinh THKP.
- Throw/chan form khi chua co package phap ly.
- Tao cac sheet lookup/rule trung gian thanh tab visible neu khong phai ho so in.
- Tu dong suy luan so khu vuc va bat wizard du an theo khu vuc.
- Tu dong gan dinh muc thay nguoi lap.
- Render ket qua tinh san bang C# thanh so chet.

## 6. Lo trinh de xuat

### V2-001 - Baseline huong moi + go bo gate khoi dong

- Chon `Du toan RPBM HoaLuNamDinh_Ver1.xlsx` lam workbook tham chieu chinh.
- Lap inventory 5 sheet in.
- Danh dau input/formula/manual region.
- Ghi checksum va tao working copy rieng.
- Lap inventory cac tab "rac"/technical cua luong hien tai va quyet dinh: bo, VeryHidden hay chuyen sang UI.
- Sua flow khoi dong de bam `Du toan` mo form ngay, khong bat map THKP/7 role.
- Missing package phai hien status trong form, khong lam crash/abort form.

### V2-101 - WorkItem + Norm Binding

- Model `EstimateWorkItem`.
- Dang ky Table/vung cong tac.
- WorkItemId ben vung.
- Form tra cuu + gan/doi/bo dinh muc.
- Batch bind nhieu dong.
- Khi package chua co: form van mo, nhung nut tra/gan dinh muc thong bao dependency can cai.
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
- Chi tao sheet theo nhom dinh muc thuc te dang dung.

### V2-401 - Gia DT TC + THKP-TC

- Link tung cong tac toi block don gia.
- Thanh tien theo khoi luong.
- THKP link tu Gia DT TC.
- Cac ty le/quy tac link toi vung tham so/can cu.
- Map/tao `THKP-TC` o day neu can, khong yeu cau tu luc mo module.

### V2-501 - Validation

- Cong tac chua gan dinh muc.
- Dinh muc/package khong ton tai.
- Thieu gia VL/NC/M.
- Cong thuc bi ghi de.
- `#REF!`, `#VALUE!`, `#N/A`.
- Block thua/khong con duoc dung.
- Package workbook khac package dang duoc app de xuat.
- Validation phan loai loi chan thao tac va canh bao; khong bien moi canh bao thanh gate khoi dong.

### V2-601 - Tuong thich va migration

- Thu workbook chua co bat ky sheet du toan nao.
- Thu workbook co 1/2/3 khu vuc.
- Thu VT/DN.
- Thu workbook doi ten sheet.
- Thu may chua cai package phap ly: workspace van mo duoc.
- Thu package pin bi missing/corrupt: hien thong bao xu ly, khong crash form.
- Mo lai tren may khong co add-in: cong thuc van doc/in duoc.
- Neu can, viet converter cho workbook legacy sau khi V2 on dinh.

### V2-701 - Thiet lap chung

- Cau hinh theo workbook, khong tao sheet settings visible.
- Map 6 output theo persistent identity.
- Auto restore norm display, validate-on-open nhe, auto-sync row, mapping-loss warning.
- Formula/Custom XML/hidden technical columns la invariant khong duoc tat.
- Auto-save co interval, khong Save As workbook moi.
- Mapping save co rollback.
- UI mo trong cung CustomTaskPane qua Ribbon, khong modal.

### V2-801 - Goi phap ly & Du lieu

- Hien package workbook dang pin: PackageId / DataVersion / checksum.
- Phan biet ready / missing / corrupt / chua cau hinh.
- Cai/chon package on-demand.
- Khong auto latest.
- Neu doi package phai preview impact + confirm + rollback.
- Tai su dung package store/offline update/effective-date/transition/migration engine hien co.
- Khong dua package scan thanh startup gate.

### V2-901 - Bao cao & Xuat in

- Chon cac sheet ho so can in/xuat.
- Kiem tra PrintArea/page setup.
- Loai technical columns ra khoi output.
- Ho tro Can/Nuoc/Bien theo sheet thuc te.
- Khong tao tab bao cao rac chi de export.

## 7. Gate nghiem thu V2 co ban

Mot luong V2 chi duoc coi la dat khi:

- Bam nut `Du toan` mo duoc workspace ngay ca khi workbook chua map THKP/role.
- Thieu package phap ly khong lam form loi/khong mo; chi cac chuc nang phu thuoc bi khoa kem huong xu ly.
- Workbook khong bi day them tab ky thuat khong can in.
- Nguoi dung lap xong khoi luong va gan dinh muc ma khong can sua code/template.
- Moi don gia duoc sinh mot lan theo ma unique.
- Tat ca ket qua tinh trong 5 sheet chinh la cong thuc/link.
- Thay gia dau/NC/VL lam Excel tu cap nhat den THKP.
- 1, 2 va 3 khu vuc cho cung ket qua theo cac dinh muc da gan.
- Insert/delete/sort dong khong mat mapping.
- Save/reopen khong mat metadata.
- Khong co `#REF!` do add-in tao.
- Workbook van xem, audit va in duoc neu add-in khong duoc cai.
