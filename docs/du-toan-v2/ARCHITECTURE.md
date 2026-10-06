# Kien truc dich - Du toan RPBM V2

## 1. So do tong the

```text
                    DATABASE / PACKAGE
          Dinh muc + hao phi + VL/NC/M + can cu
                         |
                         | tra cuu
                         v
+--------------------------------------------------+
|                 EXCEL ADD-IN V2                  |
|                                                  |
|  Quet bang cong tac                              |
|        |                                         |
|  Gan dinh muc thu cong                          |
|        |                                         |
|  Gom tai nguyen dang dung                       |
|        |                                         |
|  Sinh/cap nhat VL-NC-M                          |
|        |                                         |
|  Sinh/cap nhat DG Can / DG Nuoc / DG Bien       |
|        |                                         |
|  Link Gia DT TC                                  |
|        |                                         |
|  Link THKP-TC                                    |
|        |                                         |
|  Kiem tra cong thuc/du lieu                      |
+--------------------------------------------------+
                         |
                         v
                    EXCEL WORKBOOK
              cong thuc + link + ho so in
```

## 2. Nguyen tac bien gioi

### Database/package chiu trach nhiem

- Ma va ten dinh muc.
- Don vi tinh.
- Nhom cong tac: can/nuoc/bien/chung.
- Thuoc tinh tra cuu: mat do/khu vuc/dieu kien ap dung.
- Chi tiet hao phi VL/NC/M.
- Danh muc vat lieu, nhan cong, may.
- Thong so/quy tac can de tinh gia ca may va nhan cong.
- Source locator, van ban, version, khoang hieu luc, checksum.

### Workbook chiu trach nhiem

- Khoi luong du an.
- Cac bang tinh khoi luong cua nguoi dung.
- Gia thi truong va tham so gia cua du an.
- Cong thuc tinh NC/M/VL.
- Phan tich don gia.
- Gia du toan thi cong.
- Tong hop kinh phi.
- Toan bo sheet de kiem tra/in.

### Add-in chiu trach nhiem

- Nhan dien bang/vung cong tac.
- Quan ly metadata gan dinh muc.
- Tra cuu va gan/doi/bo dinh muc.
- Gom danh sach resource unique.
- Sinh block bang va cong thuc.
- Dong bo khi them/xoa/doi cong tac.
- Validation va bao loi.

## 3. Cong tac Excel la doi tuong trung tam

Moi dong cong tac co mot ID ben vung, khong duoc buoc vao so dong.

De xuat metadata:

```text
WorkItemId
NormCode
NormVersion
NormGroup
LaborMode (neu can)
ManagedBy = TTBMVN.EstimateV2
```

Uu tien luu metadata bang cot an trong Excel Table ket hop mot sheet he thong `xlSheetVeryHidden` de phuc hoi/audit. Khong de metadata ky thuat xuat hien trong ban in.

## 4. Quet va gan dinh muc

Luong V2:

1. Nguoi dung chon mot Table/vung cong tac va dang ky voi add-in.
2. Add-in doc cac dong cong tac, khong tu doan dinh muc.
3. Nguoi dung chon mot/nhieu dong.
4. Mo form tra cuu dinh muc.
5. Xem hao phi neu can.
6. Gan ma dinh muc da chon.
7. Metadata di theo dong khi chen/xoa/sap xep Table.

Ho tro:

- Gan mot dinh muc cho nhieu dong.
- Sao chep/dan dinh muc.
- Doi dinh muc.
- Bo dinh muc.
- Loc theo can/nuoc/bien, tu khoa, ma.
- Hien thi mat do/khu vuc nhu thong tin ho tro, khong tu dong ep chon.

## 5. Sinh VL-NC-M

Add-in lay hop cac tai nguyen cua **cac dinh muc dang duoc su dung**.

Vi du 100 dong cong tac dung 18 dinh muc thi chi sinh nhung VL/NC/M can cho 18 dinh muc do.

Gia thi truong/tham so nam trong workbook. Gia nhan cong va gia ca may phai sinh bang cong thuc theo quy tac hien hanh.

## 6. Sinh don gia

Danh sach dinh muc phai unique.

Mot ma dinh muc duoc su dung 20 lan trong `Gia DT TC` chi co mot block phan tich don gia.

Nhom output:

- `DG Can`: tat ca dinh muc can.
- `DG Nuoc`: tat ca dinh muc nuoc.
- `DG Bien`: chi sinh neu co dinh muc bien.

Khong tao `DG Can KV1`, `DG Can KV2`... Khu vuc da duoc the hien boi variant dinh muc ma nguoi dung gan.

## 7. Link Gia DT TC

Moi dong cong tac lay VL/NC/M don vi tu block don gia tuong ung.

Thanh tien la cong thuc:

```text
VL thanh tien = Khoi luong * VL don vi
NC thanh tien = Khoi luong * NC don vi
M  thanh tien = Khoi luong * M don vi
```

Khong tinh ket qua trong C# roi ghi `Range.Value` cho cac cot ket qua.

## 8. Link THKP-TC

`THKP-TC` lay tong tu `Gia DT TC` va tiep tuc tinh cac khoan theo quy tac/ty le.

Ty le phap ly khong nen nhung truc tiep thanh hang so trong cong thuc neu co the. Uu tien link toi vung tham so co nguon/can cu.

## 9. Cong thuc va dia chi

Uu tien:

- Excel Table.
- Structured Reference.
- Named Range/Name.
- XLOOKUP/INDEX-MATCH phu hop host.

Chi dung dia chi cell tuyet doi khi mau in/merge bat buoc; dia chi phai duoc FormulaBuilder/TemplateEngine sinh dong, khong hard-code rai rac trong code.

## 10. Tai su dung kien truc cu

Nen giu:

- `ExcelAddIn1.Core` va nguyen tac business logic tach Interop.
- RegulationPackage bat bien + version/checksum.
- Effective-date resolver va transition rule.
- Package store/offline update.
- Worksheet role metadata.
- ExcelWriteContext/batch write.
- Validation/diagnostics.
- Co che rollback/checkpoint khi sua workbook lon.

Can don gian hoa/refactor:

- Estimate Workspace hien tai thanh luong quet/gan dinh muc.
- Bat buoc 7 worksheet role truoc khi vao module: V2 nen lazy/on-demand theo sheet thuc te.
- PriceProfile phai bam vao cac o/bang gia trong workbook va sinh formula, khong bien workbook thanh ket qua render tu engine.
- UnitRate/EstimateAppendix writers phai uu tien formula/link thay vi value snapshot.

## 11. Tieu chi bat buoc

- Khong co ket qua so chet.
- Chen/xoa/sort dong cong tac khong lam mat mapping.
- Doi ten sheet khong lam mat role.
- 1/2/3 khu vuc deu dung cung mot kien truc.
- Workbook mo lai van tinh va in duoc.
- Khong con `#REF!` do add-in sinh.
- Ho so cu khong bi tu dong nang package phap ly.
