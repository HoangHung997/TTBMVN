# V2-802 - Quan ly va bien soan goi phap ly

Ngay: 2026-10-06. Trang thai: DONE implementation; build/Core va UI smoke PASS;
full Excel regression va doi chieu tai chinh van pending.

## Diem vao va luong su dung

Trong tong quan Tro ly Du toan, muc **7. Goi phap ly** mo bang goi da cai.
Bang hien Ma goi, Phien ban, Hieu luc tu va so dinh muc. Chon dong khong tu
doi goi cua workbook.

1. **Tao goi moi tu mau da chon**: sao chep du lieu goi mau, dat ma moi va
   phien ban 1.0.0. Khong tao goi rong thieu cac module bat buoc.
2. **Sua goi / dinh muc**: mo ban sao voi cung ma va phien ban tang patch.
   Cho phep sua metadata, them/sua/xoa dong trong cac module va van ban nguon.
3. **Luu ban nhap / Mo ban nhap**: luu/nap thu muc nguon, chua cai vao kho.
4. **Kiem tra goi**: tao bundle, kiem tra checksum, can cu va cau truc du lieu.
5. **Cai goi**: kiem tra lai du lieu hien tai va yeu cau xac nhan. Goi workbook
   dang pin khong thay doi khi cai.
6. **Xem tac dong doi goi / Xac nhan ap dung goi**: luong chon goi cho workbook;
   co preview, kiem tra binding va ban du phong truoc khi ap dung.
7. **Cai bundle tu thu muc / Cai goi cap nhat co chu ky**: them goi co san.
8. **Xoa phien ban goi da chon**: chuyen vao kho du phong, khong xoa vinh vien.
9. **Dat phien ban uu tien trong kho**: khong doi pin cua workbook hien tai.

## Noi dung bang sua

Sau module: Quy trinh, Dinh muc, Chi phi, Gia may, Dia ban, Tuan thu;
mot tab Van ban nguon. Ten tab duoc hien thi theo UI hien tai.
Bang module gom ma, loai ban ghi, don vi, ten du lieu, du lieu variant/hao phi/
quy tac, ma van ban, trang tu/den, muc can cu va trang thai kiem chung.

Cot Ten du lieu cho phep sua ten hien thi sang tieng Viet. Cot don vi/variant/
hao phi van la cac khoa va cu phap ky thuat cua package, khong phai trinh soan
thao dinh muc dang cay. Khong doi tuy tien khoa duoc tham chieu boi module khac.
Kiem tra cau truc khong thay the kiem chung noi dung thong tu/hao phi.

Sua ban ghi se dat lai trang thai Chua kiem chung. Sua gia tri van ban nguon
se dat lai kiem chung cua cac ban ghi. Du lieu chua kiem chung co the luu nhap
nhung khong duoc dong goi/cai. Nguoi bien soan phai doi chieu can cu truoc khi
chon trang thai kiem chung; app khong tu xac nhan tinh dung phap ly.
Ban nhap hien can cac truong so/ngay hop le, chua ho tro luu dong dang nhap do
voi truong bat buoc rong hoac sai dinh dang.

## An toan va luu tru

- Goi da cai la immutable theo ma/phien ban/checksum; sua tao phien ban moi.
- Chan xoa goi dang pin trong cac workbook mo cung phien Excel va goi uu tien.
- Khong the biet het workbook dang dong hoac mo o phien Excel khac: UI canh bao.
- Goi xoa duoc giu trong `RegulationPackages/.removed/...`; bootstrap khong
  tu cai lai goi da xoa. Co the nhap lai bundle tu thu muc du phong.
- Thu muc goc: `%LOCALAPPDATA%/TTBMVNExcelTools`.
- Ban nhap: `PackageDrafts/<guid>`; bundle kiem tra: `PackageBuilds/<guid>`.
- Khi kho khong con goi mau, nhap mot bundle/mau truoc khi tao goi moi.
- Khong sua du lieu ho so Excel thanh so chet; sau doi goi can cap nhat cac bang.

## File va kiem thu

- `ExcelAddIn1.Core/RegulationPackageAuthoring.cs`: xuat nguon va build co kiem tra.
- `ExcelAddIn1.Core/RegulationPackageStore.cs`: chan xoa, luu du phong/khoi phuc.
- `ExcelAddIn1/Winform/RegulationPackageEditorForm.cs`: bang bien soan.
- `ExcelAddIn1/Winform/EstimatePackagesPaneView.cs`: quan ly/cai/chon goi.
- `ExcelAddIn1/Winform/EstimateTaskPaneControl.cs`: muc 7.
- `ExcelAddIn1/Funtion/RegulationPackageBootstrapService.cs`: bo qua goi da xoa.
- `ExcelAddIn1.Tests/Program.cs`: authoring roundtrip, Unverified bi chan,
  ky tu tab sai bi chan, pin/uu tien chan xoa, du phong va nhap lai.

`scripts/build-release.ps1 -Platform x64 -Configuration Release`: build sach,
**83 test Core PASS**.
UI smoke trong Excel/VSTO that tren ban sao HoaLuNamDinh: muc 7 mo bang goi,
tao tu mau, mo tab Dinh muc, Kiem tra goi thanh cong, Luu ban nhap thanh cong.
Khong cai/xoa goi nguoi dung qua UI de thu. Xoa/khoi phuc duoc test bang kho tam.
File mau goc tren Y: khong bi thay doi. Full UI CRUD/pin regression van pending.
