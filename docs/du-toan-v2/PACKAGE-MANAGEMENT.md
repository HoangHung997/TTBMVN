# V2-802 - Quan ly va bien soan goi phap ly

Ngay: 2026-10-06. Trang thai: DONE implementation; build/Core va UI smoke PASS;
full Excel regression va doi chieu tai chinh van pending.
Commit implementation: `44fb906`.

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

### Cap nhat V2-803 - Bang theo cong tac (2026-10-06)

- Khong con bang phang gom tat ca ban ghi va chuoi data dai trong cac tab module.
  Moi tab chon mot ma cong tac/quy tac, hien tieu de va bang chi tiet cua ma do.
- Dinh muc 2025: hang la hao phi VL/NC/M, cot la tung loai `020.0200.1`, `.2`,
  `.3`, `.4` kem nhan mat do/dat/rung/do sau. Ma variant that giu nguyen,
  tooltip hien khoa; `.1` la thu tu hien thi, khong doi identity binding.
- Thong tin/can cu va dieu kien co tab chi tiet rieng. So hao phi dung dau thap
  phan Excel; file package van luu so theo dinh dang invariant.
- Khu vuc: GeographyZone va ProvisionalEstimateRate (suat theo khu vuc/nguon
  kinh phi), dung chung ban ghi goc, khong sao chep thanh module moi. Dia ban
  giu cac bang mat do, dia hinh, phan vung tinh/thanh.
- Chuyen ma/tab va luu nhap se commit bang dang sua; nhap sai bi chan chuyen.
- Goi 2021 chi co danh muc/so loai, chua co bang rates chi tiet: hien cac truong
  goc, khong tu tao hao phi du doan.
- Dieu chinh/rang buoc phuc tap van dung cu phap ky thuat trong bang thong so;
  chua co editor rieng tung dieu kien hay ten/nhan variant tuy chinh. Thay doi
  so luong variant phai khop cot hao phi; chua co nut them/xoa cot variant.
- Day la bo cuc theo tung cong tac, chua phai ban sao 100% tat ca bang van ban.
  Can doi chieu bang goc cu the de chot editor rieng cho tung loai quy tac.

Build Release x64 sach, **84 test Core PASS**. Test ReadFields/WriteFields cho
tat ca ban ghi 2021/2025; ma tran roundtrip giu hao phi, dieu chinh, rang buoc;
chan variant trung va hao phi am. Excel smoke PASS: mo Dinh muc/Khu vuc/Chi phi/
Gia may/Dia ban/Tuan thu, chuyen tab va Kiem tra goi. Luu nhap ma tran PASS.
Khong cai/xoa goi qua UI. Chua test day du sua cell/doi ma/bo sung variant.
Ban test rieng: `tmp/HoaLuNamDinh-package-table-test.xlsx`.

Sau module: Quy trinh, Dinh muc, Chi phi, Gia may, Dia ban, Tuan thu;
mot tab Van ban nguon. Ten tab duoc hien thi theo UI hien tai.
Bang module gom ma, loai ban ghi, don vi, ten du lieu, du lieu variant/hao phi/
quy tac, ma van ban, trang tu/den, muc can cu va trang thai kiem chung.

Trong ban bang phang ban dau, cot Ten du lieu cho phep sua ten hien thi sang tieng Viet. Cot don vi/variant/
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
