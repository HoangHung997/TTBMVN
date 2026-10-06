# TTBMVN Excel Tools 1.0.4 PILOT x64

## Sua bo cai

- Nhan dien dang ky Visual Studio `bin/.../ExcelAddIn1.vsto|vstolocal` khong thuoc VSTO cache.
- Tim va go ClickOnce subscription cu trong Apps & features theo dung product/publisher,
  ke ca khi khoa Excel da bi Visual Studio ghi de sang manifest khac.
- Khong goi `/Uninstall` cho dang ky debug, tranh VSTOInstaller exit `-401`.
- Sao luu thong tin dang ky stale vao LocalAppData truoc khi xoa rieng khoa `ExcelAddIn1`.
- Neu cai ban moi that bai, tu dong khoi phuc dang ky cu.
- Manifest cu bi mat hoac VSTO bao not-installed cung duoc don an toan.
- Uninstaller xu ly cung cac trang thai debug/stale.

## Chuc nang kem theo

- Bao gom hotfix quet Phu luc DT cua ban 1.0.2.
- Bao gom form Du toan modeless cua ban 1.0.3.

## Cai dat

- Dong tat ca cua so Excel.
- Giai nen toan bo ZIP, sau do nhap dup `Cai-dat-TTBMVN.cmd`.
