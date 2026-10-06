# Cai dat va go cai TTBMVN Excel Tools

## Pham vi ban 1.0.1

Ban pilot da nghiem thu tren Windows 11, Microsoft 365 Excel x64 va .NET Framework 4.8.1.
Office 2016/2019/2021, Windows 10 va Excel x86 chua co may that de chay, nen khong duoc
tuyen bo ho tro trong dot pilot nay.

## Truoc khi cai

1. Giai nen toan bo file ZIP vao mot thu muc co dinh.
2. Mo `release-info.json`, doi chieu version, channel va publisher thumbprint.
3. Chay `Verify-TTBMVNRelease.ps1` va chi tiep tuc khi ket qua la `PASS`.
4. Dong tat ca cua so Excel.

Co the kiem tra dieu kien cai dat ma khong thay doi may, ke ca khi Excel dang mo:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-TTBMVN.ps1 -VerifyOnly
```

Ban `PILOT` dung self-signed certificate. Certificate nay chi nen duoc tin cay tren may pilot
sau khi doi chieu thumbprint qua kenh ban giao doc lap. Day khong phai chu ky CA thuong mai.

## Cai dat

Cach truc quan: dong Excel, sau do nhap dup `Cai-dat-TTBMVN.cmd` trong thu muc da giai nen.
Khong mo truc tiep `ExcelAddIn1.vsto` khi may dang co ban cu duoc cai tu thu muc khac.

Hoac mo PowerShell tai thu muc giai nen:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-TTBMVN.ps1 `
  -TrustPilotCertificate
```

Script se import public pilot certificate vao trust store cua tai khoan, go ban `ExcelAddIn1`
cu neu no duoc cai tu vi tri khac, sau do cai manifest moi. Setting va package trong
`%LOCALAPPDATA%\TTBMVNExcelTools` khong bi xoa khi nang cap.

Sau khi cai, mo lai Excel va kiem tra ribbon TTBMVN. Neu add-in khong nap, vao Excel Options,
Add-ins, COM Add-ins va xem `TTBMVN Excel Tools` co duoc bat hay khong.

## Kich hoat

Mo `Ho tro` trong form Ho dao hoac module Du toan, copy Machine ID va nhap offline key duoc
cap rieng cho may. Khong gui license key trong goi chan doan.

## Cap nhat

- Binary add-in: chay bo cai version moi; script tu go ban cu truoc khi cai.
- Du lieu phap ly: mo `Du toan > Cap nhat`, chon file `.ttbupdate`, xem diff va cai.
- Cai package moi khong tu dong doi package da pin trong workbook cu. Dung `Chuyen goi` de
  preview va tao backup truoc khi migrate.

## Go cai

Dong Excel, sau do chay:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Uninstall-TTBMVN.ps1
```

De go them public pilot certificate khoi trust store cua tai khoan:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Uninstall-TTBMVN.ps1 `
  -RemovePilotTrust
```

Go add-in khong tu dong xoa setting, log, license va package tai
`%LOCALAPPDATA%\TTBMVNExcelTools`. Sao luu roi xoa thu cong neu can xoa du lieu nguoi dung.
