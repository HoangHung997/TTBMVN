# Gate phat hanh DT-606

## Tieu chi

- Build Release x64 sach va Core test PASS.
- Deployment/application manifest co chu ky hop le va architecture amd64.
- `setup.exe` co Authenticode signer khop publisher trong `release-info.json`.
- Khong co PFX/private key trong release.
- Toan bo file co SHA-256; verifier fail neu thua, thieu hoac sai hash.
- Goi update, public update key va workbook mau duoc luu trong artifact.
- Co tai lieu cai, go, cap nhat, release notes va compatibility.
- DT-605 acceptance va semantic checkpoint cuoi PASS.

## Kenh ky

- `PILOT`: self-signed Code Signing certificate, private key non-exportable trong Windows
  CurrentUser certificate store. Setup co chu ky nhung trust status tren may sach la
  `UnknownError/UntrustedRoot` cho den khi public cert duoc import.
- `EXTERNAL`: dat `TTBMVN_SIGNING_CERT_THUMBPRINT` tro den Code Signing certificate con han
  co private key. Ban thuong mai con can CA cong khai, timestamp RFC 3161 va SmartScreen
  reputation; khong duoc doi nhan PILOT thanh commercial chi bang sua text.

## Lenh dong goi

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\publish-release.ps1 `
  -Version 1.0.0 -Overwrite
```

Artifact chi duoc ban giao khi `Verify-TTBMVNRelease.ps1` va release acceptance deu PASS.

## Ket qua 1.0.0 PILOT

- Build Release x64 va 68/68 Core test: `PASS`.
- Deployment/application manifest RSA-SHA256: `PASS`.
- ZIP checksum/file inventory/workbook sample: `PASS`.
- Tamper release bi verifier tu choi: `PASS`.
- Install/Uninstall `-VerifyOnly`: `PASS`.
- DT-605 acceptance: `14/14 PASS`.
- DT-606 semantic checkpoint: `PASS`.
- Private PFX trong workspace/release: `ABSENT`.
- Actual VSTO install/uninstall mutation: `NOT_RUN` tren may build vi Excel nguoi dung dang mo;
  can chay pilot deployment tren may test sach truoc khi mo rong pham vi phan phoi.
