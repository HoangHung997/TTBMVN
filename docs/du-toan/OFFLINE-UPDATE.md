# Goi cap nhat offline

## Pham vi DT-601

Goi `.ttbupdate` chi van chuyen mot regulation package da seal. Add-in xac minh goi
truoc khi hien preview hoac cai dat; khong co che do bo qua signature/checksum.

## Trust boundary

- Trust root phat hanh: `TTBMVN-OFFLINE-2026-01`, RSA 3072-bit, SHA-256/PKCS#1 v1.5.
- Add-in/Core chi chua public key trong `OfflineUpdateTrustCatalog`; public artifact o
  `data/regulations/update-keys/TTBMVN-OFFLINE-2026-01.ttbpub`.
- Private key khong co file trong workspace. No nam trong Windows CAPI user key container
  `TTBMVN_Offline_Update_Publisher_2026_01`, duoc tao voi `UseNonExportableKey`.
- Chi `ExcelAddIn1.RegulationTool` mo key container de ky. Add-in khong co API ky va khong
  nhan private key.
- Mat key container nghia la khong the phat hanh tiep bang key nay. Moi truong phat hanh
  thuong mai phai co backup/HSM va thu tuc xoay key rieng truoc release chinh thuc.

## Dinh dang

Archive ZIP khong co directory entry va chi gom:

1. `update.ttbmanifest`: UTF-8/LF canonical, schema 1.
2. `update.signature`: chu ky detached tren chinh byte manifest.
3. `payload/manifest.ttbmanifest`.
4. Sau module trong `payload/modules/*.data`.

Manifest chua update/package identity, minimum app version, signing key ID, thuat toan,
do dai va SHA-256 cua tung payload. Thu tu field va file la deterministic; timestamp ZIP
duoc co dinh de cung input/key tao cung noi dung logic.

## Fail-closed

Verifier tu choi truoc khi giai nen/cai dat neu gap mot trong cac dieu kien:

- archive qua lon, entry thua/thieu/trung, path traversal, reparse point;
- manifest sai schema, khong canonical hoac catalog khong dung 7 payload bat buoc;
- signing key khong thuoc trust store hoac signature sai;
- length/SHA-256 payload sai;
- package doc lai khong hop le hoac identity/checksum khong khop manifest;
- app thap hon minimum version, downgrade hoac cung version khac checksum.

Package cung version va checksum tra `AlreadyInstalled`; no khong duoc cai trung. Version
cu va moi van co the song song trong package store, con rollback se chon package da cai.

## Lenh phat hanh

Tao key mot lan tren may phat hanh:

```powershell
ExcelAddIn1.RegulationTool.exe keygen `
  TTBMVN_Offline_Update_Publisher_2026_01 `
  TTBMVN-OFFLINE-2026-01.ttbpub `
  TTBMVN-OFFLINE-2026-01
```

Tao va kiem tra goi:

```powershell
ExcelAddIn1.RegulationTool.exe build-update <bundle> <output.ttbupdate> 1.0.0.0 `
  TTBMVN_Offline_Update_Publisher_2026_01 <public-key.ttbpub>
ExcelAddIn1.RegulationTool.exe verify-update <output.ttbupdate> `
  <public-key.ttbpub> 1.0.0.0 [package-store]
```

Khong sao chep key container vao may build thong thuong. Viec ky release phai chay tren
may phat hanh rieng va luu checksum goi vao release manifest.
