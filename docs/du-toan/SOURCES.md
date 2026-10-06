# Nguon phap ly va tham khao

Cap nhat danh sach nay truoc moi dot nhap package phap ly. Chi dung nguon chinh thuc de xac nhan gia tri; bai viet phan mem chi dung tham khao workflow.

## Nguon BQP chinh thuc

### Thong tu 101/2025/TT-BQP

- Ban hanh: 13/09/2025.
- Hieu luc: 28/10/2025.
- Trang thai khi kiem tra 03/08/2026: con hieu luc.
- Link thuoc tinh: https://vbpl.vn/TW/Pages/vbpq-thuoctinh.aspx?ItemID=182364&dvid=13
- Link toan van: https://vbpl.vn/boquocphong/Pages/vbpq-toanvan.aspx?ItemID=182364
- PDF BQP: https://mod.gov.vn/wcm/connect/ead80087-4eed-4bdd-b739-0961956cf8e2/TT101.2025BQP.pdf?CACHEID=ROOTWORKSPACE-ead80087-4eed-4bdd-b739-0961956cf8e2-pBxSjyH&MOD=AJPERES

### Van ban hop nhat sau TT101/2025

- 94/VBHN-BQP: TT195/2019 va phan sua doi lien quan.
  - https://vbpl.vn/TW/Pages/vbpq-thuoctinh-hopnhat.aspx?ItemID=184657
- 95/VBHN-BQP: quy trinh ky thuat TT121/2021 + TT101/2025.
  - https://vbpl.vn/TW/Pages/vbpq-thuoctinh-hopnhat.aspx?ItemID=184660
- 96/VBHN-BQP: don gia ca may TT122/2021 + TT101/2025.
  - https://vbpl.vn/boquocphong/Pages/vbpq-thuoctinh-hopnhat.aspx?ItemID=184661&dvid=314
- 97/VBHN-BQP: dinh muc va quan ly chi phi TT123/2021 + TT101/2025.
  - https://vbpl.vn/boquocphong/Pages/vbpq-thuoctinh-hopnhat.aspx?ItemID=184663&dvid=314
- 98/VBHN-BQP: QCVN 01:2022/BQP va phan sua doi lien quan.
  - https://vbpl.vn/TW/Pages/vbpq-thuoctinh-hopnhat.aspx?ItemID=184665

### Raw evidence DT-302

- `data/regulations/raw/2025/TT101-2025-BQP.pdf`: 69 trang, SHA-256
  `E445F63CFDB515CB187780247423610A5A29DD0BBFDB231EBC21D858516DD240`.
- `data/regulations/raw/2025/VBHN94-2025-BQP.pdf`: 30 trang, SHA-256
  `D375AAC8F0C0B016D0C8B1E8CC65140560DF23D13AB3B499DE581DC5276F8935`.
- `data/regulations/raw/2025/VBHN95-2025-BQP.pdf`: 99 trang, SHA-256
  `2D5207B562EC63A24BD0B9D38C86E748AF6F7A4009195250517B83EFE22A1094`.
- `data/regulations/raw/2025/VBHN96-2025-BQP.pdf`: 35 trang, SHA-256
  `10A25B90BC95F8EB8330C9CD6D2423AD0DF46746D95DB7D33C7A2157228AE0B1`.
- `data/regulations/raw/2025/VBHN97-2025-BQP.pdf`: 52 trang, SHA-256
  `92B6C42996E554A908661B0395CE7286D41DDD69C17859B40993DC46300B69E0`.
- `data/regulations/raw/2025/VBHN98-2025-BQP.pdf`: 76 trang, SHA-256
  `62CC32BDA9ADED4310AB183180D6D4D861E41FEA94852848425EA0C19BAFE0A2`.
- Text-layer manifest: `data/regulations/text/2025/text-manifest.tsv`.
- Package/audit: `data/regulations/packages/BQP-RPBM-2025`.
- Package checksum hien tai sau khi DT-303 hoan thien du lieu Bang 01/03:
  `D58CAC64422E14151FD68F1C92944FAB7A9859908165C59F92089E7F13998B7C`.
- MachineRate module checksum:
  `6352542DC1F488FEA3D6B5D05D16E37D80464FEC91C09A0486FA4E7839BE8CBE`.

### Nguon BXD hien hanh dung cho K2/K5

- TT36/2026/TT-BXD, hieu luc 01/07/2026, bai bo TT11/2021/TT-BXD.
  - https://vanban.chinhphu.vn/?docid=218629&pageid=27160
  - Phu luc III bang 3.7 giu bang ty le K2.
- TT38/2026/TT-BXD, hieu luc 01/07/2026, thay the TT12/2021/TT-BXD.
  - https://vanban.chinhphu.vn/?docid=218632&pageid=27160&typegroupid=6
  - Phu luc VIII bang 2.24 la bang K5 hien hanh, co them moc 5.000/8.000/10.000 ty.
- Raw PDF va checksum: `data/regulations/raw/2026/README.md`.
- Audit DT-304: `data/regulations/packages/BQP-RPBM-2025/audit/DT-304-norm-cost-review.md`.
- Audit/Gate DT-305: `data/regulations/packages/BQP-RPBM-2025/audit/DT-305-gate-review.md`.
- Package checksum hien tai sau DT-305:
  `C23A3AE526841473AEAB6E36CCC76B0663A56C5FE63D3351F7AFBFA4639447F2`.
- Ban va du lieu hien hanh tu DT-406: `BQP-RPBM-2025@2.0.1`, checksum
  `E2537C08E7156414585BEFC446E74DD51E79AB8D67A3774B1F5250E7EECCDA90`;
  sua `MAT-RED-FLAG-LARGE` theo VBHN97/2025 trang 25. Checksum 2.0.0 o tren duoc
  giu lam bang chung Gate B3 lich su.

### Van ban nen nam 2021

- TT121/2021/TT-BQP: quy trinh ky thuat.
  - https://vbpl.vn/TW/Pages/vbpq-toanvan.aspx?ItemID=149773&Keyword=
- TT122/2021/TT-BQP: don gia ca may va thiet bi.
  - https://vbpl.vn/TW/Pages/vbpq-toanvan.aspx?ItemID=149812
- TT123/2021/TT-BQP: dinh muc du toan va quan ly chi phi.
  - https://vbpl.vn/boquocphong/Pages/vbpq-toanvan.aspx?ItemID=149809

### Raw evidence DT-301

- `data/regulations/raw/2021/TT121-2021-BQP.pdf`: 99 trang, SHA-256
  `9C0C9870EFD49FA49F7CC345B9D6D538D62AB5917EC9B68AE0B0C1E167363BA3`.
- `data/regulations/raw/2021/TT122-2021-BQP.pdf`: 28 trang, SHA-256
  `C5F78850F99411518267445B83F97124843666EEECDE06D5946C3DC24D0407EC`.
- `data/regulations/raw/2021/TT123-2021-BQP.pdf`: 39 trang, SHA-256
  `DE436E3BE81CE8B62A089D5B36C8A7A931BDB9EAC468D65B110455CB1F0DCF2E`.
- OCR co manifest tai `data/regulations/ocr/2021/ocr-manifest.tsv`; chi dung tim trang vi
  PDF la ban scan. Gia tri phat hanh phai doi chieu anh PDF.
- Package/audit: `data/regulations/packages/BQP-RPBM-2021`.
- Package checksum:
  `BBE95EFDF3B2DA0780CAFEEF828C9418644452655C733C1BF2417F234C3E8008`.

## Nguyen tac xac minh van ban moi

- Truoc moi release du lieu, tim kiem lai CSDL VBPL cua Bo Quoc phong.
- Kiem tra trang thai hieu luc, lich su sua doi, van ban thay the va van ban hop nhat.
- Ghi ngay kiem tra va nguoi kiem tra vao manifest package.
- Khong ket luan `moi nhat` chi tu ten/nam van ban.

## Xac minh cho resolver DT-203

Kiem tra lai ngay 03/08/2026 tren CSDL VBPL chinh thuc:

- TT121/2021/TT-BQP, Dieu 2: hieu luc tu 05/11/2021.
- TT122/2021/TT-BQP, Dieu 7 va Dieu 8: ho so da phe duyet truoc ngay hieu luc giu
  don gia ca may da phe duyet; hieu luc tu 05/11/2021.
- TT123/2021/TT-BQP, Dieu 5 va Dieu 6: ho so da phe duyet truoc ngay hieu luc giu
  dinh muc/chi phi da phe duyet; hieu luc tu 05/11/2021.
- TT101/2025/TT-BQP, Dieu 6 va Dieu 7: phuong an ky thuat thi cong/du toan da duoc
  phe duyet truoc ngay hieu luc tiep tuc theo ban da phe duyet; hieu luc 28/10/2025.

Quyet dinh resolver:

- Ho so da phe duyet truoc 28/10/2025 duoc de xuat package 2021 theo transition rule
  `BQP-TT101-2025-ARTICLE-6`.
- Ho so chua phe duyet duoc resolve tai ngay tinh/kiem tra, khong tu dong bao luu chi vi
  ngay lap nam truoc moc hieu luc.
- Package `Withdrawn` khong duoc de xuat; package cu can cho transition ma khong co trong
  kho phai tra loi ro `RequiredTransitionPackageUnavailable`.

## Tham khao phan mem du toan

### G8

- Workflow cap nhat phan mem, dinh muc, du lieu dia phuong va chuyen doi ho so cu:
  - https://phanmemg8.vn/cap-nhat-thong-tu-moi-tren-phan-mem-du-toan-g8/

### F1

- Tach cap nhat phien ban phan mem va cap nhat don gia/du lieu:
  - https://dutoanf1.com.vn/bao-gia-phan-mem-du-toan-f1/
- Workflow du toan moi va chu dong tra lai cong tac cho ho so cu:
  - https://dutoanf1.com.vn/ap-dung-gia-nhan-cong-ca-may-theo-37-2026-tt-bxd-vao-bo-dinh-muc-38-2026-tt-bxd-tren-du-toan-f1/

Du lieu cua G8/F1 khong phai nguon phap ly cho module RPBM; chi dung de hoc cach quan ly cap nhat, migration va trai nghiem nguoi dung.
