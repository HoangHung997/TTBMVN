# Ma tran tuong thich

Cap nhat: 2026-08-03. Trang thai khong co may that la `NOT_RUN`, khong duoc quy doi thanh
PASS tu ket qua compile hoac suy luan API.

## Cau hinh da chay

| ID | Windows | Office/Excel | Bitness | Excel decimal/group/list | Trang thai |
|---|---|---|---|---|---|
| C-01 | Windows 11 Pro 10.0.26200 | Microsoft 365 16.0.20228.20124, vi-vn | x64 | `,` / `.` / `;` | PASS |
| C-02 | Nhu C-01 | Nhu C-01 | x64 | custom `,` / `.`; FormulaR1C1 multi-arg | PASS |
| C-03 | Nhu C-01 | Nhu C-01 | x64 | custom `.` / `,`; FormulaR1C1 multi-arg | PASS |
| C-04 | Build only | VSTO/.NET 4.8.1 | Any CPU | 68 Core test | PASS_BUILD_ONLY |
| C-05 | Build only | VSTO/.NET 4.8.1 | x64 | 68 Core test | PASS |

Bang chung may hien tai: `tmp/DT-603-compatibility-current.json`.

## Cau hinh chua co host

| ID | Windows | Office | Bitness | Trang thai | Dieu kien chuyen PASS |
|---|---|---|---|---|---|
| C-10 | Windows 10 22H2 | Office 2016 MSI | x86 | NOT_RUN | Cai/go cai, mo workbook, full smoke tren may that |
| C-11 | Windows 10 22H2 | Office 2016 MSI | x64 | NOT_RUN | Nhu C-10 |
| C-12 | Windows 10/11 | Office 2019 | x86 | NOT_RUN | Nhu C-10 |
| C-13 | Windows 10/11 | Office 2019 | x64 | NOT_RUN | Nhu C-10 |
| C-14 | Windows 10/11 | Office 2021 LTSC | x86 | NOT_RUN | Nhu C-10 |
| C-15 | Windows 10/11 | Office 2021 LTSC | x64 | NOT_RUN | Nhu C-10 |
| C-16 | Windows 10/11 | Microsoft 365 | x86 | NOT_RUN | Nhu C-10 |
| C-17 | Windows 10 | Microsoft 365 | x64 | NOT_RUN | Nhu C-10 |

`Release|Any CPU` build dat chi la bang chung binary khong hard-pin x64 o configuration
do; no khong chung minh VSTO/COM da chay tren Excel x86.

## Chinh sach ban phat hanh hien tai

- Cau hinh da chung nhan cho pilot: Windows 11 x64 + Microsoft 365 Excel x64 + .NET 4.8.1.
- Office 2016/2019/2021 va Excel x86 chua duoc tuyen bo ho tro thuong mai cho den khi cac
  dong `NOT_RUN` co bang chung may that.
- Interop types duoc embed va VSTO project target Office 15.0+, nhung day chi la kha nang
  binary, khong thay the acceptance tren host.
- Decimal/group separator lay tu `Application.International`; `FormulaR1C1` dung cu phap
  invariant, khong ghi `FormulaLocal`, nen khong phu thuoc list separator khi ghi cong thuc.

## Kich ban bat buoc moi host

1. Clean install VSTO va xac minh ribbon load.
2. Mo checkpoint, role/profile/package pin va E27.
3. Mo Ho dao, nhap H/V bang decimal separator cua Excel va preview/run tren working copy.
4. Mo Du toan, tra cuu, bang gia, don gia, phu luc, THKP, validation va audit.
5. Verify `.ttbupdate`, install/restart/rollback activation; workbook pin khong tu doi.
6. Chay uninstall/reinstall va mo lai workbook cu.
7. Tao support package va luu screenshot/log/build/Office inventory.
