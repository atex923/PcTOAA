# PcTOAA

KOHAKU 預算書轉換容器。

## 目前正式版

- `PsTOAA_V4_5_6.py`
- `PsTOAA_V4_5_6.pyw`

目前正式版：V4.5.6

視窗標題：`KOHAKU預算書轉換容器V4.5.6`

匯出預設檔名：`AA3118_工程編號_YYYYMMDDHHMMSS`

## 使用方式

命令列執行：

```bash
python3 PsTOAA_V4_5_6.py
```

Windows 可使用 `.pyw` 啟動無 console 視窗版：

```bash
python3 PsTOAA_V4_5_6.pyw
```

## Nuitka 打包

建議使用 Python 3.10 進行 Nuitka 打包，並明確包含 Excel 讀寫套件：

```powershell
python -m nuitka `
  --standalone `
  --onefile `
  --assume-yes-for-downloads `
  --windows-console-mode=disable `
  --enable-plugin=pyside6 `
  --include-package=pandas `
  --include-package=openpyxl `
  --include-package=xlrd `
  --nofollow-import-to=*.tests `
  --noinclude-pytest-mode=nofollow `
  --noinclude-unittest-mode=nofollow `
  --windows-icon-from-ico="C:\Users\newat\Downloads\icon_result (2).ico" `
  --output-filename=PsTOAA_V4_5_6.exe `
  PsTOAA_V4_5_6.pyw
```

## 版本進版規則

每次修改程式都必須進版第三碼，例如 `V4.5.6` 的下一次修改為 `V4.5.7`。

根目錄只保留目前正式版的 `.py` 與 `.pyw`。

`.pyw` 必須與 `.py` 保持完整同步碼，不使用精簡 launcher。

舊版 `.py` 與 `.pyw` 必須移到 `history/versions/`，避免 A 電腦與 B 電腦同步時誤用舊版或互相覆蓋。

## 歷史版本

舊版程式檔保留於 `history/versions/`。

版本整理紀錄請參考 `history/README.md`。
