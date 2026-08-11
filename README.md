# Buyer Reports 自動整理工具

由 **DPS 原始** 與 **PP 原始** 資料自動生成「整理後」報表。
數值一律以原始檔為唯一來源，人工整理版只在 `--compare` 模式下用來對帳。

## 資料整理規則與前置條件

### DPS

1. 檔名需包含 `DPS` 字眼，且原始資料工作表名稱需為 `DPS原始`。
2. 表頭列的欄位名稱必須一致，且需包含：`Line`、`W/O`、`AVTC P/N`。
3. 日期欄必須是 Excel 日期格式；若日期是用公式累加產生也可以，但檔案需經 Excel 計算並儲存，讓程式讀得到日期值。
4. 正常格式下，每個日期應該有兩欄，也就是 `D` / `N` 兩班。程式預設會把同一天的欄位加總。
5. 料號欄必須使用 `AVTC P/N` 這欄。
6. `AVTC P/N` 結尾帶 `*` 的料號預設會被排除。

### PP

1. 檔名需包含 `PP` 字眼，且檔案內必須有 pivot cache。
2. pivot cache 內必須有這些欄位，且名稱需一致：`Plan`、`Customer`、`Model`、`AVTC FG Part Number`。
3. `Plan` 欄位的值預設必須是 `Production Input`，程式才會納入計算。
4. 週別與月份欄位需符合 Excel 樞紐表 / pivot cache 的格式。
5. 跨月但同週的項目會合併加總，例如 `WK36` 和 `WK36 Sep` 會寫進同一個 `WK36`。
6. 程式會抓 pivot cache 的更新日所在週再減一，作為輸出起始週。例如更新日是 `2026-07-31`，屬於第 31 週，所以程式會從 `WK30` 開始。
7. 沒有週明細的月份會直接以月份欄位作為輸出。

## 目錄結構

```
buyer-reports/
├── build_exe.bat               # Windows 打包腳本（產生 exe）
├── generate_buyer_reports.py   # 主程式
├── intput/                     # 輸入：把來源 Excel 放這裡
│   ├── TV DPS Jul 31-Ver 2.xlsx
│   └── AVTC TV MNT VC PP 20260730 update (002).xlsx
├── output/                     # 輸出：執行時自動建立（已列入 .gitignore）
│   ├── DPS整理後.xlsx
│   └── PP整理後.xlsx
├── requirements.txt
├── Windows執行檔(exe)使用說明.txt
└── README.md
```

## 安裝與執行

```bash
pip install -r requirements.txt

# 最常用：自動抓 intput/ 下的檔案，產出到 output/
python generate_buyer_reports.py

# 產出後順便跟來源檔內既有的人工整理版逐格對帳
python generate_buyer_reports.py --compare
```

輸入檔以檔名關鍵字自動辨識（含 `DPS` → DPS 檔，含 `PP` → PP 檔）；
同類有多個檔案時取修改時間最新的那個，並會提示。也可用 `--dps` / `--pp` 明確指定。
每次執行都會寫入 `output/run.log`，方便回查成功訊息或錯誤原因。

## Windows 執行檔

正式給 Windows 使用者時，建議交付整個資料夾，而不是只給單一 exe：

```
BuyerReports/
├── BuyerReports.exe
├── intput/
├── output/
└── Windows執行檔(exe)使用說明.txt
```

使用者只要把 DPS / PP Excel 放進 `intput/`，再雙擊 `BuyerReports.exe`，
完成後到 `output/` 取得 `DPS整理後.xlsx` 與 `PP整理後.xlsx`。
exe 會以自身所在資料夾為根目錄，因此整包資料夾可搬到其他位置使用。

若要打包 exe，請先在 Windows 安裝 Python 3，確認 `py -3 --version` 或
`python --version` 其中一個可執行，然後執行：

```bat
build_exe.bat
```

打包完成後產物會在 `release/BuyerReports/`。腳本會自動建立 `intput/`、
`output/`，並複製 `Windows執行檔(exe)使用說明.txt`。

## 資料規律（本工具依據的規則）

### DPS

| 項目 | 規則 |
|---|---|
| 來源工作表 | `DPS原始` |
| 表頭 | 自動尋找同時含 `Line` / `W/O` / `AVTC P/N` 的那一列（目前是第 9 列） |
| 日期欄 | 表頭為日期型別的欄位。正常格式下，每個日期應該有兩欄，也就是 `D` / `N` 兩班。程式預設會把同一天的欄位加總 |
| 列標籤 | 料號欄必須使用 `AVTC P/N`；`AVTC P/N` 結尾帶 `*` 的料號預設會被排除（與人工整理一致，可用 `--include-star-parts` 保留） |
| 排序 | 模擬 Excel 樞紐：數字型標籤在前（依數值），文字型在後 |
| 空值 | 數量為 0 時留白，不填 0（與人工版一致） |
| total | 寫成 `=SUM(...)` 公式 |

**末欄彙總桶**：人工版是 Excel 樞紐日期群組的產物，最後一個日期欄其實是
「> 前一日」的彙總桶（本次檔案為 `9/4`，等於 9/4~9/26 全部相加）。
`--dps-tail-cutoff` 控制此行為：

- `auto`（預設）— 若來源檔內還有人工版 `DPS整理後` 或舊版 `DPS整理后`，就沿用它的末欄日期
- `none` — 每個日期各自成欄（未來只拿到原始檔時的預設結果）
- `YYYY-MM-DD` — 指定彙總桶起始日

### PP

PP 檔內**沒有原始資料工作表**：逐料號明細只存在於樞紐快取
（`xl/pivotCache/`，來源指向外部活頁簿的 `Raw Data`）。本工具直接解析快取。

| 項目 | 規則 |
|---|---|
| 篩選 | `Plan` 欄位的值預設必須是 `Production Input`，程式才會納入計算（可用 `--pp-plan` 改） |
| 列 | `AVTC FG Part Number`，期間內合計為 0 者不輸出 |
| 欄 | 由可見樞紐報表的欄位版面自動推導 |
| 重複週次 | 跨月重複的週（如 `WK36` 同時在 Aug/Sep）**相加為單一欄** |
| 月小計欄 | 已有週明細的月份，其月欄是小計，不重複輸出 |
| 年度合計欄 | 略過 |
| 列序 | 依可見樞紐報表的客戶顯示順序，同客戶內依料號排序 |

**期間欄怎麼決定的**：程式讀取可見樞紐報表中描述版面的那一列
（`WK27 Jul | WK28 | … | Oct-26 | Nov26FCST | 2026 TOTAL | Jan'27 FCST | …`），
從起始週開始往右取，跳過小計欄。因此「週明細到哪個月為止、之後改用月預測」
會跟著來源檔自動調整，不需改程式。

起始週預設為**報表基準日所在週的前一週**（保留上一週作參照）；
報表基準日預設取樞紐快取的更新日期。可用 `--pp-start-week` / `--pp-report-date` 覆寫。

## 常用參數

| 參數 | 說明 |
|---|---|
| `--input-dir` / `--out-dir` | 輸入 / 輸出資料夾 |
| `--dps` / `--pp` | 明確指定來源檔 |
| `--skip-dps` / `--skip-pp` | 只跑其中一種報表 |
| `--include-star-parts` | 保留結尾帶 `*` 的 DPS 料號 |
| `--dps-tail-cutoff` | `auto` / `none` / `YYYY-MM-DD` |
| `--pp-plan` | Plan 篩選值，預設 `Production Input` |
| `--pp-start-week` | `auto` 或週數 |
| `--pp-base-year` | 主年度兩位數，例 `26` |
| `--pp-report-date` | 報表基準日 `YYYY-MM-DD` |
| `--compare` | 與來源檔內的人工整理版逐格對帳 |
| `--quiet` | 只輸出錯誤訊息 |
| `--no-pause` | Windows exe 模式下完成後不等待按 Enter |

## 對帳結果（2026-07 版來源檔）

```
=== 對帳：DPS ===
  人工版 143 列 / 本次產出 143 列
  數值差異：0 格

=== 對帳：PP ===
  人工版 109 列 / 本次產出 111 列
  ▲ 人工版漏記（原始有、人工無）共 2 筆：
      + P75WWCV22EO-.CG75VT1  {'WK32': 1241.0}
      + P98WWCV22EO-.CG98VT1  {'WK30': 6.0}
  數值差異：2 格
      ! M27USSS22ESH.HG356-1 | WK43 | 人工=1600 原始=0
      ! M27USSS22ESH.HG35C-1 | WK43 | 人工=0 原始=1600
```

DPS 完全一致。PP 的 3 處差異均為人工整理疏漏，本工具產出的是原始檔的正確值：

1. 漏記 `P75WWCV22EO-.CG75VT1`（CVTE VS）WK32 = 1,241 pcs
2. 漏記 `P98WWCV22EO-.CG98VT1`（CVTE VS）WK30 = 6 pcs
3. WK43 的 1,600 pcs 記到相鄰料號 `…HG356-1`，應為 `…HG35C-1`

> 兩筆漏記都是 CVTE VS 客戶——人工版的客戶區塊直接從 Telly 跳到 Ubiquiti，
> 整個 CVTE 群組被跳過。每月產出後建議用 `--compare` 複核。

## 內建檢核

每次執行都會列印：來源檔、表頭列、日期欄數與區間、被排除的 `*` 料號清單與數量、
樞紐快取的更新日期與更新者、Plan 篩選筆數、期間欄清單、輸出列數與合計。
另外會在下列情況示警：

- 某個日期的欄數不是 2（D/N 兩班）
- 表頭未被判定為日期、但底下有數量的欄位（避免靜默漏算）
- 樞紐快取更新日距今超過 45 天（可能是舊快照）
- 找不到可見樞紐版面，改用推導模式

## 版面說明

輸出刻意貼近人工版：DPS 為 72 欄 x（143+1）列、`total` 以 `=SUM()` 公式呈現；
PP 的 `total` 同樣保留在第 31 欄（AE），前面留 3 個空白欄，與人工版一致。
差別只有一處：PP 的 A/B/C 欄補上了 `Customer` / `AVTC FG Part Number` / `Model`
標題（人工版該三格是空的）。

產出檔會自動在標題列套用篩選下拉（AutoFilter）。若來源檔內已有人工整理版
工作表（DPS：`DPS整理後` / `DPS整理后`；PP：`PP整理後` / `整理后PP`），
程式會參考該工作表複製欄寬、列高、字型、填色、框線、對齊與數字格式。
若未來只拿到原始資料、沒有人工整理版工作表，程式仍會用內建基本版面產出報表，
並保留篩選下拉功能。

## 已知限制

- PP 的數字來自樞紐快取**快照**。若提供者交檔前未 refresh 樞紐，工具讀到的會是舊數字。
  最穩健的做法是直接取得 `Raw Data` 原始活頁簿，就能完全繞開快取解析。
- DPS 表頭日期在原始檔中是公式，需由 Excel 存檔（含快取值）才讀得到。
  若偵測到此問題會示警並列出受影響欄位。
