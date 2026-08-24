# Buyer Reports 自動整理工具

由 **DPS 來源工作表** 與 **PP 樞紐快取** 自動生成「整理後」、`DPS+PP` 與 `CTB` 報表。
數值一律以原始檔為唯一來源，人工整理版只在 `--compare` 模式下用來對帳。
目前支援 AVTC / RAKEN 兩組輸入資料夾；RAKEN 的 DPS 會合併多份 DPS 檔，
同料號同日期累加。

## 資料整理規則與前置條件

### DPS

1. 檔名需包含 `DPS` 字眼；程式會依活頁簿順序抓第一張名稱包含 `DPS` 的工作表作為來源。
2. 表頭列需包含：`Line`、`W/O`，以及 `buyer_reports.ini` 內設定的其中一個 DPS 料號欄別名。
3. 日期欄必須是 Excel 日期格式；若日期是用公式累加產生也可以，但檔案需經 Excel 計算並儲存，讓程式讀得到日期值。
4. 正常格式下，每個日期應該有兩欄，也就是 `D` / `N` 兩班。程式預設會把同一天的欄位加總。
5. 料號欄預設可使用 `AVTC P/N`、`P/N` 或 `Model`。
6. 料號結尾帶 `*` 的資料預設會被排除。
7. AVTC DPS 維持「第一份格式正確檔案」的原本邏輯；RAKEN DPS 會合併資料夾內所有格式正確的 DPS 檔。

### PP

1. 檔名需包含 `PP` 字眼。
2. 程式會依活頁簿順序抓第一張名稱包含 `PP` 或 `Data` 且含樞紐分析表的工作表，作為欄位版面與客戶排序來源。
3. 該樞紐分析表對應的 pivot cache 內必須有 `Plan`、`Customer`、`Model`，並有名稱包含 `Part Number` 的料號欄位。
4. `Plan` 欄位的值預設必須是 `Production Input`，程式才會納入計算。
5. 週別與月份欄位需符合 Excel 樞紐表 / pivot cache 的格式。
6. 跨月但同週的項目會合併加總，例如 `WK36` 和 `WK36 Sep` 會寫進同一個 `WK36`。
7. 程式會抓 pivot cache 的更新日所在週再減一，作為可見版面起始週。例如更新日是 `2026-07-31`，屬於第 31 週，所以可見版面會從 `WK30` 開始銜接。
8. 起始週以前、同年度的 WK 週欄會從 pivot cache 補齊輸出，即使來源樞紐欄位清單未勾選也會顯示在整理後檔案中，但不納入 `total`。
9. 來源樞紐報表中被隱藏的時間欄位也會納入計算，並在整理後檔案中顯示出來。
10. 沒有週明細的月份會直接以月份欄位作為輸出。

### DPS+PP

1. `DPS+PP.xlsx` 會在同一個客戶資料夾內同時找到 DPS 與 PP 來源後產出。
2. 週別採 buyer week：週六到週五。例如 `2026-08-08` 到 `2026-08-14` 是 `WK33`。
3. 自動判斷目前週時，週六到週四用執行日所在 buyer week；週五會提前使用隔天週六的新 buyer week。例如 `2026-08-15` 到 `2026-08-20` 會用 `WK34`，`2026-08-21` 會提前用 `WK35`。
4. AVTC 預設 DPS 保留「含目前 buyer week 共 5 週」；例如目前 `WK34`，DPS 到 `WK38`，PP 從 `WK39` 開始。
5. RAKEN 預設 DPS 保留「含目前 buyer week 共 3 週」；例如目前 `WK34`，DPS 到 `WK36`，PP 從 `WK37` 開始。Windows exe 雙擊執行時，若 `input/RAKEN` 有 Excel，會先跳出視窗讓使用者選擇 RAKEN 保留 2 / 3 / 4 週。
6. DPS 若在 PP 接續週含之後仍有數字，會併入 DPS 保留週的最後一天。例如 AVTC 的 `WK39` 後數字會併入 `WK38` 最後一天。
7. PP 接續週開始後的週欄與月 FCST 欄會沿用 PP pivot cache 的數字。
8. 若來源檔內有 `BOM1` 工作表，程式會用 `BOM1` 的 B 欄料號對應 H 欄 vendor 寫入 `BOM`；若沒有，`BOM` 欄留空。

### CTB

1. 每個客戶資料夾若同時有 CTB 所需來源 sheet，且本次 `DPS+PP.xlsx` 成功產出，會額外產出 `CTB.xlsx`。
2. `BOM1` 用來把 `DPS+PP` 的成品需求依 `USE` 展開成子件需求。
3. `Shortage` 用來取得 `QN` 欄，也就是來源檔的 `Overshortage1`，作為 CTB 的 `OVER SHORTAGE` / balance 起始 shortage 基準。
4. `open po` 用 `Item + Supplier Site` 產生 key，對應 `Quantity Due` 與 `Need By Date`，產出 ETA / PO Remain 列。
5. CTB 的 D 欄 key 由 C 欄 `Part No` + E 欄 `Code` 組成；E 欄 `Code` 來自 `open po` 的 `Supplier Site`。
6. ETA 日期會依 `open po` 的 `Need By Date` 放到相同或下一個可用期間欄；若人工曾調整 ETA 日期，自動產出日期可能與人工版不同。
7. Balance 由程式逐期計算：上一期 balance + 上一期 ETA - 本期 demand - 本期 other。
8. 若同一客戶資料夾內另有原始 `CTB` sheet，程式會沿用該 sheet 的表頭、日期欄、列順序與格式作為版型，但 A:M 資料列會依自動化規則重建，不直接抄原始 CTB 的人工欄位；若沒有原始 `CTB` sheet，仍會使用程式新建版面輸出。

## 目錄結構

```
buyer-reports/
├── build_exe.bat               # Windows 打包腳本（產生 exe）
├── generate_buyer_reports.py   # 入口檔
├── buyer_reports.ini           # sheet / 欄位偵測關鍵字設定
├── buyer_reports/              # 主程式模組
│   ├── common.py               # 共用工具、log、Excel helper、Windows exe 判斷
│   ├── dps.py                  # DPS 解析與輸出
│   ├── pp.py                   # PP 樞紐快取解析、期間推導與輸出
│   ├── dps_pp.py               # DPS+PP 整合輸出
│   ├── ctb.py                  # CTB 整合輸出
│   ├── compare.py              # 對帳工具
│   └── runner.py               # CLI、找檔與執行流程
├── input/                      # 輸入：把來源 Excel 放這裡
│   ├── AVTC/
│   │   ├── AVTC 的 DPS 檔
│   │   ├── AVTC 的 PP 檔
│   │   ├── 含 BOM1 / open po 的檔案（若要產 CTB）
│   │   └── 含 Shortage 的檔案（若要產 CTB）
│   └── RAKEN/
│       ├── RAKEN 的 DPS 檔，可放多份
│       ├── RAKEN 的 PP 檔
│       ├── 含 BOM1 / open po 的檔案（若要產 CTB）
│       └── 含 Shortage 的檔案（若要產 CTB）
├── output/                     # 輸出：執行時自動建立（已列入 .gitignore）
│   ├── AVTC/
│   │   ├── DPS整理後.xlsx
│   │   ├── PP整理後.xlsx
│   │   ├── DPS+PP.xlsx
│   │   └── CTB.xlsx
│   └── RAKEN/
│       ├── DPS整理後.xlsx
│       ├── PP整理後.xlsx
│       ├── DPS+PP.xlsx
│       └── CTB.xlsx
├── requirements.txt
├── Windows執行檔(exe)使用說明.txt
└── README.md
```

## 安裝與執行

```bash
pip install -r requirements.txt

# 最常用：自動抓 input/AVTC、input/RAKEN 下的檔案，產出到 output/
python generate_buyer_reports.py

# 產出後順便跟來源檔內既有的人工整理版逐格對帳
python generate_buyer_reports.py --compare
```

輸入檔以檔名關鍵字自動辨識（含 `DPS` → DPS 檔，含 `PP` → PP 檔）。
AVTC 同類有多個檔案時會依修改時間由新到舊嘗試，第一份格式正確者產出；
RAKEN DPS 會把所有格式正確的 DPS 檔合併，格式不符的檔案會被略過並列出警告。
也可用 `--dps` / `--pp` 明確指定。
每次執行都會在本次處理的客戶輸出資料夾寫入 `log`，例如
`output/RAKEN/log`，方便回查成功訊息或錯誤原因。
同一客戶同時有可用 DPS 與 PP 時，還會額外產出 `DPS+PP.xlsx`。
若同一客戶資料夾也有 `BOM1`、`open po`、`Shortage` 工作表，會再產出 `CTB.xlsx`。
若完全沒有 CTB 來源 sheet，程式只會產出 DPS / PP / DPS+PP，不會嘗試產 CTB；若只有部分 CTB 來源，CTB 會略過並在 log 中顯示原因。

## 偵測關鍵字設定

程式啟動時會讀取專案根目錄或 exe 同層的 `buyer_reports.ini`。若檔案不存在，
會使用內建預設值：DPS 抓名稱包含 `DPS` 的 sheet，PP 抓名稱包含 `PP` 或
`Data` 且含樞紐分析表的 sheet。DPS 料號欄可接受 `AVTC P/N`、`P/N`、`Model`，
PP 樞紐快取的料號欄會抓名稱包含 `Part Number` 的欄位；若有多個符合欄位，
會優先選擇包含 `FG` 的欄位。
RAKEN 預設會略過 DPS 中日期數量全為 0 的列，避免尾端備註或簽核文字被當成料號；
同時只輸出到最後一個有數量的日期，尾端全空日期欄會被略過。

```ini
[sheet_detection]
dps_sheet_keywords = DPS
pp_sheet_keywords = PP, Data

[dps]
part_number_headers = AVTC P/N, P/N, Model

[pp]
part_number_field_keywords = Part Number

[customers]
names = AVTC, RAKEN

[customer.AVTC]
dps_mode = first_valid
pp_mode = first_valid
dps_drop_zero_total_rows = false
dps_trim_trailing_zero_dates = false
dps_pp_dps_weeks_ahead = 5

[customer.RAKEN]
dps_mode = merge_all
pp_mode = first_valid
dps_drop_zero_total_rows = true
dps_trim_trailing_zero_dates = true
dps_pp_dps_weeks_ahead = 3
```

未來若 PP sheet 又有新命名，例如 `Pivot`，或 DPS 料號欄又有新表頭，可直接改成：

```ini
[sheet_detection]
pp_sheet_keywords = PP, Data, Pivot

[dps]
part_number_headers = AVTC P/N, P/N, Model, Buyer P/N
```

Windows 使用者只要修改 exe 同層的 `buyer_reports.ini` 即可，不需要重新 build。
其中 `dps_mode = first_valid` 代表取第一份格式正確的 DPS，`dps_mode = merge_all`
代表合併該客戶資料夾內所有格式正確的 DPS。
`dps_drop_zero_total_rows = true` 代表 DPS 來源列若所有日期欄都沒有非 0 數量，
就不輸出該列；RAKEN 預設開啟，AVTC 預設關閉。
`dps_trim_trailing_zero_dates = true` 代表 DPS 輸出只列到最後一個有非 0 數量的日期；
只會裁掉尾端連續空白日期，中間空白日期仍保留。RAKEN 預設開啟，AVTC 預設關閉。
`dps_pp_dps_weeks_ahead` 代表 `DPS+PP` 中 DPS 要保留幾週，會包含目前 buyer week；
目前 buyer week 在自動模式下會套用週五提前使用下一週的規則。

## Windows 執行檔

正式給 Windows 使用者時，建議交付整個資料夾，而不是只給單一 exe：

```
BuyerReports/
├── BuyerReports.exe
├── _internal/
├── buyer_reports.ini
├── input/
│   ├── AVTC/
│   └── RAKEN/
├── output/
│   ├── AVTC/
│   └── RAKEN/
└── Windows執行檔(exe)使用說明.txt
```

使用者只要把 AVTC 檔案放進 `input/AVTC/`、RAKEN 檔案放進 `input/RAKEN/`，
再雙擊 `BuyerReports.exe`。完成後到 `output/AVTC/`、`output/RAKEN/`
取得各自的 `DPS整理後.xlsx` 與 `PP整理後.xlsx`。
若 DPS 與 PP 來源都可用，也會產出 `DPS+PP.xlsx`。
若該客戶資料夾內另有 `BOM1`、`open po`、`Shortage` 工作表，也會產出 `CTB.xlsx`。
若 `input/RAKEN/` 內有 Excel，exe 啟動後會先跳出 RAKEN `DPS+PP` 的 DPS 保留週數選擇視窗，
可選含本週共 2 / 3 / 4 週；這只影響 `output/RAKEN/DPS+PP.xlsx`。
exe 會以自身所在資料夾為根目錄，因此整包資料夾可搬到其他位置使用；
請務必先完整解壓縮，不要在 zip 壓縮檔視窗內直接執行。
若 SmartScreen 顯示「Windows 已保護你的電腦」，請確認來源可信後點
「更多資訊」→「仍要執行」。
執行時 console 會顯示階段式進度條；若某一種報表因格式錯誤失敗，另一種報表仍會繼續嘗試產出。
若失敗報表的同名輸出檔已存在，摘要會提醒該檔可能是前次執行留下的舊檔。

若要打包 exe，請先在 Windows 安裝 Python 3，確認 `py -3 --version` 或
`python --version` 其中一個可執行，然後執行：

```bat
build_exe.bat
```

打包完成後產物會在 `release/BuyerReports/`。腳本會自動建立 `input/`、
`output/`，並複製使用說明與 `buyer_reports.ini`。若要提供 zip 給使用者，
可自行壓縮要交付的 release 資料夾；使用者仍需要先完整解壓縮後再執行。

## 資料規律（本工具依據的規則）

### DPS

| 項目 | 規則 |
|---|---|
| 來源工作表 | 依活頁簿順序取第一張名稱包含 `buyer_reports.ini` 內 DPS 關鍵字的工作表，預設為 `DPS` |
| 表頭 | 自動尋找同時含 `Line` / `W/O`，且含有 `buyer_reports.ini` 內 DPS 料號欄別名的那一列 |
| 日期欄 | 表頭為日期型別的欄位。正常格式下，每個日期應該有兩欄，也就是 `D` / `N` 兩班。程式預設會把同一天的欄位加總 |
| 列標籤 | 使用選定的 DPS 料號欄；料號結尾帶 `*` 者預設會被排除（與人工整理一致，可用 `--include-star-parts` 保留） |
| 排序 | 模擬 Excel 樞紐：數字型標籤在前（依數值），文字型在後 |
| 空值 | 數量為 0 時留白，不填 0（與人工版一致） |
| 尾端空白日期 | 可由 `dps_trim_trailing_zero_dates` 控制；RAKEN 預設略過最後有數量日期之後的全空日期欄 |
| total | 由程式計算後以數值寫入，不再使用 Excel 公式 |
| 客戶模式 | AVTC 預設 `first_valid` 且保留 0 數量列；RAKEN 預設 `merge_all`，會合併所有可用 DPS 檔，同料號同日期累加，並略過日期數量全為 0 的列 |

**末欄彙總桶**：人工版是 Excel 樞紐日期群組的產物，最後一個日期欄其實是
「> 前一日」的彙總桶（本次檔案為 `9/4`，等於 9/4~9/26 全部相加）。
`--dps-tail-cutoff` 控制此行為：

- `auto`（預設）— 若來源檔內還有人工版 `DPS整理後` 或舊版 `DPS整理后`，就沿用它的末欄日期
- `none` — 每個日期各自成欄（未來只拿到原始檔時的預設結果）
- `YYYY-MM-DD` — 指定彙總桶起始日

### PP

PP 檔內**沒有原始資料工作表**：逐料號明細只存在於樞紐快取
（`xl/pivotCache/`，來源指向外部活頁簿的 `Raw Data`）。本工具會先找第一張
名稱包含 `buyer_reports.ini` 內 PP 關鍵字且含樞紐分析表的工作表，預設關鍵字為
`PP`、`Data`。程式會拿該工作表判斷欄位版面與客戶順序，再直接解析該樞紐分析表
對應的快取取得數字。

| 項目 | 規則 |
|---|---|
| 篩選 | `Plan` 欄位的值預設必須是 `Production Input`，程式才會納入計算（可用 `--pp-plan` 改） |
| 列 | 使用名稱符合 `buyer_reports.ini` 內 PP 料號欄關鍵字的欄位，預設抓 `Part Number`；期間內合計為 0 者不輸出 |
| 欄 | 起始週以前的同年度 WK 從 pivot cache 補齊；起始週以後由選定 PP 工作表上的樞紐報表欄位版面自動推導 |
| 重複週次 | 跨月重複的週（如 `WK36` 同時在 Aug/Sep）**相加為單一欄** |
| 月小計欄 | 已有週明細的月份，其月欄是小計，不重複輸出 |
| 樞紐清單未勾選週欄 | pivot cache 中同年度、起始週以前的 WK 週欄會補齊輸出，但不納入 `total` |
| 隱藏欄 | 來源樞紐報表中 hidden 的 WK / FCST 時間欄仍會納入，整理後會強制顯示 |
| 年度合計欄 | 略過 |
| 列序 | 依選定 PP 工作表上的樞紐報表客戶顯示順序，同客戶內依料號排序 |
| total | 由程式計算後以數值寫入，不再使用 Excel 公式；PP 只從起始週開始累加，欄名會寫成如 `total(WK31起)` |

**期間欄怎麼決定的**：輸出期間分兩段處理。第一段會從 pivot cache 補齊主年度中
早於起始週的所有 WK 週欄；第二段讀取選定 PP 工作表上描述樞紐版面的那一列
（`WK27 Jul | WK28 | … | Oct-26 | Nov26FCST | 2026 TOTAL | Jan'27 FCST | …`），
從起始週開始往右取，跳過小計欄。因此來源樞紐清單未勾選的歷史週可以補回，
而 `total` 仍只從起始週開始累加；「週明細到哪個月為止、之後改用月預測」
也仍會跟著來源檔自動調整，不需改程式。
樞紐快取週欄目前支援 `WK31 26'Aug` 與 `WK31 Jul '26` 兩種格式。

起始週預設為**報表基準日所在週的前一週**（保留上一週作參照），它是可見版面
銜接點，不代表輸出檔第一個 WK 一定從該週開始；報表基準日預設取樞紐快取的
更新日期。可用 `--pp-start-week` / `--pp-report-date` 覆寫。

### DPS+PP

`DPS+PP.xlsx` 會直接讀取 DPS 原始解析結果與 PP pivot cache，而不是把
`DPS整理後.xlsx` / `PP整理後.xlsx` 兩個檔案再相加。這樣可以套用整合專用規則：

| 項目 | 規則 |
|---|---|
| 週別 | buyer week，週六到週五 |
| 目前週 | 預設依執行日推算；週六到週四用當週，週五提前用下一週。可用 `--dps-pp-current-week` 覆寫 |
| AVTC | DPS 保留含目前週共 5 週，後續改用 PP |
| RAKEN | 預設 DPS 保留含目前週共 3 週，後續改用 PP；Windows exe 可在啟動畫面選 2 / 3 / 4 週 |
| DPS 後段數字 | 若 DPS 在 PP 接續週含之後仍有數字，併入 DPS 保留週的最後一天 |
| PP 接續欄 | 只取 cutover 之後的 PP 週欄與月 FCST 欄 |
| DPS 尾端空白日期 | 若客戶設定啟用 `dps_trim_trailing_zero_dates`，只會裁掉 DPS 日期區段尾端全空日期欄，不會改變 PP 接續週 |
| BOM | 若來源檔有 `BOM1`，用 B 欄料號對應 H 欄 vendor；沒有則留空 |
| total | 由程式加總整份 `DPS+PP` 期間欄，以數值寫入 |

### CTB

`CTB.xlsx` 會以同一客戶資料夾內的來源 sheet 產出，不限定 AVTC 或 RAKEN。

| 項目 | 規則 |
|---|---|
| DPS+PP | 必須先在本次執行成功產出 `output/<客戶>/DPS+PP.xlsx` |
| CTB 來源完整性 | 若完全沒有 CTB 來源 sheet，程式不會嘗試產 CTB；若只有部分 CTB 來源，CTB 會略過並在 log 中顯示原因 |
| BOM1 | 來源檔需有 `BOM1` sheet，且表頭需有 `Child P/N`、`USE`；程式用 Child P/N 前一欄作為成品料號欄 |
| open po | 來源檔需有 `open po` sheet，且表頭需有 `Item`、`Quantity Due`；若有 `料号+厂商`、`Supplier Site`、`Need By Date` 會一併使用 |
| Shortage | 來源檔需有 `Shortage` sheet，且表頭需有 `Part No`；程式固定優先讀取 `QN` 欄作為 `OVER SHORTAGE` / balance 起始值，並將 `HLD`、`BOR MM`、`PO_REMAIN` 等欄位寫入輔助頁供追溯 |
| 原始 CTB | 可選。若來源檔有 `CTB` sheet，輸出會套用它的表頭、日期欄、列順序與格式；A:M 資料列不直接抄原值，而是依下方欄位規則重建 |
| Demand | `DPS+PP` 成品需求 × `BOM1` USE，依子件料號彙總 |
| ETA | `open po` 的 `Quantity Due` 依 `Need By Date` 放到相同或下一個可用期間欄 |
| Balance | 使用 `Shortage!QN` 作起始 shortage 基準，再逐期加 ETA、扣 Demand / other |

CTB A:M 欄位目前依下列規則輸出：

| 欄位 | 規則 |
|---|---|
| A Category | 保留欄位與 title，資料列空白 |
| B model | 由同一個 `BOM1` child part 的 `Model + USE` 組成，例如 `32C*1/32J*1`；找不到來源則空白 |
| C Part No | CTB row 的主料號 / 子件料號 |
| D key | ETA row 使用 `C + E`，用來對 `open po` 的 A 欄 key |
| E Code | ETA row 取 `open po` 的 `Supplier Site`；找不到來源則空白 |
| F vendor | 只有在 `BOM1` vendor 是乾淨來源時填入；若 vendor 是反查 CTB 的公式或找不到來源，則空白 |
| G / H / I | 保留欄位與 title，資料列空白 |
| J OVER SHORTAGE | Balance row 取 `Shortage` sheet 的 `QN` 欄 |
| K PO Remain | ETA row 加總同 key 的 `open po` Quantity Due；Balance row 依 balance 公式邏輯計算 |
| L total | ETA row 加總該 row 的 ETA 日期區數量；其他 row 依目前計算邏輯留空或重算 |
| M ETA目標 | 保留 row type，例如 `Demand` / `ETA` / `other` / `Balance1` |

## 常用參數

| 參數 | 說明 |
|---|---|
| `--input-dir` / `--out-dir` | 輸入 / 輸出資料夾 |
| `--dps` / `--pp` | 明確指定來源檔 |
| `--skip-dps` / `--skip-pp` / `--skip-dps-pp` / `--skip-ctb` | 略過指定報表 |
| `--include-star-parts` | 保留結尾帶 `*` 的 DPS 料號 |
| `--dps-tail-cutoff` | `auto` / `none` / `YYYY-MM-DD` |
| `--pp-plan` | Plan 篩選值，預設 `Production Input` |
| `--pp-start-week` | `auto` 或週數 |
| `--pp-base-year` | 主年度兩位數，例 `26` |
| `--pp-report-date` | 報表基準日 `YYYY-MM-DD` |
| `--dps-pp-current-week` | `DPS+PP` 目前週：`auto` 或週數；`auto` 時週五提前用下一週 |
| `--raken-dps-pp-weeks` | RAKEN `DPS+PP` 的 DPS 保留週數：`2` / `3` / `4`，會覆寫 INI 與 Windows 選擇視窗 |
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

每次執行都會列印：來源檔、實際使用的 DPS / PP 工作表、表頭列、日期欄數與區間、
被排除的 `*` 料號清單與數量、樞紐快取的更新日期與更新者、Plan 篩選筆數、
期間欄清單、`DPS+PP` cutover 週、輸出列數與合計。
另外會在下列情況示警：

- 單一 DPS / PP 來源檔格式錯誤，已略過並嘗試下一個候選檔或下一種報表
- 某個日期的欄數不是 2（D/N 兩班）
- 表頭未被判定為日期、但底下有數量的欄位（避免靜默漏算）
- 樞紐快取更新日距今超過 45 天（可能是舊快照）
- 找不到選定 PP 工作表上的樞紐版面，改用推導模式

## 版面說明

輸出刻意貼近人工版，標題與料號欄以文字格式寫入，數量欄與 total 以數值格式寫入；
DPS / PP 的 `total` 都由程式先算好後寫入，不再依賴 Excel `=SUM()` 公式。
PP 的 `total` 仍保留在期間欄後方，前面留 3 個空白欄，與人工版一致；
欄名會標示起算欄，例如 `total(WK31起)`。
差別只有一處：PP 的 A/B/C 欄補上了 `Customer` / 選定料號欄 / `Model`
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
