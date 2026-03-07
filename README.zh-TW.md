# xlex

<p align="center">
  <img src="logo.png" alt="xlex logo" width="640">
</p>

<p align="center">
  <strong>專為 AI Agent 設計的 Excel CLI 工具 — 讓 Copilot、Cursor、Claude 等編碼代理讀寫與操作 Excel 檔案。</strong>
</p>

<p align="center">
  <a href="https://github.com/yen0304/xlex/actions/workflows/ci.yml"><img src="https://github.com/yen0304/xlex/actions/workflows/ci.yml/badge.svg" alt="CI"></a>
  <a href="https://codecov.io/gh/yen0304/xlex"><img src="https://codecov.io/gh/yen0304/xlex/graph/badge.svg" alt="codecov"></a>
  <a href="https://www.npmjs.com/package/xlex"><img src="https://img.shields.io/npm/v/xlex.svg" alt="npm"></a>
  <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT"></a>
  <a href="https://blog.rust-lang.org/2024/07/25/Rust-1.80.0.html"><img src="https://img.shields.io/badge/MSRV-1.80-blue.svg" alt="MSRV"></a>
</p>

<p align="center">
  <a href="README.md">English</a> | 繁體中文
</p>

## 為什麼需要 xlex？

AI 編碼代理（Copilot、Cursor、Claude Code 等）可以執行 CLI 指令，但無法直接開啟 Excel 檔案。xlex 填補了這個缺口 — agent 使用簡單的 CLI 指令即可讀取、寫入、設定樣式與轉換 `.xlsx` 檔案，無需任何 SDK 或函式庫整合。

## 功能特色

- **Agent 友善**：結構化 JSON 輸出、明確的結束代碼、支援模擬執行（dry-run）
- **內建 Skill 文件**：附帶隨時可用的 [agent skill 文件](docs/skills/xlex-agent/)，讓 agent 了解所有指令
- **工作階段管理**：類似 Git 的 `open → batch → commit` 工作流程，適合多步驟編輯
- **批次寫入**：行程內批次執行 — 單次開啟/儲存循環，專為 AI agent 設計
- **串流架構**：處理高達 200MB 的檔案而不會耗盡記憶體
- **多種輸出格式**：Text、JSON、CSV、NDJSON
- **模板系統**：支援 `{{placeholder}}` 語法的變數替換
- **匯入/匯出**：支援 CSV、JSON、YAML、TSV、Markdown
- **跨平台**：macOS、Linux、Windows — 單一執行檔，零相依

## 安裝

### Shell 腳本 (Linux/macOS)

```bash
curl -fsSL https://raw.githubusercontent.com/yen0304/xlex/main/install.sh | bash
```

### npm

```bash
# 全域安裝
npm install -g xlex

# 或直接用 npx（無需安裝）
npx xlex info report.xlsx
```

### 從原始碼編譯

```bash
git clone https://github.com/yen0304/xlex.git
cd xlex
cargo build --release
# 二進制檔位於 target/release/xlex
```

### 下載二進制檔

從 [Releases 頁面](https://github.com/yen0304/xlex/releases) 下載預編譯的二進制檔。

## 快速開始

```bash
# 顯示活頁簿資訊
xlex info report.xlsx

# 取得儲存格值
xlex cell get report.xlsx Sheet1 A1

# 設定儲存格值
xlex cell set report.xlsx Sheet1 A1 "Hello, World!"

# 匯出為 CSV
xlex export csv report.xlsx -s Sheet1 > data.csv

# 從 JSON 匯入
xlex import json data.json output.xlsx

# 模板處理
xlex template apply template.xlsx output.xlsx -D name="John" -D date="2026-01-15"

# 互動模式
xlex interactive
```

## 工作階段管理

xlex 提供**類似 Git 的工作流程**進行多步驟編輯：`open → batch → commit`。這是 AI agent 推薦的使用方式 — 不需要互動式 REPL。

```bash
# 開啟檔案進行編輯（建立工作階段）
xlex open report.xlsx

# 對工作階段套用批次指令
xlex batch -c 'cell set Sheet1 A1 "Hello"' -c 'cell set Sheet1 B1 42'

# 檢查工作階段狀態
xlex status

# 將變更儲存回原始檔案
xlex commit

# 或捨棄變更
xlex close
```

### 批次寫入

`batch` 指令在**單次開啟/儲存循環**中執行多個寫入操作 — 不產生子程序、不重複檔案 I/O：

```bash
# 使用 -c 行內指令
xlex batch report.xlsx -c 'cell set Sheet1 A1 "標題"' -c 'row append Sheet1 a b c'

# 從腳本檔案執行
xlex batch report.xlsx -s commands.txt

# 從 stdin 管線輸入
echo 'cell set Sheet1 A1 "Hello"' | xlex batch report.xlsx

# 搭配使用中的工作階段（不需要檔案參數）
xlex open report.xlsx
xlex batch -c 'cell set Sheet1 A1 "Hello"' -c 'sheet add NewSheet'
xlex commit
```

**支援的批次指令：**
- `cell set <sheet> <ref> <value>` — 設定儲存格值
- `cell clear <sheet> <ref>` — 清除儲存格
- `cell formula <sheet> <ref> <formula>` — 設定公式
- `row append <sheet> <values...>` — 附加列
- `row insert <sheet> <row>` — 插入空白列
- `row delete <sheet> <row>` — 刪除列
- `sheet add <name>` — 新增工作表
- `sheet remove <name>` — 移除工作表
- `sheet rename <old> <new>` — 重新命名工作表

### 互動式 REPL

用於互動式探索大型檔案（唯讀）：

```bash
# 啟動 REPL 工作階段
xlex repl report.xlsx

# 在 REPL 模式中：
session> help      # 顯示可用指令
session> info      # 顯示活頁簿資訊
session> sheets    # 列出所有工作表
session> cell Sheet1 A1        # 取得儲存格值
session> cell Sheet1 B2:D5     # 取得範圍值
session> row Sheet1 1          # 取得列資料
session> exit      # 退出 REPL
```

**優勢：**
- 檔案只在工作階段啟動時載入一次
- 後續指令即時執行
- 適合互動式探索大型活頁簿
- 支援 JSON 輸出 `--format json`

## AI Agent 整合

xlex 內建 **agent skill 文件**，教導 AI 編碼代理完整的指令集。將它們放入你的專案，agent 就能立即操作 Excel 檔案。

```
docs/skills/xlex-agent/
├── SKILL.md                    # 核心概覽 — 從這裡開始
└── references/
    ├── commands.md             # 完整 CLI 指令參考
    └── examples.md             # 實際工作流程範例
```

**相容 agent**：GitHub Copilot、Cursor、Claude Code、Windsurf，或任何支援 skill/instruction 文件的 agent。

**使用方式**：將 `docs/skills/xlex-agent/` 複製到你的專案中。Agent 在處理 Excel 任務時會自動發現並遵循 skill 文件。

**Agent 能用 xlex 做什麼**：
- 讀寫儲存格、列、欄、範圍
- 建立活頁簿、管理工作表
- 套用樣式、格式設定、條件式規則
- 匯入/匯出 CSV、JSON、YAML、Markdown
- 使用變數替換處理模板
- 執行公式與計算

無需 MCP 伺服器、無需 SDK、無需執行環境 — 只需 agent 透過終端機呼叫的 CLI 指令。

## 指令參考

### 活頁簿操作

```bash
xlex info <file>              # 顯示活頁簿資訊
xlex validate <file>          # 驗證活頁簿結構
xlex create <file> [sheets]   # 建立新活頁簿
xlex clone <src> <dest>       # 複製活頁簿
xlex stats <file>             # 顯示統計資訊
xlex props <file> [key]       # 取得/設定屬性
```

### 工作表操作

```bash
xlex sheet list <file>                    # 列出所有工作表
xlex sheet add <file> <name>              # 新增工作表
xlex sheet remove <file> <name>           # 移除工作表
xlex sheet rename <file> <old> <new>      # 重新命名工作表
xlex sheet copy <file> <src> <dest>       # 複製工作表
xlex sheet move <file> <name> <pos>       # 移動工作表到指定位置
xlex sheet hide <file> <name>             # 隱藏工作表
xlex sheet unhide <file> <name>           # 取消隱藏工作表
xlex sheet info <file> <name>             # 顯示工作表資訊
xlex sheet active <file> [name]           # 取得/設定使用中的工作表
```

### 儲存格操作

```bash
xlex cell get <file> <sheet> <ref>            # 取得儲存格值
xlex cell set <file> <sheet> <ref> <value>    # 設定儲存格值
xlex cell formula <file> <sheet> <ref> <formula>  # 設定公式
xlex cell clear <file> <sheet> <ref>          # 清除儲存格
xlex cell type <file> <sheet> <ref>           # 取得儲存格類型
xlex cell batch <file>                        # 從 stdin 批次操作
xlex cell comment get <file> <sheet> <ref>    # 取得儲存格註解
xlex cell comment set <file> <sheet> <ref> <text>  # 設定註解
xlex cell link get <file> <sheet> <ref>       # 取得超連結
xlex cell link set <file> <sheet> <ref> <url> # 設定超連結
```

### 列操作

```bash
xlex row get <file> <sheet> <row>                 # 取得列資料
xlex row append <file> <sheet> <values...>        # 附加一列
xlex row insert <file> <sheet> <row>              # 插入列
xlex row delete <file> <sheet> <row>              # 刪除列
xlex row copy <file> <sheet> <src> <dest>         # 複製列
xlex row move <file> <sheet> <src> <dest>         # 移動列
xlex row height <file> <sheet> <row> [height]     # 取得/設定高度
xlex row hide <file> <sheet> <row>                # 隱藏列
xlex row unhide <file> <sheet> <row>              # 取消隱藏列
xlex row find <file> <sheet> <pattern>            # 搜尋列
```

### 欄操作

```bash
xlex column get <file> <sheet> <col>              # 取得欄資料
xlex column insert <file> <sheet> <col>           # 插入欄
xlex column delete <file> <sheet> <col>           # 刪除欄
xlex column copy <file> <sheet> <src> <dest>      # 複製欄
xlex column move <file> <sheet> <src> <dest>      # 移動欄
xlex column width <file> <sheet> <col> [width]    # 取得/設定寬度
xlex column hide <file> <sheet> <col>             # 隱藏欄
xlex column unhide <file> <sheet> <col>           # 取消隱藏欄
xlex column header <file> <sheet> <col>           # 取得欄標題
xlex column find <file> <sheet> <pattern>         # 搜尋欄
xlex column stats <file> <sheet> <col>            # 欄統計資訊
```

### 範圍操作

```bash
xlex range get <file> <sheet> <range>             # 取得範圍資料
xlex range copy <file> <sheet> <src> <dest>       # 複製範圍
xlex range move <file> <sheet> <src> <dest>       # 移動範圍
xlex range clear <file> <sheet> <range>           # 清除範圍
xlex range fill <file> <sheet> <range> <value>    # 填充範圍
xlex range merge <file> <sheet> <range>           # 合併儲存格
xlex range unmerge <file> <sheet> <range>         # 取消合併儲存格
xlex range style <file> <sheet> <range> [opts]    # 套用樣式
xlex range border <file> <sheet> <range> [opts]   # 套用框線
xlex range name <file> <name> <range>             # 定義命名範圍
xlex range names <file>                           # 列出命名範圍
xlex range validate <file> <sheet> <range> <rule> # 驗證資料
xlex range sort <file> <sheet> <range> [opts]     # 排序範圍
```

### 匯入/匯出

```bash
# 匯出
xlex export csv <file> [-s sheet]             # 匯出為 CSV
xlex export tsv <file> [-s sheet]             # 匯出為 TSV
xlex export json <file> [-s sheet] [--header] # 匯出為 JSON
xlex export markdown <file> [-s sheet]        # 匯出為 Markdown
xlex export yaml <file> [-s sheet]            # 匯出為 YAML
xlex export ndjson <file> [-s sheet]          # 匯出為 NDJSON
xlex export meta <file>                       # 匯出中繼資料

# 匯入
xlex import csv <source> <dest>               # 匯入 CSV
xlex import tsv <source> <dest>               # 匯入 TSV
xlex import json <source> <dest>              # 匯入 JSON
xlex import ndjson <source> <dest>            # 匯入 NDJSON

# 轉換
xlex convert <source> <dest>                  # 自動偵測格式
```

### 公式操作

```bash
xlex formula get <file> <sheet> <cell>            # 取得公式
xlex formula set <file> <sheet> <cell> <formula>  # 設定公式
xlex formula list <file> <sheet>                  # 列出所有公式
xlex formula eval <file> <sheet> <formula>        # 計算公式
xlex formula check <file>                         # 檢查錯誤
xlex formula validate <formula>                   # 驗證語法
xlex formula stats <file>                         # 公式統計
xlex formula refs <file> <sheet> <cell>           # 顯示參照
xlex formula replace <file> <sheet> <find> <replace>  # 替換參照
xlex formula circular <file>                      # 偵測循環參照
xlex formula calc sum <file> <sheet> <range>      # 計算總和
xlex formula calc avg <file> <sheet> <range>      # 計算平均值
xlex formula calc count <file> <sheet> <range>    # 計算數量
xlex formula calc min <file> <sheet> <range>      # 取得最小值
xlex formula calc max <file> <sheet> <range>      # 取得最大值
```

### 模板操作

```bash
xlex template apply <template> <output> -D key=value  # 套用模板
xlex template init <output>                           # 建立新模板
xlex template list <template>                         # 列出佔位符
xlex template validate <template> --vars vars.json    # 驗證模板
xlex template create <source> <output>                # 從既有檔案建立
xlex template preview <template> --vars vars.json     # 預覽渲染結果
```

### 樣式操作

```bash
xlex style list <file>                            # 列出所有樣式
xlex style get <file> <id>                        # 取得樣式詳情
xlex style apply <file> <sheet> <range> <id>      # 套用樣式
xlex style copy <file> <sheet> <src> <dest>       # 複製樣式
xlex style clear <file> <sheet> <range>           # 清除樣式
xlex style condition <file> <sheet> <range> [opts]  # 條件式格式設定
xlex style freeze <file> <sheet> [opts]           # 凍結窗格
xlex style preset list                            # 列出預設樣式
xlex style preset apply <file> <sheet> <range> <preset>  # 套用預設樣式
```

## 輸出格式

使用 `-f` 或 `--format` 指定輸出格式：

```bash
xlex info report.xlsx -f json    # JSON 輸出
xlex info report.xlsx -f csv     # CSV 輸出
xlex info report.xlsx -f text    # 文字輸出（預設）
```

## 全域選項

```
-q, --quiet        僅顯示錯誤訊息
-v, --verbose      啟用詳細輸出
-f, --format       輸出格式（text、json、csv、ndjson）
    --no-color     停用彩色輸出
    --color        強制彩色輸出
    --json-errors  以 JSON 格式輸出錯誤
    --dry-run      模擬執行，不實際變更
-o, --output       將輸出寫入檔案
```

## 結束代碼

| 代碼 | 說明 |
|------|------|
| 0    | 成功 |
| 1    | 一般錯誤 |
| 2    | 無效參數 |
| 3    | 找不到檔案 |
| 4    | 權限不足 |
| 5    | 無效檔案格式 |
| 6    | 找不到工作表 |
| 7    | 儲存格參照錯誤 |

## 工作階段 & 批次指令

```bash
xlex open <file>                  # 開啟檔案進行編輯（建立工作階段）
xlex commit                       # 將工作階段變更儲存回原始檔案
xlex close                        # 捨棄工作階段變更並關閉
xlex status                       # 顯示目前工作階段狀態
xlex batch [file] -c <cmd>        # 執行行內批次指令
xlex batch [file] -s <script>     # 從腳本檔案執行批次指令
xlex repl <file>                  # 啟動互動式 REPL（唯讀）
```

## 工具指令

```bash
xlex completion <shell>           # 產生 shell 自動補全（bash、zsh、fish、powershell）
xlex config show                  # 顯示設定
xlex config get <key>             # 取得設定值
xlex config set <key> <value>     # 設定值
xlex alias list                   # 列出指令別名
xlex alias add <name> <command>   # 新增別名
xlex examples [command]           # 顯示指令範例
xlex man                          # 產生 man page
xlex version                      # 顯示版本資訊
```

## 函式庫使用

```rust
use xlex_core::{Workbook, CellRef, CellValue};

// 開啟活頁簿
let mut workbook = Workbook::open("report.xlsx")?;

// 取得儲存格值
let value = workbook.get_cell("Sheet1", &CellRef::parse("A1")?)?;
println!("A1: {}", value);

// 設定儲存格值
workbook.set_cell("Sheet1", CellRef::parse("B1")?, CellValue::Number(42.0))?;

// 儲存變更
workbook.save()?;
```

## 貢獻

請參閱 [CONTRIBUTING.md](CONTRIBUTING.md) 了解貢獻指南。

## 授權條款

MIT 授權條款 - 詳見 [LICENSE](LICENSE)。
