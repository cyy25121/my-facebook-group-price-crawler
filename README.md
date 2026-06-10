# Facebook 社團遊戲片價格爬蟲

爬取指定 Facebook 社團一個月內的貼文,解析 PS4/PS5(及 Switch)遊戲片售價,
追蹤同一遊戲的行情變化,並保存每篇貼文當下的內容快照(可偵測後續編輯改價)。

## 架構

| 檔案 | 用途 |
|---|---|
| `crawler.py` | Playwright 開啟獨立 profile 的 Chrome,攔截 Facebook GraphQL 回應,取得貼文 post_id / 精確發文時間 / 全文,append 到 `data/posts.jsonl` |
| `analyze.py` | 解析快照 → 抽出 (遊戲, 價格) → 輸出 `data/listings.csv`、`data/report.xlsx` |
| `games.py` | 遊戲名稱別名字典(可自行擴充)+ 非價格脈絡關鍵字 |
| `config.json` | 社團清單、爬取天數(預設 30)、捲動上限 |

## 使用

```bash
# 第一次:建環境(已完成則略過)
python3 -m venv .venv && .venv/bin/pip install playwright openpyxl && .venv/bin/playwright install chromium

# 爬取(第一次會開瀏覽器視窗,請手動登入 Facebook;登入狀態存在 profile/)
.venv/bin/python crawler.py            # 全部社團
.venv/bin/python crawler.py --group buyswitchandps --days 14

# 分析
.venv/bin/python analyze.py
```

## 資料設計

- `data/posts.jsonl`:append-only。每次執行 crawler 都寫入「當下看到的貼文內容」快照
  (含 `scraped_at` 批次時間)。同一篇貼文若被作者編輯,下次爬取會留下新快照,
  `analyze.py` 會比對各批次的價格差異,輸出「編輯改價」清單。
- `data/listings.csv`:每筆解析出的售價(日期、社團、遊戲、平台、價格、貼文連結)。
- `data/report.xlsx`:Excel 報告,六張工作表 —
  **總覽**(整體數據)、**行情統計**(遊戲 × 平台:筆數/最低/中位/最高/最新,全部為 Excel 公式,
  改「明細」資料會自動重算)、**週走勢**(各遊戲每週中位價)、**明細**(全部售價,含貼文超連結)、
  **編輯改價**(跨批次比對)、**未配對**(供擴充 games.py 的貼文清單)。

## 解析規則摘要(analyze.py)

- 平台偵測:行內 token > 貼文【平台】欄位 > 全文,支援 PS5/PS4/PS3/Switch/Switch 2(NS2)/PSV/Xbox。
- 噪音排除:運費/賣貨便/福袋/攜帶包/主機/同捆/特仕機/合售等行不參與配對;>6000 視為主機價。
- 清單格式:支援「1. 品名」「• 101、品名」項次、品名與價格分行(往後看 1~2 行)。
- 系列名+副標題(「刺客教條:幻象」「戰神諸神黃昏」)只記特定作品,不重複記系列。
- 同作者同日同遊戲同價的跨團轉貼只計一筆。

## 注意

- 請僅以自己的帳號、合理頻率使用;爬蟲行為可能違反 Facebook 服務條款,帳號有被限制的風險。
- 社團貼文多為二手交易,價格解析採啟發式規則,異常值請以 `url` 回到原文確認。
