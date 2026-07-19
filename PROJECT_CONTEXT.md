# 專案脈絡與規則摘要 — 我的財富管理 App

> 開新對話時，把這份貼進開頭（或用 `@PROJECT_CONTEXT.md` 引用），AI 就能快速接手。

## 1. 核心架構
- **前端**：單一檔 `portfolio_1.html`（~3700 行，HTML+CSS+JS 全包在一個檔），部署在 GitHub Pages。
- **後端**：Google Apps Script `apps_script_程式碼.gs`（~645 行），負責代抓台股股價 + 當雲端資料庫。
- **線上網址**：https://0900459044s-lang.github.io/portfolio/portfolio_1.html ・repo：`0900459044s-lang/portfolio`
- **資料流**：
  - App 資料存 localStorage，並雙向同步到 Google Sheet 的 `user_data` 分頁（B2 儲存格 = 一整包 JSON），透過 GAS Web App 的 `?action=read` / POST `action=write`。
  - **股價流**：GAS 從 user_data 的 trades 推算持股 → 抓 TWSE(STOCK_DAY)/TPEx(tradingStock) 收盤 → 寫進 Sheet 的 `prices`(報價)、`benchmark`(0050大盤)、`mv_history`(每日總市值) 分頁。App 讀已發佈 CSV(SHEETS_URL)，或呼叫 `?action=history`(市值+大盤)、`?action=refreshprices`(一鍵抓價)。
  - GAS 每天 14:45 觸發器自動更新。

## 2. 技術棧
- 純前端：原生 HTML/CSS/JS + **Chart.js**(CDN)。**無框架、無 build、無 node**。
- 後端：Google Apps Script（doGet/doPost、UrlFetchApp、SpreadsheetApp）。
- **PWA**：manifest.json + sw.js + icon-192/512/180.png（可安裝成手機/桌面 App）。
- Git 認證：token 存 macOS 鑰匙圈（remote 無明文 token）。

## 3. 已完成功能
**核心**：FIFO 已實現損益、融資追蹤(自備款/融資金/現金水位)、XIRR、Modified Dietz 金額加權報酬、財務健康分數、HHI 分散度、現金流預測、帳戶轉帳自動沖銷、JSON 備份還原、淨資產全景(股票+銀行+負債+實體投資)。
**量化分析**(交易分析頁)：大盤比較(0050等值線+贏輸標示)、風險指標(最大回撤/年化波動/Sharpe/日勝率)、回撤水下圖、交易行為分析(勝率/盈虧比/獲利因子/期望值/連勝連敗)、月報酬率熱力表、Beta/Alpha/相關性、個股戰績表、每日損益日曆、年度統計(報稅)、CSV 匯出。
**體驗**：Yahoo 式持股明細(點開→批次/成交/股息三分頁)、股息綁定個股+試算、目標價/停損提醒、自訂股票名稱、報價防火牆(±45%擋壞資料)、新增交易「上次記到哪」提示、個股跳 Yahoo/鉅亨連結、休市日過濾。
**即時性**：盤中即時報價(TWSE MIS API，GAS 代抓，`?action=live`)，盤後自動走收盤價路徑。
**可靠性**：雲端同步衝突偵測(存檔前用 `updatedAt` 比對雲端版本，被其他裝置改過會跳確認)、核心計算回歸測試(`tests/core_calc_test.js`，用 JavaScriptCore 跑)。

## 3.5 2026-07-20 新增（本輪）
- **產業分類 sector**：{代號:產業} 存 localStorage(pf_sectors)+雲端；種子分類 SECTOR_MAP；持倉明細內🏷️下拉可改。分散度面板新增「產業集中度」HHI（抓「個股看似分散、實則單一產業重壓」）。
- **持股地圖 Treemap**（股票頁）：slice-and-dice，列高=產業市值、格寬=個股市值、色=損益，renderSectorTreemap()。
- **除息行事曆 divEvents**：即將除息預估入帳 + 除息後填息率(填息%=(現價−除息參考)/現金股利)。renderDivCalendar()/openDivEventModal()。
- **配股（股票股利）**：交易類別新增「配股」，computeAll 存零成本股(price=0)攤薄均價、賣出時 FIFO 實現。有回歸測試。GAS getHoldingsFromUserData 也已同步認「配股」加股數。
- **融資追繳 Email 預警**：GAS checkMarginCall()——每日 updateClosePrices 後，用最新收盤總市值÷marginBalance 算整戶維持率，<140% 寄 MailApp 信（同日不重複；回到門檻以上重置）。選單「測試融資警示信」可試寄。**需使用者重貼 GAS + 存檔**（觸發器跑編輯器碼，免重新部署 Web App）。

## 4. 接下來的開發目標
- 個股迷你走勢 sparkline / 風險報酬散布圖（**需 GAS 開始逐檔記每日收盤到新分頁**才有歷史，目前只有總市值/0050 歷史）
- 二代健保補充費 & 股利所得課稅試算、券商庫存對帳快照
- 目標價達標 Email 通知(GAS 觸發器)、買賣點疊在個股 K 線
- 股息：使用者決定**從現在起才記**(不補歷史)；融資利息領到時記「費用」現金交易
- GAS 每日快照備份(user_data → backup 分頁，滾動保留 30 天) — **GAS 碼已寫好，待使用者重新部署**
- (討論中)GAS Web App 加金鑰保護資料安全

## 5. 寫 code 時要遵守的規則（重要！）
1. **全程繁體中文**（使用者是台灣人，個人台股投資者、有用融資）。
2. **改 `INIT_TRADES`/`INIT_CASH_TXNS` 常數沒用**——那只是「localStorage 為空時的種子」。真實資料在雲端/localStorage。要改資料：讓使用者在 App UI 操作，或直接 `curl` GAS Web App 讀改寫 user_data(B2 JSON)。
3. **每次改完必驗語法**（這台沒 node）：抽出最大 `<script>` → `osascript -l JavaScript` + `Function(code)` 只解析。
4. **分批 commit + push**，訊息寫清楚，結尾加 `Co-Authored-By: Claude`。GitHub Pages 生效要 1-2 分鐘 + Cmd+Shift+R。
5. **所有報酬/回撤計算一律排除入出金**（Modified Dietz 日報酬：`(mv - prev - flow + div)/(prev + max(0,flow))`），共用 `dailyReturnSeries()`。市值歷史用 `tradingDaysOnly()` 去重濾休市日。**絕不用「市值當報酬」**(賣股換現金不是虧損)。
6. **雲端同步新欄位**：要同時加到 `saveToCloud` 的 payload 和 `loadFromCloud` 的賦值兩邊。
7. **GAS 改動**要使用者手動重貼 + 建立新版本部署（貼碼用 GitHub 網頁的複製鈕，**pbcopy 到不了他的剪貼簿**）。GAS 選單跑的是編輯器碼，網站用的是「已部署版本」，兩者不同。
8. **防呆優先**：NaN/超賣/壞價/日期時區(用 `localDateKey()` 本地日期，勿用 toISOString UTC)。
9. **手機響應式**(≤640px 媒體查詢)，寬表格包 `.tbl-scroll`+min-width，頁面 `overflow-x:clip`。
10. **帶使用者操作要一步一步**，每步給「肉眼檢查點」(如行號、選單多了某項)。使用者不熟技術。

## 6. 關鍵檔案/變數
- `computeAll()` FIFO 引擎(trades→holdings/realizedList，lots 帶買入日期+融資 mps)
- `dailyReturnSeries()` 共用日報酬序列(風險/回撤/日曆共用)
- `renderPerfCard()` 績效卡+大盤比較、`renderRiskCard()`、`renderTxAnalysis()`(交易分析各卡)
- `saveToCloud()`/`loadFromCloud()` 雲端同步、`WEB_APP_URL` 常數
- 雲端同步欄位：trades/banks/debts/bizProjects/bizIncome/cashflows/mvHistory/nwHistory/cashSettings/cashTxns/transfers/customNames/priceTargets/**sectors**/**divEvents**/**marginBalance**/sheetsUrl
  （`marginBalance` 是給 GAS 算融資維持率用；`sectors`={代號:產業}；`divEvents`=除息行事曆陣列）
