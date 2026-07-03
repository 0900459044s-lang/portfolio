/**
 * ════════════════════════════════════════════════════════════
 *  台股報價更新 — 盤後版（修正版）
 *  每天 14:45 後自動執行一次，抓取當日收盤價
 *
 *  修正內容：
 *  1. 月初抓不到昨收 → 自動回頭抓上個月最後一個交易日
 *  2. 時間觸發器執行時 getUi() 拋錯 → 改用 safeAlert，
 *     觸發器模式下只寫 Log 不彈窗
 * ════════════════════════════════════════════════════════════
 */

// ── 選單 ──
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 股票工具")
    .addItem("更新今日收盤價", "updateClosePrices")
    .addItem("檢查 user_data 內容", "debugUserData")
    .addItem("設定每日自動更新", "setupDailyTrigger")
    .addItem("回補歷史市值（一次性）", "backfillHistory")
    .addToUi();
}

// ── 安全彈窗：由觸發器執行時沒有 UI，改寫 Log ──
function safeAlert(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log("[自動執行] " + msg);
  }
}

// ── 中文名稱對照表 ──
var NAME_MAP = {
  "0050":"元大台灣50","0056":"元大高股息","006208":"富邦台50",
  "009812":"元大龍頭正2","2201":"裕隆","2308":"台達電",
  "2327":"國巨","2330":"台積電","2383":"台光電","2408":"南亞科",
  "2449":"京元電子","2454":"聯發科","2472":"立隆電","3026":"禾伸堂",
  "3037":"欣興","3042":"晶技","3481":"群創","6257":"詮欣","6442":"博錸",
  "3163":"波若威","3260":"威剛","5274":"信驊","5314":"世紀",
  "5425":"台半","6187":"萬潤","6640":"均華","6664":"群翊",
  "8027":"鈦昇科","8064":"東捷"
};
function getChineseName(sym) {
  var code = sym.split(":")[1];
  return NAME_MAP[code] || sym;
}

// ── 從 user_data 讀取持股 ──
function getHoldingsFromUserData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("user_data");
  if (!sheet) return [];
  var raw = sheet.getRange("B2").getValue();
  if (!raw) return [];
  var data;
  try { data = JSON.parse(raw); } catch(e) { return []; }
  var trades = data.trades || [];
  if (!trades.length) return [];
  var qtyMap = {};
  trades.forEach(function(t) {
    var sym = t[1], action = t[2], qty = Number(t[3]) || 0;
    if (!qtyMap[sym]) qtyMap[sym] = 0;
    if (action === "買進") qtyMap[sym] += qty;
    else if (action === "賣出") qtyMap[sym] -= qty;
  });
  var holdings = [];
  for (var sym in qtyMap) {
    if (qtyMap[sym] > 0) {
      holdings.push({
        sym: sym.split(":")[1],
        market: sym.startsWith("TWO:") ? "otc" : "tse",
        key: sym,
        qty: qtyMap[sym]
      });
    }
  }
  return holdings;
}

// ── 讀取全部交易記錄（回補歷史用）──
function getTradesFromUserData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("user_data");
  if (!sheet) return [];
  var raw = sheet.getRange("B2").getValue();
  if (!raw) return [];
  try { return (JSON.parse(raw).trades) || []; } catch(e) { return []; }
}

function debugUserData() {
  var holdings = getHoldingsFromUserData();
  var msg = "持股清單（" + holdings.length + " 檔）：\n\n" + holdings.map(function(h){ return h.key; }).join("\n");
  if (!holdings.length) msg = "⚠️ 找不到持股！請確認 user_data 分頁 B2 有資料。";
  safeAlert(msg);
}

function getDateInfo() {
  var tw = new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Taipei" }));
  var y = tw.getFullYear(), m = tw.getMonth() + 1, d = tw.getDate();
  return {
    yyyymmdd: "" + y + pad(m) + pad(d),
    rocSlash: (y - 1911) + "/" + pad(m),
    twDate: tw
  };
}
function pad(n) { return n < 10 ? "0" + n : "" + n; }

// ── 取得上個月的日期資訊（給月初補昨收用）──
function getPrevMonthInfo(di) {
  var y = di.twDate.getFullYear(), m = di.twDate.getMonth(); // getMonth() 已是 0-based，直接用就是上個月
  if (m === 0) { y -= 1; m = 12; } else { /* m 保持 1~11 */ }
  // 用上個月的最後一天當查詢日期
  var lastDay = new Date(y, m, 0).getDate();
  return {
    yyyymmdd: "" + y + pad(m) + pad(lastDay),
    rocSlash: (y - 1911) + "/" + pad(m)
  };
}

// ── 抓上市收盤價（TWSE）──
function fetchTseClose(sym, di) {
  try {
    var url = "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=" + di.yyyymmdd + "&stockNo=" + sym + "&response=json";
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var data = JSON.parse(res.getContentText());
    if (data.stat === "OK" && data.data && data.data.length > 0) {
      var last  = data.data[data.data.length - 1];
      var close = parseFloat(last[6].replace(/,/g, ""));
      var prev  = null;
      if (data.data.length > 1) {
        prev = parseFloat(data.data[data.data.length - 2][6].replace(/,/g, ""));
      } else {
        // 月初：本月只有一筆資料，回頭抓上個月最後一個交易日
        var pm = getPrevMonthInfo(di);
        var url2 = "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=" + pm.yyyymmdd + "&stockNo=" + sym + "&response=json";
        var res2 = UrlFetchApp.fetch(url2, { muteHttpExceptions: true });
        var data2 = JSON.parse(res2.getContentText());
        if (data2.stat === "OK" && data2.data && data2.data.length > 0) {
          prev = parseFloat(data2.data[data2.data.length - 1][6].replace(/,/g, ""));
        }
      }
      if (close > 0) return { price: close, prev: prev };
    }
  } catch(e) { Logger.log("TSE close error " + sym + ": " + e.message); }
  return { price: null, prev: null };
}

// ── 抓上櫃收盤價（TPEx 新版 API）──
function otcMonthRows(sym, dateStr) {
  var url = "https://www.tpex.org.tw/www/zh-tw/afterTrading/tradingStock?code=" + sym + "&date=" + dateStr + "&response=json";
  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { "User-Agent": "Mozilla/5.0" } });
  var data = JSON.parse(res.getContentText());
  if (data.tables && data.tables[0] && data.tables[0].data) return data.tables[0].data;
  return [];
}

function fetchOtcClose(sym, di) {
  try {
    var y = di.twDate.getFullYear(), m = di.twDate.getMonth() + 1;
    var rows = otcMonthRows(sym, y + "/" + pad(m) + "/01");
    if (rows.length > 0) {
      var close = parseFloat(rows[rows.length - 1][6].replace(/,/g, ""));
      var prev  = null;
      if (rows.length > 1) {
        prev = parseFloat(rows[rows.length - 2][6].replace(/,/g, ""));
      } else {
        // 月初：回頭抓上個月最後一個交易日
        var py = m === 1 ? y - 1 : y, pm2 = m === 1 ? 12 : m - 1;
        var rows2 = otcMonthRows(sym, py + "/" + pad(pm2) + "/01");
        if (rows2.length > 0) prev = parseFloat(rows2[rows2.length - 1][6].replace(/,/g, ""));
      }
      if (close > 0) return { price: close, prev: prev };
    }
  } catch(e) { Logger.log("OTC close error " + sym + ": " + e.message); }
  return { price: null, prev: null };
}

// ── 主要更新函數 ──
function updateClosePrices() {
  var startTime = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("prices") || ss.insertSheet("prices");
  var holdings = getHoldingsFromUserData();

  if (!holdings.length) {
    safeAlert("⚠️ 找不到持股資料！\n請先在網頁新增交易記錄並同步到雲端。");
    return;
  }

  // 確認現在是否已過收盤（13:30 後）— 只在手動執行時詢問
  var di = getDateInfo();
  var hour = di.twDate.getHours();
  var min  = di.twDate.getMinutes();
  var isAfterClose = hour > 13 || (hour === 13 && min >= 30);
  var isWeekend = di.twDate.getDay() === 0 || di.twDate.getDay() === 6;

  if (!isAfterClose && !isWeekend) {
    try {
      var ui = SpreadsheetApp.getUi();
      var resp = ui.alert(
        "⚠️ 目前尚未收盤",
        "現在是 " + hour + ":" + pad(min) + "，台股尚未收盤（13:30）。\n收盤前資料可能不完整。\n\n要繼續更新嗎？",
        ui.ButtonSet.YES_NO
      );
      if (resp !== ui.Button.YES) return;
    } catch (e) {
      // 觸發器模式下不彈窗，直接繼續（觸發器排在 14:45，正常不會走到這裡）
      Logger.log("[自動執行] 未收盤時段觸發，直接繼續更新");
    }
  }

  // 讀取舊資料做 fallback
  var oldData = {};
  var existing = sheet.getDataRange().getValues();
  for (var oi = 1; oi < existing.length; oi++) {
    var oKey = existing[oi][0];
    if (oKey) oldData[oKey] = { price: parseFloat(existing[oi][2])||null, prev: parseFloat(existing[oi][3])||null };
  }

  sheet.getRange("A1:H1").setValues([["代號","名稱","收盤價","昨收","漲跌","漲幅%","更新日期","交易日"]]);
  sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#1e2330").setFontColor("#e3b341");

  var results = {};
  var failedSyms = [], staleSyms = [];

  // 逐檔抓收盤價（盤後 API 穩定，不需要重試）
  holdings.forEach(function(h) {
    var r = h.market === "tse"
      ? fetchTseClose(h.sym, di)
      : fetchOtcClose(h.sym, di);

    if (r.price !== null) {
      // 昨收抓不到時，沿用舊表的昨收（總比空白好）
      if (r.prev === null && oldData[h.key] && oldData[h.key].prev) {
        r.prev = oldData[h.key].prev;
      }
      results[h.key] = r;
    } else if (oldData[h.key] && oldData[h.key].price) {
      // fallback 用上一次的資料
      results[h.key] = oldData[h.key];
      staleSyms.push(h.key);
    } else {
      results[h.key] = { price: null, prev: null };
      failedSyms.push(h.key);
    }
  });

  var now = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");
  var dateOnly = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd");

  var rows = holdings.map(function(h) {
    var r = results[h.key] || {};
    var price = r.price != null ? r.price : "";
    var prev  = r.prev  != null ? r.prev  : "";
    var diff  = (price !== "" && prev !== "") ? Math.round((price - prev) * 100) / 100 : "";
    var pct   = (price !== "" && prev !== "" && prev > 0) ? Math.round((price - prev) / prev * 10000) / 100 : "";
    return [h.key, getChineseName(h.key), price, prev, diff, pct, now, dateOnly];
  });

  sheet.getRange(2, 1, rows.length, 8).setValues(rows);
  var lastRow = sheet.getLastRow();
  if (lastRow > rows.length + 1) {
    sheet.getRange(rows.length + 2, 1, lastRow - rows.length - 1, 8).clearContent();
  }
  rows.forEach(function(row, i) {
    var diff = row[4];
    var bg = diff > 0 ? "#d4edda" : diff < 0 ? "#f8d7da" : "#ffffff";
    sheet.getRange(i + 2, 1, 1, 8).setBackground(bg);
  });

  // 記錄今日投資組合總市值到歷史（方案A：每日自動累積）
  var totalMV = 0;
  holdings.forEach(function(h) {
    var r2 = results[h.key];
    if (r2 && r2.price) totalMV += h.qty * r2.price;
  });
  if (totalMV > 0) recordHistory(dateOnly, Math.round(totalMV));

  var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
  Logger.log("✅ 盤後更新完成，耗時 " + elapsed + " 秒");

  var msg = "✅ 盤後收盤價更新完成！\n耗時 " + elapsed + " 秒\n更新時間：" + now + "\n\n";
  msg += "成功更新：" + (rows.length - failedSyms.length - staleSyms.length) + "/" + rows.length + " 檔";
  if (staleSyms.length > 0) msg += "\n⏳ 沿用昨日資料：" + staleSyms.join("、");
  if (failedSyms.length > 0) msg += "\n⚠️ 完全失敗：" + failedSyms.join("、");

  safeAlert(msg);
}

// ── 設定每日自動更新（14:45 執行）──
function setupDailyTrigger() {
  // 刪除所有舊的觸發器
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  // 設定每天 14:45（台灣時間）執行
  // GAS 使用 UTC，台灣是 UTC+8，所以 14:45 台灣時間 = 06:45 UTC
  ScriptApp.newTrigger("updateClosePrices")
    .timeBased()
    .atHour(6)   // UTC 6 = 台灣 14
    .nearMinute(45)
    .everyDays(1)
    .create();

  safeAlert(
    "✅ 已設定每日盤後自動更新！\n\n" +
    "執行時間：每天台灣時間 14:45（收盤後 15 分鐘）\n\n" +
    "注意：週六日不會有新資料（非交易日），\n" +
    "系統會沿用上一個交易日的收盤價。"
  );

  // 立即執行一次
  updateClosePrices();
}

// ══════════════════════════════════════════════
// 投資組合市值歷史（mv_history 分頁）
// ══════════════════════════════════════════════
var HISTORY_SHEET = "mv_history";

function getHistorySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET);
    sheet.getRange("A1:B1").setValues([["日期","總市值"]]).setFontWeight("bold");
  }
  return sheet;
}

// 同日覆蓋，不同日新增（保持按日期排序）
function recordHistory(dateStr, mv) {
  var sheet = getHistorySheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var d = data[i][0];
    var key = (d instanceof Date) ? Utilities.formatDate(d, "Asia/Taipei", "yyyy/MM/dd") : String(d);
    if (key === dateStr) { sheet.getRange(i + 1, 2).setValue(mv); return; }
  }
  sheet.appendRow([dateStr, mv]);
}

// ── 方案B：一次性回補歷史市值 ──
// 從第一筆交易日回算到今天：每個交易日的持股 × 當日收盤價
function backfillHistory() {
  var startTime = new Date().getTime();
  var trades = getTradesFromUserData();
  if (!trades.length) { safeAlert("⚠️ 沒有交易資料"); return; }
  trades.sort(function(a, b) { return a[0].localeCompare(b[0]); });

  // 1. 收集每檔股票的持有月份區間
  var firstDate = trades[0][0]; // yyyy/MM/dd
  var today = new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Taipei" }));

  // 2. 抓每檔股票、每個月的每日收盤價
  var priceMap = {}; // priceMap[sym][yyyy/MM/dd] = close
  var symsEver = {};
  trades.forEach(function(t) { symsEver[t[1]] = true; });

  var months = [];
  var cur = new Date(firstDate.replace(/\//g, "-") + "T00:00:00");
  cur.setDate(1);
  while (cur <= today) {
    months.push({ y: cur.getFullYear(), m: cur.getMonth() + 1 });
    cur.setMonth(cur.getMonth() + 1);
  }

  var fetchCount = 0;
  for (var symKey in symsEver) {
    var code = symKey.split(":")[1];
    var isOtc = symKey.indexOf("TWO:") === 0;
    priceMap[symKey] = {};
    months.forEach(function(mo) {
      try {
        var rows;
        if (isOtc) {
          rows = otcMonthRows(code, mo.y + "/" + pad(mo.m) + "/01");
          rows.forEach(function(row) {
            // 民國日期 115/07/01 → 2026/07/01
            var p = row[0].split("/");
            var key = (parseInt(p[0], 10) + 1911) + "/" + p[1] + "/" + p[2];
            var c = parseFloat(String(row[6]).replace(/,/g, ""));
            if (c > 0) priceMap[symKey][key] = c;
          });
        } else {
          var url = "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=" + mo.y + pad(mo.m) + "01&stockNo=" + code + "&response=json";
          var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
          var data = JSON.parse(res.getContentText());
          if (data.stat === "OK" && data.data) {
            data.data.forEach(function(row) {
              var p = row[0].split("/");
              var key = (parseInt(p[0], 10) + 1911) + "/" + p[1] + "/" + p[2];
              var c = parseFloat(String(row[6]).replace(/,/g, ""));
              if (c > 0) priceMap[symKey][key] = c;
            });
          }
        }
        fetchCount++;
        Utilities.sleep(120); // 溫和一點，避免被官方 API 擋
      } catch(e) { Logger.log("backfill fetch error " + symKey + " " + mo.y + "/" + mo.m + ": " + e.message); }
    });
  }

  // 3. 取得所有交易日（任何一檔有價的日期）
  var allDates = {};
  for (var s in priceMap) { for (var d in priceMap[s]) allDates[d] = true; }
  var dateList = Object.keys(allDates).sort();

  // 4. 逐日回算持股 × 收盤價
  var ti = 0;
  var qty = {}; // 目前持股
  var lastClose = {}; // 停牌時沿用上一個收盤
  var rows = [];
  dateList.forEach(function(date) {
    // 套用當日（含）之前的所有交易
    while (ti < trades.length && trades[ti][0] <= date) {
      var t = trades[ti];
      var sym = t[1], action = t[2], q = Number(t[3]) || 0;
      if (!qty[sym]) qty[sym] = 0;
      if (action === "買進") qty[sym] += q;
      else if (action === "賣出") qty[sym] -= q;
      ti++;
    }
    var mv = 0;
    for (var sym2 in qty) {
      if (qty[sym2] <= 0) continue;
      var c = (priceMap[sym2] && priceMap[sym2][date]) || lastClose[sym2];
      if (priceMap[sym2] && priceMap[sym2][date]) lastClose[sym2] = priceMap[sym2][date];
      if (c) mv += qty[sym2] * c;
    }
    if (mv > 0) rows.push([date, Math.round(mv)]);
  });

  // 5. 整批寫入（覆蓋整個 mv_history）
  var sheet = getHistorySheet();
  sheet.clearContents();
  sheet.getRange("A1:B1").setValues([["日期","總市值"]]).setFontWeight("bold");
  if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);

  var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
  var msg = "✅ 歷史回補完成！\n" +
    "API 呼叫：" + fetchCount + " 次\n" +
    "回補交易日：" + rows.length + " 天（" + (rows[0] ? rows[0][0] : "-") + " ~ " + (rows.length ? rows[rows.length-1][0] : "-") + "）\n" +
    "耗時 " + elapsed + " 秒";
  Logger.log(msg);
  safeAlert(msg);
}

// ── Web App 同步（保留，讓網頁可以讀寫資料）──
const DATA_SHEET = "user_data";

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify(handleRequest(e)))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  return ContentService
    .createTextOutput(JSON.stringify(handleRequest(e)))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleRequest(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(DATA_SHEET);
    if (!sheet) sheet = ss.insertSheet(DATA_SHEET);

    if (e.parameter && e.parameter.action === 'read') {
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) return { ok: true, data: null };
      var raw = data[1][1];
      return { ok: true, data: raw ? JSON.parse(raw) : null };
    }

    // 投資組合市值歷史（給網頁畫走勢圖用）
    if (e.parameter && e.parameter.action === 'history') {
      var hs = ss.getSheetByName(HISTORY_SHEET);
      if (!hs) return { ok: true, history: [] };
      var hd = hs.getDataRange().getValues();
      var out = [];
      for (var hi = 1; hi < hd.length; hi++) {
        var d = hd[hi][0];
        var key = (d instanceof Date) ? Utilities.formatDate(d, "Asia/Taipei", "yyyy/MM/dd") : String(d);
        var mv = Number(hd[hi][1]) || 0;
        if (key && mv > 0) out.push({ date: key.replace(/\//g, '-'), mv: mv });
      }
      return { ok: true, history: out };
    }

    if (e.postData && e.postData.contents) {
      var payload = JSON.parse(e.postData.contents);
      if (payload.action === 'write' && payload.data) {
        sheet.getRange("A1").setValue("最後更新");
        sheet.getRange("B1").setValue("資料");
        sheet.getRange("A2").setValue(new Date().toLocaleString("zh-TW", { timeZone: "Asia/Taipei" }));
        sheet.getRange("B2").setValue(JSON.stringify(payload.data));
        return { ok: true, message: "已儲存" };
      }
    }
    return { ok: false, message: "未知的請求" };
  } catch(err) {
    return { ok: false, message: err.message };
  }
}
