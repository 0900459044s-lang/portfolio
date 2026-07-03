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
        key: sym
      });
    }
  }
  return holdings;
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
