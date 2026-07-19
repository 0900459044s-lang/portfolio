#!/usr/bin/env osascript -l JavaScript
// ═══════════════════════════════════════════════════════════════════════════
// 核心財務計算 — 回歸測試（無 node 環境，用 macOS 內建 JavaScriptCore 跑）
//
// 跑法：   osascript -l JavaScript tests/core_calc_test.js
//         （在專案根目錄 /Users/chenhuansheng/Desktop/portfolio 下執行）
//
// 原理：直接從 portfolio_1.html「即時抽出」真實函式原始碼來測（零漂移——
//       改了 App 的算式，這裡測到的就是新版），不複製一份邏輯。
//       只測純計算引擎：computeAll(FIFO 損益/持股/融資)、computeXIRR(年化)、
//       renderXIRR(融資來回單現金流)、breakEvenPrice、twTick、limitUp/Down。
//
// 加新測試：在最下面 TESTS 區塊照樣 push 一組 {name, fn} 即可。
// ═══════════════════════════════════════════════════════════════════════════
'use strict';
ObjC.import('Foundation');

function say(s){
  var h=$.NSFileHandle.fileHandleWithStandardOutput;
  h.writeData($.NSString.alloc.initWithString(s+'\n').dataUsingEncoding($.NSUTF8StringEncoding));
}

// ── 讀 HTML 原始碼 ────────────────────────────────────────────────────────
function readFile(p){
  var s=$.NSString.stringWithContentsOfFileEncodingError(p,$.NSUTF8StringEncoding,null);
  if(!s) return null;
  return ObjC.unwrap(s);
}

// 從一段 JS 原始碼中抽出「function NAME(...) {...}」完整原文。
// 用小型掃描器跳過字串/樣板/註解，只在「真正的程式碼」裡數大括號配對，
// 避免被字串裡的 { } 或 // 註解誤導。（此專案的樣板字面內沒有裸大括號，
// 所以把反引號整段當不透明字串處理即可，最穩。）
function extractFn(src, name){
  var re=new RegExp('function\\s+'+name+'\\s*\\(','g');
  var m=re.exec(src);
  if(!m) throw new Error('找不到函式 '+name+'（App 是否改名/刪除了？）');
  var i=src.indexOf('{', m.index);
  if(i<0) throw new Error('函式 '+name+' 找不到起始 {');
  var depth=0, j=i;
  for(; j<src.length; j++){
    var c=src[j];
    if(c==='"'||c==="'"||c==='`'){          // 跳過字串/樣板（含跳脫）
      var quote=c; j++;
      while(j<src.length && src[j]!==quote){ if(src[j]==='\\') j++; j++; }
      continue;
    }
    if(c==='\\'){ j++; continue; }   // 跳脫字元（只會出現在正則字面裡，如 /\//g）— 跳過下一字元，避免 \/ 被當成 // 註解
    if(c==='/'&&src[j+1]==='/'){ while(j<src.length && src[j]!=='\n') j++; continue; }   // 行註解
    if(c==='/'&&src[j+1]==='*'){ j+=2; while(j<src.length && !(src[j]==='*'&&src[j+1]==='/')) j++; j++; continue; } // 區塊註解
    if(c==='{') depth++;
    else if(c==='}'){ depth--; if(depth===0){ return src.slice(m.index, j+1); } }
  }
  throw new Error('函式 '+name+' 大括號未配對');
}

// ── 沙盒：把抽出的函式 eval 進來，並提供它們需要的全域/DOM stub ──────────────
// computeAll 會寫 realizedList/holdings/marginBalance/window.oversellSyms；
// renderXIRR 會讀 document.getElementById、trades、cashTxns、holdings、marginBalance。
var trades=[], cashTxns=[], realizedList=[], holdings={}, marginBalance=0;
var window={};
var _dom={};   // id -> 假元素
var document={ getElementById:function(id){ if(!_dom[id]) _dom[id]={textContent:'',className:''}; return _dom[id]; } };
// console.warn（computeAll 超賣時會呼叫）——測試時收集起來、不吵
var _warns=[];
var console={ warn:function(){ _warns.push(Array.prototype.join.call(arguments,' ')); }, log:function(){} };

var html=readFile('portfolio_1.html') || readFile('../portfolio_1.html');
if(!html) throw new Error('讀不到 portfolio_1.html — 請在專案根目錄執行');

['breakEvenPrice','twTick','limitUp','limitDown','computeAll','computeXIRR','renderXIRR'].forEach(function(name){
  // eslint-disable-next-line no-eval
  (0,eval)(extractFn(html, name));
});

// ── 迷你斷言工具 ──────────────────────────────────────────────────────────
var passed=0, failed=0, fails=[];
function approx(a,b,tol){ return Math.abs(a-b) <= (tol||1e-6); }
function assert(cond, msg){ if(!cond) throw new Error(msg||'斷言失敗'); }
function eq(a,b,msg){ if(a!==b) throw new Error((msg||'')+`（預期 ${b}，實得 ${a}）`); }
function near(a,b,tol,msg){ if(!approx(a,b,tol)) throw new Error((msg||'')+`（預期 ≈${b}，實得 ${a}）`); }

// 注意：不可命名為 run()——osascript 會把頂層 run() 當自動進入點在腳本跑完後再呼叫一次
function test(name, fn){
  // 每個測試前重置沙盒全域，互不污染
  trades=[]; cashTxns=[]; realizedList=[]; holdings={}; marginBalance=0; window={}; _dom={}; _warns=[];
  try{ fn(); passed++; say('  ✅ '+name); }
  catch(e){ failed++; fails.push(name+' → '+e.message); say('  ❌ '+name+' → '+e.message); }
}

// ═══════════════════════════════════════════════════════════════════════════
// TESTS
// ═══════════════════════════════════════════════════════════════════════════
say('── 純函式 ──');

test('breakEvenPrice：均價100 → 損益兩平價含來回手續費+證交稅', function(){
  // 買進手續費0.1425% + 賣出手續費0.1425% + 證交稅0.3% ⇒ 100/(1-0.001425-0.003)
  near(breakEvenPrice(100), 100/(1-0.001425-0.003), 1e-6, '損益兩平價');
});

test('twTick：不同價位的最小跳動單位', function(){
  eq(twTick(9.99), 0.01, '<10');
  eq(twTick(25),   0.05, '10~50');
  eq(twTick(88),   0.1,  '50~100');
  eq(twTick(300),  0.5,  '100~500');
  eq(twTick(800),  1,    '500~1000');
  eq(twTick(1200), 5,    '>=1000');
});

test('limitUp/limitDown：漲跌停對齊 tick', function(){
  // 前收 100 → 漲停 110（tick 0.5，110 已對齊）、跌停 90
  eq(limitUp(100),   110, '漲停');
  eq(limitDown(100), 90,  '跌停');
  assert(limitUp(0)===null && limitDown(0)===null, '前收<=0 應回 null');
});

say('── computeAll（FIFO 引擎）──');

test('FIFO 已實現損益：分批買、一次賣，本金/手續費正確配對', function(){
  // [date, sym, action, qty, price, fee, tax, margin]
  trades=[
    ['2024/01/02','2330','買進',1000,500,712,0,0],
    ['2024/02/02','2330','買進',1000,600,855,0,0],
    ['2024/03/02','2330','賣出',1500,700,1496,3150,0],
  ];
  computeAll();
  eq(realizedList.length, 1, '應有一筆已實現');
  var r=realizedList[0];
  // 賣 1500：吃掉第一批 1000@500 + 第二批 500@600 ⇒ 買進本金 500000+300000=800000
  eq(r.qty, 1500, '賣出股數');
  var gross=1500*700 - (1000*500+500*600);     // 1050000-800000=250000
  near(r.gross, gross, 1e-6, '毛損益');
  near(r.net, gross-1496-3150, 1e-6, '淨損益(扣賣出手續費+稅)');
  // 剩餘持股：第二批剩 500@600
  eq(holdings['2330'].qty, 500, '剩餘股數');
  near(holdings['2330'].avgCost, 600, 1e-6, '剩餘均價');
});

test('超賣防呆：賣出量>持有，超賣部分不計損益且發警告', function(){
  trades=[
    ['2024/01/02','2454','買進',100,1000,143,0,0],
    ['2024/02/02','2454','賣出',300,1100,472,990,0],   // 只有100股有對應買進
  ];
  computeAll();
  eq(realizedList[0].qty, 100, '只認 100 股');
  assert(realizedList[0].oversell===true, '應標記 oversell');
  assert(window.oversellSyms.indexOf('2454')>=0, 'oversellSyms 應含 2454');
  assert(_warns.length>0, '應有超賣警告');
});

test('同日先賣後買：買進排序在賣出之前，不誤判超賣', function(){
  // 資料順序故意把賣出寫在前面，但同日買進應先進 FIFO
  trades=[
    ['2024/01/02','1101','賣出',100,50,36,15,0],
    ['2024/01/02','1101','買進',100,45,32,0,0],
  ];
  computeAll();
  eq(realizedList.length, 1, '應成功配對一筆');
  assert(realizedList[0].oversell!==true, '同日買進在前，不該超賣');
  near(realizedList[0].gross, 100*50-100*45, 1e-6, '毛損益 500');
});

test('融資：marginBalance＝各批未償融資本金加總（股數加權）', function(){
  trades=[
    ['2024/01/02','2603','買進',1000,100,143,0,60000],  // 融資 60000（每股60）
    ['2024/02/02','2603','買進',1000,120,171,0,72000],  // 融資 72000（每股72）
    ['2024/03/02','2603','賣出',1000,130,185,390,0],    // FIFO 賣掉第一批 → 償還 60000
  ];
  computeAll();
  near(marginBalance, 72000, 1e-6, '賣掉第一批後，只剩第二批 72000 融資');
  near(holdings['2603'].margin, 72000, 1e-6, 'holdings.margin 同步');
});

say('── computeXIRR / renderXIRR ──');

test('computeXIRR：-100 → 一年後 +110 ≈ 10%', function(){
  var f=[
    {date:new Date('2024-01-01'), amount:-100},
    {date:new Date('2025-01-01'), amount:110},
  ];
  near(computeXIRR(f), 0.10, 2e-3, 'XIRR 約 10%');
});

test('computeXIRR：無正負相間 → 回 null', function(){
  var f=[
    {date:new Date('2024-01-01'), amount:-100},
    {date:new Date('2025-01-01'), amount:-50},
  ];
  eq(computeXIRR(f), null, '全負無解');
});

test('renderXIRR 融資來回單（回歸）：借款不得被當報酬灌水', function(){
  // 全部平倉 ⇒ holdings 空、equity=0 ⇒ 不 push「現在」的期末流 ⇒ XIRR 只剩固定日期交易流，可精確斷言。
  // 融資 60%：買 1000@100 自備 40000；一年後賣 1000@110，還券商 60000 ⇒ 淨入 50000。
  // 正確：-40000 → +50000 ≈ +25.0%。 若回歸成舊 bug（賣出流沒扣還款）→ 會變 +175%。
  trades=[
    ['2024/01/01','9999','買進',1000,100,0,0,60000],
    ['2025/01/01','9999','賣出',1000,110,0,0,0],
  ];
  cashTxns=[];
  computeAll();                       // 產生 holdings(空) / marginBalance(0)
  eq(Object.keys(holdings).length, 0, '應已全部平倉');
  renderXIRR();
  var txt=_dom['sk-xirr'].textContent;  // 形如 "+25.0%"
  var val=parseFloat(txt);
  assert(!isNaN(val), 'sk-xirr 應為百分比字串，實得「'+txt+'」');
  near(val, 25.0, 0.3, '融資來回單 XIRR 應 ≈+25%，不是把借款當獲利的爆高值（'+txt+'）');
});

// ═══════════════════════════════════════════════════════════════════════════
say('');
say('═══════════════════════════════════════');
say(`結果：${passed} 通過、${failed} 失敗`);
if(failed>0){
  say('失敗清單：');
  fails.forEach(function(f){ say('  • '+f); });
  throw new Error(`有 ${failed} 個測試失敗`);   // 讓 shell 拿到非 0 結束碼
}
say('全部通過 ✅');
