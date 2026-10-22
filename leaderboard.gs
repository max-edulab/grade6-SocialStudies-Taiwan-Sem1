/**
 * 社會任務：排行榜（Google Apps Script）
 * 使用方式：
 *   1. 開一個新的 Google 試算表 → 選單「擴充功能」→「Apps Script」
 *   2. 把這個檔案的內容全部貼上（取代原本的程式碼），按儲存
 *   3. 右上角「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *      執行身分：我　／　誰可以存取：任何人 → 按「部署」，第一次要按「授權存取」
 *   4. 複製「網頁應用程式網址」（https://script.google.com/macros/s/…/exec）
 * 成績會寫在試算表的「board」工作表，可以直接在這裡查看，或刪掉不恰當的名字。
 * 排行榜有快取：直接在試算表刪掉名字後，最晚 6 小時會更新；想馬上更新，執行一次 clearCache 函式即可。
 * 更新程式後：「部署」→「管理部署作業」→ 編輯（鉛筆）→ 版本選「新版本」→ 部署，網址不會變。
 */
const SHEET = 'board';
const KEEP = 20;   // 每種排行回傳前幾名
// 三種排行：分數、累積題數越大越好；全對時間越小越好。也順便擋掉不合理的數字
const RULES = {
  score: { ok: v => v >= 0 && v <= 10000 },
  total: { ok: v => v >= 0 && v <= 1000000 },
  time:  { ok: v => v >= 20 && v <= 36000, lower: true }
};

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(['board', 'playerId', 'name', 'pet', 'value', 'updated']); }
  return sh;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
// 排行榜存在快取裡（最多 6 小時），讀的時候不用每次翻試算表；有人上傳成績就更新快取
const CACHE_KEY = 'top-v1';
function top_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CACHE_KEY);
  if (hit) return JSON.parse(hit);
  const out = topFresh_();
  cache.put(CACHE_KEY, JSON.stringify(out), 21600);
  return out;
}
function topFresh_() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const out = {};
  Object.keys(RULES).forEach(b => {
    const r = RULES[b];
    out[b] = rows.filter(x => x[0] === b)
      .map(x => ({ pid: String(x[1]), name: String(x[2]), pet: String(x[3]), value: Number(x[4]) }))
      .sort((a, c) => r.lower ? a.value - c.value : c.value - a.value)
      .slice(0, KEEP);
  });
  return out;
}
// 讀排行榜
function doGet() {
  return json_(top_());
}
// 上傳成績：{ pid, name, pet, entries:[{ board, value }] }，每個人在每種排行只留最好的一筆
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const pid = String(d.pid || '').slice(0, 40);
    const name = String(d.name || '').replace(/[<>]/g, '').trim().slice(0, 12) || '無名氏';
    const pet = String(d.pet || '🐕').slice(0, 4);
    if (!pid) return json_({ error: 'pid' });
    const sh = sheet_();
    const rows = sh.getDataRange().getValues();
    (d.entries || []).forEach(en => {
      const r = RULES[en.board], v = Number(en.value);
      if (!r || !isFinite(v) || !r.ok(v)) return;
      const i = rows.findIndex((x, k) => k > 0 && x[0] === en.board && String(x[1]) === pid);
      if (i < 0) {
        sh.appendRow([en.board, pid, name, pet, v, new Date()]);
        rows.push([en.board, pid, name, pet, v]);
      } else {
        const old = Number(rows[i][4]);
        const better = r.lower ? v < old : v > old;
        if (better) { sh.getRange(i + 1, 5, 1, 2).setValues([[v, new Date()]]); rows[i][4] = v; }
      }
    });
    // 名字和小夥伴以最新的為準
    rows.forEach((x, k) => {
      if (k > 0 && String(x[1]) === pid && (x[2] !== name || x[3] !== pet)) sh.getRange(k + 1, 3, 1, 2).setValues([[name, pet]]);
    });
    const out = topFresh_();
    CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(out), 21600);
    return json_(out);
  } finally {
    lock.releaseLock();
  }
}
// 手動清除快取（在試算表刪掉不恰當的名字後執行一次）
function clearCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}
