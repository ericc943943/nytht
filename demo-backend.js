/* 自動產生，請勿直接修改：示範模式用的後端。
 * 內容是 backend/Code.gs 原封不動，外面包一層「假的 Google 試算表」，資料存在這台電腦的瀏覽器裡。
 * 改了 Code.gs 之後執行 python3 backend/build-demo.py 重新產生。 */
window.DemoBackend = (function () {
  "use strict";
  const KEY = "tsuhui-demo-v2";
  let db = null;
  try { db = JSON.parse(localStorage.getItem(KEY) || "null"); } catch (e) {}
  if (!db || !db.sheets) db = { sheets: {}, order: [] };
  const persist = () => { try { localStorage.setItem(KEY, JSON.stringify(db)); } catch (e) {} };

  /* ---- 假的 SpreadsheetApp ---- */
  function sheetObj(name) {
    const data = () => db.sheets[name];
    const range = (r, c, nr, nc) => {
      const self = {
        getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => {
          const row = data()[r - 1 + i] || []; const v = row[c - 1 + j]; return v === undefined || v === null ? "" : v; })),
        getValue: () => self.getValues()[0][0],
        setValues: (vals) => { vals.forEach((row, i) => { while (data().length < r + i) data().push([]); row.forEach((v, j) => { data()[r - 1 + i][c - 1 + j] = v; }); }); return self; },
        setValue: (v) => self.setValues([[v]]),
        setFontWeight: () => self, setNumberFormat: () => self,
      };
      return self;
    };
    return {
      getName: () => name,
      getLastRow: () => data().length,
      appendRow: (row) => { data().push(row.slice()); },
      setFrozenRows: () => {},
      getRange: (a, b, c, d) => typeof a === "string" ? (a === "A1" ? range(1, 1, 1, 1) : { setNumberFormat() {} }) : range(a, b, c || 1, d || 1),
    };
  }
  const ss = {
    getSheetByName: (n) => db.sheets[n] ? sheetObj(n) : null,
    insertSheet: (n) => { db.sheets[n] = []; db.order.push(n); return sheetObj(n); },
    getSheets: () => db.order.map(sheetObj),
    deleteSheet: (s) => { delete db.sheets[s.getName()]; db.order = db.order.filter(x => x !== s.getName()); },
  };
  const SpreadsheetApp = { getActiveSpreadsheet: () => ss };
  const cacheStore = {};
  const CacheService = { getScriptCache: () => ({
    get: k => (k in cacheStore ? cacheStore[k] : null), put: (k, v) => { cacheStore[k] = v; }, remove: k => { delete cacheStore[k]; } }) };
  const LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
  const Session = { getScriptTimeZone: () => "America/New_York" };
  const Logger = { log() {} };
  const ContentService = { createTextOutput: (t) => ({ text: t, setMimeType() { return this; } }), MimeType: { JSON: "json" } };

  // SHA-256（同步版，跟 Apps Script 的 computeDigest 結果一樣）
  function sha256Bytes(str) {
    const bytes = Array.from(new TextEncoder().encode(str));
    const K = [0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2];
    let H = [0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19];
    const l = bytes.length * 8; bytes.push(0x80); while (bytes.length % 64 !== 56) bytes.push(0);
    for (let i = 7; i >= 0; i--) bytes.push(Math.floor(l / Math.pow(2, i * 8)) & 0xff);
    const rotr = (x, n) => (x >>> n) | (x << (32 - n));
    for (let o = 0; o < bytes.length; o += 64) {
      const w = new Array(64);
      for (let i = 0; i < 16; i++) w[i] = (bytes[o+4*i] << 24) | (bytes[o+4*i+1] << 16) | (bytes[o+4*i+2] << 8) | bytes[o+4*i+3];
      for (let i = 16; i < 64; i++) { const s0 = rotr(w[i-15],7) ^ rotr(w[i-15],18) ^ (w[i-15] >>> 3), s1 = rotr(w[i-2],17) ^ rotr(w[i-2],19) ^ (w[i-2] >>> 10); w[i] = (w[i-16] + s0 + w[i-7] + s1) | 0; }
      let [a,b,c,d,e,f,g,h] = H;
      for (let i = 0; i < 64; i++) {
        const t1 = (h + (rotr(e,6) ^ rotr(e,11) ^ rotr(e,25)) + ((e & f) ^ (~e & g)) + K[i] + w[i]) | 0;
        const t2 = ((rotr(a,2) ^ rotr(a,13) ^ rotr(a,22)) + ((a & b) ^ (a & c) ^ (b & c))) | 0;
        h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
      }
      H = H.map((v, i) => (v + [a,b,c,d,e,f,g,h][i]) | 0);
    }
    const out = []; H.forEach(v => { for (let i = 3; i >= 0; i--) { const x = (v >>> (i * 8)) & 0xff; out.push(x > 127 ? x - 256 : x); } });
    return out;
  }
  const Utilities = {
    DigestAlgorithm: { SHA_256: "sha256" }, Charset: { UTF_8: "utf8" },
    computeDigest: (alg, str) => sha256Bytes(str),
    getUuid: () => (crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2) + Date.now().toString(36)),
    formatDate: (d, tz, fmt) => {
      const p = new Intl.DateTimeFormat("en-US", { timeZone: tz, year: "2-digit", month: "2-digit", day: "2-digit" }).formatToParts(d);
      const g = t => p.find(x => x.type === t).value; return g("year") + g("month") + g("day");
    },
  };

  /* ================= 以下是 backend/Code.gs ================= */
/**
 * 紐約花蓮慈惠堂｜線上點燈和法會報名系統｜後端（Google Apps Script）
 *
 * 用法：在堂裡 Google 帳號建一份空白試算表 →「擴充功能」→「Apps Script」
 *      → 把這整份貼上 → 先執行一次 setup → 部署成「網頁應用程式」。
 * 詳細步驟見 backend/設定說明.txt。
 *
 * 試算表分頁（setup 會自動建立）：
 *   報名  一筆報名一列；「完整資料」欄是 JSON，後台靠它顯示，請勿手動修改
 *   帳號  義工帳號；PIN 只存加密後的值
 *   設定  報名項目開關（A1 的 JSON），由後台「設定」修改
 *
 * 示範模式：網站根目錄的 demo-backend.js 是用這份檔案產生的（在瀏覽器裡跑同一套規則），
 * 改了這份之後請重新產生（見 backend/build-demo.py）。
 */

/* ================= 設定 ================= */
const PRICES = {
  lamp: { "個人光明燈": 120, "太歲燈": 120, "五路財神燈": 120, "文昌燈": 120 },  // 點燈：每盞
  lotus: 60,     // 法會：蓮花福慧燈，每位
  chaodu: 160,   // 法會：超渡牌位，每牌位
};
// 燈位：每種燈牆上的排列。燈號 = 代碼-排排個個，例：A01-0512 ＝ 光明燈第 5 排第 12 個
const SEATS = {
  "個人光明燈": { code: "A01", perRow: 30, rows: 20 },
  "文昌燈":     { code: "A02", perRow: 20, rows: 20 },
  "太歲燈":     { code: "A03", perRow: 30, rows: 19 },
  "五路財神燈": { code: "A05", perRow: 30, rows: 23 },
};
const DEFAULT_SETTINGS = {
  lamp: { status: "open" },
  events: [
    { id: "2027-zhongyuan", name: "瑤池金母聖誕祈福消災超渡暨普渡大法會", date: "", status: "hidden" },
    { id: "2027-spring-lidou", name: "春季祈安禮斗法會", date: "", status: "hidden" },
    { id: "2027-autumn-lidou", name: "秋季祈安禮斗法會", date: "", status: "hidden" },
  ],
};
const FIRST_ADMIN = { name: "管理員", pin: "8888" };   // 第一次登入會被要求改 PIN
const MAX_FAILS = 8;          // PIN 或查詢連續打錯幾次就暫時鎖住
const LOCK_MINUTES = 15;

const SHEET = { reg: "報名", users: "帳號", settings: "設定" };
const REG_HEAD = ["報名編號", "報名時間", "表單", "場次", "聯絡人", "電話", "住址", "應繳", "已繳清", "收款時間", "經手人", "語言", "備註", "收款紀錄", "送出代碼", "完整資料", "已取消", "燈位"];
const C = { ref: 1, time: 2, form: 3, event: 4, name: 5, phone: 6, address: 7, total: 8, paid: 9, paidAt: 10, paidBy: 11, lang: 12, note: 13, log: 14, cid: 15, json: 16, cancelled: 17, seats: 18 };
const USER_HEAD = ["id", "姓名", "身分", "PIN（加密）", "啟用", "待改PIN"];

/* ================= 第一次使用：建立分頁與管理員帳號 ================= */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const make = (name, head) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (head) {   // 標題列：新建或欄位有增加時都重寫
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
      sh.setFrozenRows(1);
    }
    return sh;
  };
  const reg = make(SHEET.reg, REG_HEAD);
  reg.getRange("A:A").setNumberFormat("@");   // 報名編號、電話當文字，避免被改成數字或日期
  reg.getRange("F:F").setNumberFormat("@");
  const users = make(SHEET.users, USER_HEAD);
  if (users.getLastRow() < 2) users.appendRow(["u-admin", FIRST_ADMIN.name, "admin", hashPin_(FIRST_ADMIN.pin), true, true]);
  const st = make(SHEET.settings);
  if (!st.getRange("A1").getValue()) st.getRange("A1").setValue(JSON.stringify(DEFAULT_SETTINGS));
  const def = ss.getSheetByName("工作表1") || ss.getSheetByName("Sheet1");
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  Logger.log("完成。管理員第一次登入 PIN：" + FIRST_ADMIN.pin + "（登入後會被要求改 PIN）");
}

/* ================= 進入點 ================= */
function doGet(e) {
  const p = e.parameter || {};
  try {
    switch (p.action) {
      case "settings": { const s = getSettings_(); s.seatsLeft = seatsLeft_(); return out_(s); }
      case "users":    return out_({ users: readUsers_().map(pub_) });
      case "login": {
        if (isLocked_(p.uid)) return out_({ ok: false, error: "locked" });
        const u = checkLogin_(p.uid, p.pin);
        return out_(u ? { ok: true, user: pub_(u) } : { ok: false, error: "auth" });
      }
      case "list": {
        if (!checkLogin_(p.uid, p.pin)) return out_({ ok: false, error: "auth" });
        return out_({ ok: true, rows: listRows_() });
      }
      case "lookup": return out_(lookup_(p.ref, p.phone));    // 信眾用編號＋電話查自己的報名
      default: return out_({ ok: true, service: "tsuhui-signup" });
    }
  } catch (err) { return out_({ ok: false, error: String(err && err.message || err) }); }
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: "bad-json" }); }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (!body.action) return out_(submit_(body));           // 報名頁送出
    if (body.action === "update") return out_(update_(body)); // 信眾修改自己的報名
    const me = checkLogin_(body.uid, body.pin);
    if (!me) return out_({ ok: false, error: "auth" });
    switch (body.action) {
      case "setPaid":      return out_(setPaid_(me, body.ref, !!body.paid));
      case "setCancelled": return out_(setCancelled_(me, body.ref, !!body.cancelled));
      case "assignSeats":  return out_(assignWaitlist_(me, body.ref));
      case "saveSettings": if (me.role !== "admin") return out_({ ok: false, error: "auth" });
                           return out_(saveSettings_(body.settings));
      case "userOp":       return out_(userOp_(me, body));
      default:             return out_({ ok: false, error: "unknown-action" });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally { lock.releaseLock(); }
}

/* ================= 報名 ================= */
function submit_(d) {
  if (!d || !d.contact || !d.form) return { ok: false, error: "bad-data" };
  const st = getSettings_();
  if (d.form === "lamp" && st.lamp.status !== "open") return { ok: false, error: "closed" };
  if (d.form !== "lamp") { const ev = st.events.find(x => x.id === d.eventId); if (!ev || ev.status !== "open") return { ok: false, error: "closed" }; }

  const sh = sheet_(SHEET.reg);
  const last = sh.getLastRow();
  // 同一份報名重送（網路斷掉後再按一次）：回傳原本的報名，不重複建立
  if (d.clientId && last > 1) {
    const ids = sh.getRange(2, C.cid, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (ids[i][0] === d.clientId) {
      const rec = readRecord_(sh.getRange(i + 2, 1, 1, REG_HEAD.length).getValues()[0]);
      return { ok: true, ref: rec.ref, record: publicRecord_(rec), duplicate: true };
    }
  }
  // 報名編號由這裡決定，確保不重複（格式：L/F + yymmdd + - + 4 碼）
  const existing = new Set(last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : []);
  const tz = Session.getScriptTimeZone() || "America/New_York";
  const prefix = (d.form === "lamp" ? "L" : "F") + Utilities.formatDate(new Date(), tz, "yyMMdd") + "-";
  let ref = "";
  for (let i = 0; !ref && i < 200; i++) { const c = prefix + ("000" + Math.floor(Math.random() * 10000)).slice(-4); if (!existing.has(c)) ref = c; }
  if (!ref) return { ok: false, error: "ref" };

  cleanData_(d);
  d.totals = calcTotals_(d);                     // 金額一律在這裡重新計算
  if (d.form === "lamp") assignSeats_(d, usedSeats_(""), null);
  d.ref = ref; d.paid = false; d.cancelled = false; d.submittedAt = new Date().toISOString();
  const evName = d.form === "lamp" ? "點燈" : ((st.events.find(x => x.id === d.eventId) || {}).name || d.formTitle || "法會");
  sh.appendRow([ref, new Date(), d.form === "lamp" ? "點燈" : "法會", evName, d.contact.name || "", d.contact.phone || "", d.contact.address || "",
                d.totals.all, false, "", "", d.lang === "en" ? "英文" : "中文", d.note || "", "[]", d.clientId || "", JSON.stringify(d), false, seatText_(d)]);
  return { ok: true, ref: ref, record: publicRecord_(d) };
}

// 只留下報名需要的欄位（金額、燈位、繳費狀態一律由後端決定）
function cleanData_(d) {
  delete d.totals; delete d.paid; delete d.paidAt; delete d.paidBy; delete d.payLog; delete d.cancelled; delete d.sample;
  (d.lamps || []).forEach(p => { delete p.seats; if (!p.pid) p.pid = Utilities.getUuid().slice(0, 8); });
}

function calcTotals_(d) {
  if (d.form === "lamp") {
    const byType = {}; let all = 0;
    (d.lamps || []).forEach(p => (p.lampTypes || []).forEach(t => {
      const price = PRICES.lamp[t]; if (price == null) return;
      byType[t] = byType[t] || { name: t, count: 0, amount: 0 };
      byType[t].count++; byType[t].amount += price; all += price;
    }));
    return { byType: byType, all: all };
  }
  const lamp = (d.lamps || []).length * PRICES.lotus;
  const chaodu = (d.chaodu || []).length * PRICES.chaodu;
  const meal = (d.meals || []).reduce((a, m) => a + Math.max(0, Number(m.amount) || 0), 0);
  return { lamp: lamp, chaodu: chaodu, meal: meal, all: lamp + chaodu + meal };
}

/* ================= 燈位 ================= */
const seatId_ = (cfg, r, c) => cfg.code + "-" + ("0" + r).slice(-2) + ("0" + c).slice(-2);

// 目前被佔用的燈位（已取消的不算；exceptRef 是正在修改的那一筆）
function usedSeats_(exceptRef) {
  const used = new Set(), sh = sheet_(SHEET.reg), n = sh.getLastRow() - 1;
  if (n < 1) return used;
  sh.getRange(2, 1, n, REG_HEAD.length).getValues().forEach(r => {
    if (r[C.cancelled - 1] === true || String(r[0]) === String(exceptRef)) return;
    let d = {}; try { d = JSON.parse(r[C.json - 1]); } catch (e) { return; }
    (d.lamps || []).forEach(p => Object.keys(p.seats || {}).forEach(t => { if (p.seats[t]) used.add(p.seats[t]); }));
  });
  return used;
}
function nextFreeSeat_(type, used) {
  const cfg = SEATS[type]; if (!cfg) return null;
  for (let r = 1; r <= cfg.rows; r++) for (let c = 1; c <= cfg.perRow; c++) {
    const id = seatId_(cfg, r, c); if (!used.has(id)) return id;
  }
  return null;   // 額滿：候補
}
// 分配燈位。prev 是修改前的資料：同一個人（pid）同一種燈原本有燈位、且沒被別人佔走，就保留原位
function assignSeats_(d, used, prev) {
  const old = {};
  ((prev && prev.lamps) || []).forEach(p => { if (p.pid) old[p.pid] = p.seats || {}; });
  (d.lamps || []).forEach(p => {
    const keep = old[p.pid] || {}, seats = {};
    (p.lampTypes || []).forEach(t => {
      if (!SEATS[t]) return;
      if (keep[t] && !used.has(keep[t])) seats[t] = keep[t];
      else seats[t] = nextFreeSeat_(t, used);
      if (seats[t]) used.add(seats[t]);
    });
    p.seats = seats;
  });
}
function seatsLeft_() {
  const used = usedSeats_(""), left = {};
  Object.keys(SEATS).forEach(t => {
    const cfg = SEATS[t], prefix = cfg.code + "-";
    let n = 0; used.forEach(id => { if (id.indexOf(prefix) === 0) n++; });
    left[t] = cfg.rows * cfg.perRow - n;
  });
  return left;
}
const SHORT_ = { "個人光明燈": "光明燈", "五路財神燈": "財神燈" };
function seatText_(d) {
  return (d.lamps || []).filter(p => p.seats && Object.keys(p.seats).length)
    .map(p => p.name + "：" + Object.keys(p.seats).map(t => (SHORT_[t] || t) + " " + (p.seats[t] || "候補")).join("、")).join("\n");
}

/* ================= 讀寫一筆報名 ================= */
function readRecord_(r) {
  let d = {}; try { d = JSON.parse(r[C.json - 1]); } catch (e) {}
  let log = []; try { log = JSON.parse(r[C.log - 1] || "[]"); } catch (e) {}
  d.ref = String(r[C.ref - 1]);
  d.paid = r[C.paid - 1] === true;
  d.paidAt = r[C.paidAt - 1] ? new Date(r[C.paidAt - 1]).toISOString() : null;
  d.paidBy = r[C.paidBy - 1] || null;
  d.cancelled = r[C.cancelled - 1] === true;
  d.payLog = log;
  return d;
}
function findRow_(ref) {
  const sh = sheet_(SHEET.reg), n = sh.getLastRow() - 1;
  if (n < 1 || !ref) return null;
  const refs = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (String(refs[i][0]) === String(ref).trim().toUpperCase()) {
    const row = i + 2;
    return { sh: sh, row: row, rec: readRecord_(sh.getRange(row, 1, 1, REG_HEAD.length).getValues()[0]) };
  }
  return null;
}
// 把修改後的報名內容寫回（繳費、取消、收款紀錄另外處理）
function writeData_(f, d) {
  const keep = {}; ["paid", "paidAt", "paidBy", "payLog", "cancelled"].forEach(k => { keep[k] = d[k]; delete d[k]; });
  f.sh.getRange(f.row, C.name, 1, 4).setValues([[d.contact.name || "", d.contact.phone || "", d.contact.address || "", d.totals.all]]);
  f.sh.getRange(f.row, C.note).setValue(d.note || "");
  f.sh.getRange(f.row, C.json).setValue(JSON.stringify(d));
  f.sh.getRange(f.row, C.seats).setValue(seatText_(d));
  Object.assign(d, keep);
}
function addLog_(f, entry) {
  const log = f.rec.payLog || [];
  log.push(entry);
  f.sh.getRange(f.row, C.log).setValue(JSON.stringify(log));
  f.rec.payLog = log;
}
// 給信眾看的資料：不含收款紀錄和內部欄位
function publicRecord_(d) {
  const o = JSON.parse(JSON.stringify(d));
  delete o.payLog; delete o.clientId; delete o.paidBy;
  return o;
}

function listRows_() {
  const sh = sheet_(SHEET.reg), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, REG_HEAD.length).getValues().map(readRecord_);
}

/* ================= 信眾查詢、修改 ================= */
const digits_ = s => String(s || "").replace(/\D/g, "");
function phoneMatch_(stored, given) {
  const a = digits_(stored), b = digits_(given);
  return b.length >= 7 && (a === b || a.slice(-10) === b.slice(-10));
}
// 用報名編號＋電話找自己的報名；同一個編號連續打錯會暫時鎖住
function lookup_(ref, phone) {
  const cache = CacheService.getScriptCache(), key = "lk-" + String(ref || "").toUpperCase();
  const fails = Number(cache.get(key) || 0);
  if (fails >= MAX_FAILS) return { ok: false, error: "locked" };
  const f = findRow_(ref);
  if (!f || !phoneMatch_(f.rec.contact && f.rec.contact.phone, phone)) {
    cache.put(key, String(fails + 1), LOCK_MINUTES * 60);
    return { ok: false, error: "not-found" };
  }
  if (fails) cache.remove(key);
  return { ok: true, record: publicRecord_(f.rec) };
}
// 繳費前可以修改；修改後重新計算金額，原本的燈位盡量保留
function update_(b) {
  const look = lookup_(b.ref, b.phone);
  if (!look.ok) return look;
  const f = findRow_(b.ref), old = f.rec;
  if (old.cancelled) return { ok: false, error: "cancelled" };
  if (old.paid) return { ok: false, error: "paid" };
  const d = b.data || {};
  if (!d.contact) return { ok: false, error: "bad-data" };
  cleanData_(d);
  // 不能改的欄位沿用原本的
  ["ref", "form", "eventId", "formTitle", "submittedAt", "clientId", "lang"].forEach(k => { d[k] = old[k]; });
  if (b.data.lang) d.lang = b.data.lang;
  d.updatedAt = new Date().toISOString();
  d.totals = calcTotals_(d);
  if (d.form === "lamp") assignSeats_(d, usedSeats_(old.ref), old);
  Object.assign(d, { paid: old.paid, paidAt: old.paidAt, paidBy: old.paidBy, payLog: old.payLog, cancelled: old.cancelled });
  writeData_(f, d);
  addLog_(f, { action: "edit", by: "信眾", at: d.updatedAt, amount: d.totals.all });
  d.payLog = f.rec.payLog;
  return { ok: true, record: publicRecord_(d) };
}

/* ================= 後台：收款、取消、分配燈位 ================= */
function setPaid_(me, ref, paid) {
  const f = findRow_(ref);
  if (!f) return { ok: false, error: "not-found" };
  if (paid && f.rec.cancelled) return { ok: false, error: "cancelled" };
  const now = new Date(), amount = f.rec.totals ? f.rec.totals.all : f.sh.getRange(f.row, C.total).getValue();
  f.sh.getRange(f.row, C.paid, 1, 3).setValues([[paid, paid ? now : "", paid ? me.name : ""]]);
  addLog_(f, { action: paid ? "paid" : "unpaid", by: me.name, byId: me.id, at: now.toISOString(), amount: amount });
  return { ok: true, record: { paid: paid, paidAt: paid ? now.toISOString() : null, paidBy: paid ? me.name : null, paidById: paid ? me.id : null, payLog: f.rec.payLog } };
}
// 取消報名會把燈位放出來；恢復時重新分配（原位還空著就用原位）
function setCancelled_(me, ref, cancelled) {
  const f = findRow_(ref);
  if (!f) return { ok: false, error: "not-found" };
  if (cancelled && f.rec.paid) return { ok: false, error: "paid" };
  const d = f.rec;
  if (!cancelled && d.form === "lamp") { assignSeats_(d, usedSeats_(d.ref), d); writeData_(f, d); }
  f.sh.getRange(f.row, C.cancelled).setValue(cancelled);
  d.cancelled = cancelled;
  addLog_(f, { action: cancelled ? "cancel" : "restore", by: me.name, byId: me.id, at: new Date().toISOString() });
  return { ok: true, record: d };
}
// 候補的燈：有空位時補上燈位（已經有燈位的不變）
function assignWaitlist_(me, ref) {
  const f = findRow_(ref);
  if (!f) return { ok: false, error: "not-found" };
  const d = f.rec;
  if (d.cancelled) return { ok: false, error: "cancelled" };
  assignSeats_(d, usedSeats_(d.ref), d);
  writeData_(f, d);
  addLog_(f, { action: "seat", by: me.name, byId: me.id, at: new Date().toISOString() });
  return { ok: true, record: d };
}

/* ================= 設定 ================= */
function getSettings_() {
  try { const v = sheet_(SHEET.settings).getRange("A1").getValue(); if (v) { const s = JSON.parse(v); if (s && s.events) return s; } } catch (e) {}
  return JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
}
function saveSettings_(s) {
  if (!s || !s.lamp || !Array.isArray(s.events)) return { ok: false, error: "bad-data" };
  delete s.seatsLeft;
  s.updatedAt = new Date().toISOString();
  sheet_(SHEET.settings).getRange("A1").setValue(JSON.stringify(s));
  return { ok: true };
}

/* ================= 帳號 ================= */
function readUsers_() {
  const sh = sheet_(SHEET.users), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, USER_HEAD.length).getValues().map((r, i) => ({
    row: i + 2, id: String(r[0]), name: String(r[1]), role: r[2] === "admin" ? "admin" : "volunteer",
    pin: String(r[3]), active: r[4] === true, mustChange: r[5] === true,
  }));
}
function writeUser_(u) {
  const vals = [[u.id, u.name, u.role, u.pin, !!u.active, !!u.mustChange]];
  const sh = sheet_(SHEET.users);
  if (u.row) sh.getRange(u.row, 1, 1, USER_HEAD.length).setValues(vals); else sh.appendRow(vals[0]);
}
const pub_ = u => ({ id: u.id, name: u.name, role: u.role, active: u.active, mustChange: u.mustChange });

// 跟後台網頁用同一種加密方式：SHA-256("tsuhui:" + PIN)
function hashPin_(pin) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, "tsuhui:" + String(pin), Utilities.Charset.UTF_8);
  return b.map(x => ((x + 256) % 256).toString(16).padStart(2, "0")).join("");
}

// 登入檢查；連續打錯太多次會暫時鎖住這個帳號
function isLocked_(uid) { return Number(CacheService.getScriptCache().get("fail-" + uid) || 0) >= MAX_FAILS; }
function checkLogin_(uid, pin) {
  if (!uid || !pin) return null;
  const cache = CacheService.getScriptCache(), key = "fail-" + uid;
  const fails = Number(cache.get(key) || 0);
  if (fails >= MAX_FAILS) return null;
  const u = readUsers_().find(x => x.id === String(uid) && x.active);
  if (u && u.pin === hashPin_(pin)) { if (fails) cache.remove(key); return u; }
  cache.put(key, String(fails + 1), LOCK_MINUTES * 60);
  return null;
}

// 前端送來的欄位：uid／pin 是登入用；newPin 是要設定的新 PIN
function userOp_(me, b) {
  const list = readUsers_(), meU = list.find(x => x.id === me.id);
  const pinOK = p => /^\d{4,8}$/.test(String(p || ""));
  if (b.op === "changeMyPin") {
    if (meU.pin !== hashPin_(b.cur)) return { ok: false, error: "cur" };
    if (!pinOK(b.newPin)) return { ok: false, error: "pin" };
    if (hashPin_(b.newPin) === meU.pin) return { ok: false, error: "same" };
    meU.pin = hashPin_(b.newPin); meU.mustChange = false; writeUser_(meU);
    return { ok: true };
  }
  if (me.role !== "admin") return { ok: false, error: "auth" };
  if (b.op === "add") {
    const name = String(b.name || "").trim();
    if (!name || !pinOK(b.newPin)) return { ok: false, error: "bad-data" };
    if (list.some(x => x.name === name)) return { ok: false, error: "dup" };
    writeUser_({ id: "u-" + Date.now().toString(36), name: name, role: b.role === "admin" ? "admin" : "volunteer", pin: hashPin_(b.newPin), active: true, mustChange: true });
    return { ok: true };
  }
  const u = list.find(x => x.id === b.id);
  if (!u) return { ok: false, error: "not-found" };
  if (b.op === "resetPin") {
    if (!pinOK(b.newPin)) return { ok: false, error: "pin" };
    u.pin = hashPin_(b.newPin); u.mustChange = true; writeUser_(u);
    CacheService.getScriptCache().remove("fail-" + u.id);   // 重設 PIN 時順便解鎖
    return { ok: true };
  }
  if (b.op === "setActive") {
    if (u.id === me.id) return { ok: false, error: "self" };
    u.active = !!b.active; writeUser_(u);
    return { ok: true };
  }
  return { ok: false, error: "unknown-op" };
}

/* ================= 工具 ================= */
function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error("找不到「" + name + "」分頁，請先執行 setup");
  return sh;
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

  /* ================= Code.gs 結束 ================= */

  // 第一次打開示範版：建立分頁、示範帳號和幾筆範例報名
  function seed() {
    setup();
    const users = sheet_(SHEET.users);
    users.getRange(2, 6).setValue(false);   // 示範版管理員不用改 PIN
    [["u-wang", "王師姐", "1234"], ["u-li", "李師兄", "5678"]].forEach(([id, n, pin]) => users.appendRow([id, n, "volunteer", hashPin_(pin), true, false]));
    const B = (y, m, d, h) => ({ lunar: { year: y, month: m, leap: false, day: d, hour: h || "" }, solar: null, entered: "lunar" });
    const samples = [
      { contact: { name: "王秀英", phone: "(718) 555-0142", address: "範例地址" }, lamps: [
        { pid: "s1", name: "王秀英", gender: "女", birth: B(1958, 8, 12, "辰"), zodiac: "狗", age_xu: 70, lampTypes: ["個人光明燈", "太歲燈"] },
        { pid: "s2", name: "李建國", gender: "男", birth: B(1955, 3, 3), zodiac: "羊", age_xu: 73, lampTypes: ["個人光明燈"] }] },
      { contact: { name: "張雅婷", phone: "(347) 555-0198", address: "範例地址" }, lamps: [
        { pid: "s3", name: "張雅婷", gender: "女", birth: B(1986, 11, 2, "午"), zodiac: "虎", age_xu: 42, lampTypes: ["五路財神燈"] },
        { pid: "s4", name: "陳柏宇", gender: "男", birth: B(2012, 6, 18), zodiac: "龍", age_xu: 16, lampTypes: ["文昌燈", "個人光明燈"] },
        { pid: "s5", name: "陳品妍", gender: "女", birth: B(2015, 1, 9), zodiac: "羊", age_xu: 13, lampTypes: ["文昌燈"] }] },
    ];
    const refs = samples.map(s => submit_(Object.assign({ form: "lamp", eventId: null, formTitle: "啟點光明燈暨安奉太歲", note: "", lang: "zh", chaodu: [], meals: [] }, s)).ref);
    // 標成範例，第一筆設為已繳清
    const sh = sheet_(SHEET.reg);
    for (let r = 2; r <= sh.getLastRow(); r++) { const d = JSON.parse(sh.getRange(r, C.json).getValue()); d.sample = true; sh.getRange(r, C.json).setValue(JSON.stringify(d)); }
    setPaid_({ id: "u-admin", name: "管理員" }, refs[0], true);
    db.seeded = true;
  }
  if (!db.seeded) { seed(); persist(); }

  // 報名頁和後台可能同時開著：每次都先讀最新的資料，避免互相蓋掉
  const refresh = () => { try { const v = JSON.parse(localStorage.getItem(KEY) || "null"); if (v && v.sheets) db = v; } catch (e) {} };
  return {
    get(q) { refresh(); const r = JSON.parse(doGet({ parameter: q || {} }).text); persist(); return r; },
    post(b) { refresh(); const r = JSON.parse(doPost({ postData: { contents: JSON.stringify(b) } }).text); persist(); return r; },
    reset() { db = { sheets: {}, order: [] }; seed(); persist(); },
  };
})();
