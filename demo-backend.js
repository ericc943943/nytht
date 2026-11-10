/* 自動產生，請勿直接修改：示範模式用的後端。
 * 內容是 backend/Code.gs 原封不動，外面包一層「假的 Google 試算表」，資料存在這台電腦的瀏覽器裡。
 * 改了 Code.gs 之後執行 python3 backend/build-demo.py 重新產生。 */
window.DemoBackend = (function () {
  "use strict";
  const KEY = "tsuhui-demo-v8";
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
      deleteRow: (r) => { data().splice(r - 1, 1); },
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
  lamp: { "光明燈": 120, "太歲燈": 120, "財神燈": 120, "文昌燈": 120 },  // 點燈：每盞
  lotus: 60,     // 法會：蓮花福慧燈，每位
  chaodu: 160,   // 法會：超渡牌位，每牌位
};
// 燈位：每種燈牆上的排列。順序照燈種代碼（01、02、03、05）。
// 燈號 = 英文字母＋燈種代碼（給燈號查詢機用）＋排＋第幾個，例：G010512 ＝ 光明燈第 5 排第 12 個。
// 字母：G 光明、W 文昌、T 太歲、C 財神。單據上燈號前面都會寫燈的名稱。
const SEATS = {
  "光明燈": { letter: "G", label: "G", code: "01", perRow: 30, rows: 60 },
  "文昌燈":     { letter: "W", label: "W", code: "02", perRow: 20, rows: 60 },
  "太歲燈":     { letter: "T", label: "T", code: "03", perRow: 30, rows: 39 },
  "財神燈": { letter: "C", label: "C", code: "05", perRow: 30, rows: 70 },
};
const DEFAULT_LAMP_YEAR = 2027;   // 點燈年度：之後在後台「設定」修改
const DEFAULT_SETTINGS = {
  lamp: { status: "open", year: DEFAULT_LAMP_YEAR },
  events: [
    { id: "2027-zhongyuan", name: "瑤池金母聖誕祈福消災超渡暨普渡大法會", date: "", status: "hidden" },
    { id: "2027-spring-lidou", name: "春季祈安禮斗法會", date: "", status: "hidden" },
    { id: "2027-autumn-lidou", name: "秋季祈安禮斗法會", date: "", status: "hidden" },
  ],
};
const FIRST_ADMIN = { name: "管理員", pin: "8888" };   // 第一次登入會被要求改 PIN
const MAX_FAILS = 8;          // PIN 或查詢連續打錯幾次就暫時鎖住
const LOCK_MINUTES = 15;

const SHEET = { reg: "報名", users: "帳號", settings: "設定", audit: "紀錄" };
const REG_HEAD = ["報名編號", "報名時間", "表單", "場次", "聯絡人", "電話", "住址", "應繳", "已繳清", "收款時間", "經手人", "語言", "備註", "收款紀錄", "送出代碼", "完整資料", "已取消", "燈位", "點燈年度", "付款方式", "支票戶名", "支票號碼", "支票金額", "支票付", "現金付", "支票備註"];
const C = { ref: 1, time: 2, form: 3, event: 4, name: 5, phone: 6, address: 7, total: 8, paid: 9, paidAt: 10, paidBy: 11, lang: 12, note: 13, log: 14, cid: 15, json: 16, cancelled: 17, seats: 18, year: 19, method: 20, checkName: 21, checkNo: 22, checkAmt: 23, checkPart: 24, cashPart: 25, checkNote: 26 };
const USER_HEAD = ["id", "姓名", "身分", "PIN（加密）", "啟用", "待改PIN"];
const AUDIT_HEAD = ["時間", "經手人", "動作", "對象", "內容"];

/* ================= 身分與權限 =================
   admin 管理員（舊的身分，等於全部權限；交接時拆成堂主＋網站管理員）
   owner 堂主：全部。deputy 副堂主：跟堂主一樣，但不能動堂主的帳號、不能指定堂主或副堂主。finance 財務：看全部帳、改別人的收款。web 網站管理員：設定、義工帳號，看不到總帳。volunteer 義工。
   錢（總帳、別人經手的付款細節、改回別人收的款）＝ admin／owner／finance
   設定、帳號 ＝ admin／owner／web（web 只能管義工帳號）
   指定財務、網站管理員 ＝ admin／owner／deputy；指定堂主、副堂主 ＝ admin／owner。堂主、副堂主各限一位 */
const ROLES = ["admin", "owner", "deputy", "finance", "web", "volunteer"];
const ROLE_ZH = { admin: "管理員", owner: "堂主", deputy: "副堂主", finance: "財務", web: "網站管理員", volunteer: "義工" };
const canMoney_ = u => ["admin", "owner", "deputy", "finance"].includes(u.role);
const canSettings_ = u => ["admin", "owner", "deputy", "web"].includes(u.role);
const canAccounts_ = u => ["admin", "owner", "deputy", "web"].includes(u.role);
const canAssignRoles_ = u => ["admin", "owner", "deputy"].includes(u.role);
const isBoss_ = u => ["admin", "owner", "deputy"].includes(u.role) && u.active;

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
      case "settings": { const s = getSettings_(); s.lamp.year = lampYear_(s); s.seatsLeft = seatsLeft_(s.lamp.year); return out_(s); }
      case "users":    return out_({ users: readUsers_().map(pub_) });
      case "login": {
        if (isLocked_(p.uid)) return out_({ ok: false, error: "locked" });
        const u = checkLogin_(p.uid, p.pin);
        return out_(u ? { ok: true, user: pub_(u) } : { ok: false, error: "auth" });
      }
      case "list": {
        const me = checkLogin_(p.uid, p.pin);
        if (!me) return out_({ ok: false, error: "auth" });
        return out_({ ok: true, me: pub_(me), features: ["payMethod", "roles"], rows: listRows_().map(d => viewFor_(me, d)) });
      }
      case "findCheck": {  // 收支票時：這張支票是不是已經用在別的報名表
        const me = checkLogin_(p.uid, p.pin);
        if (!me) return out_({ ok: false, error: "auth" });
        return out_({ ok: true, checks: findChecks_(p.no, p.except) });
      }
      case "audit": {      // 帳號與設定的變動紀錄
        const me = checkLogin_(p.uid, p.pin);
        if (!me || me.role === "volunteer") return out_({ ok: false, error: "auth" });
        return out_({ ok: true, rows: readAudit_() });
      }
      case "lookup": return out_(lookup_(p.ref, p.phone));    // 信眾用編號＋電話查自己的報名
      case "findRefs": return out_(findRefs_(p.phone, p.name)); // 忘記報名編號：電話＋聯絡人姓名找回編號
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
      case "setPaid":      return out_(setPaid_(me, body.ref, !!body.paid, body));
      case "setCancelled": return out_(setCancelled_(me, body.ref, !!body.cancelled));
      case "assignSeats":  return out_(assignWaitlist_(me, body.ref));
      case "setCard":      return out_(setCard_(me, body.ref, !!body.given));
      case "editPay":      return out_(setPaid_(me, body.ref, true, body, true));
      case "setCheckNote": if (!canMoney_(me)) return out_({ ok: false, error: "auth" });
                           return out_(setCheckNote_(me, body.no, body.name, body.note));
      case "saveSettings": if (!canSettings_(me)) return out_({ ok: false, error: "auth" });
                           return out_(saveSettings_(body.settings, me));
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
  // 報名編號由這裡決定：L（點燈）／F（法會）＋ yymmdd ＋ - ＋當天流水號，例：L261031-001
  const tz = Session.getScriptTimeZone() || "America/New_York";
  const prefix = (d.form === "lamp" ? "L" : "F") + Utilities.formatDate(new Date(), tz, "yyMMdd") + "-";
  let maxNo = 0;
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(r => {
    const v = String(r[0]); if (v.indexOf(prefix) === 0) maxNo = Math.max(maxNo, Number(v.slice(prefix.length)) || 0);
  });
  const ref = prefix + ("00" + (maxNo + 1)).slice(-Math.max(3, String(maxNo + 1).length));

  cleanData_(d);
  d.totals = calcTotals_(d);                     // 金額一律在這裡重新計算
  if (d.form === "lamp") {
    d.lampYear = lampYear_(st);                  // 點燈年度以後台設定為準
    applyAges_(d, d.lampYear);
    assignSeats_(d, usedSeats_("", d.lampYear), null);
  }
  d.ref = ref; d.paid = false; d.cancelled = false; d.submittedAt = new Date().toISOString();
  const evName = d.form === "lamp" ? "點燈" : ((st.events.find(x => x.id === d.eventId) || {}).name || d.formTitle || "法會");
  sh.appendRow([ref, new Date(), d.form === "lamp" ? "點燈" : "法會", evName, d.contact.name || "", d.contact.phone || "", d.contact.address || "",
                d.totals.all, false, "", "", d.lang === "en" ? "英文" : "中文", d.note || "", "[]", d.clientId || "", JSON.stringify(d), false, seatText_(d), d.lampYear || ""]);
  return { ok: true, ref: ref, record: publicRecord_(d) };
}

// 只留下報名需要的欄位（金額、燈位、繳費狀態一律由後端決定）
function cleanData_(d) {
  // 追加報名：記下原本的報名編號（格式不對就不記）
  if (d.addTo != null) { d.addTo = String(d.addTo).trim().toUpperCase(); if (!/^[LF]\d{6}-\d{3,}$/.test(d.addTo)) delete d.addTo; }
  normLamps_(d);
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

/* ================= 點燈年度 ================= */
const lampYear_ = st => Number((st && st.lamp && st.lamp.year) || DEFAULT_LAMP_YEAR);
const recYear_ = d => Number(d.lampYear || DEFAULT_LAMP_YEAR);   // 沒有年度的舊資料當作預設年度
// 歲數照點燈那一年的農曆年算（出生那年算 1 歲）
function applyAges_(d, year) {
  (d.lamps || []).forEach(p => { const y = p.birth && p.birth.lunar && Number(p.birth.lunar.year); p.age_xu = y ? year - y + 1 : null; });
}

/* ================= 燈位 ================= */
const seatId_ = (cfg, r, c) => cfg.letter + cfg.code + ("0" + r).slice(-2) + ("0" + c).slice(-2);
// 顯示用（目前直接顯示字母，例 G010512）
const seatLabel_ = id => { if (!id) return ""; const t = Object.keys(SEATS).find(k => SEATS[k].letter === id[0]); return t ? SEATS[t].label + id.slice(1) : id; };

// 目前被佔用的燈位（只算同一個點燈年度；已取消的不算；exceptRef 是正在修改的那一筆）
function usedSeats_(exceptRef, year) {
  const used = new Set(), sh = sheet_(SHEET.reg), n = sh.getLastRow() - 1;
  if (n < 1) return used;
  sh.getRange(2, 1, n, REG_HEAD.length).getValues().forEach(r => {
    if (r[C.cancelled - 1] === true || String(r[0]) === String(exceptRef)) return;
    let d = {}; try { d = normLamps_(JSON.parse(r[C.json - 1])); } catch (e) { return; }
    if (d.form !== "lamp" || recYear_(d) !== Number(year)) return;
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
function seatsLeft_(year) {
  const used = usedSeats_("", year), left = {};
  Object.keys(SEATS).forEach(t => {
    const cfg = SEATS[t], prefix = cfg.letter + cfg.code;
    let n = 0; used.forEach(id => { if (id.indexOf(prefix) === 0) n++; });
    left[t] = cfg.rows * cfg.perRow - n;
  });
  return left;
}
const SHORT_ = {};   // 燈名已經都是三個字
// 舊資料的燈名（2026-11 以前存的是全名）讀進來時換成三個字
const LAMP_RENAME_ = { "個人光明燈": "光明燈", "五路財神燈": "財神燈" };
function normLamps_(d) {
  const ren = t => LAMP_RENAME_[t] || t;
  (d.lamps || []).forEach(p => {
    if (p.lampTypes) p.lampTypes = p.lampTypes.map(ren);
    if (p.seats) { const o = {}; Object.keys(p.seats).forEach(k => { o[ren(k)] = p.seats[k]; }); p.seats = o; }
  });
  if (d.totals && d.totals.byType) { const o = {}; Object.keys(d.totals.byType).forEach(k => { const x = d.totals.byType[k]; o[ren(k)] = Object.assign({}, x, { name: ren(x.name || k) }); }); d.totals.byType = o; }
  return d;
}
function seatText_(d) {
  return (d.lamps || []).filter(p => p.seats && Object.keys(p.seats).length)
    .map(p => p.name + "：" + Object.keys(SEATS).filter(t => t in p.seats).map(t => (SHORT_[t] || t) + " " + (seatLabel_(p.seats[t]) || "候補")).join("、")).join("\n");
}

/* ================= 讀寫一筆報名 ================= */
function readRecord_(r) {
  let d = {}; try { d = JSON.parse(r[C.json - 1]); } catch (e) {}
  normLamps_(d);
  let log = []; try { log = JSON.parse(r[C.log - 1] || "[]"); } catch (e) {}
  d.ref = String(r[C.ref - 1]);
  d.paid = r[C.paid - 1] === true;
  d.paidAt = r[C.paidAt - 1] ? new Date(r[C.paidAt - 1]).toISOString() : null;
  d.paidBy = r[C.paidBy - 1] || null;
  d.cancelled = r[C.cancelled - 1] === true;
  // 付款方式：現金、支票、支票＋現金。支票記支票戶名、號碼、整張支票的金額，和這筆用支票付了多少、現金付了多少
  // 以前收的（沒記方式）都算現金；手動在試算表補「支票」但沒填金額的，當成整筆用支票付
  const pm = String(r[C.method - 1] || ""), due = Number(r[C.total - 1]) || 0;
  d.payMethod = d.paid ? (pm === "支票" ? "check" : pm === "支票＋現金" ? "mix" : "cash") : null;
  const ck = d.payMethod === "check" || d.payMethod === "mix";
  d.checkName = ck ? String(r[C.checkName - 1] || "") : null;
  d.checkNo = ck ? String(r[C.checkNo - 1] || "") : null;
  d.checkPart = ck ? (r[C.checkPart - 1] === "" ? due : Number(r[C.checkPart - 1]) || 0) : (d.paid ? 0 : null);
  d.cashPart = d.paid ? (ck ? (r[C.cashPart - 1] === "" ? due - d.checkPart : Number(r[C.cashPart - 1]) || 0) : due) : null;
  d.checkAmt = ck ? (Number(r[C.checkAmt - 1]) || null) : null;
  d.checkNote = ck ? String(r[C.checkNote - 1] || "") : null;
  d.payLog = log;
  // 點燈卡：看處理紀錄裡最後一次「發點燈卡／改回未領」
  let card = null; log.forEach(l => { if (l.action === "card") card = l; else if (l.action === "uncard") card = null; });
  d.cardAt = card ? card.at : null; d.cardBy = card ? card.by : null;
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
  const keep = {}; ["paid", "paidAt", "paidBy", "payLog", "cancelled", "payMethod", "checkName", "checkNo", "checkPart", "cashPart", "checkAmt", "checkNote"].forEach(k => { keep[k] = d[k]; delete d[k]; });
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
// 這筆是不是我收的款（看最後一次「標記已繳清」是誰）
function paidByMe_(me, d) {
  const last = (d.payLog || []).filter(x => x.action === "paid").pop();
  return last ? (last.byId ? last.byId === me.id : last.by === me.name) : d.paidBy === me.name;
}
// 後台清單：看不到總帳的人（網站管理員、義工），別人收的款不給經手人和付款細節
function viewFor_(me, d) {
  if (canMoney_(me) || !d.paid || paidByMe_(me, d)) return d;
  const o = JSON.parse(JSON.stringify(d));
  delete o.paidBy; delete o.paidById; delete o.payMethod; delete o.checkName; delete o.checkNo;
  delete o.checkPart; delete o.cashPart; delete o.checkAmt; delete o.checkNote;
  o.payLog = (o.payLog || []).map(x => (x.action === "paid" || x.action === "unpaid" || x.action === "editpay") && x.byId !== me.id && x.by !== me.name
    ? { action: x.action, at: x.at, amount: x.amount } : x);
  return o;
}
// 帳號、設定的變動寫進「紀錄」分頁
function audit_(me, action, target, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET.audit);
  if (!sh) { sh = ss.insertSheet(SHEET.audit); sh.getRange(1, 1, 1, AUDIT_HEAD.length).setValues([AUDIT_HEAD]).setFontWeight("bold"); sh.setFrozenRows(1); }
  sh.appendRow([new Date(), me.name, action, target || "", detail || ""]);
}
function readAudit_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.audit);
  if (!sh || sh.getLastRow() < 2) return [];
  const n = sh.getLastRow() - 1, from = Math.max(2, sh.getLastRow() - 499);
  return sh.getRange(from, 1, sh.getLastRow() - from + 1, AUDIT_HEAD.length).getValues().map(r => ({
    at: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), by: String(r[1]), action: String(r[2]), target: String(r[3]), detail: String(r[4]) })).reverse();
}
// 給信眾看的資料：不含收款紀錄和內部欄位
function publicRecord_(d) {
  const o = JSON.parse(JSON.stringify(d));
  delete o.payLog; delete o.clientId; delete o.paidBy; delete o.cardBy; delete o.checkName;
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
// 用報名編號＋電話找自己的報名。同一個編號、或同一支電話連續打錯，都會暫時鎖住
// （編號是連號，所以也要擋「拿一支電話一個一個試編號」）
function lookup_(ref, phone) {
  const cache = CacheService.getScriptCache();
  const keys = ["lk-" + String(ref || "").trim().toUpperCase(), "lkp-" + digits_(phone).slice(-10)];
  const fails = keys.map(k => Number(cache.get(k) || 0));
  if (fails.some(n => n >= MAX_FAILS)) return { ok: false, error: "locked" };
  const f = findRow_(ref);
  if (!f || !phoneMatch_(f.rec.contact && f.rec.contact.phone, phone)) {
    keys.forEach((k, i) => cache.put(k, String(fails[i] + 1), LOCK_MINUTES * 60));
    return { ok: false, error: "not-found" };
  }
  keys.forEach((k, i) => { if (fails[i]) cache.remove(k); });
  return { ok: true, record: publicRecord_(f.rec) };
}
// 忘記報名編號：電話＋聯絡人姓名都對，才列出編號（只給編號和表單名稱，不給其他資料）。
// 同一支電話連續找不到會暫時鎖住
function findRefs_(phone, name) {
  const norm = x => String(x || "").replace(/\s+/g, "").toLowerCase();
  const key = "frp-" + digits_(phone).slice(-10), cache = CacheService.getScriptCache();
  const fails = Number(cache.get(key) || 0);
  if (fails >= MAX_FAILS) return { ok: false, error: "locked" };
  if (digits_(phone).length < 7 || !norm(name)) return { ok: false, error: "not-found" };
  // 姓名：聯絡人、點燈的人、供齋的人、超渡的陽世報恩人，任何一個對上都算（不分大小寫、空格）
  const names = d => [d.contact && d.contact.name, ...(d.lamps || []).map(p => p.name), ...(d.meals || []).map(m => m.name), ...(d.chaodu || []).map(c => c.sponsor)];
  const hits = listRows_().filter(d => phoneMatch_(d.contact && d.contact.phone, phone) && names(d).some(n => n && norm(n) === norm(name)))
    .map(d => ({ ref: d.ref, form: d.form, eventId: d.eventId || null, formTitle: d.formTitle || "", lampYear: d.lampYear || null,
                 submittedAt: d.submittedAt, paid: !!d.paid, cancelled: !!d.cancelled }))
    .sort((a, b) => String(b.submittedAt).localeCompare(String(a.submittedAt)));
  if (!hits.length) { cache.put(key, String(fails + 1), LOCK_MINUTES * 60); return { ok: false, error: "not-found" }; }
  if (fails) cache.remove(key);
  return { ok: true, refs: hits };
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
  ["ref", "form", "eventId", "formTitle", "submittedAt", "clientId", "lang", "lampYear", "addTo"].forEach(k => { d[k] = old[k]; });
  if (b.data.lang) d.lang = b.data.lang;
  d.updatedAt = new Date().toISOString();
  d.totals = calcTotals_(d);
  if (d.form === "lamp") { applyAges_(d, recYear_(d)); assignSeats_(d, usedSeats_(old.ref, recYear_(d)), old); }
  Object.assign(d, { paid: old.paid, paidAt: old.paidAt, paidBy: old.paidBy, payLog: old.payLog, cancelled: old.cancelled });
  writeData_(f, d);
  addLog_(f, { action: "edit", by: "信眾", at: d.updatedAt, amount: d.totals.all });
  d.payLog = f.rec.payLog;
  return { ok: true, record: publicRecord_(d) };
}

/* ================= 後台：收款、取消、分配燈位 ================= */
// 收款：現金或支票。支票要填支票戶名、號碼、金額；一張支票可以付好幾張報名表（輸入同一個號碼＋戶名）。
// 支票剩下的錢不夠付這一筆，其餘收現金（付款方式記「支票＋現金」）。寫在「付款方式」到「支票備註」七欄，也記進收款紀錄
const ckKey_ = (no, name) => String(no || "").trim() + "|" + String(name || "").replace(/\s+/g, "").toLowerCase();
function findChecks_(no, exceptRef) {
  no = String(no || "").trim(); if (!no) return [];
  const sh = sheet_(SHEET.reg), n = sh.getLastRow() - 1; if (n < 1) return [];
  const map = {};
  sh.getRange(2, 1, n, REG_HEAD.length).getValues().forEach(r => {
    if (r[C.paid - 1] !== true || r[C.cancelled - 1] === true || String(r[C.checkNo - 1]).trim() !== no) return;
    if (exceptRef && String(r[C.ref - 1]) === exceptRef) return;
    const d = readRecord_(r); if (d.payMethod !== "check" && d.payMethod !== "mix") return;
    const k = ckKey_(d.checkNo, d.checkName);
    const c = map[k] = map[k] || { no: d.checkNo, name: d.checkName, amount: 0, used: 0, refs: [], note: "" };
    c.used += d.checkPart || 0; c.amount = Math.max(c.amount, d.checkAmt || 0); c.refs.push({ ref: d.ref, part: d.checkPart }); if (d.checkNote) c.note = d.checkNote;
  });
  return Object.keys(map).map(k => { const c = map[k]; if (!c.amount) c.amount = c.used; c.left = Math.round((c.amount - c.used) * 100) / 100; return c; });
}
// 付款方式的文字（寫在處理紀錄「修改收款」用）
function payDesc_(d) {
  if (d.payMethod === "check" || d.payMethod === "mix") return "支票 #" + d.checkNo + " " + d.checkName + " $" + d.checkPart + (d.payMethod === "mix" ? "＋現金 $" + d.cashPart : "");
  return "現金";
}
// edit＝修改收款：已繳清的報名只改付款方式和支票資料，收款日期和經手人不變
function setPaid_(me, ref, paid, pay, edit) {
  const f = findRow_(ref);
  if (!f) return { ok: false, error: "not-found" };
  if (edit && (!f.rec.paid || f.rec.cancelled)) return { ok: false, error: "unpaid" };
  if (edit && !canMoney_(me) && !paidByMe_(me, f.rec)) return { ok: false, error: "not-mine" };
  if (paid && f.rec.cancelled) return { ok: false, error: "cancelled" };
  if (!paid && f.rec.paid && !canMoney_(me) && !paidByMe_(me, f.rec)) return { ok: false, error: "not-mine" };   // 別人收的款，要堂主或財務才能改回
  const now = new Date(), amount = f.rec.totals ? f.rec.totals.all : Number(f.sh.getRange(f.row, C.total).getValue()) || 0;
  const check = paid && pay && pay.method === "check";
  let ckName = "", ckNo = "", ckAmt = "", ckPart = "", cashPart = "", note = "";
  if (check) {
    ckName = String(pay.checkName || "").trim().slice(0, 60); ckNo = String(pay.checkNo || "").trim().slice(0, 30);
    if (!ckName || !ckNo) return { ok: false, error: "check" };
    const same = findChecks_(ckNo, f.rec.ref).find(c => ckKey_(c.no, c.name) === ckKey_(ckNo, ckName));
    if (same) { ckName = same.name; ckAmt = same.amount; note = same.note; }
    else { ckAmt = Math.round(Number(pay.checkAmount) * 100) / 100; if (!(ckAmt > 0)) return { ok: false, error: "check-amount" }; }
    const left = same ? same.left : ckAmt;
    if (!(left > 0)) return { ok: false, error: "check-used" };
    // 使用金額：這筆從支票扣多少（不能超過支票剩下的錢、也不能超過應繳）；沒給就盡量用支票付
    const want = Math.round(Number(pay.checkPart) * 100) / 100;
    ckPart = want > 0 ? Math.min(want, left, amount) : Math.min(left, amount);
    cashPart = Math.round((amount - ckPart) * 100) / 100;
  }
  const method = !paid ? "" : !check ? "現金" : cashPart > 0 ? "支票＋現金" : "支票";
  if (!edit) f.sh.getRange(f.row, C.paid, 1, 3).setValues([[paid, paid ? now : "", paid ? me.name : ""]]);
  if (f.sh.getRange(1, C.checkNote).getValue() !== REG_HEAD[C.checkNote - 1] || f.sh.getRange(1, C.checkName).getValue() !== REG_HEAD[C.checkName - 1])   // 舊的試算表：補上（或改）新欄位的標題
    f.sh.getRange(1, C.method, 1, 7).setValues([REG_HEAD.slice(C.method - 1, C.method + 6)]).setFontWeight("bold");
  f.sh.getRange(f.row, C.checkNo).setNumberFormat("@");
  f.sh.getRange(f.row, C.method, 1, 7).setValues([[method, ckName, ckNo, ckAmt, ckPart, paid && !check ? amount : cashPart, note]]);
  const entry = { action: edit ? "editpay" : paid ? "paid" : "unpaid", by: me.name, byId: me.id, at: now.toISOString(), amount: amount };
  if (edit) entry.from = payDesc_(f.rec);
  if (paid) entry.method = !check ? "cash" : cashPart > 0 ? "mix" : "check";
  if (check) { entry.checkName = ckName; entry.checkNo = ckNo; entry.checkAmt = ckAmt; entry.checkPart = ckPart; entry.cashPart = cashPart; }
  addLog_(f, entry);
  // 收款時義工記下支票多出來的錢的用途（寫到這張支票付過的每一筆）
  const extra = check ? String(pay.checkNote || "").trim().slice(0, 100) : "";
  if (extra && extra !== note) {
    const all = f.sh.getRange(2, 1, f.sh.getLastRow() - 1, REG_HEAD.length).getValues();
    all.forEach((r, i) => { if (String(r[C.checkNo - 1]).trim() === ckNo && ckKey_(ckNo, r[C.checkName - 1]) === ckKey_(ckNo, ckName)) f.sh.getRange(i + 2, C.checkNote).setValue(extra); });
    audit_(me, "支票註記", "#" + ckNo + " " + ckName, extra);
  }
  let rec = readRecord_(f.sh.getRange(f.row, 1, 1, REG_HEAD.length).getValues()[0]);
  if (edit) {   // 處理紀錄補上「改成什麼」
    const log = rec.payLog || []; const last = log[log.length - 1]; if (last && last.action === "editpay") { last.to = payDesc_(rec); f.sh.getRange(f.row, C.log).setValue(JSON.stringify(log)); }
    return { ok: true, record: { payMethod: rec.payMethod, checkName: rec.checkName, checkNo: rec.checkNo, checkAmt: rec.checkAmt, checkPart: rec.checkPart, cashPart: rec.cashPart, checkNote: rec.checkNote, payLog: log } };
  }
  return { ok: true, record: { paid: paid, paidAt: paid ? now.toISOString() : null, paidBy: paid ? me.name : null, paidById: paid ? me.id : null,
    payMethod: rec.payMethod, checkName: rec.checkName, checkNo: rec.checkNo, checkAmt: rec.checkAmt, checkPart: rec.checkPart, cashPart: rec.cashPart, checkNote: rec.checkNote, payLog: rec.payLog } };
}
// 支票多出來的錢：在支票明細註記用途（寫到這張支票付過的每一筆）
function setCheckNote_(me, no, name, note) {
  no = String(no || "").trim(); note = String(note || "").trim().slice(0, 100);
  const sh = sheet_(SHEET.reg), n = sh.getLastRow() - 1; if (n < 1 || !no) return { ok: false, error: "not-found" };
  const rows = sh.getRange(2, 1, n, REG_HEAD.length).getValues(); let hit = 0;
  rows.forEach((r, i) => { if (String(r[C.checkNo - 1]).trim() === no && ckKey_(no, r[C.checkName - 1]) === ckKey_(no, name)) { sh.getRange(i + 2, C.checkNote).setValue(note); hit++; } });
  if (!hit) return { ok: false, error: "not-found" };
  audit_(me, "支票註記", "#" + no + " " + name, note || "（清除）");
  return { ok: true };
}
// 發點燈卡：只有已繳清的點燈報名可以標記；紀錄在處理紀錄裡
function setCard_(me, ref, given) {
  const f = findRow_(ref);
  if (!f) return { ok: false, error: "not-found" };
  if (given && (f.rec.cancelled || !f.rec.paid || f.rec.form !== "lamp")) return { ok: false, error: "unpaid" };
  const at = new Date().toISOString();
  addLog_(f, { action: given ? "card" : "uncard", by: me.name, byId: me.id, at: at });
  return { ok: true, record: { cardAt: given ? at : null, cardBy: given ? me.name : null, payLog: f.rec.payLog } };
}
// 取消報名會把燈位放出來；恢復時重新分配（原位還空著就用原位）
function setCancelled_(me, ref, cancelled) {
  const f = findRow_(ref);
  if (!f) return { ok: false, error: "not-found" };
  if (cancelled && f.rec.paid) return { ok: false, error: "paid" };
  const d = f.rec;
  if (!cancelled && d.form === "lamp") { assignSeats_(d, usedSeats_(d.ref, recYear_(d)), d); writeData_(f, d); }
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
  assignSeats_(d, usedSeats_(d.ref, recYear_(d)), d);
  writeData_(f, d);
  addLog_(f, { action: "seat", by: me.name, byId: me.id, at: new Date().toISOString() });
  return { ok: true, record: d };
}

/* ================= 設定 ================= */
function getSettings_() {
  try { const v = sheet_(SHEET.settings).getRange("A1").getValue(); if (v) { const s = JSON.parse(v); if (s && s.events) { s.lamp = s.lamp || { status: "open" }; if (!s.lamp.year) s.lamp.year = DEFAULT_LAMP_YEAR; return s; } } } catch (e) {}
  return JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
}
function saveSettings_(s, me) {
  if (!s || !s.lamp || !Array.isArray(s.events)) return { ok: false, error: "bad-data" };
  const old = getSettings_();
  const y = Number(s.lamp.year);
  if (!(y >= 2020 && y <= 2100)) return { ok: false, error: "year" };
  s.lamp.year = y;
  // 每年建議點太歲燈的生肖（後台勾選；沒設定就用預設規則）：只留合法的生肖
  if (s.lamp.taisui && typeof s.lamp.taisui === "object") {
    const Z = "鼠牛虎兔龍蛇馬羊猴雞狗豬", t = {};
    Object.keys(s.lamp.taisui).forEach(k => { const v = s.lamp.taisui[k]; if (/^\d{4}$/.test(k) && Array.isArray(v)) t[k] = v.filter(z => Z.indexOf(z) >= 0 && String(z).length === 1); });
    s.lamp.taisui = t;
  } else delete s.lamp.taisui;
  delete s.seatsLeft;
  s.updatedAt = new Date().toISOString();
  sheet_(SHEET.settings).getRange("A1").setValue(JSON.stringify(s));
  if (me) { const d = settingsDiff_(old, s); if (d) audit_(me, "修改報名設定", "", d); }
  return { ok: true };
}

// 設定改了什麼（寫進紀錄用）
function settingsDiff_(a, b) {
  const ST = { open: "開放報名", paused: "暫停", hidden: "不顯示" }, out = [];
  const al = a.lamp || {}, bl = b.lamp || {};
  if (al.status !== bl.status) out.push("點燈報名：" + (ST[al.status] || al.status || "—") + " → " + (ST[bl.status] || bl.status));
  if (Number(al.year) !== Number(bl.year)) out.push("點燈年度：" + (al.year || "—") + " → " + bl.year);
  const at = JSON.stringify((al.taisui || {})[bl.year] || null), bt = JSON.stringify((bl.taisui || {})[bl.year] || null);
  if (at !== bt) out.push(bl.year + " 太歲生肖：" + (((bl.taisui || {})[bl.year]) ? bl.taisui[bl.year].join("") : "恢復預設"));
  const ae = {}; (a.events || []).forEach(e => ae[e.id] = e);
  (b.events || []).forEach(e => { const o = ae[e.id];
    if (!o) out.push("新增法會「" + e.name + "」（" + (ST[e.status] || e.status) + "）");
    else { if (o.name !== e.name) out.push("法會改名：" + o.name + " → " + e.name);
           if (o.status !== e.status) out.push("「" + e.name + "」：" + (ST[o.status] || o.status) + " → " + (ST[e.status] || e.status));
           if ((o.date || "") !== (e.date || "")) out.push("「" + e.name + "」日期：" + (e.date || "（空白）")); }
    delete ae[e.id]; });
  Object.keys(ae).forEach(id => out.push("移除法會「" + ae[id].name + "」"));
  return out.join("；");
}

/* ================= 帳號 ================= */
function readUsers_() {
  const sh = sheet_(SHEET.users), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, USER_HEAD.length).getValues().map((r, i) => ({
    row: i + 2, id: String(r[0]), name: String(r[1]), role: ROLES.includes(String(r[2])) ? String(r[2]) : "volunteer",
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
    audit_(me, "改自己的 PIN", me.name, "");
    return { ok: true };
  }
  if (!canAccounts_(me)) return { ok: false, error: "auth" };
  // 網站管理員只能管義工；指定堂主、財務、網站管理員要堂主（或管理員）
  const roleOK = role => ROLES.includes(role) && (role === "volunteer" || canAssignRoles_(me)) && (role !== "admin" || me.role === "admin")
    && (!["owner", "deputy"].includes(role) || ["admin", "owner"].includes(me.role));
  // 堂主只能有一位。現任堂主把堂主交給別人時，自己改成義工（由新堂主重新指定）；其他人要換堂主，得先把原堂主改成別的身分
  const ownerNow = exceptId => list.find(x => x.role === "owner" && x.active && x.id !== exceptId);
  const deputyNow = exceptId => list.find(x => x.role === "deputy" && x.active && x.id !== exceptId);
  const handOver = (toName) => { meU.role = "volunteer"; writeUser_(meU); audit_(me, "移交堂主", toName, me.name + " 改為義工"); };
  if (b.op === "add") {
    const name = String(b.name || "").trim(), role = String(b.role || "volunteer");
    if (!name || !pinOK(b.newPin)) return { ok: false, error: "bad-data" };
    if (!roleOK(role)) return { ok: false, error: "role-auth" };
    if (list.some(x => x.name === name)) return { ok: false, error: "dup" };
    const cur = role === "owner" ? ownerNow("") : null;
    if (cur && cur.id !== me.id) return { ok: false, error: "one-owner", owner: cur.name };
    if (role === "deputy" && deputyNow("")) return { ok: false, error: "one-deputy" };
    writeUser_({ id: "u-" + Date.now().toString(36), name: name, role: role, pin: hashPin_(b.newPin), active: true, mustChange: true });
    audit_(me, "新增帳號", name, ROLE_ZH[role]);
    if (cur) { handOver(name); return { ok: true, handedOver: true }; }
    return { ok: true };
  }
  const u = list.find(x => x.id === b.id);
  if (!u) return { ok: false, error: "not-found" };
  if (u.role !== "volunteer" && !canAssignRoles_(me) && u.id !== me.id) return { ok: false, error: "role-auth" };   // 網站管理員不能動其他身分的帳號
  if (["owner", "deputy", "admin"].includes(u.role) && me.role === "deputy" && u.id !== me.id) return { ok: false, error: "role-auth" };   // 副堂主不能動堂主的帳號
  if (b.op === "resetPin") {
    if (!pinOK(b.newPin)) return { ok: false, error: "pin" };
    u.pin = hashPin_(b.newPin); u.mustChange = true; writeUser_(u);
    CacheService.getScriptCache().remove("fail-" + u.id);   // 重設 PIN 時順便解鎖
    audit_(me, "重設 PIN", u.name, "");
    return { ok: true };
  }
  // 最後一位（還能登入的）堂主／管理員不能被停用、改身分或刪除，免得沒有人能管
  const lastAdmin = isBoss_(u) && list.filter(isBoss_).length <= 1;
  if (b.op === "setActive") {
    if (u.id === me.id) return { ok: false, error: "self" };
    if (!b.active && lastAdmin) return { ok: false, error: "last-admin" };
    if (b.active && u.role === "owner" && ownerNow(u.id)) return { ok: false, error: "one-owner", owner: ownerNow(u.id).name };
    if (b.active && u.role === "deputy" && deputyNow(u.id)) return { ok: false, error: "one-deputy" };
    u.active = !!b.active; writeUser_(u);
    audit_(me, u.active ? "重新啟用帳號" : "停用帳號", u.name, "");
    return { ok: true };
  }
  if (b.op === "edit") {     // 改名字、改身分
    const name = String(b.name || "").trim();
    const role = ROLES.includes(String(b.role)) ? String(b.role) : u.role;
    if (!name) return { ok: false, error: "bad-data" };
    if (list.some(x => x.id !== u.id && x.name === name)) return { ok: false, error: "dup" };
    const changes = [];
    if (role !== u.role) {
      if (u.id === me.id) return { ok: false, error: "self-role" };
      if (!roleOK(role)) return { ok: false, error: "role-auth" };
      if (lastAdmin && role !== "admin" && role !== "owner") return { ok: false, error: "last-admin" };
      changes.push("身分：" + ROLE_ZH[u.role] + " → " + ROLE_ZH[role]);
    }
    const cur = role === "owner" && u.role !== "owner" ? ownerNow(u.id) : null;
    if (cur && cur.id !== me.id) return { ok: false, error: "one-owner", owner: cur.name };
    if (role === "deputy" && u.role !== "deputy" && deputyNow(u.id)) return { ok: false, error: "one-deputy" };
    if (name !== u.name) changes.push("姓名：" + u.name + " → " + name);
    const oldName = u.name;
    u.name = name; u.role = role; writeUser_(u);
    if (changes.length) audit_(me, "修改帳號", oldName, changes.join("；"));
    if (cur) { handOver(u.name); return { ok: true, user: pub_(u), handedOver: true }; }
    return { ok: true, user: pub_(u) };
  }
  if (b.op === "delete") {   // 只能刪沒經手過任何報名的帳號（建錯的）；經手過的請用停用
    if (u.id === me.id) return { ok: false, error: "self" };
    if (lastAdmin) return { ok: false, error: "last-admin" };
    if (userUsed_(u)) return { ok: false, error: "used" };
    sheet_(SHEET.users).deleteRow(u.row);
    CacheService.getScriptCache().remove("fail-" + u.id);
    audit_(me, "刪除帳號", u.name, ROLE_ZH[u.role]);
    return { ok: true };
  }
  return { ok: false, error: "unknown-op" };
}
// 這個帳號有沒有經手過報名（收款、取消、排燈位）
function userUsed_(u) {
  const sh = sheet_(SHEET.reg), n = sh.getLastRow() - 1;
  if (n < 1) return false;
  const paidBy = sh.getRange(2, C.paidBy, n, 1).getValues(), logs = sh.getRange(2, C.log, n, 1).getValues();
  for (let i = 0; i < n; i++) {
    if (String(paidBy[i][0]) === u.name) return true;
    let log = []; try { log = JSON.parse(logs[i][0] || "[]"); } catch (e) {}
    if (log.some(x => x.byId === u.id || (!x.byId && x.by === u.name))) return true;
  }
  return false;
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
    [["u-wang", "王師姐", "volunteer", "1234"], ["u-li", "李師兄", "volunteer", "5678"], ["u-head", "堂主", "owner", "6666"], ["u-dep", "副堂主", "deputy", "7777"], ["u-chen", "陳師姐", "finance", "2468"], ["u-lin", "林網管", "web", "1357"]].forEach(([id, n, role, pin]) => users.appendRow([id, n, role, hashPin_(pin), true, false]));
    const B = (y, m, d, h) => ({ lunar: { year: y, month: m, leap: false, day: d, hour: h || "" }, solar: null, entered: "lunar" });
    const samples = [
      { contact: { name: "王秀英", phone: "(718) 555-0142", address: "範例地址" }, lamps: [
        { pid: "s1", name: "王秀英", gender: "女", birth: B(1958, 8, 12, "辰"), zodiac: "狗", age_xu: 70, lampTypes: ["光明燈", "太歲燈"] },
        { pid: "s2", name: "李建國", gender: "男", birth: B(1955, 3, 3), zodiac: "羊", age_xu: 73, lampTypes: ["光明燈"] }] },
      { contact: { name: "張雅婷", phone: "(347) 555-0198", address: "範例地址" }, lamps: [
        { pid: "s3", name: "張雅婷", gender: "女", birth: B(1986, 11, 2, "午"), zodiac: "虎", age_xu: 42, lampTypes: ["財神燈"] },
        { pid: "s4", name: "陳柏宇", gender: "男", birth: B(2012, 6, 18), zodiac: "龍", age_xu: 16, lampTypes: ["文昌燈", "光明燈"] },
        { pid: "s5", name: "陳品妍", gender: "女", birth: B(2015, 1, 9), zodiac: "羊", age_xu: 13, lampTypes: ["文昌燈"] }] },
      // 未繳費（測試收款用）：黃家兩張報名表，可以試一張支票付兩筆
      { contact: { name: "黃志明", phone: "(718) 555-0263", address: "範例地址" }, lamps: [
        { pid: "s6", name: "黃志明", gender: "男", birth: B(1969, 4, 21, "巳"), zodiac: "雞", age_xu: 59, lampTypes: ["光明燈", "財神燈"] },
        { pid: "s7", name: "林淑芬", gender: "女", birth: B(1972, 9, 5), zodiac: "鼠", age_xu: 56, lampTypes: ["光明燈"] }] },
      { contact: { name: "黃美華", phone: "(917) 555-0377", address: "範例地址" }, lamps: [
        { pid: "s8", name: "黃美華", gender: "女", birth: B(1991, 7, 15), zodiac: "羊", age_xu: 37, lampTypes: ["太歲燈", "光明燈"] }] },
    ];
    const refs = samples.map(s => submit_(Object.assign({ form: "lamp", eventId: null, formTitle: "啟點光明燈暨安奉太歲", note: "", lang: "zh", chaodu: [], meals: [] }, s)).ref);
    // 標成範例，第一筆設為已繳清（現金），第二筆用支票，第三、四筆未繳費
    const sh = sheet_(SHEET.reg);
    for (let r = 2; r <= sh.getLastRow(); r++) { const d = JSON.parse(sh.getRange(r, C.json).getValue()); d.sample = true; sh.getRange(r, C.json).setValue(JSON.stringify(d)); }
    setPaid_({ id: "u-admin", name: "管理員" }, refs[0], true);
    setPaid_({ id: "u-admin", name: "管理員" }, refs[1], true, { method: "check", checkName: "CHANG YA TING", checkNo: "1052", checkAmount: 480 });
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
