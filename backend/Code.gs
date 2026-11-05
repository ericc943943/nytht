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
// 燈位：每種燈牆上的排列。
// 燈號 = 英文字母＋燈種代碼（給燈號查詢機用）＋排＋第幾個，例：G010512 ＝ 光明燈第 5 排第 12 個。
// 字母：G 光明、W 文昌、T 太歲、C 財神。單據上燈號前面都會寫燈的名稱。
const SEATS = {
  "個人光明燈": { letter: "G", label: "G", code: "01", perRow: 30, rows: 60 },
  "文昌燈":     { letter: "W", label: "W", code: "02", perRow: 20, rows: 60 },
  "太歲燈":     { letter: "T", label: "T", code: "03", perRow: 30, rows: 39 },
  "五路財神燈": { letter: "C", label: "C", code: "05", perRow: 30, rows: 70 },
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

const SHEET = { reg: "報名", users: "帳號", settings: "設定" };
const REG_HEAD = ["報名編號", "報名時間", "表單", "場次", "聯絡人", "電話", "住址", "應繳", "已繳清", "收款時間", "經手人", "語言", "備註", "收款紀錄", "送出代碼", "完整資料", "已取消", "燈位", "點燈年度"];
const C = { ref: 1, time: 2, form: 3, event: 4, name: 5, phone: 6, address: 7, total: 8, paid: 9, paidAt: 10, paidBy: 11, lang: 12, note: 13, log: 14, cid: 15, json: 16, cancelled: 17, seats: 18, year: 19 };
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
      case "settings": { const s = getSettings_(); s.lamp.year = lampYear_(s); s.seatsLeft = seatsLeft_(s.lamp.year); return out_(s); }
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
    let d = {}; try { d = JSON.parse(r[C.json - 1]); } catch (e) { return; }
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
const SHORT_ = { "個人光明燈": "光明燈", "五路財神燈": "財神燈" };
function seatText_(d) {
  return (d.lamps || []).filter(p => p.seats && Object.keys(p.seats).length)
    .map(p => p.name + "：" + Object.keys(p.seats).map(t => (SHORT_[t] || t) + " " + (seatLabel_(p.seats[t]) || "候補")).join("、")).join("\n");
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
  ["ref", "form", "eventId", "formTitle", "submittedAt", "clientId", "lang", "lampYear"].forEach(k => { d[k] = old[k]; });
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
function saveSettings_(s) {
  if (!s || !s.lamp || !Array.isArray(s.events)) return { ok: false, error: "bad-data" };
  const y = Number(s.lamp.year);
  if (!(y >= 2020 && y <= 2100)) return { ok: false, error: "year" };
  s.lamp.year = y;
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
