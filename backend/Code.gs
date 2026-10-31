/**
 * 紐約花蓮慈惠堂｜線上點燈和法會報名系統｜後端（Google Apps Script）
 *
 * 用法：在堂裡 Google 帳號建一份空白試算表 →「擴充功能」→「Apps Script」
 *      → 把這整份貼上 → 先執行一次 setup → 部署成「網頁應用程式」。
 * 詳細步驟見 backend/設定說明.txt。
 *
 * 試算表分頁（setup 會自動建立）：
 *   報名  一筆報名一列；最後一欄「完整資料」是 JSON，後台靠它顯示，請勿手動修改
 *   帳號  義工帳號；PIN 只存加密後的值
 *   設定  報名項目開關（A1 的 JSON），由後台「設定」修改
 */

/* ================= 設定 ================= */
const PRICES = {
  lamp: { "個人光明燈": 120, "太歲燈": 120, "五路財神燈": 120, "文昌燈": 120 },  // 點燈：每盞
  lotus: 60,     // 法會：蓮花福慧燈，每位
  chaodu: 160,   // 法會：超渡牌位，每牌位
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
const MAX_FAILS = 8;          // PIN 連續打錯幾次就暫時鎖住
const LOCK_MINUTES = 15;

const SHEET = { reg: "報名", users: "帳號", settings: "設定" };
const REG_HEAD = ["報名編號", "報名時間", "表單", "場次", "聯絡人", "電話", "住址", "應繳", "已繳清", "收款時間", "經手人", "語言", "備註", "收款紀錄", "送出代碼", "完整資料"];
const USER_HEAD = ["id", "姓名", "身分", "PIN（加密）", "啟用", "待改PIN"];

/* ================= 第一次使用：建立分頁與管理員帳號 ================= */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const make = (name, head) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (head && sh.getLastRow() === 0) { sh.appendRow(head); sh.setFrozenRows(1); sh.getRange(1, 1, 1, head.length).setFontWeight("bold"); }
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
      case "settings": return out_(getSettings_());
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
    if (!body.action) return out_(submit_(body));          // 報名頁送出
    const me = checkLogin_(body.uid, body.pin);
    if (!me) return out_({ ok: false, error: "auth" });
    switch (body.action) {
      case "setPaid":      return out_(setPaid_(me, body.ref, !!body.paid));
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
  // 同一份報名重送（網路斷掉後再按一次）：回傳原本的編號，不重複建立
  if (d.clientId && last > 1) {
    const ids = sh.getRange(2, 15, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (ids[i][0] === d.clientId) {
      return { ok: true, ref: String(sh.getRange(i + 2, 1).getValue()), duplicate: true };
    }
  }
  // 報名編號由這裡決定，確保不重複（格式同報名頁：L/F + yymmdd + - + 4 碼）
  const existing = new Set(last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : []);
  const tz = Session.getScriptTimeZone() || "America/New_York";
  const prefix = (d.form === "lamp" ? "L" : "F") + Utilities.formatDate(new Date(), tz, "yyMMdd") + "-";
  let ref = d.ref && String(d.ref).indexOf(prefix) === 0 && !existing.has(d.ref) ? String(d.ref) : "";
  for (let i = 0; !ref && i < 200; i++) { const c = prefix + ("000" + Math.floor(Math.random() * 10000)).slice(-4); if (!existing.has(c)) ref = c; }
  if (!ref) return { ok: false, error: "ref" };

  // 金額一律在這裡重新計算，不採用網頁送來的數字
  d.totals = calcTotals_(d);
  d.ref = ref; d.paid = false; d.submittedAt = new Date().toISOString();
  const evName = d.form === "lamp" ? "點燈" : ((st.events.find(x => x.id === d.eventId) || {}).name || d.formTitle || "法會");
  sh.appendRow([ref, new Date(), d.form === "lamp" ? "點燈" : "法會", evName, d.contact.name || "", d.contact.phone || "", d.contact.address || "",
                d.totals.all, false, "", "", d.lang === "en" ? "英文" : "中文", d.note || "", "[]", d.clientId || "", JSON.stringify(d)]);
  return { ok: true, ref: ref, totals: d.totals };
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

function listRows_() {
  const sh = sheet_(SHEET.reg), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, REG_HEAD.length).getValues().map(r => {
    let d = {}; try { d = JSON.parse(r[15]); } catch (e) {}
    let log = []; try { log = JSON.parse(r[13] || "[]"); } catch (e) {}
    d.ref = String(r[0]);
    d.paid = r[8] === true;
    d.paidAt = r[9] ? new Date(r[9]).toISOString() : null;
    d.paidBy = r[10] || null;
    d.payLog = log;
    return d;
  });
}

function setPaid_(me, ref, paid) {
  const sh = sheet_(SHEET.reg), n = sh.getLastRow() - 1;
  if (n < 1) return { ok: false, error: "not-found" };
  const refs = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (String(refs[i][0]) === String(ref)) {
    const row = i + 2, now = new Date();
    let log = []; try { log = JSON.parse(sh.getRange(row, 14).getValue() || "[]"); } catch (e) {}
    const amount = sh.getRange(row, 8).getValue();
    log.push({ action: paid ? "paid" : "unpaid", by: me.name, byId: me.id, at: now.toISOString(), amount: amount });
    sh.getRange(row, 9, 1, 3).setValues([[paid, paid ? now : "", paid ? me.name : ""]]);
    sh.getRange(row, 14).setValue(JSON.stringify(log));
    return { ok: true, record: { paid: paid, paidAt: paid ? now.toISOString() : null, paidBy: paid ? me.name : null, paidById: paid ? me.id : null, payLog: log } };
  }
  return { ok: false, error: "not-found" };
}

/* ================= 設定 ================= */
function getSettings_() {
  try { const v = sheet_(SHEET.settings).getRange("A1").getValue(); if (v) { const s = JSON.parse(v); if (s && s.events) return s; } } catch (e) {}
  return DEFAULT_SETTINGS;
}
function saveSettings_(s) {
  if (!s || !s.lamp || !Array.isArray(s.events)) return { ok: false, error: "bad-data" };
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
