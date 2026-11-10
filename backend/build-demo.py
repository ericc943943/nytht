#!/usr/bin/env python3
"""用 backend/Code.gs 產生網站根目錄的 demo-backend.js（示範模式在瀏覽器裡跑同一套後端規則）。
用法：python3 backend/build-demo.py
"""
import pathlib

root = pathlib.Path(__file__).resolve().parent.parent
code = (root / "backend" / "Code.gs").read_text(encoding="utf-8")

shell = r'''/* 自動產生，請勿直接修改：示範模式用的後端。
 * 內容是 backend/Code.gs 原封不動，外面包一層「假的 Google 試算表」，資料存在這台電腦的瀏覽器裡。
 * 改了 Code.gs 之後執行 python3 backend/build-demo.py 重新產生。 */
window.DemoBackend = (function () {
  "use strict";
  const KEY = "tsuhui-demo-v9";
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
  // 指令碼屬性：存在示範資料裡（重新整理後還在）
  const PropertiesService = { getScriptProperties: () => ({
    getProperty: k => ((db.props || {})[k] ?? null), setProperty: (k, v) => { (db.props = db.props || {})[k] = String(v); }, deleteProperty: k => { if (db.props) delete db.props[k]; } }) };
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
/*__CODE__*/
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
    setPaid_({ id: "u-admin", name: "管理員" }, refs[1], true, { method: "check", checkName: "CHANG YA TING", checkNo: "1052", checkAmount: 600 });
    // 示範版的核對碼固定：王秀英 1111、張雅婷 2222、黃志明 3333、黃美華 4444（張雅婷那張支票還有 $120 餘額，可以試核對碼）
    for (let r = 2; r <= sh.getLastRow(); r++) sh.getRange(r, C.code).setValue(String(r - 1).repeat(4));
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
'''

out = shell.replace("/*__CODE__*/", code)
(root / "demo-backend.js").write_text(out, encoding="utf-8")
print("wrote", root / "demo-backend.js", len(out), "bytes")
