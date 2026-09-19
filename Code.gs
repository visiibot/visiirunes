/**
 * THE RUNIC SYSTEM — Google Apps Script backend
 * ------------------------------------------------
 * Deploy this bound to a Google Sheet. It stores every student's rune
 * inventory, crafted items, one-time milestone PIN log, and now each
 * student's personal 4-digit account PIN, serving it all to the
 * classroom web app over HTTPS so every device stays in sync.
 *
 * SETUP (one time):
 *  1. Create a new Google Sheet. Extensions > Apps Script.
 *  2. Delete the placeholder code and paste this whole file in.
 *  3. In the function dropdown at the top, choose "setup" and click Run.
 *     The first run will ask you to authorize it — allow it.
 *     This creates the Students / Pins / Config sheets with headers.
 *     (Safe to re-run "setup" any time you update this script — it only
 *     adds missing columns/sheets, it never deletes data.)
 *  4. Deploy > New deployment > gear icon > "Web app".
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     Click Deploy, authorize again if asked, then copy the Web App URL
 *     (it ends in /exec). Paste that URL into the Runic System app's
 *     backend settings (the gear icon in the header).
 *  5. Whenever you edit this script, make a NEW deployment (or "Manage
 *     deployments" > edit > new version) for changes to take effect.
 *
 * UPGRADING an existing deployment to add per-adventurer PINs:
 *  1. Replace the old Code.gs contents with this file.
 *  2. Run "setup" once more (adds the new PIN columns to your existing
 *     Students sheet without touching anyone's rune counts).
 *  3. Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

/* ---------------- Game data (mirrors the frontend) ---------------- */

const RUNES = [
  { code: "TIR", prob: 14, points: 2 },
  { code: "ITH", prob: 14, points: 2 },
  { code: "NEF", prob: 12, points: 5 },
  { code: "TAL", prob: 12, points: 5 },
  { code: "ORT", prob: 10, points: 9 },
  { code: "THUL", prob: 10, points: 9 },
  { code: "SHAEL", prob: 8, points: 14 },
  { code: "IO", prob: 8, points: 14 },
  { code: "LUM", prob: 4, points: 20 },
  { code: "JAH", prob: 4, points: 20 },
  { code: "BER", prob: 2, points: 27 },
  { code: "ZOD", prob: 2, points: 27 },
];
const RUNE_CODES = RUNES.map(function (r) { return r.code; });
const ALL_INV_CODES = RUNE_CODES.concat(["GOD"]);
const RUNE_POINTS = {};
RUNES.forEach(function (r) { RUNE_POINTS[r.code] = r.points; });

const RUNEWORDS = [
  { id: "enigma", name: "Enigma", recipe: ["BER", "NEF", "THUL", "ORT"] },
  { id: "cling", name: "Cling", recipe: ["IO", "NEF", "ITH"] },
  { id: "prism", name: "Prism", recipe: ["TIR", "NEF", "TAL"] },
  { id: "rizz", name: "Silver Tongue of Rizz", recipe: ["SHAEL", "IO"] },
  { id: "quicksilver", name: "Quicksilver", recipe: ["LUM", "ORT", "ITH"] },
  { id: "charger", name: "Charger", recipe: ["BER", "THUL", "TAL", "SHAEL"] },
  { id: "precious", name: "Precious", recipe: ["ZOD", "LUM", "THUL"] },
  { id: "hive", name: "Hive", recipe: ["BER", "LUM", "ORT", "ITH"] },
  { id: "starboard", name: "Starboard", recipe: ["IO", "NEF", "TIR"] },
  { id: "carotene", name: "Blade of Carotene", recipe: ["IO", "NEF", "TIR"] },
  { id: "folly", name: "Folly", recipe: ["ZOD", "SHAEL", "ORT"] },
  { id: "insolence", name: "Insolence", recipe: ["LUM", "SHAEL", "ORT", "TAL"] },
  { id: "sage", name: "Sage", recipe: ["TIR", "NEF", "ORT", "SHAEL"] },
  { id: "lark", name: "Lark", recipe: ["ITH", "TAL", "THUL", "IO"] },
  { id: "inferno", name: "Inferno", recipe: ["ZOD", "IO", "THUL"] },
  { id: "socks", name: "Socks of Blocks", recipe: ["ZOD", "JAH", "IO", "TAL"] },
  { id: "north", name: "Invigorated by the North", recipe: ["LUM", "SHAEL", "TAL"] },
  { id: "bramble", name: "Bramble", recipe: ["SHAEL", "ORT", "NEF", "TIR"] },
  { id: "crescent", name: "Crescent Moon", recipe: ["BER", "SHAEL", "TAL", "ITH"] },
  { id: "harmony", name: "Harmony", recipe: ["LUM", "IO", "TAL", "TIR"] },
  { id: "summarmaekir", name: "Summarmækir", recipe: ["GOD", "ZOD", "JAH", "IO"] },
  { id: "ransnet", name: "Ran's Net", recipe: ["GOD", "BER", "LUM", "SHAEL"] },
  { id: "herofeller", name: "Herofeller", recipe: ["TIR", "TAL", "ORT", "IO"] },
  { id: "coolth", name: "Coolth", recipe: ["TIR", "THUL", "SHAEL", "LUM"] },
  { id: "misfortune", name: "Amulet of Misfortune", recipe: ["TIR", "SHAEL", "THUL", "NEF"] },
  { id: "stuffed", name: "Stuffed Buddy", recipe: ["ITH", "NEF", "THUL", "IO"] },
];
const SOUR_KEY_BASE = ["SHAEL", "IO", "JAH", "BER"];

const PIN_LOCKOUT_MS = 60000; // 60s lockout after too many wrong PIN attempts
const PIN_MAX_FAILS = 5;

/* ---------------- Setup ---------------- */

function setup() {
  var studentHeaders = ["Name"].concat(ALL_INV_CODES).concat([
    "Crafted", "PinHash", "PinSalt", "PinFailCount", "PinLockUntil"
  ]);
  var sh = ensureSheet_("Students", studentHeaders);
  ensureColumns_(sh, studentHeaders);
  ensureSheet_("Pins", ["Code", "Used", "UsedBy", "CreatedAt", "UsedAt"]);
  ensureSheet_("Config", ["Key", "Value"]);
}

function ensureSheet_(name, headers) {
  var sh = SS.getSheetByName(name);
  if (!sh) sh = SS.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}

function ensureColumns_(sh, headers) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var existing = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var toAdd = headers.filter(function (h) { return existing.indexOf(h) === -1; });
  if (toAdd.length) {
    sh.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd]);
  }
}

/* ---------------- HTTP entry points ---------------- */

function doGet(e) {
  return jsonOut_({ ok: true, data: "The Runic System API is running." });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var data;
    switch (action) {
      case "checkName": data = apiCheckName(body.name); break;
      case "claimPin": data = apiClaimPin(body.name, body.authPin); break;
      case "verifyPin": data = apiVerifyPin(body.name, body.authPin); break;
      case "getStudent": data = apiGetStudent(body.name); break;
      case "roll": data = apiRoll(body.name, body.code, body.authPin); break;
      case "transform": data = apiTransform(body.name, body.target, body.selection, body.authPin); break;
      case "craft": data = apiCraft(body.name, body.itemId, body.authPin); break;
      case "craftSour": data = apiCraftSour(body.name, body.mult, body.authPin); break;
      case "teacherSetup": data = apiTeacherSetup(body.pass); break;
      case "teacherLogin": data = apiTeacherLogin(body.pass); break;
      case "teacherGeneratePin": data = apiTeacherGeneratePin(body.pass); break;
      case "teacherGrantGod": data = apiTeacherGrantGod(body.pass, body.name); break;
      case "teacherListPins": data = apiTeacherListPins(body.pass); break;
      case "teacherListStudents": data = apiTeacherListStudents(body.pass); break;
      case "teacherResetPin": data = apiTeacherResetPin(body.pass, body.name); break;
      case "teacherReset": data = apiTeacherReset(body.pass); break;
      default: throw new Error("Unknown action: " + action);
    }
    return jsonOut_({ ok: true, data: data });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err.message || err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- Student row helpers ---------------- */

function studentsSheet_() { return SS.getSheetByName("Students"); }

function requireName_(name) {
  if (!name || !String(name).trim()) throw new Error("Missing student name.");
}

function findStudentRowIndex_(sh, name) {
  var values = sh.getDataRange().getValues();
  var key = String(name).trim().toLowerCase();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === key) return i + 1;
  }
  return -1;
}

function ensureStudentRow_(sh, name) {
  var row = findStudentRowIndex_(sh, name);
  if (row === -1) {
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var newRow = headers.map(function (h) {
      if (h === "Name") return String(name).trim();
      if (h === "Crafted") return "[]";
      if (h === "PinHash" || h === "PinSalt" || h === "PinLockUntil") return "";
      if (h === "PinFailCount") return 0;
      return 0; // rune counts + GOD
    });
    sh.appendRow(newRow);
    row = sh.getLastRow();
  }
  return row;
}

// Public view of a student — never includes PinHash/PinSalt.
function readStudentPublic_(sh, row) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var values = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  var obj = {};
  headers.forEach(function (h, i) { obj[h] = values[i]; });
  var inventory = {};
  ALL_INV_CODES.forEach(function (c) { inventory[c] = Number(obj[c]) || 0; });
  var crafted = [];
  try { crafted = JSON.parse(obj.Crafted || "[]"); } catch (e) { crafted = []; }
  return { name: obj.Name, inventory: inventory, crafted: crafted, hasPin: !!obj.PinHash };
}

// Internal-only auth fields — never returned to the client.
function readStudentAuth_(sh, row) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var values = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  var obj = {};
  headers.forEach(function (h, i) { obj[h] = values[i]; });
  return {
    hash: obj.PinHash || "",
    salt: obj.PinSalt || "",
    failCount: Number(obj.PinFailCount) || 0,
    lockUntil: obj.PinLockUntil || ""
  };
}

function writeInventoryDelta_(sh, row, deltas) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Object.keys(deltas).forEach(function (code) {
    var col = headers.indexOf(code) + 1;
    if (col > 0) {
      var cell = sh.getRange(row, col);
      var cur = Number(cell.getValue()) || 0;
      cell.setValue(cur + deltas[code]);
    }
  });
}

function writeFields_(sh, row, fields) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Object.keys(fields).forEach(function (key) {
    var col = headers.indexOf(key) + 1;
    if (col > 0) sh.getRange(row, col).setValue(fields[key]);
  });
}

function appendCrafted_(sh, row, entry) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = headers.indexOf("Crafted") + 1;
  var cell = sh.getRange(row, col);
  var list = [];
  try { list = JSON.parse(cell.getValue() || "[]"); } catch (e) {}
  list.push(entry);
  cell.setValue(JSON.stringify(list));
}

/* ---------------- Per-adventurer PIN auth ---------------- */

function validateAuthPin_(pin) {
  if (!/^\d{4}$/.test(String(pin || ""))) throw new Error("PINs are four digits.");
}

function hashPin_(pin, salt) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ":" + pin, Utilities.Charset.UTF_8);
  return bytes.map(function (b) {
    var v = b < 0 ? b + 256 : b;
    var s = v.toString(16);
    return s.length === 1 ? "0" + s : s;
  }).join("");
}

// Verifies pin for the student at `row`, handling lockout/fail-count.
// Throws on any failure (wrong pin, locked out, no pin set yet).
function checkPin_(sh, row, pin) {
  validateAuthPin_(pin);
  var auth = readStudentAuth_(sh, row);
  if (!auth.hash) throw new Error("No PIN is set for this adventurer yet. Refresh and set one first.");
  var now = Date.now();
  if (auth.lockUntil && now < Number(auth.lockUntil)) {
    var secs = Math.ceil((Number(auth.lockUntil) - now) / 1000);
    throw new Error("Too many wrong attempts. Try again in " + secs + "s.");
  }
  var hash = hashPin_(pin, auth.salt);
  if (hash !== auth.hash) {
    var fails = auth.failCount + 1;
    var updates = { PinFailCount: fails };
    if (fails >= PIN_MAX_FAILS) {
      updates.PinLockUntil = String(now + PIN_LOCKOUT_MS);
      updates.PinFailCount = 0;
    }
    writeFields_(sh, row, updates);
    throw new Error("Incorrect PIN.");
  }
  writeFields_(sh, row, { PinFailCount: 0, PinLockUntil: "" });
}

function apiCheckName(name) {
  requireName_(name);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = studentsSheet_();
    var row = ensureStudentRow_(sh, name);
    var auth = readStudentAuth_(sh, row);
    return { exists: true, hasPin: !!auth.hash };
  } finally { lock.releaseLock(); }
}

function apiClaimPin(name, authPin) {
  requireName_(name);
  validateAuthPin_(authPin);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = studentsSheet_();
    var row = ensureStudentRow_(sh, name);
    var auth = readStudentAuth_(sh, row);
    if (auth.hash) throw new Error("This adventurer already has a PIN. Enter it, or ask your teacher to reset it.");
    var salt = Utilities.getUuid();
    var hash = hashPin_(authPin, salt);
    writeFields_(sh, row, { PinHash: hash, PinSalt: salt, PinFailCount: 0, PinLockUntil: "" });
    return readStudentPublic_(sh, row);
  } finally { lock.releaseLock(); }
}

function apiVerifyPin(name, authPin) {
  requireName_(name);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = studentsSheet_();
    var row = ensureStudentRow_(sh, name);
    checkPin_(sh, row, authPin);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------------- Student-facing actions ---------------- */

function apiGetStudent(name) {
  requireName_(name);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = studentsSheet_();
    var row = ensureStudentRow_(sh, name);
    return readStudentPublic_(sh, row);
  } finally { lock.releaseLock(); }
}

function weightedRoll_() {
  var total = RUNES.reduce(function (s, r) { return s + r.prob; }, 0);
  var roll = Math.random() * total;
  for (var i = 0; i < RUNES.length; i++) {
    if (roll < RUNES[i].prob) return RUNES[i].code;
    roll -= RUNES[i].prob;
  }
  return RUNES[RUNES.length - 1].code;
}

function findPinRow_(sh, code) {
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(code).trim()) return i + 1;
  }
  return -1;
}

function apiRoll(name, code, authPin) {
  requireName_(name);
  if (!/^\d{6}$/.test(String(code || ""))) throw new Error("Milestone codes are six digits.");
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = studentsSheet_();
    var row = ensureStudentRow_(sh, name);
    checkPin_(sh, row, authPin);

    var pinsSh = SS.getSheetByName("Pins");
    var pinRow = findPinRow_(pinsSh, code);
    if (pinRow === -1) throw new Error("That code doesn't exist. Check with your teacher.");
    var used = pinsSh.getRange(pinRow, 2).getValue();
    if (used === true || used === "TRUE") {
      var usedBy = pinsSh.getRange(pinRow, 3).getValue();
      throw new Error("That code was already redeemed by " + usedBy + ".");
    }

    var drawn = [weightedRoll_(), weightedRoll_(), weightedRoll_()];
    var deltas = {};
    drawn.forEach(function (c) { deltas[c] = (deltas[c] || 0) + 1; });
    writeInventoryDelta_(sh, row, deltas);
    pinsSh.getRange(pinRow, 2, 1, 3).setValues([[true, String(name).trim(), new Date().toISOString()]]);
    return { drawn: drawn, student: readStudentPublic_(sh, row) };
  } finally { lock.releaseLock(); }
}

function apiTransform(name, target, selection, authPin) {
  requireName_(name);
  if (!RUNE_POINTS[target]) throw new Error("Unknown target rune.");
  selection = selection || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = studentsSheet_();
    var row = ensureStudentRow_(sh, name);
    checkPin_(sh, row, authPin);

    var student = readStudentPublic_(sh, row);
    var total = 0;
    Object.keys(selection).forEach(function (code) {
      var qty = Number(selection[code]) || 0;
      if (qty < 0) throw new Error("Invalid selection.");
      if (qty > (student.inventory[code] || 0)) throw new Error("You don't own that many " + code + " runes.");
      total += qty * (RUNE_POINTS[code] || 0);
    });
    if (total < RUNE_POINTS[target]) throw new Error("Selected runes aren't worth enough yet.");
    var deltas = {};
    Object.keys(selection).forEach(function (code) {
      var qty = Number(selection[code]) || 0;
      if (qty > 0) deltas[code] = (deltas[code] || 0) - qty;
    });
    deltas[target] = (deltas[target] || 0) + 1;
    writeInventoryDelta_(sh, row, deltas);
    return readStudentPublic_(sh, row);
  } finally { lock.releaseLock(); }
}

function apiCraft(name, itemId, authPin) {
  requireName_(name);
  var item = null;
  for (var i = 0; i < RUNEWORDS.length; i++) { if (RUNEWORDS[i].id === itemId) { item = RUNEWORDS[i]; break; } }
  if (!item) throw new Error("Unknown runeword.");
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = studentsSheet_();
    var row = ensureStudentRow_(sh, name);
    checkPin_(sh, row, authPin);

    var student = readStudentPublic_(sh, row);
    var need = {};
    item.recipe.forEach(function (c) { need[c] = (need[c] || 0) + 1; });
    Object.keys(need).forEach(function (code) {
      if ((student.inventory[code] || 0) < need[code]) throw new Error("Missing runes for " + item.name + ".");
    });
    var deltas = {};
    Object.keys(need).forEach(function (code) { deltas[code] = -need[code]; });
    writeInventoryDelta_(sh, row, deltas);
    appendCrafted_(sh, row, { name: item.name, at: new Date().toISOString() });
    return readStudentPublic_(sh, row);
  } finally { lock.releaseLock(); }
}

function apiCraftSour(name, mult, authPin) {
  requireName_(name);
  mult = Number(mult) || 1;
  if ([1, 2, 3].indexOf(mult) === -1) throw new Error("Invalid size.");
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = studentsSheet_();
    var row = ensureStudentRow_(sh, name);
    checkPin_(sh, row, authPin);

    var student = readStudentPublic_(sh, row);
    var need = {};
    SOUR_KEY_BASE.forEach(function (c) { need[c] = mult; });
    Object.keys(need).forEach(function (code) {
      if ((student.inventory[code] || 0) < need[code]) throw new Error("Missing runes for Sour Key.");
    });
    var deltas = {};
    Object.keys(need).forEach(function (code) { deltas[code] = -need[code]; });
    writeInventoryDelta_(sh, row, deltas);
    var label = mult === 1 ? "Small" : mult === 2 ? "Medium" : "Large";
    appendCrafted_(sh, row, { name: "Sour Key (" + label + ")", at: new Date().toISOString() });
    return readStudentPublic_(sh, row);
  } finally { lock.releaseLock(); }
}

/* ---------------- Teacher actions ---------------- */

function getConfig_(key) {
  var sh = SS.getSheetByName("Config");
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) { if (values[i][0] === key) return values[i][1]; }
  return null;
}
function setConfig_(key, value) {
  var sh = SS.getSheetByName("Config");
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
}
function requireTeacher_(pass) {
  var existing = getConfig_("TeacherPass");
  if (!existing || String(existing) !== String(pass)) throw new Error("Incorrect teacher passphrase.");
}

function apiTeacherSetup(pass) {
  if (!pass || String(pass).length < 3) throw new Error("Passphrase must be at least 3 characters.");
  var existing = getConfig_("TeacherPass");
  if (existing) throw new Error("A passphrase is already set.");
  setConfig_("TeacherPass", pass);
  return { ok: true };
}

function apiTeacherLogin(pass) {
  var existing = getConfig_("TeacherPass");
  if (!existing) throw new Error("No passphrase set yet — use setup first.");
  if (String(existing) !== String(pass)) throw new Error("Incorrect passphrase.");
  return { ok: true };
}

function apiTeacherGeneratePin(pass) {
  requireTeacher_(pass);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = SS.getSheetByName("Pins");
    var code;
    do { code = String(Math.floor(100000 + Math.random() * 900000)); } while (findPinRow_(sh, code) !== -1);
    sh.appendRow([code, false, "", new Date().toISOString(), ""]);
    return { code: code };
  } finally { lock.releaseLock(); }
}

function apiTeacherGrantGod(pass, name) {
  requireTeacher_(pass);
  requireName_(name);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = studentsSheet_();
    var row = ensureStudentRow_(sh, name);
    writeInventoryDelta_(sh, row, { GOD: 1 });
    return readStudentPublic_(sh, row);
  } finally { lock.releaseLock(); }
}

function apiTeacherListPins(pass) {
  requireTeacher_(pass);
  var sh = SS.getSheetByName("Pins");
  var values = sh.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    rows.push({
      code: String(values[i][0]),
      used: values[i][1] === true || values[i][1] === "TRUE",
      usedBy: values[i][2],
      createdAt: values[i][3],
      usedAt: values[i][4]
    });
  }
  return rows.reverse();
}

function apiTeacherListStudents(pass) {
  requireTeacher_(pass);
  var sh = studentsSheet_();
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var obj = {};
    headers.forEach(function (h, idx) { obj[h] = values[i][idx]; });
    var total = 0;
    ALL_INV_CODES.forEach(function (c) { total += Number(obj[c]) || 0; });
    var crafted = [];
    try { crafted = JSON.parse(obj.Crafted || "[]"); } catch (e) {}
    rows.push({ name: obj.Name, total: total, craftedCount: crafted.length, hasPin: !!obj.PinHash });
  }
  return rows;
}

function apiTeacherResetPin(pass, name) {
  requireTeacher_(pass);
  requireName_(name);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = studentsSheet_();
    var row = ensureStudentRow_(sh, name);
    writeFields_(sh, row, { PinHash: "", PinSalt: "", PinFailCount: 0, PinLockUntil: "" });
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function apiTeacherReset(pass) {
  requireTeacher_(pass);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ["Students", "Pins"].forEach(function (name) {
      var sh = SS.getSheetByName(name);
      var lastRow = sh.getLastRow();
      if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
    });
    return { ok: true };
  } finally { lock.releaseLock(); }
}
