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
 *
 * UPGRADING an existing deployment to add teacher-managed items + the
 * God Rune transform:
 *  1. Replace the old Code.gs contents with this file, and the old
 *     index.html with the new one.
 *  2. Run "setup" once more. This creates a new "Runewords" sheet tab and,
 *     since it's empty on first run, fills it with the same runeword
 *     catalog that used to be hardcoded — nobody's crafted items or rune
 *     counts are touched.
 *  3. Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
 *  4. From then on, add/edit/hide/delete runewords from the Teacher tab in
 *     the app itself — no more code edits needed for game-content changes.
 *  5. Any student can now forge a God Rune under the Transform tab by
 *     destroying one of every ordinary rune at once. "Grant a God Rune" is
 *     gone from the Teacher tab, since it's no longer how students get one.
 *
 * UPGRADING an existing deployment to add Power Words (whole-class group
 * rewards):
 *  1. Replace the old Code.gs contents with this file, and the old
 *     index.html with the new one.
 *  2. Run "setup" once more. This adds a new "PowerWordRedemptions" sheet
 *     tab; nobody's rune counts, PINs, or milestone codes are touched.
 *  3. Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
 *  4. From the Teacher tab, set a Power Word (3–12 characters) any time you
 *     want to reward the whole class at once — it's distinct from a
 *     one-time Milestone Code: any adventurer can redeem it (once each) on
 *     the Roll page for ONE random rune, and it expires 12 hours after
 *     you create it.
 *
 * UPGRADING further to add a teacher-replaceable site logo: no new sheet
 * tab is needed (it reuses Config: LogoImage / LogoUpdatedAt), so just
 * redeploy Code.gs and push the new index.html. Any teacher can upload a
 * JPEG/PNG/BMP/WEBP/GIF from the Teacher tab; it replaces the header glyph
 * for every visitor until changed again or cleared back to the default.
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

// This used to be the live source of truth for runewords, duplicated by
// hand in index.html. It's now only a one-time seed: on first "setup" run,
// these rows are copied into a "Runewords" sheet tab, which becomes the
// real source of truth from then on (see Section 9 of the tech guide —
// this note supersedes it). The teacher manages runewords from the Teacher
// tab in the app (add/edit/hide/delete), no code edits or redeploys needed
// for game-content changes anymore.
const RUNEWORD_SEED = [
  { id: "enigma", name: "Enigma", slot: "Chest Armor", recipe: ["BER", "NEF", "THUL", "ORT"],
    effect: "Once per long rest, cast Teleport on yourself.",
    quote: "It's a riddle wrapped in a mystery inside an enigma…" },
  { id: "cling", name: "Cling", slot: "Gloves", recipe: ["IO", "NEF", "ITH"],
    effect: "Once per long rest, firmly attach your gloves to an object for 30 minutes. Passively +2 to Sleight of Hand checks.",
    quote: "Loss of loving connection is coded by the human brain into a primal panic response." },
  { id: "prism", name: "Prism", slot: "Necklace", recipe: ["TIR", "NEF", "TAL"],
    effect: "Once per long rest, change your eye colour until you use it again.",
    quote: "Life is a painting, and you are the artist." },
  { id: "rizz", name: "Silver Tongue of Rizz", slot: "Eyewear", recipe: ["SHAEL", "IO"],
    effect: "Once per long rest, add 1d10 to a single Charisma-based check.",
    quote: "One man's ways may be as good as another's, but we all like our own best." },
  { id: "quicksilver", name: "Quicksilver", slot: "Harness / Bridle", recipe: ["LUM", "ORT", "ITH"],
    effect: "Any mount wearing this has its move speed doubled.",
    quote: "Faster, faster, until the thrill of speed overcomes the fear of death." },
  { id: "charger", name: "Charger", slot: "Harness / Bridle", recipe: ["BER", "THUL", "TAL", "SHAEL"],
    effect: "When attached to a mount, the rider cannot be knocked off.",
    quote: "When you have exhausted all possibilities, remember this: you haven't." },
  { id: "precious", name: "Precious", slot: "Ring", recipe: ["ZOD", "LUM", "THUL"],
    effect: "Once per long rest, cast Invisibility on yourself. Each combat turn, roll a d20 — a 7 or less adds a Short-Term Madness effect for the fight.",
    quote: "What makes something precious? Losing and finding it." },
  { id: "hive", name: "Hive", slot: "Boots", recipe: ["BER", "LUM", "ORT", "ITH"],
    effect: "Once per long rest, shrink to the size of an ant for 3 minutes.",
    quote: "People, chained by monotony, afraid to think, clinging to certainties... they live like ants." },
  { id: "starboard", name: "Starboard", slot: "Oar / Paddle", recipe: ["IO", "NEF", "TIR"],
    effect: "Once per day, discharge lightning into water within 30 feet. Targets: DC15 Con save, [caster level]d8 damage, half on success.",
    quote: "What is a soul? It's like electricity - we don't really know what it is, but it's a force that can light the world." },
  { id: "carotene", name: "Blade of Carotene", slot: "Sword", recipe: ["IO", "NEF", "TIR"],
    effect: "A sword of carrots. Grants Darkvision. Crits on 19 or 20 in the dark, plus 2d6 extra damage in the dark.",
    quote: "" },
  { id: "folly", name: "Folly", slot: "Scarf", recipe: ["ZOD", "SHAEL", "ORT"],
    effect: "Increases Strength to 666,666 for one second — then you instantly die.",
    quote: "Man's wisdom is his best friend; folly his worst enemy." },
  { id: "insolence", name: "Insolence", slot: "Weapon", recipe: ["LUM", "SHAEL", "ORT", "TAL"],
    effect: "Three times per long rest, gain advantage on attacks against characters over 150 years old.",
    quote: "Youth is insolent; it is its right – its necessity; all assertion in this world of doubts is a defiance, is an insolence…" },
  { id: "sage", name: "Sage", slot: "Earrings", recipe: ["TIR", "NEF", "ORT", "SHAEL"],
    effect: "Once per day, understand (but not speak) the language of animals for 3 minutes.",
    quote: "The wise understand that knowing animals allows us to better know ourselves." },
  { id: "lark", name: "Lark", slot: "Scarf", recipe: ["ITH", "TAL", "THUL", "IO"],
    effect: "Once per day, speak to (but not understand) animals for 3 minutes.",
    quote: "I may look normal, but believe me, I talk to animals and wait for them to reply." },
  { id: "inferno", name: "Inferno", slot: "Dish", recipe: ["ZOD", "IO", "THUL"],
    effect: "Once per long rest, whoever eats from this dish can cast Dragon's Breath — DC15 Con save or 2d6 fire damage.",
    quote: "I may look normal, but believe me, I talk to animals and wait for them to reply." },
  { id: "socks", name: "Socks of Blocks", slot: "Socks", recipe: ["ZOD", "JAH", "IO", "TAL"],
    effect: "Once per long rest, convert an item smaller than a microwave into brick-sized, reassemblable blocks for 1 hour.",
    quote: "Socks of blocks and blocks of socks." },
  { id: "north", name: "Invigorated by the North", slot: "Pendant / Necklace", recipe: ["LUM", "SHAEL", "TAL"],
    effect: "In cold environments, you take no exhaustion.",
    quote: "Brrr, it's NOT cold in here." },
  { id: "bramble", name: "Bramble", slot: "Leg Armor", recipe: ["SHAEL", "ORT", "NEF", "TIR"],
    effect: "Attackers who strike you take 1d4 piercing damage. Once per day, cast Entangle.",
    quote: "Just like the dead limbs and overgrowth of trees, it is superfluous. Let it all go." },
  { id: "crescent", name: "Crescent Moon", slot: "Musical Instrument", recipe: ["BER", "SHAEL", "TAL", "ITH"],
    effect: "Once per long rest, cast Summon Spirit Wolf.",
    quote: "For the strength of the pack is the wolf. And the strength of the wolf is the pack." },
  { id: "harmony", name: "Harmony", slot: "Head Armor", recipe: ["LUM", "IO", "TAL", "TIR"],
    effect: "Once per long rest, cast Mass Healing Word on targets of your choosing.",
    quote: "Music doesn't get in. Music is already in. Music simply uncovers what is there." },
  { id: "summarmaekir", name: "Summarmækir", slot: "Weapon", recipe: ["GOD", "ZOD", "JAH", "IO"],
    effect: "+3 sword. Can be thrown and guided (+½ exhaustion to the thrower on a hit). Once per long rest, cast Sunburst.",
    quote: "Yes, child, this weapon WAS held by the god Frey. It warms the wielder like a Summer's day." },
  { id: "ransnet", name: "Ran's Net", slot: "Weapon", recipe: ["GOD", "BER", "LUM", "SHAEL"],
    effect: "Once per day, cast a magic net on a target (DC20 Dex save to escape). Each turn restrained, the target suffers Vampiric Touch before acting.",
    quote: "Ethical fishing at its finest. — Senlin" },
  { id: "herofeller", name: "Herofeller", slot: "Scabbard", recipe: ["TIR", "TAL", "ORT", "IO"],
    effect: "The first attack with a weapon drawn from this scabbard deals an extra 1d4 poison damage for 10 turns. Refills on a long rest.",
    quote: "This doesn't feel sanitary." },
  { id: "coolth", name: "Coolth", slot: "Pillow", recipe: ["TIR", "THUL", "SHAEL", "LUM"],
    effect: "Perpetually cold on both sides. Once per 3 days, treat a short rest as a long rest.",
    quote: "I may look normal, but believe me, I talk to animals and wait for them to reply." },
  { id: "misfortune", name: "Amulet of Misfortune", slot: "Necklace", recipe: ["TIR", "SHAEL", "THUL", "NEF"],
    effect: "Once per long rest, create a blazing light and cast Light Binding.",
    quote: "I saw the light, I saw the light. No more darkness, no more night." },
  { id: "stuffed", name: "Stuffed Buddy", slot: "Doll", recipe: ["ITH", "NEF", "THUL", "IO"],
    effect: "Tuck this doll in for 3 nights running with your runes inside — it wakes as your familiar.",
    quote: "Who's my lil guy? Who's my lil' buddy?" },
];
const SOUR_KEY_BASE = ["SHAEL", "IO", "JAH", "BER"];

const RUNEWORD_HEADERS = ["Id", "Name", "Slot", "Recipe", "Effect", "Quote", "Image", "Hidden", "UpdatedAt"];
// Keeps each Runewords!Image cell safely under a Google Sheet cell's ~50,000
// character limit. The frontend compresses/resizes photos before upload to
// stay under this, but the backend re-checks in case a caller skips that.
const MAX_IMAGE_CHARS = 45000;

const PIN_LOCKOUT_MS = 60000; // 60s lockout after too many wrong PIN attempts
const PIN_MAX_FAILS = 5;

// A Power Word is a *shared* code the whole class can redeem, unlike a
// Milestone Code (Pins tab), which is one-time-use by a single student.
// Redeeming it grants exactly one random rune, not three. See Section on
// Power Words in the tech guide for the full design.
const POWER_WORD_DURATION_MS = 12 * 60 * 60 * 1000; // 12 hours

// The site logo shown top-left of the header. Teacher-uploaded, stored as a
// data: URI in Config (LogoImage), same size-cap reasoning as runeword
// photos (Section 5.4 of the tech guide / MAX_IMAGE_CHARS above), just with
// a slightly higher ceiling of its own. Empty/unset means "use the built-in
// SVG glyph baked into index.html". Google Sheets' hard cap is ~50,000
// characters per cell — do not raise this past ~48,000 without also moving
// to a different storage approach (e.g. Drive + a stored file URL).
const LOGO_MAX_CHARS = 48000;
const LOGO_ALLOWED_MIME = ["image/jpeg", "image/jpg", "image/png", "image/bmp", "image/webp", "image/gif"];

/* ---------------- Setup ---------------- */

function setup() {
  var studentHeaders = ["Name"].concat(ALL_INV_CODES).concat([
    "Crafted", "PinHash", "PinSalt", "PinFailCount", "PinLockUntil"
  ]);
  var sh = ensureSheet_("Students", studentHeaders);
  ensureColumns_(sh, studentHeaders);
  ensureSheet_("Pins", ["Code", "Used", "UsedBy", "CreatedAt", "UsedAt"]);
  ensureSheet_("Config", ["Key", "Value"]);
  ensureSheet_("PowerWordRedemptions", ["Word", "Name", "RedeemedAt"]);
  seedRunewords_();
}

// Creates the Runewords sheet (if missing) and, only if it's empty, fills
// it with the built-in RUNEWORD_SEED catalog. Safe to re-run: once any
// rows exist (including teacher-added or teacher-edited ones), this never
// touches them again.
function seedRunewords_() {
  var sh = ensureSheet_("Runewords", RUNEWORD_HEADERS);
  ensureColumns_(sh, RUNEWORD_HEADERS);
  if (sh.getLastRow() <= 1) {
    RUNEWORD_SEED.forEach(function (item) {
      sh.appendRow([
        item.id, item.name, item.slot, JSON.stringify(item.recipe),
        item.effect || "", item.quote || "", "", false, new Date().toISOString()
      ]);
    });
  }
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

/* ---------------- Runeword (item) row helpers ---------------- */
// The Runewords sheet is now the single source of truth for runeword game
// content (name, slot, recipe, description, quote, photo, hidden flag).
// The Teacher tab's "Manage Items" panel calls the teacher* actions below
// to add/edit/hide/delete rows here — no code edits or redeploys needed.

function runewordsSheet_() { return SS.getSheetByName("Runewords"); }

function rowToRuneword_(headers, values) {
  var obj = {};
  headers.forEach(function (h, i) { obj[h] = values[i]; });
  var recipe = [];
  try { recipe = JSON.parse(obj.Recipe || "[]"); } catch (e) { recipe = []; }
  return {
    id: obj.Id,
    name: obj.Name,
    slot: obj.Slot,
    recipe: recipe,
    effect: obj.Effect || "",
    quote: obj.Quote || "",
    image: obj.Image || "",
    hidden: obj.Hidden === true || obj.Hidden === "TRUE"
  };
}

// includeHidden=false (default) is what students see (Craft, Codex).
// includeHidden=true is what the Teacher tab's item manager sees.
function readRunewords_(includeHidden) {
  var sh = runewordsSheet_();
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue; // skip any blank row
    var item = rowToRuneword_(headers, values[i]);
    if (!includeHidden && item.hidden) continue;
    out.push(item);
  }
  return out;
}

function findRunewordRowIndex_(sh, id) {
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function writeRunewordFields_(sh, row, fields) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Object.keys(fields).forEach(function (key) {
    var col = headers.indexOf(key) + 1;
    if (col > 0) sh.getRange(row, col).setValue(fields[key]);
  });
}

function slugify_(name) {
  return String(name).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") || "item";
}

function uniqueRunewordId_(sh, base) {
  var id = base, n = 2;
  while (findRunewordRowIndex_(sh, id) !== -1) { id = base + "-" + n; n++; }
  return id;
}

function validateRecipe_(recipe) {
  if (!Array.isArray(recipe) || recipe.length === 0) throw new Error("Pick at least one rune for the recipe.");
  recipe.forEach(function (code) {
    code = String(code).toUpperCase();
    if (ALL_INV_CODES.indexOf(code) === -1) throw new Error("Unknown rune in recipe: " + code);
  });
}

function validateImage_(image) {
  if (image && String(image).length > MAX_IMAGE_CHARS) {
    throw new Error("That photo is too large even after compression — try a smaller image.");
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
      case "redeemPowerWord": data = apiRedeemPowerWord(body.name, body.word, body.authPin); break;
      case "transform": data = apiTransform(body.name, body.target, body.selection, body.authPin); break;
      case "craft": data = apiCraft(body.name, body.itemId, body.authPin); break;
      case "craftSour": data = apiCraftSour(body.name, body.mult, body.authPin); break;
      case "getRunewords": data = apiGetRunewords(); break;
      case "getLogo": data = apiGetLogo(); break;
      case "teacherSetup": data = apiTeacherSetup(body.pass); break;
      case "teacherLogin": data = apiTeacherLogin(body.pass); break;
      case "teacherGeneratePin": data = apiTeacherGeneratePin(body.pass); break;
      case "teacherSetPowerWord": data = apiTeacherSetPowerWord(body.pass, body.word); break;
      case "teacherGetPowerWord": data = apiTeacherGetPowerWord(body.pass); break;
      case "teacherSetLogo": data = apiTeacherSetLogo(body.pass, body.image); break;
      case "teacherClearLogo": data = apiTeacherClearLogo(body.pass); break;
      case "teacherListPins": data = apiTeacherListPins(body.pass); break;
      case "teacherListRunewords": data = apiTeacherListRunewords(body.pass); break;
      case "teacherAddRuneword": data = apiTeacherAddRuneword(body.pass, body); break;
      case "teacherEditRuneword": data = apiTeacherEditRuneword(body.pass, body.id, body); break;
      case "teacherSetRunewordHidden": data = apiTeacherSetRunewordHidden(body.pass, body.id, body.hidden); break;
      case "teacherDeleteRuneword": data = apiTeacherDeleteRuneword(body.pass, body.id); break;
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

/* ---------------- Power Words (shared, whole-class redemption) ---------------- */
// Distinct from a Milestone Code (Pins tab): a Power Word is one shared
// word the teacher sets for the whole class, any adventurer can redeem it
// (once each), and each redemption grants exactly ONE random rune, not
// three. It auto-expires 12 hours after the teacher creates it. The
// currently-active word lives in Config (PowerWord / PowerWordExpiresAt);
// who has already redeemed it lives in the PowerWordRedemptions sheet, one
// row per (word, name) pair, so the same student can't redeem the same
// word twice but every other student still can.

function powerWordRedemptionsSheet_() { return SS.getSheetByName("PowerWordRedemptions"); }

function validatePowerWordText_(word) {
  var w = String(word || "").trim();
  if (w.length < 3 || w.length > 12) throw new Error("Power Words must be 3–12 characters.");
  return w;
}

function hasRedeemedPowerWord_(word, name) {
  var sh = powerWordRedemptionsSheet_();
  var values = sh.getDataRange().getValues();
  var wKey = String(word).trim().toLowerCase();
  var nKey = String(name).trim().toLowerCase();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === wKey &&
        String(values[i][1]).trim().toLowerCase() === nKey) return true;
  }
  return false;
}

function apiRedeemPowerWord(name, word, authPin) {
  requireName_(name);
  var entered = validatePowerWordText_(word);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = studentsSheet_();
    var row = ensureStudentRow_(sh, name);
    checkPin_(sh, row, authPin);

    var current = getConfig_("PowerWord");
    var expiresAt = Number(getConfig_("PowerWordExpiresAt")) || 0;
    if (!current) throw new Error("There's no Power Word active right now. Ask your teacher.");
    if (Date.now() >= expiresAt) throw new Error("That Power Word has expired. Ask your teacher for a new one.");
    if (String(current).trim().toLowerCase() !== entered.toLowerCase()) throw new Error("Incorrect Power Word.");
    if (hasRedeemedPowerWord_(current, name)) throw new Error(String(name).trim() + " has already redeemed this Power Word.");

    var drawn = weightedRoll_();
    var deltas = {};
    deltas[drawn] = 1;
    writeInventoryDelta_(sh, row, deltas);
    powerWordRedemptionsSheet_().appendRow([current, String(name).trim(), new Date().toISOString()]);
    return { drawn: drawn, student: readStudentPublic_(sh, row) };
  } finally { lock.releaseLock(); }
}

function apiTransform(name, target, selection, authPin) {
  requireName_(name);
  if (target !== "GOD" && !RUNE_POINTS[target]) throw new Error("Unknown target rune.");
  selection = selection || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = studentsSheet_();
    var row = ensureStudentRow_(sh, name);
    checkPin_(sh, row, authPin);

    var student = readStudentPublic_(sh, row);

    // God Runes are no longer hand-granted by the teacher. Instead, a
    // student forges one by destroying exactly one of each of the twelve
    // ordinary runes at once — this is a fixed recipe, not a point-value
    // trade like an ordinary Transform.
    if (target === "GOD") {
      var godDeltas = {};
      RUNE_CODES.forEach(function (code) {
        var qty = Number(selection[code]) || 0;
        if (qty < 1) throw new Error("You need one of every ordinary rune to forge a God Rune.");
        if (qty > (student.inventory[code] || 0)) throw new Error("You don't own that many " + code + " runes.");
        godDeltas[code] = -1;
      });
      godDeltas.GOD = (godDeltas.GOD || 0) + 1;
      writeInventoryDelta_(sh, row, godDeltas);
      return readStudentPublic_(sh, row);
    }

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
  var items = readRunewords_(false); // hidden items can't be crafted
  for (var i = 0; i < items.length; i++) { if (items[i].id === itemId) { item = items[i]; break; } }
  if (!item) throw new Error("Unknown or unavailable runeword.");
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

// Creates/overwrites the single active Power Word for the whole class.
// Setting a new one always replaces whatever word (expired or not) was
// active before — there's only ever one live Power Word at a time.
function apiTeacherSetPowerWord(pass, word) {
  requireTeacher_(pass);
  var w = validatePowerWordText_(word);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var expiresAt = Date.now() + POWER_WORD_DURATION_MS;
    setConfig_("PowerWord", w);
    setConfig_("PowerWordExpiresAt", String(expiresAt));
    setConfig_("PowerWordCreatedAt", String(Date.now()));
    return { word: w, expiresAt: expiresAt };
  } finally { lock.releaseLock(); }
}

// Status for the Teacher tab: the current word (if any), when it expires,
// whether it's still active, and how many adventurers have redeemed it so
// far — read-only, no lock needed.
function apiTeacherGetPowerWord(pass) {
  requireTeacher_(pass);
  var word = getConfig_("PowerWord") || "";
  var expiresAt = Number(getConfig_("PowerWordExpiresAt")) || 0;
  var active = !!word && Date.now() < expiresAt;
  var redemptions = 0;
  if (word) {
    var sh = powerWordRedemptionsSheet_();
    var values = sh.getDataRange().getValues();
    var wKey = String(word).trim().toLowerCase();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim().toLowerCase() === wKey) redemptions++;
    }
  }
  return { word: word, expiresAt: expiresAt, active: active, redemptions: redemptions };
}

/* ---------------- Runeword (item) management — student-facing ---------------- */

// Public: returns every visible runeword. Called by the frontend on load
// to populate the Craft grid and Codex compendium, instead of keeping its
// own hardcoded copy.
function apiGetRunewords() {
  return readRunewords_(false);
}

/* ---------------- Site logo (teacher-uploaded, whole-site) ---------------- */

function validateLogoImage_(image) {
  var s = String(image || "");
  if (!s) throw new Error("No image was received.");
  if (s.length > LOGO_MAX_CHARS) throw new Error("That image is too large even after compression — try a smaller file.");
  var m = /^data:([a-zA-Z0-9.+-]+\/[a-zA-Z0-9.+-]+);base64,/.exec(s);
  if (!m || LOGO_ALLOWED_MIME.indexOf(m[1].toLowerCase()) === -1) {
    throw new Error("Logo must be a JPEG, PNG, BMP, WEBP, or GIF image.");
  }
}

// Public: returns the current custom logo (a data: URI), or "" if the
// teacher hasn't set one, in which case the frontend falls back to its
// built-in SVG glyph. Called by every visitor on page load.
function apiGetLogo() {
  return { image: getConfig_("LogoImage") || "" };
}

// Replaces the site-wide logo. Whatever is uploaded here is shown to every
// visitor, on every device, until a teacher changes or clears it — there is
// only ever one live logo, same one-active-thing pattern as the Power Word.
function apiTeacherSetLogo(pass, image) {
  requireTeacher_(pass);
  validateLogoImage_(image);
  setConfig_("LogoImage", image);
  setConfig_("LogoUpdatedAt", String(Date.now()));
  return { image: image };
}

// Reverts to the default built-in SVG glyph.
function apiTeacherClearLogo(pass) {
  requireTeacher_(pass);
  setConfig_("LogoImage", "");
  return { ok: true };
}

/* ---------------- Runeword (item) management — teacher-facing ---------------- */

function apiTeacherListRunewords(pass) {
  requireTeacher_(pass);
  return readRunewords_(true);
}

function apiTeacherAddRuneword(pass, body) {
  requireTeacher_(pass);
  var name = String(body.name || "").trim();
  if (!name) throw new Error("Give the item a name.");
  var slot = String(body.slot || "").trim();
  if (!slot) throw new Error("Give the item a slot (e.g. Ring, Boots).");
  var recipe = (body.recipe || []).map(function (c) { return String(c).toUpperCase(); });
  validateRecipe_(recipe);
  var image = body.image ? String(body.image) : "";
  validateImage_(image);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = runewordsSheet_();
    var id = uniqueRunewordId_(sh, slugify_(name));
    sh.appendRow([
      id, name, slot, JSON.stringify(recipe),
      String(body.effect || "").trim(), String(body.quote || "").trim(),
      image, false, new Date().toISOString()
    ]);
    return readRunewords_(true);
  } finally { lock.releaseLock(); }
}

function apiTeacherEditRuneword(pass, id, body) {
  requireTeacher_(pass);
  if (!id) throw new Error("Missing item id.");
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = runewordsSheet_();
    var row = findRunewordRowIndex_(sh, id);
    if (row === -1) throw new Error("That item no longer exists.");
    var fields = {};
    if (body.name !== undefined) {
      var name = String(body.name).trim();
      if (!name) throw new Error("Give the item a name.");
      fields.Name = name;
    }
    if (body.slot !== undefined) {
      var slot = String(body.slot).trim();
      if (!slot) throw new Error("Give the item a slot.");
      fields.Slot = slot;
    }
    if (body.recipe !== undefined) {
      var recipe = (body.recipe || []).map(function (c) { return String(c).toUpperCase(); });
      validateRecipe_(recipe);
      fields.Recipe = JSON.stringify(recipe);
    }
    if (body.effect !== undefined) fields.Effect = String(body.effect).trim();
    if (body.quote !== undefined) fields.Quote = String(body.quote).trim();
    if (body.image !== undefined) {
      var image = String(body.image || "");
      validateImage_(image);
      fields.Image = image;
    }
    fields.UpdatedAt = new Date().toISOString();
    writeRunewordFields_(sh, row, fields);
    return readRunewords_(true);
  } finally { lock.releaseLock(); }
}

function apiTeacherSetRunewordHidden(pass, id, hidden) {
  requireTeacher_(pass);
  if (!id) throw new Error("Missing item id.");
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = runewordsSheet_();
    var row = findRunewordRowIndex_(sh, id);
    if (row === -1) throw new Error("That item no longer exists.");
    writeRunewordFields_(sh, row, { Hidden: !!hidden, UpdatedAt: new Date().toISOString() });
    return readRunewords_(true);
  } finally { lock.releaseLock(); }
}

// Permanently removes an item from the catalog. Students who already
// crafted it keep it in their Crafted list (that's just a name+timestamp
// snapshot, not a live reference), but no one can craft it again.
function apiTeacherDeleteRuneword(pass, id) {
  requireTeacher_(pass);
  if (!id) throw new Error("Missing item id.");
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = runewordsSheet_();
    var row = findRunewordRowIndex_(sh, id);
    if (row === -1) throw new Error("That item no longer exists.");
    sh.deleteRow(row);
    return readRunewords_(true);
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

// Wipes all student and milestone-code data, plus any active Power Word
// and its redemption history. Irreversible. (Does not touch Runewords.)
function apiTeacherReset(pass) {
  requireTeacher_(pass);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ["Students", "Pins", "PowerWordRedemptions"].forEach(function (name) {
      var sh = SS.getSheetByName(name);
      var lastRow = sh.getLastRow();
      if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
    });
    setConfig_("PowerWord", "");
    setConfig_("PowerWordExpiresAt", "");
    return { ok: true };
  } finally { lock.releaseLock(); }
}
