// Samora farmer database: receives the form, serves the impact website,
// and backs everything up to Google Drive.
//
// One-time setup (in the Google Sheet: Extensions > Apps Script):
//   1. Paste this file, save, choose "setup" in the toolbar and press Run.
//      It creates the "Samora Farmer Database" Drive folder holding this Sheet,
//      a Backups folder, and a PRIVATE "Settings & Security" sheet with the PINs.
//   2. Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
//   3. Put the web-app URL into SUBMIT_URL (form) and DATA_URL (impact website).

const SHEET_NAME = "Wafugaji";
const COLUMNS = [
  "id", "muda", "mkusanyaji", "jina", "jinsia", "umri", "mahali", "simu",
  "wanyama_kwa_safari", "safari_kwa_mwaka", "aina_ya_wanyama",
  "watu_nyumbani", "sehemu_kwa_mnunuzi", "sehemu_kwa_samora",
  "lat", "lng", "chanzo_mahali"
];
const MAX_WRONG_PINS = 5;   // then the website locks for everyone...
const LOCK_MINUTES = 15;    // ...for this long
const TANZANIA = { south: -12.1, north: -0.9, west: 29.2, east: 40.6 };

// ---------- setup and settings ----------

function setup() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const folder = props.getProperty("FOLDER_ID")
    ? DriveApp.getFolderById(props.getProperty("FOLDER_ID"))
    : DriveApp.createFolder("Samora Farmer Database");
  DriveApp.getFileById(ss.getId()).moveTo(folder);

  const backups = props.getProperty("BACKUP_ID")
    ? DriveApp.getFolderById(props.getProperty("BACKUP_ID"))
    : folder.createFolder("Backups");

  let settings;
  if (props.getProperty("SETTINGS_ID")) {
    settings = SpreadsheetApp.openById(props.getProperty("SETTINGS_ID"));
  } else {
    settings = SpreadsheetApp.create("Samora – Settings & Security (PRIVATE)");
    DriveApp.getFileById(settings.getId()).moveTo(folder);
    const sh = settings.getSheets()[0];
    sh.setName("Settings");
    sh.getRange("A:C").setNumberFormat("@");
    sh.getRange(1, 1, 9, 3).setValues([
      ["Setting", "Value", "Notes"],
      ["Admin PIN", newPin_(), "Sees everything, including phone numbers. Change it here any time."],
      ["Funder PIN", newPin_(), "Sees impact and map. No phone numbers; map pins blurred to about 1 km."],
      ["Script web-app URL", ScriptApp.getService().getUrl() || "(paste after Deploy > Web app)", "Goes into SUBMIT_URL (form) and DATA_URL (website)."],
      ["Form website", "(paste once live)", "The address the QR code points to."],
      ["Impact website", "(paste once live)", "Share with funders together with the Funder PIN."],
      ["Database Sheet", ss.getUrl(), "Where every form lands. Fix mistakes here by hand."],
      ["Backups folder", backups.getUrl(), "A dated CSV copy is saved here every night."],
      ["Who can see this file", "Samora admins only", "Anyone who can open this file can read the PINs."]
    ]);
    sh.getRange("A1:C1").setFontWeight("bold");
    sh.autoResizeColumns(1, 3);
  }

  props.setProperties({ FOLDER_ID: folder.getId(), BACKUP_ID: backups.getId(), SETTINGS_ID: settings.getId() });
  sheet_();

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "backup")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("backup").timeBased().everyDays(1).atHour(2).inTimezone("Africa/Dar_es_Salaam").create();
  backup();

  Logger.log("Done. Drive folder: " + folder.getUrl());
}

function newPin_() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function setting_(name) {
  const id = PropertiesService.getScriptProperties().getProperty("SETTINGS_ID");
  if (!id) throw new Error("Run setup() first.");
  const row = SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getDisplayValues()
    .find(r => r[0] === name);
  return row ? String(row[1]).trim() : "";
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(COLUMNS);
  // Plain text, or Sheets turns "2-3" into a date and "+255…" into a number
  sh.getRange(1, 1, sh.getMaxRows(), COLUMNS.length).setNumberFormat("@");
  // Add any new columns to a sheet made by an older version of this script
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  COLUMNS.forEach(c => { if (!head.includes(c)) sh.getRange(1, sh.getLastColumn() + 1).setValue(c); });
  return sh;
}

// ---------- web requests ----------

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  return json_(body.action === "read" ? read_(body.pin) : save_(body));
}

function doGet() {
  return json_({ ok: true });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// One form submission from the field
function save_(row) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    // A phone may resend a form if the signal dropped before the reply came back
    const seen = row.id && sh.getRange("A:A").createTextFinder(row.id).matchEntireCell(true).findNext();
    if (seen) return { ok: true, duplicate: true };
    Object.assign(row, locate_(row));
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    sh.appendRow(head.map(c => row[c] ?? ""));
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// GPS from the phone if taken, otherwise look up the typed place name
function locate_(row) {
  if (inTanzania_(row.lat, row.lng)) return { chanzo_mahali: "gps" };
  try {
    const res = Maps.newGeocoder().setRegion("tz").geocode(row.mahali + ", Tanzania");
    const loc = res.status === "OK" && res.results[0].geometry.location;
    if (loc && inTanzania_(loc.lat, loc.lng)) return { lat: loc.lat, lng: loc.lng, chanzo_mahali: "jina la mahali" };
  } catch (err) {}
  return { lat: "", lng: "", chanzo_mahali: "" };
}

function inTanzania_(lat, lng) {
  lat = Number(lat); lng = Number(lng);
  return lat && lng && lat > TANZANIA.south && lat < TANZANIA.north && lng > TANZANIA.west && lng < TANZANIA.east;
}

// Data for the impact website, checked against the PINs in the Settings sheet
function read_(pin) {
  const cache = CacheService.getScriptCache();
  const wrong = Number(cache.get("wrong_pins") || 0);
  if (wrong >= MAX_WRONG_PINS) return { error: "locked" };

  pin = String(pin || "").trim();
  const role = !pin ? null
    : pin === setting_("Admin PIN") ? "admin"
    : pin === setting_("Funder PIN") ? "funder"
    : null;
  if (!role) {
    cache.put("wrong_pins", String(wrong + 1), LOCK_MINUTES * 60);
    return { error: "wrong_pin" };
  }

  const [head, ...values] = sheet_().getDataRange().getDisplayValues();
  const rows = values.map(r => {
    const o = Object.fromEntries(head.map((h, i) => [h, r[i]]));
    if (role === "funder") {
      delete o.simu;
      // About 1 km, so funders never see a farmer's exact home
      if (o.lat) o.lat = Number(o.lat).toFixed(2);
      if (o.lng) o.lng = Number(o.lng).toFixed(2);
    }
    return o;
  });
  return { role, rows };
}

// ---------- nightly backup ----------

function backup() {
  const id = PropertiesService.getScriptProperties().getProperty("BACKUP_ID");
  if (!id) throw new Error("Run setup() first.");
  const values = sheet_().getDataRange().getDisplayValues();
  const csv = values.map(r => r.map(v => '"' + String(v).replace(/"/g, '""') + '"').join(",")).join("\n");
  const stamp = Utilities.formatDate(new Date(), "Africa/Dar_es_Salaam", "yyyy-MM-dd");
  DriveApp.getFolderById(id).createFile(`samora_wafugaji_${stamp}.csv`, "﻿" + csv, MimeType.CSV);
}
