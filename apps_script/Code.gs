/**
 * Backend of the end-of-action annotation site. Paste into the Apps Script project of the results Google Sheet.
 *
 * Tabs: key (clip table and Drive file ids, built by setup), annotators (name, code; you fill these in),
 * sessions (each annotator's shuffled queue and marks), annotations (the results, one row per annotator and clip).
 * After every mark the annotations tab is also written to annotations.csv in Drive, next to extensions.csv.
 * The page only ever receives the action name, timing and a Drive file id: never the clip name or version.
 */

// Paste the id of the Drive folder that holds extensions.csv and the action folders (or a folder above it).
const ROOT_FOLDER_ID = "PASTE_DRIVE_FOLDER_ID_HERE";
// Paste the Google Cloud API key restricted to the Drive API and this site's address. The page needs it to fetch videos.
const DRIVE_API_KEY = "PASTE_API_KEY_HERE";

const KEY_COLUMNS = ["clip_id", "action", "original_path", "synthesized_path", "original_end_s", "new_end_s", "extension_s",
                     "original_end_frame", "new_end_frame", "fps_exact", "original_file_id", "synthesized_file_id"];
const COLUMNS = ["annotator", "action", "clip_id", "original_path", "synthesized_path", "original_end_s", "synthesized_end_s",
                 "extension_s", "annotation_original_s", "annotation_original_frame", "annotation_synthesized_s",
                 "annotation_synthesized_frame", "annotated_action_original", "annotated_action_synthesized", "flag_original",
                 "flag_synthesized", "annotated_at"];
const FLAGS = ["", "no_clear_end", "video_problem"];

// ---------------------------------------------------------------- setup, run by hand from the Apps Script editor

/** Builds the key tab from extensions.csv and the Drive folders. Safe to re-run: sessions and results are kept. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const files = {};
  walkFolder_(DriveApp.getFolderById(ROOT_FOLDER_ID), "", files);
  const csvPath = Object.keys(files).find(p => p.endsWith("/extensions.csv"));
  if (!csvPath) throw new Error("extensions.csv not found under the Drive folder");
  const table = Utilities.parseCsv(files[csvPath].getBlob().getDataAsString("UTF-8"));
  const head = table[0];
  const rows = [], missing = [];
  table.slice(1).forEach(values => {
    const r = {};
    head.forEach((h, i) => { r[h] = values[i]; });
    const ids = ["original", "synthesized"].map(version => {
      const name = String(r[version + "_path"]).split(/[\\/]/).pop();
      const tail = `/${r.action}/${version} ${r.action}/${name}`;
      const hit = Object.keys(files).filter(p => p.endsWith(tail));
      if (hit.length !== 1) missing.push(`${r.clip_id} ${version} (${hit.length} found)`);
      return hit.length === 1 ? files[hit[0]].getId() : "";
    });
    rows.push(KEY_COLUMNS.map(c => c === "original_file_id" ? ids[0] : c === "synthesized_file_id" ? ids[1] : r[c]));
  });
  if (missing.length) throw new Error(`${missing.length} videos are not in Drive yet, nothing was changed: ${missing.join(", ")}`);
  writeTab_(tab_(ss, "key"), KEY_COLUMNS, rows);
  ensureTab_(ss, "annotators", ["name", "code"]);
  ensureTab_(ss, "sessions", ["annotator", "created", "items"]);
  ensureTab_(ss, "annotations", COLUMNS);
  const csvName = "annotations.csv";
  const existing = Object.keys(files).filter(p => p === csvPath.replace("extensions.csv", csvName));
  const out = existing.length ? files[existing[0]] : files[csvPath].getParents().next().createFile(csvName, "", "text/csv");
  PropertiesService.getScriptProperties().setProperty("ANNOTATIONS_CSV_ID", out.getId());
  writeCsv_(ss);
  Logger.log(`Key built: ${rows.length} clips.`);
}

/** Fills in a random access code for every annotator name that has none. */
function makeCodes() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("annotators");
  const values = sheet.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() && !String(values[r][1]).trim()) {
      sheet.getRange(r + 1, 2, 1, 1).setValues([[Utilities.getUuid().replace(/-/g, "").slice(0, 10)]]);
    }
  }
}

function walkFolder_(folder, path, files) {
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    const sub = subs.next();
    if (sub.getName() !== "_work") walkFolder_(sub, `${path}/${sub.getName()}`, files);
  }
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    files[`${path}/${f.getName()}`] = f;
  }
}

function tab_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureTab_(ss, name, columns) {
  const sheet = tab_(ss, name);
  if (sheet.getLastRow() === 0) writeTab_(sheet, columns, []);
}

// Plain text everywhere: otherwise Sheets reads a frame rate like 25/1 as a date.
function writeTab_(sheet, columns, rows) {
  sheet.clear();
  const all = [columns].concat(rows);
  const range = sheet.getRange(1, 1, all.length, columns.length);
  range.setNumberFormat("@");
  range.setValues(all);
  sheet.setFrozenRows(1);
}

// ---------------------------------------------------------------- requests from the page

function doPost(e) {
  let out;
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse(e.postData.contents);
    lock.waitLock(25000);
    out = handle_(body);
  } catch (err) {
    out = {error: err.message};
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(body) {
  const ss = SpreadsheetApp.getActive();
  const annotator = annotatorOf_(ss, body.code);
  const clips = loadKey_(ss);
  const session = loadSession_(ss, annotator, clips);
  const items = session.items;
  if (body.op === "session") {
    const isDone = it => it.f !== null || Boolean(it.g);
    const position = items.findIndex(it => !isDone(it));
    return {annotator, total: items.length, done: items.filter(isDone).length, position: position < 0 ? items.length : position,
            drive_key: DRIVE_API_KEY};
  }
  const i = Number(body.i);
  if (!Number.isInteger(i) || i < 0 || i >= items.length) throw new Error("No such video");
  const item = items[i];
  const clip = clips[item.c];
  const version = item.v === "o" ? "original" : "synthesized";
  const frames = Number(clip[version === "original" ? "original_end_frame" : "new_end_frame"]) + 1;
  if (body.op === "item") {
    const [num, den] = String(clip.fps_exact).split("/").map(Number);
    const actions = Array.from(new Set(Object.values(clips).map(c => c.action))).sort();
    // Deliberately no clip id, version or path: the page must not know which version it shows.
    return {i, total: items.length, action: clip.action, actions, fps_num: num, fps_den: den, n_frames: frames, frame: item.f,
            flag: item.g, action_label: item.a, file_id: clip[version + "_file_id"]};
  }
  if (body.op === "mark") {
    const frame = body.frame === null || body.frame === undefined ? null : Number(body.frame);
    const flag = body.flag || "";
    if (!FLAGS.includes(flag)) throw new Error(`Unknown flag ${flag}`);
    if (frame === null ? !flag : !Number.isInteger(frame) || frame < 0 || frame >= frames) throw new Error("Bad mark");
    item.f = frame;
    item.g = flag;
    item.a = cleanLabel_(body.action_label, clip.action);
    item.t = new Date().toISOString().slice(0, 19);
    saveSession_(ss, session);
    upsertAnnotation_(ss, session, clip);
    writeCsv_(ss);
    return {ok: true};
  }
  throw new Error("Unknown request");
}

function annotatorOf_(ss, code) {
  code = String(code || "").trim();
  if (!code) throw new Error("Enter your access code");
  const row = ss.getSheetByName("annotators").getDataRange().getValues().slice(1)
    .find(r => String(r[1]).trim() === code && String(r[0]).trim());
  if (!row) throw new Error("That access code is not valid");
  return String(row[0]).trim();
}

function loadKey_(ss) {
  const values = ss.getSheetByName("key").getDataRange().getValues();
  const clips = {};
  values.slice(1).forEach(v => {
    const c = {};
    values[0].forEach((h, i) => { c[h] = v[i]; });
    clips[c.clip_id] = c;
  });
  return clips;
}

function loadSession_(ss, annotator, clips) {
  const sheet = ss.getSheetByName("sessions");
  const values = sheet.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (values[r][0] === annotator) return {annotator, row: r + 1, items: JSON.parse(values[r][2])};
  }
  const items = buildQueue_(Object.keys(clips).sort());
  sheet.appendRow([annotator, new Date().toISOString().slice(0, 19), JSON.stringify(items)]);
  return {annotator, row: sheet.getLastRow(), items};
}

function saveSession_(ss, session) {
  ss.getSheetByName("sessions").getRange(session.row, 3, 1, 1).setValues([[JSON.stringify(session.items)]]);
}

function shuffle_(a) {
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

/** One version of every clip in the first half, the other in the second. Which comes first is balanced. */
function buildQueue_(clipIds) {
  const ids = shuffle_(clipIds.slice());
  const originalFirst = {};
  ids.forEach((c, i) => { originalFirst[c] = i % 2 === 0; });
  const gap = Math.max(1, Math.floor(ids.length / 3));
  for (let attempt = 0; attempt < 10000; attempt++) {
    const a = shuffle_(ids.slice()), b = shuffle_(ids.slice());
    const pos = {};
    a.forEach((c, i) => { pos[c] = i; });
    if (b.every((c, i) => a.length + i - pos[c] >= gap)) {
      const make = (c, v) => ({c, v, f: null, g: "", a: "", t: null});
      return a.map(c => make(c, originalFirst[c] ? "o" : "s")).concat(b.map(c => make(c, originalFirst[c] ? "s" : "o")));
    }
  }
  throw new Error("Could not build an order that keeps the two versions apart");
}

/** An annotator's corrected action name, or "" when it is the action the clip is filed under. */
function cleanLabel_(label, action) {
  const clean = String(label || "").replace(/[^\p{L}\p{N}_ /&-]+/gu, " ").replace(/\s+/g, " ")
    .replace(/^[ /&-]+|[ /&-]+$/g, "").slice(0, 60);
  return clean === action ? "" : clean;
}

function upsertAnnotation_(ss, session, clip) {
  const [num, den] = String(clip.fps_exact).split("/").map(Number);
  const row = {annotator: session.annotator, action: clip.action, clip_id: clip.clip_id, original_path: clip.original_path,
               synthesized_path: clip.synthesized_path, original_end_s: clip.original_end_s, synthesized_end_s: clip.new_end_s,
               extension_s: clip.extension_s};
  const stamps = [];
  [["original", "o"], ["synthesized", "s"]].forEach(([version, v]) => {
    const m = session.items.find(it => it.c === clip.clip_id && it.v === v);
    const marked = m.f !== null || Boolean(m.g);
    row[`annotation_${version}_frame`] = m.f === null ? "" : m.f;
    row[`annotation_${version}_s`] = m.f === null ? "" : Math.round(m.f * den / num * 1000) / 1000;
    row[`annotated_action_${version}`] = marked ? (m.a || clip.action) : "";
    row[`flag_${version}`] = m.g;
    if (m.t) stamps.push(m.t);
  });
  row.annotated_at = stamps.sort().pop() || "";
  const values = COLUMNS.map(c => row[c]);
  const sheet = ss.getSheetByName("annotations");
  const rows = sheet.getDataRange().getValues();
  const at = rows.findIndex(r => r[0] === session.annotator && r[2] === clip.clip_id);
  const range = sheet.getRange(at < 0 ? rows.length + 1 : at + 1, 1, 1, COLUMNS.length);
  range.setNumberFormat("@");
  range.setValues([values]);
}

/** Rewrites annotations.csv in Drive from the annotations tab. doPost holds the lock, so two saves cannot collide. */
function writeCsv_(ss) {
  const cell = v => /[",\n\r]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : String(v);
  const text = ss.getSheetByName("annotations").getDataRange().getValues().map(r => r.map(cell).join(",")).join("\r\n") + "\r\n";
  DriveApp.getFileById(PropertiesService.getScriptProperties().getProperty("ANNOTATIONS_CSV_ID")).setContent(text);
}
