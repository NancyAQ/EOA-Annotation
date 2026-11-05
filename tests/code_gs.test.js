// Runs apps_script/Code.gs under Node against in-memory stand-ins for Sheets and Drive: node tests/code_gs.test.js
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

class Sheet {
  constructor() { this.data = []; }
  getLastRow() { return this.data.length; }
  clear() { this.data = []; }
  setFrozenRows() {}
  appendRow(v) { this.data.push(v.slice()); }
  getDataRange() { return {getValues: () => this.data.map(r => r.slice())}; }
  getRange(row, col, nr, nc) {
    const s = this;
    return {
      setNumberFormat() { return this; },
      setValues(v) {
        for (let r = 0; r < nr; r++) {
          while (s.data.length < row + r) s.data.push([]);
          for (let c = 0; c < nc; c++) s.data[row - 1 + r][col - 1 + c] = v[r][c];
        }
        return this;
      },
    };
  }
}
class Book {
  constructor() { this.tabs = {}; }
  getSheetByName(n) { return this.tabs[n] || null; }
  insertSheet(n) { return (this.tabs[n] = new Sheet()); }
}
const iter = a => { let i = 0; return {hasNext: () => i < a.length, next: () => a[i++]}; };
const written = {};
const file = (name, id, text = "", parent) => {
  const f = {getName: () => name, getId: () => id, getBlob: () => ({getDataAsString: () => text}), setContent: t => { written[id] = t; },
             getParents: () => iter([parent])};
  written[id] = text;
  return f;
};
const folder = (name, subs = [], files = []) => {
  const d = {getName: () => name, getFolders: () => iter(subs), getFiles: () => iter(files),
             createFile(n, content) { const f = file(n, "csv-out", content); files.push(f); return f; }};
  files.forEach(f => { f.getParents = () => iter([d]); });
  return d;
};

const FPS = {A: "25/1", B: "30000/1001"};
function driveTree(omit) {
  let csv = "action,clip_id,original_path,synthesized_path,original_end_s,new_end_s,extension_s,original_end_frame,new_end_frame,fps,fps_exact\n";
  const actions = [];
  for (const a of ["A", "B"]) {
    const o = [], s = [];
    for (let k = 0; k < 5; k++) {
      const stem = `000${k}`;
      csv += `${a},${a}/${stem},D:\\x\\${stem}.MP4,D:\\x\\${stem}_synthesized.mp4,4,10,6,99,249,25,${FPS[a]}\n`;
      o.push(file(`${stem}.MP4`, `id-${a}${k}-o`));
      if (`${a}/${stem}` !== omit) s.push(file(`${stem}_synthesized.mp4`, `id-${a}${k}-s`));
    }
    actions.push(folder(a, [folder(`original ${a}`, [], o), folder(`synthesized ${a}`, [], s)]));
  }
  const work = folder("_work", [], [file("0000_synthesized.mp4", "must-not-be-used")]);
  return folder("EOA", [folder("MAS_extended", actions.concat(work), [file("extensions.csv", "csv", csv)])]);
}

function load(book, tree) {
  const props = {};
  const ctx = vm.createContext({
    console,
    SpreadsheetApp: {getActive: () => book},
    DriveApp: {getFolderById: () => tree, getFileById: id => ({setContent: t => { written[id] = t; }})},
    PropertiesService: {getScriptProperties: () => ({setProperty: (k, v) => { props[k] = v; }, getProperty: k => props[k]})},
    Logger: {log() {}},
    LockService: {getScriptLock: () => ({waitLock() {}, releaseLock() {}})},
    Utilities: {parseCsv: t => t.trim().split("\n").map(l => l.split(",")), getUuid: () => "abcdef12-3456-7890-abcd-ef1234567890"},
    ContentService: {MimeType: {JSON: "json"}, createTextOutput: s => ({setMimeType: () => s})},
  });
  vm.runInContext(fs.readFileSync(path.join(__dirname, "..", "apps_script", "Code.gs"), "utf8"), ctx);
  return ctx;
}


const book = new Book();
let ctx = load(book, driveTree("A/0002"));
assert.throws(() => vm.runInContext("setup()", ctx), /1 videos are not in Drive yet.*A\/0002 synthesized/s);
assert.strictEqual(book.getSheetByName("key"), null, "a failed setup must change nothing");

ctx = load(book, driveTree());
const post = body => JSON.parse(vm.runInContext("doPost", ctx)({postData: {contents: JSON.stringify(body)}}));
vm.runInContext("setup()", ctx);
assert.strictEqual(book.tabs.key.data.length, 11);
const keyHead = book.tabs.key.data[0];
assert.ok(!book.tabs.key.data.flat().includes("must-not-be-used"), "_work must be skipped");
assert.strictEqual(book.tabs.key.data[1][keyHead.indexOf("fps_exact")], "25/1");

const CODE = "PASTE_ACCESS_CODE_HERE";
assert.match(post({op: "session", name: "Test Person", code: "wrong"}).error, /not valid/);
assert.match(post({op: "session", name: "Test Person", code: ""}).error, /not valid/);
assert.match(post({op: "session", name: "  =!! ", code: CODE}).error, /Enter your name/);
assert.match(post({op: "session", code: CODE}).error, /Enter your name/);
assert.strictEqual(book.tabs.sessions.data.length, 1, "refused requests must not create a session");

let s = post({op: "session", name: "Test Person", code: CODE});
assert.deepStrictEqual([s.annotator, s.total, s.done, s.position], ["Test Person", 20, 0, 0]);
assert.strictEqual(s.drive_key, "PASTE_API_KEY_HERE");
assert.strictEqual(post({op: "session", name: "Test Person", code: "wrong"}).drive_key, undefined, "no key without a valid code");
assert.strictEqual(post({op: "session", name: "test  person", code: CODE}).annotator, "Test Person", "capitals and spacing resume the same session");
assert.strictEqual(post({op: "session", name: "=נסי אבו-ק", code: CODE}).annotator, "נסי אבו-ק", "non-Latin names are kept, formula characters are not");
book.tabs.sessions.data.pop();
assert.strictEqual(book.tabs.sessions.data.length, 2);
assert.strictEqual(post({op: "session", name: "Test Person", code: CODE}).total, 20);
assert.strictEqual(book.tabs.sessions.data.length, 2, "the same annotator must keep one session");

// Queue: every clip twice, once per version, in separate halves and at least a third of the clips apart.
const items = JSON.parse(book.tabs.sessions.data[1][2]);
assert.strictEqual(new Set(items.map(it => it.c + it.v)).size, 20);
items.forEach((it, i) => {
  const j = items.findIndex((o, k) => k !== i && o.c === it.c);
  assert.notStrictEqual(it.v, items[j].v);
  assert.ok(Math.abs(i - j) >= 3);
});

// The page never receives anything that identifies the clip or version.
const it0 = post({op: "item", name: "Test Person", code: CODE, i: 0});
for (const k of ["clip_id", "version", "original_path", "synthesized_path", "c", "v"]) assert.ok(!(k in it0));
assert.ok(!JSON.stringify(it0).match(/D:|\/000|_synthesized/));
assert.deepStrictEqual(it0.actions, ["A", "B"]);
assert.match(it0.file_id, /^id-[AB]\d-[os]$/);
assert.strictEqual(it0.n_frames, items[0].v === "o" ? 100 : 250);
assert.match(post({op: "item", name: "Test Person", code: CODE, i: 20}).error, /No such/);

// Marks: validation, seconds, label cleaning, upsert into one row per clip.
assert.match(post({op: "mark", name: "Test Person", code: CODE, i: 0, frame: null, flag: ""}).error, /Bad mark/);
assert.match(post({op: "mark", name: "Test Person", code: CODE, i: 0, frame: 100000, flag: ""}).error, /Bad mark/);
assert.match(post({op: "mark", name: "Test Person", code: CODE, i: 0, frame: 5, flag: "nope"}).error, /Unknown flag/);
assert.ok(post({op: "mark", name: "Test Person", code: CODE, i: 0, frame: 50, flag: "", action_label: "=Tighten bolt!"}).ok);
const A = book.tabs.annotations.data, H = A[0];
assert.strictEqual(A.length, 2);
const clip = items[0].c, ver = items[0].v === "o" ? "original" : "synthesized", other = ver === "original" ? "synthesized" : "original";
const fps = clip.startsWith("A") ? 25 : 30000 / 1001;
assert.strictEqual(A[1][H.indexOf(`annotation_${ver}_frame`)], 50);
assert.strictEqual(A[1][H.indexOf(`annotation_${ver}_s`)], Math.round(50 / fps * 1000) / 1000);
assert.strictEqual(A[1][H.indexOf(`annotated_action_${ver}`)], "Tighten bolt");
assert.strictEqual(A[1][H.indexOf(`annotation_${other}_frame`)], "");
assert.strictEqual(A[1][H.indexOf(`annotated_action_${other}`)], "");
assert.strictEqual(post({op: "session", name: "Test Person", code: CODE}).position, 1);
assert.strictEqual(post({op: "item", name: "Test Person", code: CODE, i: 0}).action_label, "Tighten bolt");

const pair = items.findIndex((o, k) => k !== 0 && o.c === clip);
assert.ok(post({op: "mark", name: "Test Person", code: CODE, i: pair, frame: null, flag: "no_clear_end", action_label: clip[0]}).ok);
assert.strictEqual(A.length, 2, "both versions of a clip share one row");
assert.strictEqual(A[1][H.indexOf(`flag_${other}`)], "no_clear_end");
assert.strictEqual(A[1][H.indexOf(`annotated_action_${other}`)], clip[0]);
assert.ok(post({op: "mark", name: "Test Person", code: CODE, i: 0, frame: 60, flag: ""}).ok);
assert.strictEqual(A.length, 2);
assert.strictEqual(A[1][H.indexOf(`annotation_${ver}_frame`)], 60);
assert.strictEqual(A[1][H.indexOf(`annotated_action_${ver}`)], clip[0], "clearing the label returns to the filed action");

// annotations.csv in Drive follows the tab, and awkward characters are quoted.
const csvLines = () => written["csv-out"].trim().split("\r\n");
assert.strictEqual(csvLines()[0], H.join(","));
assert.strictEqual(csvLines().length, A.length);
assert.ok(csvLines()[1].includes(",60,"));
A.push(['a,b', 'say "hi"', 'plain']);
vm.runInContext("writeCsv_(SpreadsheetApp.getActive())", ctx);
assert.strictEqual(csvLines().pop(), '"a,b","say ""hi""",plain');
A.pop();

// A second annotator gets a separate queue and row; re-running setup keeps results.
assert.ok(post({op: "mark", name: "Second", code: CODE, i: 0, frame: 1, flag: ""}).ok);
assert.strictEqual(A.length, 3);
vm.runInContext("setup()", ctx);
assert.strictEqual(book.tabs.annotations.data.length, 3);
assert.strictEqual(csvLines().length, 3, "re-running setup finds or recreates the csv with the results");
assert.strictEqual(book.tabs.sessions.data.length, 3);
assert.strictEqual(post({op: "session", name: "Test Person", code: CODE}).done, 2);
console.log("all backend checks passed");
