/*
 * app.js - the page: a port of gui.py (Setup / Results / History) and gui_meritto.py (Meritto tab).
 * All processing is done by core.js, exactly as in the Python version; this file only wires the UI.
 */
import * as C from "./core.js";
import * as py from "./py.js";
import { openBook } from "./excel.js";
import { read_pdf } from "./pdfread.js";
import { build_report, build_meritto_report } from "./report.js";
import { Database } from "./storage.js";

const $ = (id) => document.getElementById(id);
const NOT_MAPPED = "(not mapped)";
const NOT_USED = "(not used)";
const AUTO = "(automatic)";

// ------------------------------------------------------------------ libraries
let pdfjs = null, PDF_ASSETS = {};
const XLSX = globalThis.XLSX, ExcelJS = globalThis.ExcelJS;

async function loadLibraries() {
  const missing = [];
  if (!XLSX) missing.push("SheetJS (vendor/xlsx.full.min.js)");
  if (!ExcelJS) missing.push("ExcelJS (vendor/exceljs.min.js)");
  try {
    pdfjs = await import("../vendor/pdfjs/pdf.min.mjs");
    pdfjs.GlobalWorkerOptions.workerSrc = new URL("../vendor/pdfjs/pdf.worker.min.mjs", import.meta.url).href;
    PDF_ASSETS = {
      cMapUrl: new URL("../vendor/pdfjs/cmaps/", import.meta.url).href,
      standardFontDataUrl: new URL("../vendor/pdfjs/standard_fonts/", import.meta.url).href,
    };
  } catch (e) {
    missing.push(`pdf.js (vendor/pdfjs): ${e.message}`);
  }
  if (missing.length) {
    const b = $("loadError");
    b.hidden = false;
    b.textContent = `Could not load: ${missing.join("; ")}. Open the page through a web server (e.g. GitHub Pages), not as a local file.`;
  }
}

// ------------------------------------------------------------------ small DOM helpers
function el(tag, attrs = {}, ...kids) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === null || v === undefined || v === false) continue;
    if (k === "class") e.className = v;
    else if (k === "text") e.textContent = v;
    else if (k.startsWith("on")) e.addEventListener(k.slice(2), v);
    else e.setAttribute(k, v === true ? "" : v);
  }
  for (const k of kids.flat()) if (k !== null && k !== undefined) e.append(k instanceof Node ? k : String(k));
  return e;
}
const disp = (v) => (v === null || v === undefined ? "" : String(v));

/** Fill a <table>: heads = [label], rows = [{cells:[...], cls, data}]. */
function fillTable(table, heads, rows, emptyText) {
  table.replaceChildren();
  table.append(el("thead", {}, el("tr", {}, heads.map((h) => el("th", { text: h })))));
  const tb = el("tbody");
  if (!rows.length && emptyText) tb.append(el("tr", {}, el("td", { class: "empty", colspan: heads.length, text: emptyText })));
  for (const r of rows) {
    const tr = el("tr", { class: r.cls || null });
    r.cells.forEach((c) => tr.append(el("td", { text: disp(c) })));
    tr._data = r.data;
    tb.append(tr);
  }
  table.append(tb);
  return tb;
}

function setStatus(text) { $("status").textContent = text; }
function busy(on) { $("progress").hidden = !on; }

// ------------------------------------------------------------------ dialogs
const dlg = $("dialog"), dlgForm = $("dialogForm");
function showDialog(title, body, buttons = [["OK", "ok"]]) {
  return new Promise((resolve) => {
    dlgForm.replaceChildren(el("h2", { text: title }), ...[].concat(body),
      el("div", { class: "buttons" }, buttons.map(([label, val, primary]) => el("button", { value: val, class: primary ? "primary" : null, text: label }))));
    dlg.onclose = () => resolve(dlg.returnValue);
    dlg.returnValue = "";
    dlg.showModal();
  });
}
const info = (title, msg) => showDialog(title, el("pre", { text: msg }));
const ask = async (title, msg) => (await showDialog(title, el("pre", { text: msg }), [["Cancel", "no"], ["Yes", "yes", true]])) === "yes";

/** Run a UI action and show any error instead of hiding it. */
function guard(fn) {
  return async (...a) => {
    try { await fn(...a); } catch (e) {
      console.error(e);
      busy(false);
      await info("Something went wrong", e && e.message ? e.message : String(e));
    }
  };
}

function download(buf, name) {
  const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const a = el("a", { href: url, download: name });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}
const stamp = () => py.nowStamp("%Y%m%d_%H%M");

// ------------------------------------------------------------------ shared settings (gui.App)
const S = {
  get mode() { return $("mode").value; },
  get titles() { return $("titles").checked; },
};
function settings() {
  return C.Settings({
    day_first: $("dateOrder").value === "dmy",
    match_threshold: Number($("match").value),
    review_threshold: Number($("review").value),
    ignore_titles: S.titles,
    ocr: false,
    mode: S.mode,
  });
}

function thresh(which) {
  let m = Math.round(Number($("match").value)), r = Math.round(Number($("review").value));
  if (which === "match" && r > m) r = m;
  if (which === "review" && m < r) m = r;
  $("match").value = m; $("review").value = r;
  $("matchOut").textContent = `${m}%`;
  $("reviewOut").textContent = `${r}%`;
  $("threshHelp").textContent = S.mode === C.MODE_STRICT
    ? "Strict mode: names must be identical apart from case, extra spaces and punctuation (and titles, if ignored). Thresholds are not used."
    : `Names scoring ${m}% or more count as a close match. ${r}% to ${m - 1}% are flagged for review. Below ${r}% are mismatches.`;
}

function syncShared() {
  for (const s of document.querySelectorAll(".shared-mode")) {
    s.addEventListener("change", () => {
      for (const o of document.querySelectorAll(".shared-mode")) o.value = s.value;
      thresh("match");
    });
  }
  for (const c of document.querySelectorAll(".shared-titles")) {
    c.addEventListener("change", () => { for (const o of document.querySelectorAll(".shared-titles")) o.checked = c.checked; });
  }
  $("match").addEventListener("input", () => thresh("match"));
  $("review").addEventListener("input", () => thresh("review"));
  $("dateOrder").addEventListener("change", () => page1.runDetect());
}

// ------------------------------------------------------------------ reading PDFs (both tabs)
let pdfBusy = false;
function setPdfButtons(disabled) {
  for (const id of ["pdfInput1", "folderInput1", "pdfInput2", "folderInput2", "runCompare", "mRun", "mRecheck", "rerun"]) {
    const n = $(id);
    n.disabled = disabled;
    if (n.closest(".btn")) n.closest(".btn").classList.toggle("muted", disabled);
  }
}

/** items: [{file, key}] in the order to read them. Mirrors App.add_pdfs()/_poll_pdfs(). */
async function readPdfs(target, items, prefix) {
  items = items.filter((it) => !target.seen.has(it.key));
  if (!items.length) return null;
  if (!pdfjs) throw new Error("pdf.js did not load, so PDFs cannot be read.");
  pdfBusy = true;
  setPdfButtons(true);
  busy(true);
  const errors = [];
  try {
    for (let i = 0; i < items.length; i++) {
      const { file, key } = items[i];
      try {
        const bytes = new Uint8Array(await file.arrayBuffer());
        const recs = await read_pdf(pdfjs, bytes, file.name, key, PDF_ASSETS);
        target.records.push(...recs);
        target.seen.add(key);
        setStatus(`${prefix}Read ${i + 1} of ${items.length} PDFs...`);
      } catch (e) {
        errors.push(`${file.name}: ${e && e.message ? e.message : e}`);
      }
      if (i % 5 === 4) await new Promise((r) => setTimeout(r, 0)); // keep the page responsive
    }
  } finally {
    pdfBusy = false;
    setPdfButtons(false);
    busy(false);
  }
  return errors;
}

function pdfItemsFromInput(input, folder) {
  let files = [...input.files];
  input.value = "";
  if (folder) {
    files = files.filter((f) => f.name.toLowerCase().endsWith(".pdf"));
    const key = (f) => f.webkitRelativePath || f.name;
    files.sort((a, b) => py.cmpStr(key(a), key(b))); // sorted(rglob) like the folder dialog
    return files.map((f) => ({ file: f, key: key(f) }));
  }
  return files.map((f) => ({ file: f, key: `${f.name}|${f.size}|${f.lastModified}` }));
}

async function itemsFromDrop(dt) {
  const out = [];
  const walk = async (entry) => {
    if (entry.isFile) {
      if (entry.name.toLowerCase().endsWith(".pdf")) out.push({ file: await new Promise((res, rej) => entry.file(res, rej)), key: entry.fullPath });
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      let batch;
      do {
        batch = await new Promise((res, rej) => reader.readEntries(res, rej));
        for (const e of batch) await walk(e);
      } while (batch.length);
    }
  };
  const entries = [...dt.items].map((i) => (i.webkitGetAsEntry ? i.webkitGetAsEntry() : null)).filter(Boolean);
  if (entries.length) {
    for (const e of entries) await walk(e);
    out.sort((a, b) => py.cmpStr(a.key, b.key));
    return out;
  }
  return [...dt.files].filter((f) => f.name.toLowerCase().endsWith(".pdf")).map((f) => ({ file: f, key: `${f.name}|${f.size}|${f.lastModified}` }));
}

function setupDrop(zone, onFiles) {
  zone.addEventListener("dragover", (e) => { e.preventDefault(); zone.classList.add("over"); });
  zone.addEventListener("dragleave", (e) => { if (!zone.contains(e.relatedTarget)) zone.classList.remove("over"); });
  zone.addEventListener("drop", guard(async (e) => {
    e.preventDefault();
    zone.classList.remove("over");
    await onFiles(e.dataTransfer);
  }));
}

/** Row selection for a table body: click, Ctrl/Cmd+click, Shift+click. */
function selectable(table, onChange) {
  let anchor = null;
  table.addEventListener("click", (e) => {
    const tr = e.target.closest("tbody tr");
    if (!tr || tr._data === undefined) return;
    const rows = [...table.tBodies[0].rows];
    if (e.shiftKey && anchor !== null) {
      const [a, b] = [rows.indexOf(anchor), rows.indexOf(tr)].sort((x, y) => x - y);
      rows.forEach((r, i) => r.classList.toggle("sel", i >= a && i <= b));
    } else if (e.ctrlKey || e.metaKey) {
      tr.classList.toggle("sel");
      anchor = tr;
    } else {
      rows.forEach((r) => r.classList.toggle("sel", r === tr));
      anchor = tr;
    }
    if (onChange) onChange();
  });
}
const selected = (table) => [...table.querySelectorAll("tbody tr.sel")].map((tr) => tr._data);

// ================================================================== page 1: Excel vs Marksheet
const page1 = {
  book: null, sheet: null, headerRow: -1, records: [], seen: new Set(),
  guesses: Object.fromEntries(C.MAP_FIELDS.map(([k]) => [k, C.ColumnGuess()])),
  manual: {}, result: null, runInfo: null,

  async loadExcel(file) {
    let book;
    try {
      book = openBook(XLSX, new Uint8Array(await file.arrayBuffer()), file.name);
    } catch (e) {
      return info("Could not read file", `${file.name}\n\n${e.message}`);
    }
    this.book = book;
    this.loadSheet(null);
  },

  loadSheet(sheet) {
    try {
      this.sheet = sheet === C.ALL_SHEETS ? C.load_all_sheets(this.book) : C.load_sheet(this.book, sheet);
    } catch (e) {
      return info("Could not read file", `${this.book.name}\n\n${e.message}`);
    }
    const sd = this.sheet;
    $("excelLabel").textContent = `${this.book.name}  (${sd.grid.length} rows, ${sd.ncols} columns)` + (sd.note ? `\n${sd.note}` : "");
    $("excelLabel").classList.remove("hint");
    const names = sd.sheet_names.length > 1 ? [...sd.sheet_names, C.ALL_SHEETS] : sd.sheet_names;
    $("sheetSelect").replaceChildren(...names.map((n) => el("option", { value: n, text: n })));
    $("sheetSelect").value = sd.sheet;
    $("sheetSelect").disabled = false;
    this.headerRow = sd.origin ? 0 : C.detect_header_row(sd.grid);
    $("headerRow").value = this.headerRow + 1;
    $("headerRow").disabled = false;
    this.manual = {};
    this.runDetect(false);
    setStatus(`Loaded ${this.book.name}, sheet "${sd.sheet}".`);
  },

  changeHeader() {
    if (!this.sheet) return;
    if (this.sheet.origin) { $("headerRow").value = this.headerRow + 1; return; } // combined sheets: row 1 is the header
    let n = parseInt($("headerRow").value, 10);
    if (Number.isNaN(n)) return;
    n = Math.max(0, Math.min(n, this.sheet.grid.length));
    $("headerRow").value = n;
    if (n - 1 !== this.headerRow) { this.headerRow = n - 1; this.runDetect(true); }
  },

  colChoices() {
    const out = [NOT_MAPPED];
    if (this.sheet) {
      for (let c = 0; c < this.sheet.ncols; c++) {
        const h = this.headerRow >= 0 && this.headerRow < this.sheet.grid.length ? this.sheet.grid[this.headerRow][c] : "";
        out.push(h ? `${C.col_letter(c)}: ${Array.from(h).slice(0, 28).join("")}` : C.col_letter(c));
      }
    }
    return out;
  },

  runDetect(keepManual = true) {
    if (!this.sheet) return;
    if (!keepManual) this.manual = {};
    this.guesses = C.auto_detect(this.sheet.grid, this.headerRow, this.records, settings(), this.manual);
    this.refreshMapping();
  },

  refreshMapping() {
    const box = $("mapping");
    box.replaceChildren();
    const choices = this.colChoices();
    for (const [key, label] of C.MAP_FIELDS) {
      const g = this.guesses[key];
      const sel = el("select", { "aria-label": label, onchange: () => this.onMap(key, sel.selectedIndex) },
        choices.map((c, i) => el("option", { value: i, text: c })));
      sel.selectedIndex = g.col !== null && g.col + 1 < choices.length ? g.col + 1 : 0;
      box.append(el("span", { class: "lbl", text: label + (key === "reg" ? " *" : "") }), sel,
        el("span", { class: `conf conf-${g.confidence}`, id: `conf-${key}` }), el("span", { class: "samp", id: `samp-${key}` }));
      this.confLabel(key);
      this.sampleLabel(key, g.col);
    }
    this.refreshPreview();
  },

  confLabel(key) {
    const g = this.guesses[key];
    const n = $(`conf-${key}`);
    n.textContent = { manual: "Set manually", none: "Not found" }[g.confidence] || `Auto, ${g.confidence} confidence`;
    n.className = `conf conf-${g.confidence}`;
  },

  sampleLabel(key, col) {
    const n = $(`samp-${key}`);
    if (col === null || !this.sheet) { n.textContent = ""; return; }
    const rows = C.data_rows(this.sheet.grid, this.headerRow).slice(0, 3);
    const vals = rows.map((r) => this.sheet.grid[r][col]).filter((v) => v);
    const g = this.guesses[key];
    let txt = vals.length ? "e.g. " + vals.join(", ") : "";
    if (g.reason && g.confidence !== "manual") txt += `  (${g.reason})`;
    n.textContent = py.len(txt) <= 78 ? txt : Array.from(txt).slice(0, 75).join("") + "...";
    n.title = txt;
  },

  onMap(key, i) {
    const col = i <= 0 ? null : i - 1;
    if (col === null) {
      delete this.manual[key];
      this.guesses[key] = C.ColumnGuess();
    } else {
      for (const [other, g] of Object.entries(this.guesses)) { // a column can serve one field only
        if (other !== key && g.col === col) {
          this.guesses[other] = C.ColumnGuess();
          delete this.manual[other];
        }
      }
      this.manual[key] = col;
      this.guesses[key] = C.ColumnGuess(col, "manual", 1.0, "Set manually");
    }
    this.refreshMapping();
  },

  mapping() { return Object.fromEntries(Object.entries(this.guesses).map(([k, g]) => [k, g.col])); },

  refreshPreview() {
    const t = $("preview");
    if (!this.sheet) { t.replaceChildren(); return; }
    const sd = this.sheet;
    const byCol = {};
    for (const [k, g] of Object.entries(this.guesses)) if (g.col !== null) byCol[g.col] = C.MAP_LABELS[k];
    const heads = ["Row", ...Array.from({ length: sd.ncols }, (_, i) => C.col_letter(i) + (byCol[i] ? ` = ${byCol[i]}` : ""))];
    const rows = [...(this.headerRow >= 0 ? [this.headerRow] : []), ...C.data_rows(sd.grid, this.headerRow).slice(0, 5)];
    fillTable(t, heads, rows.map((r) => ({ cells: [C.row_label(sd, r), ...sd.grid[r]], cls: r === this.headerRow ? "hdr" : null })));
  },

  // ---- PDFs
  async addPdfs(items) {
    if (pdfBusy) return info("Please wait", "PDFs are still being read.");
    const errors = await readPdfs(this, items, "");
    if (errors === null) return;
    this.refreshPdfs();
    this.runDetect(true);
    const bad = this.records.filter((r) => r.unreadable).length;
    setStatus(`${this.records.length} PDF records loaded` + (bad ? `, ${bad} unreadable.` : "."));
    if (errors.length) await info("Some PDFs could not be opened", errors.slice(0, 10).join("\n"));
    if (bad) {
      await info("Scanned PDFs?", `${bad} PDF(s) had no readable text. If they are scans, correct the values by hand ` +
        "(double-click the row), or use the desktop version with OCR.");
    }
  },

  refreshPdfs() {
    fillTable($("pdfTable1"), ["File", "Pg", "Reg. no.", "ABC ID", "Name", "Father", "Mother", "DOB"],
      this.records.map((r, i) => ({ cells: [r.file, r.page, r.fields.reg ?? "", r.fields.abc ?? "", r.fields.name ?? "", r.fields.father ?? "", r.fields.mother ?? "", r.fields.dob ?? ""], cls: r.unreadable ? "bad" : null, data: i })));
    $("pdfCount1").textContent = this.records.length ? `${this.records.length} PDF record(s) from ${this.seen.size} file(s).` : "No PDFs yet. Drop PDF files or a folder here.";
  },

  removePdfs() {
    const idx = selected($("pdfTable1")).sort((a, b) => b - a);
    for (const i of idx) this.records.splice(i, 1);
    this.seen = new Set(this.records.map((r) => r.path));
    this.refreshPdfs();
  },

  clearPdfs() {
    this.records = [];
    this.seen = new Set();
    this.refreshPdfs();
  },

  async editPdf(i) {
    const rec = this.records[i];
    const inputs = {};
    const form = el("div", { class: "form" }, C.FIELDS.flatMap(([k, label]) => {
      inputs[k] = el("input", { type: "text", value: rec.fields[k] ?? "", id: `ed-${k}` });
      return [el("label", { for: `ed-${k}`, text: label }), inputs[k]];
    }));
    const v = await showDialog(`Correct values: ${rec.file}`, [form, rec.note ? el("p", { class: "hint", text: rec.note }) : null].filter(Boolean),
      [["Cancel", "no"], ["Save", "save", true]]);
    if (v !== "save") return;
    rec.fields = Object.fromEntries(Object.entries(inputs).map(([k, n]) => [k, py.strip(n.value)]).filter(([, x]) => x));
    rec.source = "manual";
    rec.unreadable = !Object.keys(rec.fields).length;
    this.refreshPdfs();
    this.runDetect(true);
  },

  // ---- compare
  async runCompare() {
    if (!this.sheet) return info("Excel file needed", "Choose an Excel or CSV file first.");
    if (!this.records.length) return info("PDFs needed", "Add at least one PDF marksheet first.");
    const m = this.mapping();
    if (m.reg === null || m.reg === undefined) return info("Map the registration number", "Choose which Excel column holds the registration number. It is used to pair each PDF with its row.");
    const s = settings();
    try {
      this.result = C.compare(this.sheet.grid, this.headerRow, m, this.records, s, this.sheet.origin);
    } catch (e) {
      return info("Comparison failed", e.message);
    }
    this.runInfo = { created_at: py.nowStamp("%Y-%m-%d %H:%M:%S"), excel_file: this.book.name, sheet: this.sheet.sheet, header_row: this.headerRow, mapping: m, settings: s };
    let msg = "Comparison finished.";
    if ($("saveRun").checked) {
      try {
        const db = await getDb();
        const rid = await db.save_run(this.runInfo, this.result);
        this.runInfo.id = rid;
        msg += ` Saved as run ${rid}.`;
        await history.refresh();
      } catch (e) {
        await info("Could not save to History", e.message);
      }
    }
    setStatus(msg);
    results.fill();
    showTab("results");
  },
};

// ================================================================== page 2: Results
const OV_TAG = { [C.OV_VERIFIED]: "ok", [C.OV_MINOR]: "minor", [C.OV_REVIEW]: "review" };
const overallTag = (ov) => OV_TAG[ov] || "bad";
function shortStatus(f) {
  if (!f || f.status === C.ST_SKIPPED) return "-";
  return f.score === null || f.field === "dob" || f.status === "Match" || f.status === "Missing" ? f.status : `${f.status} ${py.fmtFixed(f.score, 0)}%`;
}

const results = {
  fill() {
    const r = page1.result;
    const tally = $("tally");
    tally.replaceChildren();
    if (!r) { fillTable($("resTable"), [], []); return; }
    const c = C.counts(r);
    for (const [n, label, color] of [
      [r.records.length, "Records", "var(--ink)"], [c[C.OV_VERIFIED] + c[C.OV_MINOR], "Verified", "#2E7A4D"],
      [c[C.OV_REVIEW], "Need review", "#93600E"], [c[C.OV_MISMATCH], "Mismatch", "#B03A2E"],
      [c[C.OV_NOTFOUND] + c[C.OV_NOREG], "Not in Excel / no reg. no.", "#B03A2E"],
    ]) tally.append(el("div", {}, el("b", { text: n, style: `color:${color}` }), el("span", { text: label })));
    const keep = { All: null, Verified: [C.OV_VERIFIED, C.OV_MINOR], "Needs review": [C.OV_REVIEW], "Mismatch / not found": [C.OV_MISMATCH, C.OV_NOTFOUND, C.OV_NOREG] }[$("resFilter").value];
    const rows = [];
    r.records.forEach((rec, i) => {
      if (keep && !keep.includes(rec.overall)) return;
      const f = rec.fields;
      rows.push({ cells: [i + 1, rec.pdf.file, rec.pdf.fields.reg ?? "", rec.excel_row ?? "", rec.overall, shortStatus(f.name), shortStatus(f.father), shortStatus(f.mother), shortStatus(f.dob), shortStatus(f.abc)], cls: overallTag(rec.overall), data: rec });
    });
    fillTable($("resTable"), ["#", "PDF file", "Reg. no.", "Excel row", "Overall", "Name", "Father", "Mother", "DOB", "ABC ID"], rows, "No records match this filter.");
    const um = r.unmatched_excel.length;
    $("unmatchedLabel").textContent = um ? `${um} Excel row(s) had no matching PDF (listed in the Excel report).` : "";
  },

  detail(r) {
    const body = [el("p", { text: `${r.pdf.file}  |  Reg. no. ${r.pdf.fields.reg ?? "-"}  |  Excel row ${r.excel_row ?? "-"}  |  ${r.overall}` })];
    if (r.note) body.push(el("p", { class: "warn", text: r.note }));
    const t = el("table", { class: "data" });
    fillTable(t, ["Field", "PDF", "Excel", "Result", "Match %", "Details"], C.FIELDS.map(([k, label]) => {
      const f = r.fields[k];
      return f ? { cells: [label, f.pdf_value, f.excel_value, f.status, f.score === null ? "" : py.fmtFixed(f.score, 0), f.note] }
        : { cells: [label, r.pdf.fields[k] ?? "", "", "-", "", ""] };
    }));
    body.push(el("div", { class: "tablewrap" }, t));
    return showDialog(`${r.pdf.file}: ${r.overall}`, body, [["Close", "ok"]]);
  },

  async export() {
    if (!page1.result) return info("Nothing to export", "Run a comparison first.");
    const buf = await build_report(ExcelJS, page1.runInfo, page1.result);
    const name = `verification_report_${stamp()}.xlsx`;
    download(buf, name);
    setStatus(`Report saved: ${name}`);
  },
};

// ================================================================== page 3: History
let dbPromise = null;
const getDb = () => (dbPromise ||= Database.open());

const history = {
  async refresh() {
    let runs = [];
    try { runs = await (await getDb()).list_runs(); } catch (e) { setStatus(e.message); }
    fillTable($("histTable"), ["Run", "Date", "Excel file", "PDFs", "Verified", "Need review", "Mismatch", "Not found"],
      runs.map((r) => ({ cells: [r.id, r.created_at, r.excel_file, r.pdf_count, r.n_verified, r.n_review, r.n_mismatch, r.n_notfound], data: r.id })),
      "No saved runs yet.");
  },
  selRun() {
    const s = selected($("histTable"));
    if (!s.length) { info("Choose a run", "Select a run in the list first."); return null; }
    return s[0];
  },
  async preview() {
    const s = selected($("histTable"));
    if (!s.length) return;
    const [, res] = await (await getDb()).load_run(s[0]);
    fillTable($("histRecTable"), ["Run", "PDF file / date", "Reg. no.", "Row", "Overall / note"],
      res.records.map((r) => ({ cells: [s[0], r.pdf.file, r.pdf.fields.reg ?? "", r.excel_row ?? "", py.strip(`${r.overall}  ${r.note}`)], cls: overallTag(r.overall) })));
  },
  async view() {
    const rid = this.selRun();
    if (rid === null) return;
    const [inf, res] = await (await getDb()).load_run(rid);
    page1.runInfo = inf;
    page1.result = res;
    results.fill();
    showTab("results");
    setStatus(`Showing saved run ${rid} (${inf.created_at}).`);
  },
  async export() {
    const rid = this.selRun();
    if (rid === null) return;
    const [inf, res] = await (await getDb()).load_run(rid);
    download(await build_report(ExcelJS, inf, res), `verification_run_${rid}.xlsx`);
    setStatus(`Report saved: verification_run_${rid}.xlsx`);
  },
  async remove() {
    const rid = this.selRun();
    if (rid === null) return;
    if (!(await ask("Delete run", `Delete run ${rid} and its results from History?`))) return;
    await (await getDb()).delete_run(rid);
    $("histRecTable").replaceChildren();
    await this.refresh();
  },
  async lookup() {
    const raw = $("lookup").value;
    const reg = C.norm_reg(raw);
    $("histRecTable").replaceChildren();
    if (!reg) return;
    const db = await getDb();
    let rows = await db.history_for_reg(py.strip(raw));
    if (!rows.length) rows = await db.history_for_reg(reg);
    fillTable($("histRecTable"), ["Run", "PDF file / date", "Reg. no.", "Row", "Overall / note"],
      rows.map((r) => ({ cells: [r.run_id, `${r.pdf_file}  (${r.created_at})`, reg, r.excel_row ?? "", py.strip(`${r.overall}  ${r.note || ""}`)], cls: overallTag(r.overall) })),
      "No saved results for this number");
  },
};

// ================================================================== page 4: Meritto vs Marksheet
const M_FILTERS = { All: null, "Problems only": C.M_PROBLEM, Correct: C.M_OK, "Still wrong / mismatch": [C.M_OLD, C.M_WRONG], "Not found": [C.M_NOPDF, C.M_NOFIELD], "Not checked": [C.M_MARKS, C.M_SKIP] };
const M_TAGS = { [C.M_CORRECT]: "ok", [C.M_MINOR]: "minor", [C.M_REVIEW]: "review", [C.M_NOFIELD]: "review", [C.M_OLD]: "bad", [C.M_WRONG]: "bad", [C.M_NOPDF]: "bad", [C.M_MARKS]: "skip", [C.M_SKIP]: "skip" };
const FIELD_CHOICES = Object.fromEntries(C.CHECKABLE.map((k) => [C.FIELD_LABELS[k], k]));

const meritto = {
  book: null, sheet: null, headerRow: -1, records: [], seen: new Set(),
  guesses: Object.fromEntries(C.ROLES.map(([k]) => [k, C.ColumnGuess()])), manual: {}, checks: [], runInfo: null,

  async loadExcel(file) {
    try {
      this.book = openBook(XLSX, new Uint8Array(await file.arrayBuffer()), file.name);
    } catch (e) {
      return info("Could not read file", `${file.name}\n\n${e.message}`);
    }
    await this.loadSheet(null);
  },

  async loadSheet(sheet) {
    try {
      this.sheet = C.load_sheet(this.book, sheet);
    } catch (e) {
      return info("Could not read file", `${this.book.name}\n\n${e.message}`);
    }
    const sd = this.sheet;
    $("mExcelLabel").textContent = `${this.book.name}  (${sd.grid.length} rows, ${sd.ncols} columns)`;
    $("mExcelLabel").classList.remove("hint");
    $("mSheet").replaceChildren(...sd.sheet_names.map((n) => el("option", { value: n, text: n })));
    $("mSheet").value = sd.sheet;
    $("mSheet").disabled = false;
    this.headerRow = C.m_detect_header_row(sd.grid);
    let warn = false;
    if (this.headerRow < 0) { this.headerRow = 0; warn = true; }
    $("mHeader").value = this.headerRow + 1;
    $("mHeader").disabled = false;
    this.manual = {};
    this.detect(false);
    setStatus(`Meritto file loaded: ${this.book.name}, header row ${this.headerRow + 1}.`);
    if (warn) await info("Header row", "Could not find the header row automatically. Set 'Header row' by hand.");
  },

  changeHeader() {
    if (!this.sheet) return;
    let n = parseInt($("mHeader").value, 10);
    if (Number.isNaN(n)) return;
    n = Math.max(1, Math.min(n, this.sheet.grid.length));
    $("mHeader").value = n;
    if (n - 1 !== this.headerRow) { this.headerRow = n - 1; this.detect(true); }
  },

  choices() {
    const out = [NOT_USED];
    if (this.sheet) {
      for (let c = 0; c < this.sheet.ncols; c++) {
        const h = this.headerRow >= 0 && this.headerRow < this.sheet.grid.length ? this.sheet.grid[this.headerRow][c] : "";
        out.push(h ? `${C.col_letter(c)}: ${Array.from(h).slice(0, 26).join("")}` : C.col_letter(c));
      }
    }
    return out;
  },

  detect(keepManual = true) {
    if (!this.sheet) return;
    if (!keepManual) this.manual = {};
    this.guesses = C.detect_columns(this.sheet.grid, this.headerRow, this.records, this.manual);
    this.refreshColumns();
  },

  refreshColumns() {
    const box = $("mColumns");
    box.replaceChildren();
    const ch = this.choices();
    for (const [key, label] of C.ROLES) {
      const g = this.guesses[key];
      const sel = el("select", { "aria-label": label, onchange: () => this.onCol(key, sel.selectedIndex) }, ch.map((c, i) => el("option", { value: i, text: c })));
      sel.selectedIndex = g.col !== null && g.col + 1 < ch.length ? g.col + 1 : 0;
      const conf = { manual: "Set manually", none: "Not found" }[g.confidence] || `Auto, ${g.confidence}`;
      box.append(el("span", { class: "lbl", text: label + (key === "enroll" || key === "expected" ? " *" : "") }), sel,
        el("span", { class: `conf conf-${g.confidence}`, text: conf, title: g.reason || null }));
    }
  },

  onCol(key, i) {
    const col = i <= 0 ? null : i - 1;
    if (col === null) {
      delete this.manual[key];
      this.guesses[key] = C.ColumnGuess();
    } else {
      for (const [other, g] of Object.entries(this.guesses)) { // one column serves one role
        if (other !== key && g.col === col) { this.guesses[other] = C.ColumnGuess(); delete this.manual[other]; }
      }
      this.manual[key] = col;
      this.guesses[key] = C.ColumnGuess(col, "manual", 1.0, "Set manually");
    }
    this.refreshColumns();
  },

  async addPdfs(items) {
    if (pdfBusy) return info("Please wait", "PDFs are still being read.");
    const errors = await readPdfs(this, items, "Meritto tab: ");
    if (errors === null) return;
    this.refreshPdfs();
    this.detect(true);
    const bad = this.records.filter((r) => r.unreadable).length;
    setStatus(`Meritto tab: ${this.records.length} PDF records loaded` + (bad ? `, ${bad} unreadable.` : "."));
    if (errors.length) await info("Some PDFs could not be opened", errors.slice(0, 10).join("\n"));
  },

  refreshPdfs() {
    fillTable($("pdfTable2"), ["File", "Reg. no.", "ABC ID", "Name", "Father", "Mother"],
      this.records.map((r, i) => ({ cells: [r.file, r.fields.reg ?? "", r.fields.abc ?? "", r.fields.name ?? "", r.fields.father ?? "", r.fields.mother ?? ""], cls: r.unreadable ? "bad" : null, data: i })));
    $("pdfCount2").textContent = this.records.length ? `${this.records.length} PDF record(s) from ${this.seen.size} file(s).` : "No PDFs yet. Drop PDF files or a folder here.";
  },

  clearPdfs() { this.records = []; this.seen = new Set(); this.refreshPdfs(); },

  async run() {
    if (!this.sheet) return info("Excel file needed", "Choose the Meritto Excel file first.");
    if (!this.records.length) return info("PDFs needed", "Add the marksheet PDFs first.");
    if (this.checks.some((c) => c.manual.length) && !(await ask("Discard manual changes?",
      "Running again re-reads the Excel and discards your manual adjustments.\nUse 'Re-check' to keep them. Continue?"))) return;
    const cols = Object.fromEntries(Object.entries(this.guesses).map(([k, g]) => [k, g.col]));
    let last = parseInt($("mLast").value || "0", 10);
    if (Number.isNaN(last)) last = 0;
    try {
      this.checks = C.build_checks(this.sheet.grid, this.headerRow, cols, last || null);
    } catch (e) {
      return info("Choose the columns", e.message);
    }
    const s = settings();
    C.run_checks(this.checks, this.records, s);
    this.runInfo = { created_at: py.nowStamp("%Y-%m-%d %H:%M:%S"), excel_file: this.book.name, sheet: this.sheet.sheet, header_row: this.headerRow, columns: cols, settings: s, pdf_count: this.records.length };
    this.fill();
    setStatus(`Meritto check finished: ${this.checks.length} checks.`);
  },

  recheck() {
    if (!this.checks.length) return this.run();
    const s = settings();
    C.run_checks(this.checks, this.records, s);
    if (this.runInfo) Object.assign(this.runInfo, { settings: s, pdf_count: this.records.length, created_at: py.nowStamp("%Y-%m-%d %H:%M:%S") });
    this.fill();
    setStatus("Re-checked with the current settings and PDFs (manual changes kept).");
  },

  fill() {
    const keep = M_FILTERS[$("mFilter").value];
    const rows = this.checks.filter((c) => !keep || keep.includes(c.status)).map((c) => ({
      cells: [c.excel_row, c.enrollment, c.student, C.field_label(c), c.expected, c.pdf_value, c.status + (c.manual.length ? "  (manual)" : ""), c.pdf_file],
      cls: M_TAGS[c.status] || "skip", data: c,
    }));
    fillTable($("mTable"), ["Row", "Enrollment", "Student", "Field", "Expected (Meritto)", "In PDF", "Result", "PDF file"], rows,
      this.checks.length ? "No checks match this filter." : "Run the check to see results.");
    const n = C.m_counts(this.checks);
    $("mTally").textContent = Object.entries(n).filter(([, v]) => v).map(([k, v]) => `${k}: ${v}`).join("   ") || "No checks yet.";
  },

  async detail(c) {
    const kv = el("dl", { class: "kv" }, [
      ["Excel row", c.excel_row], ["Enrollment no.", c.enrollment], ["Student", c.student], ["Remarks (Excel)", c.remark],
      ["Old value (marksheet)", c.claimed || "-"], ["Result", c.status], ["Match %", c.score === null ? "" : py.fmtFixed(c.score, 0)],
      ["Details", c.note], ["Manual changes", c.manual.join("; ") || "-"],
    ].flatMap(([k, v]) => [el("dt", { text: k }), el("dd", { text: disp(v) })]));
    const files = [...new Set(this.records.map((p) => p.path))]
      .map((p) => [p, this.records.find((r) => r.path === p).file])
      .sort((a, b) => py.cmpStr(a[1].toLowerCase(), b[1].toLowerCase()));
    const fSel = el("select", { id: "md-field" }, [el("option", { value: "", text: "" }), ...Object.keys(FIELD_CHOICES).map((l) => el("option", { value: l, text: l }))]);
    fSel.value = C.FIELD_LABELS[c.field] || "";
    const exp = el("input", { type: "text", id: "md-exp", value: c.expected });
    const forced = c.force_pdf_path ? (files.find(([p]) => p === c.force_pdf_path) || [null, null])[1] : null;
    const pSel = el("select", { id: "md-pdf" }, [AUTO, ...files.map(([, n]) => n)].map((n) => el("option", { value: n, text: n })));
    pSel.value = forced || AUTO;
    const val = el("input", { type: "text", id: "md-val", value: c.force_pdf_value ?? "" });
    const form = el("div", { class: "form" },
      el("label", { for: "md-field", text: "Field to check" }), fSel,
      el("label", { for: "md-exp", text: "Expected value" }), exp,
      el("label", { for: "md-pdf", text: "PDF to use" }), pSel,
      el("label", { for: "md-val", text: "Value in PDF (blank = read it)" }), val);
    const hint = el("p", { class: "hint", text: c.force_pdf_value === null || c.force_pdf_value === undefined ? `Automatic value read from the PDF: ${c.pdf_value || "-"}` : "A manual PDF value is in use." });
    const v = await showDialog(`Row ${c.excel_row}: ${c.student || c.enrollment}`, [kv, el("h2", { text: "Adjust by hand" }), form, hint],
      [["Cancel", "no"], ["Save and re-check", "save", true]]);
    if (v !== "save") return;
    const changes = [];
    const fld = FIELD_CHOICES[fSel.value];
    if (fld && fld !== c.field) { changes.push(`field set to ${fSel.value}`); c.field = fld; c.kind = "field"; }
    if (py.strip(exp.value) !== c.expected) {
      changes.push(`expected set to "${py.strip(exp.value)}"`);
      c.expected = py.strip(exp.value);
      c.kind = c.expected && c.field ? "field" : c.kind;
    }
    const name = pSel.value;
    const path = name === AUTO ? null : (files.find(([, n]) => n === name) || [null])[0];
    if (path !== (c.force_pdf_path ?? null)) { changes.push(path ? `PDF set to ${name}` : "PDF back to automatic"); c.force_pdf_path = path; }
    const fv = py.strip(val.value) || null;
    if (fv !== (c.force_pdf_value ?? null)) { changes.push(fv ? `PDF value set to "${fv}"` : "PDF value back to automatic"); c.force_pdf_value = fv; }
    if (changes.length) {
      c.manual.push(...changes);
      if (c.kind === "field" && !c.field) c.kind = "unsupported";
      C.evaluate(c, this.records, settings());
      this.fill();
    }
  },

  async export() {
    if (!this.checks.length) return info("Nothing to export", "Run the check first.");
    const name = `meritto_report_${stamp()}.xlsx`;
    download(await build_meritto_report(ExcelJS, this.runInfo || {}, this.checks), name);
    setStatus(`Report saved: ${name}`);
  },
};

// ================================================================== wiring
function showTab(name) {
  for (const b of document.querySelectorAll(".tabs button")) b.setAttribute("aria-selected", String(b.dataset.tab === name));
  for (const s of document.querySelectorAll(".tab")) s.hidden = s.id !== `tab-${name}`;
}

function onDbl(table, fn) {
  table.addEventListener("dblclick", guard(async (e) => {
    const tr = e.target.closest("tbody tr");
    if (tr && tr._data !== undefined) await fn(tr._data);
  }));
}

async function init() {
  await loadLibraries();
  for (const b of document.querySelectorAll(".tabs button")) b.addEventListener("click", () => showTab(b.dataset.tab));
  syncShared();
  thresh("match");
  $("mFilter").replaceChildren(...Object.keys(M_FILTERS).map((k) => el("option", { text: k })));

  // page 1
  $("excelInput").addEventListener("change", guard(async (e) => { const f = e.target.files[0]; e.target.value = ""; if (f) await page1.loadExcel(f); }));
  $("sheetSelect").addEventListener("change", guard(() => { if (page1.sheet && $("sheetSelect").value !== page1.sheet.sheet) page1.loadSheet($("sheetSelect").value); }));
  $("headerRow").addEventListener("change", guard(() => page1.changeHeader()));
  $("detectAgain").addEventListener("click", guard(() => page1.runDetect(true)));
  $("detectReset").addEventListener("click", guard(() => page1.runDetect(false)));
  $("pdfInput1").addEventListener("change", guard((e) => page1.addPdfs(pdfItemsFromInput(e.target, false))));
  $("folderInput1").addEventListener("change", guard((e) => page1.addPdfs(pdfItemsFromInput(e.target, true))));
  $("removePdf1").addEventListener("click", guard(() => page1.removePdfs()));
  $("clearPdf1").addEventListener("click", guard(() => page1.clearPdfs()));
  $("runCompare").addEventListener("click", guard(() => page1.runCompare()));
  selectable($("pdfTable1"));
  onDbl($("pdfTable1"), (i) => page1.editPdf(i));
  setupDrop(document.querySelector('[data-drop="excel"]'), async (dt) => { if (dt.files[0]) await page1.loadExcel(dt.files[0]); });
  setupDrop(document.querySelector('[data-drop="pdf1"]'), async (dt) => page1.addPdfs(await itemsFromDrop(dt)));

  // page 2
  $("resFilter").addEventListener("change", () => results.fill());
  $("rerun").addEventListener("click", guard(() => page1.runCompare()));
  $("exportReport").addEventListener("click", guard(() => results.export()));
  onDbl($("resTable"), (r) => results.detail(r));

  // page 3
  selectable($("histTable"), guard(() => history.preview()));
  onDbl($("histTable"), () => history.view());
  $("histView").addEventListener("click", guard(() => history.view()));
  $("histExport").addEventListener("click", guard(() => history.export()));
  $("histDelete").addEventListener("click", guard(() => history.remove()));
  $("histRefresh").addEventListener("click", guard(() => history.refresh()));
  $("lookupBtn").addEventListener("click", guard(() => history.lookup()));
  $("lookup").addEventListener("keydown", (e) => { if (e.key === "Enter") guard(() => history.lookup())(); });

  // page 4
  $("mExcelInput").addEventListener("change", guard(async (e) => { const f = e.target.files[0]; e.target.value = ""; if (f) await meritto.loadExcel(f); }));
  $("mSheet").addEventListener("change", guard(() => { if (meritto.sheet && $("mSheet").value !== meritto.sheet.sheet) return meritto.loadSheet($("mSheet").value); }));
  $("mHeader").addEventListener("change", guard(() => meritto.changeHeader()));
  $("mReset").addEventListener("click", guard(() => meritto.detect(false)));
  $("pdfInput2").addEventListener("change", guard((e) => meritto.addPdfs(pdfItemsFromInput(e.target, false))));
  $("folderInput2").addEventListener("change", guard((e) => meritto.addPdfs(pdfItemsFromInput(e.target, true))));
  $("clearPdf2").addEventListener("click", guard(() => meritto.clearPdfs()));
  $("mRun").addEventListener("click", guard(() => meritto.run()));
  $("mRecheck").addEventListener("click", guard(() => meritto.recheck()));
  $("mFilter").addEventListener("change", () => meritto.fill());
  $("mExport").addEventListener("click", guard(() => meritto.export()));
  onDbl($("mTable"), (c) => meritto.detail(c));
  setupDrop(document.querySelector('[data-drop="mexcel"]'), async (dt) => { if (dt.files[0]) await meritto.loadExcel(dt.files[0]); });
  setupDrop(document.querySelector('[data-drop="pdf2"]'), async (dt) => meritto.addPdfs(await itemsFromDrop(dt)));

  page1.refreshPdfs();
  meritto.refreshPdfs();
  meritto.fill();
  results.fill();
  await history.refresh();
  window.addEventListener("beforeunload", (e) => { if (pdfBusy) e.preventDefault(); });
}

init().catch((e) => {
  const b = $("loadError");
  b.hidden = false;
  b.textContent = `The page failed to start: ${e.message}`;
});

// exposed for automated browser tests only
globalThis.__mv = { page1, meritto, results, history, settings };
