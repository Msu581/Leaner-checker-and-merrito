/*
 * storage.js - port of storage.py. The SQLite file becomes an IndexedDB database in this browser
 * (nothing leaves the machine). The same four tables are kept, and save_run()/load_run() apply the
 * same transformation as the Python code, so a run reopened from History looks exactly as it did
 * when reopened from the SQLite database.
 */
import * as C from "./core.js";

const DB_NAME = "marksheet_results";
const STORES = ["runs", "records", "field_results", "unmatched_excel"];

function req(r) {
  return new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
}
function done(tx) {
  return new Promise((res, rej) => { tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error); tx.onabort = () => rej(tx.error || new Error("Transaction aborted")); });
}

export class Database {
  static async open(idb = globalThis.indexedDB) {
    if (!idb) throw new Error("This browser does not allow local storage (IndexedDB), so runs cannot be saved.");
    const open = idb.open(DB_NAME, 1);
    open.onupgradeneeded = () => {
      const db = open.result;
      db.createObjectStore("runs", { keyPath: "id", autoIncrement: true });
      db.createObjectStore("records", { keyPath: "id", autoIncrement: true }).createIndex("run_id", "run_id");
      db.createObjectStore("field_results", { keyPath: "id", autoIncrement: true }).createIndex("record_id", "record_id");
      db.createObjectStore("unmatched_excel", { keyPath: "id", autoIncrement: true }).createIndex("run_id", "run_id");
    };
    const d = new Database();
    d.db = await req(open);
    return d;
  }

  // ---- write -----------------------------------------------------------
  async save_run(info, result) {
    const c = C.counts(result);
    const tx = this.db.transaction(STORES, "readwrite");
    const S = (n) => tx.objectStore(n);
    const run_id = await req(S("runs").add({
      created_at: info.created_at || new Date().toISOString().slice(0, 19).replace("T", " "),
      excel_file: info.excel_file ?? "", sheet: info.sheet ?? "", header_row: info.header_row ?? 0,
      mapping: JSON.stringify(info.mapping || {}), settings: JSON.stringify(info.settings || {}),
      pdf_count: result.records.length,
      n_verified: c[C.OV_VERIFIED] + c[C.OV_MINOR], n_review: c[C.OV_REVIEW], n_mismatch: c[C.OV_MISMATCH],
      n_notfound: c[C.OV_NOTFOUND] + c[C.OV_NOREG],
    }));
    for (const r of result.records) {
      const rid = await req(S("records").add({
        run_id, pdf_file: r.pdf.file, pdf_path: r.pdf.path, page: r.pdf.page, pages: r.pdf.pages, source: r.pdf.source,
        pdf_reg: r.pdf.fields.reg ?? "", excel_row: r.excel_row, overall: r.overall, note: r.note,
      }));
      const rows = Object.keys(r.fields).length
        ? Object.values(r.fields).map((f) => [f.field, f.pdf_value, f.excel_value, f.status, f.score, f.note])
        : ["reg", ...C.COMPARE_FIELDS].map((k) => [k, r.pdf.fields[k] ?? "", "", "Not compared", null, ""]); // not paired: keep what the PDF said
      for (const [field, pdf_value, excel_value, status, score, note] of rows) {
        S("field_results").add({ record_id: rid, field, pdf_value, excel_value, status, score, note });
      }
    }
    for (const u of result.unmatched_excel) S("unmatched_excel").add({ run_id, excel_row: u.row, reg: u.reg, name: u.name });
    await done(tx);
    return run_id;
  }

  async delete_run(run_id) {
    const tx = this.db.transaction(STORES, "readwrite");
    const recs = await req(tx.objectStore("records").index("run_id").getAll(run_id));
    for (const r of recs) {
      const frs = await req(tx.objectStore("field_results").index("record_id").getAllKeys(r.id));
      for (const k of frs) tx.objectStore("field_results").delete(k);
      tx.objectStore("records").delete(r.id);
    }
    const um = await req(tx.objectStore("unmatched_excel").index("run_id").getAllKeys(run_id));
    for (const k of um) tx.objectStore("unmatched_excel").delete(k);
    tx.objectStore("runs").delete(run_id);
    await done(tx);
  }

  // ---- read ------------------------------------------------------------
  async list_runs() {
    const all = await req(this.db.transaction("runs").objectStore("runs").getAll());
    return all.sort((a, b) => b.id - a.id);
  }

  /** Returns [info, RunResult] rebuilt like storage.Database.load_run(). */
  async load_run(run_id) {
    const tx = this.db.transaction(STORES);
    const row = await req(tx.objectStore("runs").get(run_id));
    if (!row) throw new Error(`No run with id ${run_id}`);
    const info = { ...row, mapping: JSON.parse(row.mapping || "{}"), settings: JSON.parse(row.settings || "{}") };
    const recs = (await req(tx.objectStore("records").index("run_id").getAll(run_id))).sort((a, b) => a.id - b.id);
    const records = [];
    for (const r of recs) {
      const frows = (await req(tx.objectStore("field_results").index("record_id").getAll(r.id))).sort((a, b) => a.id - b.id);
      let frs = {};
      for (const f of frows) frs[f.field] = C.FieldResult(f.field, f.pdf_value || "", f.excel_value || "", f.status, f.score, f.note || "");
      const pdf_fields = Object.fromEntries(Object.entries(frs).filter(([, v]) => v.pdf_value).map(([k, v]) => [k, v.pdf_value]));
      const pdf = C.PdfRecord(r.pdf_file, r.pdf_path || "", r.page, r.pages || 1, pdf_fields, r.source || "text");
      if (r.overall === C.OV_NOTFOUND || r.overall === C.OV_NOREG) frs = {};
      records.push(C.RecordResult(pdf, r.overall, r.excel_row, frs, r.note || ""));
    }
    const um = (await req(tx.objectStore("unmatched_excel").index("run_id").getAll(run_id))).sort((a, b) => a.id - b.id);
    return [info, C.RunResult(records, um.map((u) => ({ row: u.excel_row, reg: u.reg, name: u.name })))];
  }

  /** Every past result for one registration number, newest run first. */
  async history_for_reg(reg) {
    const tx = this.db.transaction(["runs", "records"]);
    const runs = new Map((await req(tx.objectStore("runs").getAll())).map((r) => [r.id, r]));
    const recs = (await req(tx.objectStore("records").getAll())).filter((x) => x.pdf_reg === reg);
    recs.sort((a, b) => (b.run_id - a.run_id) || (a.id - b.id));
    return recs.map((rec) => {
      const r = runs.get(rec.run_id) || {};
      return { run_id: rec.run_id, created_at: r.created_at, excel_file: r.excel_file, pdf_file: rec.pdf_file, excel_row: rec.excel_row, overall: rec.overall, note: rec.note };
    });
  }
}
