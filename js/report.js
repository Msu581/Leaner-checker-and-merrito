/*
 * report.js - port of report.py (formatted Excel reports) using ExcelJS.
 *   build_report():         Summary | Comparison | Needs attention | Unmatched Excel
 *   build_meritto_report(): Summary | Results | Needs attention
 * Layout, formulas, colours, widths, merges, filters and dropdowns follow report.py cell for cell.
 */
import * as C from "./core.js";
import * as py from "./py.js";

const FONT = "Arial";
const INK = "16262C";
const HEAD = "0F5E6E";
const GROUP = "16262C";
const THIN = { style: "thin", color: { argb: "FFC9D4D3" } };
const BORDER = { left: THIN, right: THIN, top: THIN, bottom: THIN };

const PALETTE = {
  [C.ST_MATCH]: ["C6EFCE", "006100"], [C.ST_CLOSE]: ["E2EFDA", "375623"], [C.ST_REVIEW]: ["FFEB9C", "7F6000"],
  [C.ST_MISSING]: ["FFEB9C", "7F6000"], [C.ST_MISMATCH]: ["FFC7CE", "9C0006"], [C.ST_SKIPPED]: ["EDEDED", "595959"],
  [C.OV_VERIFIED]: ["C6EFCE", "006100"], [C.OV_MINOR]: ["E2EFDA", "375623"], [C.OV_REVIEW]: ["FFEB9C", "7F6000"],
  [C.OV_MISMATCH]: ["FFC7CE", "9C0006"], [C.OV_NOTFOUND]: ["FFC7CE", "9C0006"], [C.OV_NOREG]: ["FFC7CE", "9C0006"],
};

// ExcelJS treats a width of exactly 9 as "default" and does not write it; this prints as 9 in Excel.
const W9 = 9.000000001;
const argb = (hex) => ({ argb: "FF" + hex });
const fill = (hex) => ({ type: "pattern", pattern: "solid", fgColor: argb(hex) });
const _f = (bold = false, color = INK, size = 10) => ({ name: FONT, bold, color: argb(color), size });

/** Wraps an ExcelJS sheet and remembers what was written, for report.py's column-width rule. */
class Sheet {
  constructor(ws) { this.ws = ws; this.vals = new Map(); this.maxCol = 0; }
  cell(r, c) {
    this.maxCol = Math.max(this.maxCol, c);
    return this.ws.getCell(r, c);
  }
  /** report._widths(): longest text in the first 400 rows + 3, between 8 and the cap (default 38). */
  widths(caps, first, last) {
    for (let ci = 1; ci <= this.maxCol; ci++) {
      let longest = 0;
      for (let r = first; r <= Math.min(last, first + 400); r++) {
        const v = this.vals.get(`${r},${ci}`);
        if (v !== null && v !== undefined) {
          for (const part of pyStr(v).split("\n")) longest = Math.max(longest, py.len(part));
        }
      }
      const w = Math.max(8, Math.min(longest + 3, caps[ci] ?? 38));
      this.ws.getColumn(ci).width = w === 9 ? W9 : w;
    }
  }
}

/** str(v) as Python would print the value (scores are floats in Python). */
function pyStr(v) {
  if (v instanceof PyFloat) return py.floatRepr(v.v);
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : py.floatRepr(v);
  return String(v);
}
class PyFloat { constructor(v) { this.v = v; } }
const fl = (v) => (v === null || v === undefined ? null : new PyFloat(v));
const unwrap = (v) => (v instanceof PyFloat ? v.v : v);

function setv(sh, r, c, v) {
  const cl = sh.cell(r, c);
  sh.vals.set(`${r},${c}`, v);
  sh.maxCol = Math.max(sh.maxCol, c);
  const raw = unwrap(v);
  if (raw === null || raw === undefined) cl.value = null;
  else if (typeof raw === "string" && raw.startsWith("=")) cl.value = { formula: raw.slice(1) };
  else cl.value = raw;
  return cl;
}

function paint(cell, key, pal = PALETTE) {
  const [f, color] = pal[key];
  cell.fill = fill(f);
  cell.font = _f(true, color);
}

function body(cell, wrap = true, center = false) {
  cell.font = _f();
  cell.border = BORDER;
  cell.alignment = { vertical: "top", wrapText: wrap, horizontal: center ? "center" : "left" };
}

function header(sh, row, headers, start = 1) {
  headers.forEach((h, i) => {
    const c = setv(sh, row, start + i, h);
    c.font = _f(true, "FFFFFF");
    c.fill = fill(HEAD);
    c.border = BORDER;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  sh.ws.getRow(row).height = 30;
}

function page(sh) {
  sh.ws.pageSetup = { ...sh.ws.pageSetup, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
}

function views(frozenCell) {
  const v = { showGridLines: false };
  if (frozenCell) {
    const m = /^([A-Z]+)(\d+)$/.exec(frozenCell);
    Object.assign(v, { state: "frozen", xSplit: C.col_index(m[1]), ySplit: Number(m[2]) - 1, topLeftCell: frozenCell });
  }
  return [v];
}

function mode_text(s) {
  if (s.mode === "strict") return "Strict (only case, extra spaces and punctuation ignored; thresholds not used)";
  return "Normal (fuzzy: thresholds, initials and word order allowed)";
}

const thr = (v) => (v === undefined ? "" : typeof v === "number" ? py.floatRepr(v) : String(v)); // settings hold floats

function addValidation(ws, range, list) {
  ws.dataValidations.add(range, { type: "list", allowBlank: true, formulae: [`"${list}"`] });
}

// ====================================================================== Excel vs Marksheet
export async function build_report(ExcelJS, info, result) {
  const wb = new ExcelJS.Workbook();
  wb.calcProperties.fullCalcOnLoad = true;
  const sum = new Sheet(wb.addWorksheet("Summary", { views: views(null) }));
  const cmp = new Sheet(wb.addWorksheet("Comparison", { views: views("D3") }));
  const na = new Sheet(wb.addWorksheet("Needs attention", { views: views("A2") }));
  const ue = new Sheet(wb.addWorksheet("Unmatched Excel", { views: views("A2") }));

  // ---------------------------------------------------------------- Comparison
  const base = ["#", "PDF file", "Page", "Reg. no. (PDF)", "Excel row", "Overall result"];
  const sub = ["PDF value", "Excel value", "Result", "Match %"];
  cmp.ws.mergeCells(1, 1, 1, base.length);
  setv(cmp, 1, 1, "Record");
  let cols = base.length;
  const groups = [[1, base.length, "Record"]];
  for (const k of C.COMPARE_FIELDS) { groups.push([cols + 1, cols + sub.length, C.FIELD_LABELS[k]]); cols += sub.length; }
  for (const [a, b, label] of groups) {
    if (b > a && a !== 1) cmp.ws.mergeCells(1, a, 1, b);
    const c = setv(cmp, 1, a, label);
    for (let ci = a; ci <= b; ci++) { const x = cmp.cell(1, ci); x.fill = fill(GROUP); x.border = BORDER; }
    c.font = _f(true, "FFFFFF", 11);
    c.alignment = { horizontal: "center", vertical: "middle" };
  }
  cmp.ws.getRow(1).height = 22;
  header(cmp, 2, [...base, ...C.COMPARE_FIELDS.flatMap(() => sub)]);

  const first = 3;
  result.records.forEach((rec, i) => {
    const n = i + 1, r = first + n - 1;
    const row = [n, rec.pdf.file, rec.pdf.page, rec.pdf.fields.reg ?? "", rec.excel_row, rec.overall];
    for (const k of C.COMPARE_FIELDS) {
      const f = rec.fields[k];
      if (f) row.push(f.pdf_value, f.excel_value, f.status, fl(f.score));
      else row.push(rec.pdf.fields[k] ?? "", "", C.ST_SKIPPED, null);
    }
    row.forEach((v, j) => body(setv(cmp, r, j + 1, v), true, [1, 3, 5].includes(j + 1)));
    paint(cmp.cell(r, 6), rec.overall);
    for (let j = 0; j < C.COMPARE_FIELDS.length; j++) {
      const rc = base.length + 1 + j * sub.length;
      const st = cmp.cell(r, rc + 2);
      paint(st, row[rc + 1]);
      st.alignment = { horizontal: "center", vertical: "top" };
      const sc = cmp.cell(r, rc + 3);
      sc.numFmt = '0"%"';
      sc.alignment = { horizontal: "center", vertical: "top" };
    }
  });
  const last = first + Math.max(result.records.length, 1) - 1;
  cmp.ws.autoFilter = `A2:${C.col_letter(cmp.maxCol - 1)}${last}`;
  cmp.widths({ 1: 6, 2: 34, 4: 18, 6: 28 }, 3, last);
  for (let j = 0; j < C.COMPARE_FIELDS.length; j++) {
    const rc = base.length + 1 + j * sub.length;
    cmp.ws.getColumn(rc + 2).width = 13;
    cmp.ws.getColumn(rc + 3).width = W9;
  }
  page(cmp);
  cmp.ws.pageSetup.printTitlesRow = "1:2";

  // ---------------------------------------------------------------- Summary
  const ws = sum;
  setv(ws, 1, 1, "Marksheet verification report").font = _f(true, HEAD, 16);
  setv(ws, 2, 1, "PDF marksheets compared against an Excel record").font = _f(false, "5A6B71", 10);
  const kv = (r, k, v) => {
    const a = setv(ws, r, 1, k), b = setv(ws, r, 2, v);
    a.font = _f(true); b.font = _f();
    a.alignment = { vertical: "top", wrapText: true };
    b.alignment = { vertical: "top", wrapText: true };
    a.border = BORDER; b.border = BORDER;
  };
  const s = info.settings || {};
  let r = 4;
  setv(ws, r, 1, "Run details").font = _f(true, HEAD, 12);
  r += 1;
  const hr = (info.header_row ?? -1) + 1;
  for (const [k, v] of [
    ["Generated", py.nowStamp("%d-%m-%Y %H:%M")],
    ["Run date", info.created_at ?? ""],
    ["Excel file", info.excel_file ?? ""],
    ["Sheet", info.sheet ?? ""],
    ["Header row", hr || "none"],
    ["Close-match threshold", `${thr(s.match_threshold)}%  (name score at or above this is "Close match")`],
    ["Review threshold", `${thr(s.review_threshold)}%  (at or above: "Review"; below: "Mismatch")`],
    ["Date order", (s.day_first ?? true) ? "Day-Month-Year" : "Month-Day-Year"],
    ["Titles ignored (Mr, Smt, ...)", (s.ignore_titles ?? true) ? "Yes" : "No"],
    ["Name matching mode", mode_text(s)],
  ]) { kv(r, k, v); r += 1; }

  r += 1;
  setv(ws, r, 1, "Column mapping").font = _f(true, HEAD, 12);
  r += 1;
  header(ws, r, ["Field", "Excel column"]);
  r += 1;
  for (const [key, label] of C.MAP_FIELDS) {
    const col = (info.mapping || {})[key];
    kv(r, label, Number.isInteger(col) ? C.col_letter(col) : "Not mapped");
    r += 1;
  }

  r += 1;
  setv(ws, r, 1, "Results").font = _f(true, HEAD, 12);
  r += 1;
  header(ws, r, ["Result", "Records", "Share"]);
  r += 1;
  const rng = `Comparison!$F$3:$F$${last}`;
  const t0 = r, total_row = t0 + C.OVERALL_ORDER.length;
  for (const ov of C.OVERALL_ORDER) {
    const a = setv(ws, r, 1, ov), b = setv(ws, r, 2, `=COUNTIF(${rng},A${r})`), c = setv(ws, r, 3, `=IF($B$${total_row}=0,0,B${r}/$B$${total_row})`);
    for (const x of [a, b, c]) body(x, false);
    paint(a, ov);
    b.alignment = { horizontal: "center" };
    c.alignment = { horizontal: "center" };
    c.numFmt = "0.0%";
    r += 1;
  }
  {
    const a = setv(ws, r, 1, "Total records"), b = setv(ws, r, 2, `=SUM(B${t0}:B${r - 1})`), c = setv(ws, r, 3, `=IF(B${r}=0,0,SUM(C${t0}:C${r - 1}))`);
    for (const x of [a, b, c]) { body(x, false); x.font = _f(true); }
    b.alignment = { horizontal: "center" };
    c.alignment = { horizontal: "center" };
    c.numFmt = "0.0%";
  }
  r += 2;

  setv(ws, r, 1, "Results by field").font = _f(true, HEAD, 12);
  r += 1;
  const stat_list = [C.ST_MATCH, C.ST_CLOSE, C.ST_REVIEW, C.ST_MISSING, C.ST_MISMATCH];
  header(ws, r, ["Field", ...stat_list]);
  r += 1;
  C.COMPARE_FIELDS.forEach((k, j) => {
    const a = setv(ws, r, 1, C.FIELD_LABELS[k]);
    body(a, false);
    a.font = _f(true);
    const status_col = C.col_letter(base.length + j * sub.length + 2);
    stat_list.forEach((st, i) => body(setv(ws, r, 2 + i, `=COUNTIF(Comparison!$${status_col}$3:$${status_col}$${last},"${st}")`), false, true));
    r += 1;
  });

  r += 1;
  setv(ws, r, 1, "How to read this").font = _f(true, HEAD, 12);
  r += 1;
  for (const [k, v] of [
    [C.ST_MATCH, "Identical after ignoring case, spacing and punctuation."],
    [C.ST_CLOSE, "Name is similar enough (at or above the close-match threshold). Accepted automatically."],
    [C.ST_REVIEW, "Between the review and close-match thresholds, empty in one file, or day/month swapped. Check by hand."],
    [C.ST_MISMATCH, "Clearly different. Registration number alone does not guarantee the record is right."],
    [C.ST_SKIPPED, "That column was not mapped, so the field was not compared."],
  ]) {
    const a = setv(ws, r, 1, k), b = setv(ws, r, 2, v);
    body(a, false); body(b);
    paint(a, k);
    ws.ws.mergeCells(r, 2, r, 6);
    ws.ws.getRow(r).height = 28;
    r += 1;
  }
  ws.ws.getColumn(1).width = 34;
  ws.ws.getColumn(2).width = 44;
  for (let ci = 3; ci <= 6; ci++) ws.ws.getColumn(ci).width = 14;
  page(ws);

  // ---------------------------------------------------------------- Needs attention
  const heads = ["PDF file", "Page", "Reg. no. (PDF)", "Excel row", "Field", "PDF value", "Excel value",
    "Result", "Match %", "Details", "Reviewer decision", "Reviewer comment"];
  header(na, 1, heads);
  r = 2;
  for (const rec of result.records) {
    const probs = [];
    if (rec.overall === C.OV_NOTFOUND || rec.overall === C.OV_NOREG) {
      probs.push(["Registration no.", rec.pdf.fields.reg ?? "", "", rec.overall, null, rec.note]);
    } else {
      for (const k of C.COMPARE_FIELDS) {
        const f = rec.fields[k];
        if (f && [C.ST_REVIEW, C.ST_MISMATCH, C.ST_MISSING].includes(f.status)) probs.push([C.FIELD_LABELS[k], f.pdf_value, f.excel_value, f.status, f.score, f.note]);
      }
      if (rec.note && !probs.length) probs.push(["Record", "", "", rec.overall, null, rec.note]);
    }
    for (const [fld, pv, ev, st, sc, note] of probs) {
      const vals = [rec.pdf.file, rec.pdf.page, rec.pdf.fields.reg ?? "", rec.excel_row, fld, pv, ev, st, fl(sc), note, "", ""];
      vals.forEach((v, j) => body(setv(na, r, j + 1, v), true, [2, 4].includes(j + 1)));
      paint(na.cell(r, 8), PALETTE[st] ? st : C.OV_MISMATCH);
      na.cell(r, 8).alignment = { horizontal: "center", vertical: "top", wrapText: true };
      na.cell(r, 9).numFmt = '0"%"';
      na.cell(r, 9).alignment = { horizontal: "center", vertical: "top" };
      r += 1;
    }
  }
  if (r === 2) setv(na, 2, 1, "Nothing needs attention.").font = _f(true, "006100");
  else {
    addValidation(na.ws, `K2:K${r - 1}`, "Accepted,Corrected in Excel,Corrected in PDF,Rejected");
    na.ws.autoFilter = `A1:L${r - 1}`;
  }
  na.widths({ 1: 34, 10: 46, 12: 30 }, 2, r);
  na.ws.getColumn(11).width = 20;
  page(na);

  // ---------------------------------------------------------------- Unmatched Excel
  header(ue, 1, ["Excel row", "Reg. no.", "Student name", "Comment"]);
  result.unmatched_excel.forEach((u, i) => {
    [u.row, u.reg, u.name, "No PDF supplied for this row"].forEach((v, j) => body(setv(ue, i + 2, j + 1, v), true, j === 0));
  });
  if (!result.unmatched_excel.length) setv(ue, 2, 1, "Every Excel row had a PDF.").font = _f(true, "006100");
  ue.widths({}, 2, result.unmatched_excel.length + 1);
  page(ue);

  return wb.xlsx.writeBuffer();
}

// ====================================================================== Meritto vs Marksheet
export async function build_meritto_report(ExcelJS, info, checks) {
  const pal = {
    [C.M_CORRECT]: PALETTE[C.ST_MATCH], [C.M_MINOR]: PALETTE[C.ST_CLOSE], [C.M_REVIEW]: PALETTE[C.ST_REVIEW],
    [C.M_OLD]: PALETTE[C.ST_MISMATCH], [C.M_WRONG]: PALETTE[C.ST_MISMATCH], [C.M_NOFIELD]: PALETTE[C.ST_REVIEW],
    [C.M_NOPDF]: PALETTE[C.ST_MISMATCH], [C.M_MARKS]: PALETTE[C.ST_SKIPPED], [C.M_SKIP]: PALETTE[C.ST_SKIPPED],
  };
  const mpaint = (cell, st) => {
    const [f, color] = pal[st] || PALETTE[C.ST_SKIPPED];
    cell.fill = fill(f);
    cell.font = _f(true, color);
  };
  const heads = ["#", "Excel row", "Enrollment no.", "Student", "Field", "Expected (Meritto)", "Old value (marksheet)",
    "Value in PDF", "Result", "Match %", "PDF file", "Page", "Details", "Remarks (Excel)", "Manually adjusted"];
  const row_vals = (n, c) => [n, c.excel_row, c.enrollment, c.student, C.field_label(c), c.expected, c.claimed, c.pdf_value, c.status,
    fl(c.score), c.pdf_file, c.pdf_page, c.note, c.remark, c.manual && c.manual.length ? c.manual.join("; ") : ""];

  const wb = new ExcelJS.Workbook();
  wb.calcProperties.fullCalcOnLoad = true;
  const ws = new Sheet(wb.addWorksheet("Summary", { views: views(null) }));
  const res = new Sheet(wb.addWorksheet("Results", { views: views("D2") }));
  const na = new Sheet(wb.addWorksheet("Needs attention", { views: views("D2") }));

  header(res, 1, heads);
  checks.forEach((c, i) => {
    const n = i + 1;
    row_vals(n, c).forEach((v, j) => body(setv(res, n + 1, j + 1, v), true, [1, 2, 10, 12].includes(j + 1)));
    mpaint(res.cell(n + 1, 9), c.status);
    res.cell(n + 1, 10).numFmt = '0"%"';
  });
  const last = Math.max(checks.length, 1) + 1;
  res.ws.autoFilter = `A1:${C.col_letter(heads.length - 1)}${last}`;
  res.widths({ 1: 6, 4: 24, 6: 30, 7: 30, 8: 30, 9: 22, 11: 30, 13: 50, 14: 40 }, 2, last);
  page(res);

  header(na, 1, [...heads, "Reviewer decision", "Reviewer comment"]);
  let r = 2;
  checks.forEach((c, i) => {
    if (!C.M_PROBLEM.includes(c.status)) return;
    [...row_vals(i + 1, c), "", ""].forEach((v, j) => body(setv(na, r, j + 1, v), true, [1, 2, 10, 12].includes(j + 1)));
    mpaint(na.cell(r, 9), c.status);
    na.cell(r, 10).numFmt = '0"%"';
    r += 1;
  });
  if (r === 2) setv(na, 2, 1, "Nothing needs attention.").font = _f(true, "006100");
  else {
    addValidation(na.ws, `P2:P${r - 1}`, "Accepted,Marksheet re-issued,Meritto corrected,Rejected");
    na.ws.autoFilter = `A1:Q${r - 1}`;
  }
  na.widths({ 1: 6, 4: 24, 6: 30, 7: 30, 8: 30, 9: 22, 11: 30, 13: 50, 14: 40, 16: 22, 17: 30 }, 2, r);
  page(na);

  setv(ws, 1, 1, "Meritto vs Marksheet report").font = _f(true, HEAD, 16);
  setv(ws, 2, 1, "Checks that each marksheet now shows the value given in the Meritto/TR record").font = _f(false, "5A6B71", 10);
  const s = info.settings || {};
  const cols = info.columns || {};
  r = 4;
  const hr = (info.header_row ?? -1) + 1;
  for (const [k, v] of [
    ["Generated", py.nowStamp("%d-%m-%Y %H:%M")], ["Run date", info.created_at ?? ""],
    ["Excel file", info.excel_file ?? ""], ["Sheet", info.sheet ?? ""],
    ["Header row", hr || "none"], ["PDF records", info.pdf_count ?? ""],
    ["Name matching mode", mode_text(s)],
    ["Titles ignored (Mr, Smt, ...)", (s.ignore_titles ?? true) ? "Yes" : "No"],
    ...C.ROLES.map(([key, label]) => [label, Number.isInteger(cols[key]) ? C.col_letter(cols[key]) : "Not used"]),
  ]) {
    const a = setv(ws, r, 1, k), b = setv(ws, r, 2, v);
    a.font = _f(true); b.font = _f();
    for (const x of [a, b]) { x.border = BORDER; x.alignment = { vertical: "top", wrapText: true }; }
    r += 1;
  }
  r += 1;
  header(ws, r, ["Result", "Checks", "Share"]);
  r += 1;
  const t0 = r, total = r + C.M_ORDER.length;
  for (const st of C.M_ORDER) {
    const a = setv(ws, r, 1, st), b = setv(ws, r, 2, `=COUNTIF(Results!$I$2:$I$${last},A${r})`), c = setv(ws, r, 3, `=IF($B$${total}=0,0,B${r}/$B$${total})`);
    for (const x of [a, b, c]) body(x, false);
    mpaint(a, st);
    b.alignment = { horizontal: "center" };
    c.alignment = { horizontal: "center" };
    c.numFmt = "0.0%";
    r += 1;
  }
  {
    const a = setv(ws, r, 1, "Total checks"), b = setv(ws, r, 2, `=SUM(B${t0}:B${r - 1})`);
    for (const x of [a, b]) { body(x, false); x.font = _f(true); }
    b.alignment = { horizontal: "center" };
  }
  r += 2;
  for (const [k, v] of [
    [C.M_CORRECT, "The PDF shows the expected value."],
    [C.M_MINOR, "Normal mode only: very close to the expected value (spacing, initials, word order)."],
    [C.M_REVIEW, "Normal mode only: partly similar, or day/month swapped. Check by hand."],
    [C.M_OLD, "The PDF still shows the old value that was reported as wrong."],
    [C.M_WRONG, "The PDF shows something else."],
    [C.M_NOFIELD, "The PDF was found but this field could not be read from it."],
    [C.M_NOPDF, "No uploaded PDF has this enrollment number (or ABC ID)."],
    [C.M_MARKS, "The row is about marks, which are not compared automatically."],
    [C.M_SKIP, "Empty expected value, or a field that is not on marksheets (e.g. Aadhaar)."],
  ]) {
    const a = setv(ws, r, 1, k), b = setv(ws, r, 2, v);
    body(a, false); body(b);
    mpaint(a, k);
    r += 1;
  }
  ws.ws.getColumn(1).width = 34;
  ws.ws.getColumn(2).width = 80;
  ws.ws.getColumn(3).width = 12;
  page(ws);
  return wb.xlsx.writeBuffer();
}
