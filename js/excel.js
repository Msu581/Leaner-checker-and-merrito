/*
 * excel.js - the file-reading half of core.load_sheet()/list_sheets(), with SheetJS in place of
 * openpyxl (.xlsx/.xlsm) and pandas (.xls/.ods/.csv/.tsv/.txt).
 *
 * Each cell becomes the same string core._cell_str() produced in Python:
 *   numbers -> "251612400699" for whole numbers, format(v, ".15g") otherwise
 *   dates   -> "2006-07-05" (or "2006-07-05 10:30:00" when there is a time part), via openpyxl's
 *              from_excel() rules, only for cells whose number format openpyxl calls a date
 *   bools   -> "True"/"False";  text -> stripped;  empty -> ""
 * CSV follows pandas read_csv(sep=None, engine="python"): the delimiter is sniffed from the FIRST
 * line only with csv.Sniffer, blank lines are dropped, and a row with more fields than the first
 * row is an error.
 */
import * as py from "./py.js";
import { toOrdinal, fromOrdinal } from "./core.js";

// ------------------------------------------------------------------ cell values -> strings
function pad(n, w = 2) { return String(n).padStart(w, "0"); }

/** _cell_str() for a plain number */
export function numStr(v) {
  if (Number.isNaN(v)) return "";
  if (Number.isInteger(v)) return py.intStr(v);
  return py.fmtG15(v);
}

// openpyxl.styles.numbers
const LITERAL_GROUP = '".*?"';
const LOCALE_GROUP = "\\[(?!hh?\\]|mm?\\]|ss?\\])[^\\]]*\\]";
const STRIP_RE = new RegExp(`${LITERAL_GROUP}|${LOCALE_GROUP}`, "gsu");
const TIMEDELTA_RE = /\[hh?\](:mm(:ss(\.0*)?)?)?|\[mm?\](:ss(\.0*)?)?|\[ss?\](\.0*)?/iu;
export function is_date_format(fmt) {
  if (fmt === null || fmt === undefined) return false;
  fmt = fmt.split(";")[0].replace(STRIP_RE, "");
  return /(?<![_\\])[dmhysDMHYS]/u.test(fmt);
}
export function is_timedelta_format(fmt) {
  if (fmt === null || fmt === undefined) return false;
  return TIMEDELTA_RE.test(fmt.split(";")[0]);
}

/** openpyxl.utils.datetime.from_excel() + _cell_str(), returning the cell string. */
export function excelDateStr(value, date1904, timedelta) {
  const DAY_US = 86400e6;
  if (timedelta) {
    let us = py.pyRoundInt(value * DAY_US); // timedelta(days=value) keeps microseconds
    if (us % 1e6) us = Math.floor(us / 1e6) * 1e6 + py.pyRoundInt((us % 1e6) / 1000) * 1000;
    const days = Math.floor(us / DAY_US);
    let rest = us - days * DAY_US;
    const h = Math.floor(rest / 3600e6); rest -= h * 3600e6;
    const m = Math.floor(rest / 60e6); rest -= m * 60e6;
    const s = Math.floor(rest / 1e6); const micro = rest - s * 1e6;
    const t = `${h}:${pad(m)}:${pad(s)}${micro ? "." + pad(micro, 6) : ""}`;
    return days ? `${days} day${Math.abs(days) === 1 ? "" : "s"}, ${t}` : t;
  }
  let day = Math.floor(value);
  const fraction = value - day;
  let ms = py.pyRoundInt(fraction * 86400 * 1000);
  if (value >= 0 && value < 1 && ms < 86400000) { // a time of day
    return timeStr(ms);
  }
  if (value > 0 && value < 60 && !date1904) day += 1;
  const epochOrd = date1904 ? toOrdinal(1904, 1, 1) : toOrdinal(1899, 12, 30);
  const extraDays = Math.floor(ms / 86400000);
  ms -= extraDays * 86400000;
  const ord = epochOrd + day + extraDays;
  if (ord < 1 || ord > toOrdinal(9999, 12, 31)) return "#VALUE!"; // openpyxl: outside the limits for dates
  const d = fromOrdinal(ord);
  const date = `${pad(d.y, 4)}-${pad(d.m)}-${pad(d.d)}`;
  const H = Math.floor(ms / 3600000), M = Math.floor((ms % 3600000) / 60000), S = Math.floor((ms % 60000) / 1000);
  if (H === 0 && M === 0 && S === 0) return date;
  return `${date} ${timeStr(ms)}`;
}
function timeStr(ms) {
  const H = Math.floor(ms / 3600000), M = Math.floor((ms % 3600000) / 60000), S = Math.floor((ms % 60000) / 1000);
  const us = (ms % 1000) * 1000;
  return `${pad(H)}:${pad(M)}:${pad(S)}${us ? "." + pad(us, 6) : ""}`;
}

function xlsxCellStr(cell, date1904) {
  if (!cell) return "";
  switch (cell.t) {
    case "s": case "str": return py.strip(cell.v ?? "");
    case "b": return cell.v ? "True" : "False";
    case "e": return py.strip(cell.w ?? ""); // openpyxl keeps the error text, e.g. "#N/A"
    case "n": {
      if (cell.v === undefined || cell.v === null) return "";
      const z = cell.z && cell.z !== "General" ? String(cell.z) : null;
      if (z && is_date_format(z)) return excelDateStr(cell.v, date1904, is_timedelta_format(z));
      return numStr(cell.v);
    }
    case "d": return py.strip(String(cell.w ?? cell.v ?? ""));
    default: return cell.v === undefined || cell.v === null ? "" : py.strip(String(cell.v));
  }
}

/** pandas/xlrd: whole numbers become int, dates become datetime, errors become NaN (empty). */
function xlsCellStr(cell, date1904, XLSX) {
  if (!cell) return "";
  switch (cell.t) {
    case "s": case "str": return py.strip(cell.v ?? "");
    case "b": return cell.v ? "True" : "False";
    case "e": return "";
    case "n": {
      if (cell.v === undefined || cell.v === null) return "";
      const z = cell.z && cell.z !== "General" ? String(cell.z) : null;
      if (z && XLSX.SSF.is_date(z)) {
        const s = excelDateStr(cell.v, date1904, false);
        return s;
      }
      return numStr(cell.v);
    }
    default: return cell.v === undefined || cell.v === null ? "" : py.strip(String(cell.v));
  }
}

// ------------------------------------------------------------------ CSV (pandas python engine)
const SNIFF_RES = (() => {
  const w = `[^${py.W}\\n"']`;
  const BOL = "(?:^|(?<=\\n))", EOL = "(?:$|(?=\\n))"; // Python MULTILINE ^ and $
  return [
    new RegExp(`(?<delim>${w})(?<space> ?)(?<quote>["'])[\\s\\S]*?\\k<quote>\\k<delim>`, "gu"),
    new RegExp(`(?:${BOL}|\\n)(?<quote>["'])[\\s\\S]*?\\k<quote>(?<delim>${w})(?<space> ?)`, "gu"),
    new RegExp(`(?<delim>${w})(?<space> ?)(?<quote>["'])[\\s\\S]*?\\k<quote>(?:${EOL}|\\n)`, "gu"),
    new RegExp(`(?:${BOL}|\\n)(?<quote>["'])[\\s\\S]*?\\k<quote>(?:${EOL}|\\n)`, "gu"),
  ];
})();

function firstMax(map) { // max(d, key=d.get)
  let best = null, bv = -Infinity;
  for (const [k, v] of map) if (v > bv) { best = k; bv = v; }
  return best;
}

/** csv.Sniffer()._guess_quote_and_delimiter(data)[2] */
function guessQuoteAndDelimiter(data) {
  let matches = [];
  for (const re of SNIFF_RES) {
    matches = [...data.matchAll(re)];
    if (matches.length) break;
  }
  if (!matches.length) return null;
  const delims = new Map();
  for (const m of matches) {
    const g = m.groups;
    if (!("delim" in g)) continue;
    if (g.delim) delims.set(g.delim, (delims.get(g.delim) || 0) + 1);
  }
  if (!delims.size) return "";
  const d = firstMax(delims);
  return d === "\n" ? "" : d;
}

/** csv.Sniffer()._guess_delimiter(data)[0] */
function guessDelimiter(data) {
  const lines = data.split("\n").filter((x) => x);
  const ascii = Array.from({ length: 127 }, (_, i) => String.fromCharCode(i));
  const chunkLength = Math.min(10, lines.length);
  let iteration = 0;
  const charFrequency = new Map(), modes = new Map(), delims = new Map();
  let start = 0, end = chunkLength;
  while (start < lines.length) {
    iteration += 1;
    for (const line of lines.slice(start, end)) {
      for (const ch of ascii) {
        const meta = charFrequency.get(ch) || new Map();
        const freq = line.split(ch).length - 1;
        meta.set(freq, (meta.get(freq) || 0) + 1);
        charFrequency.set(ch, meta);
      }
    }
    for (const [ch, meta] of charFrequency) {
      let items = [...meta.entries()];
      if (items.length === 1 && items[0][0] === 0) continue;
      if (items.length > 1) {
        let mode = items[0];
        for (const it of items) if (it[1] > mode[1]) mode = it;
        items = items.filter((it) => it !== mode);
        modes.set(ch, [mode[0], mode[1] - items.reduce((a, it) => a + it[1], 0)]);
      } else modes.set(ch, items[0]);
    }
    const total = Math.min(chunkLength * iteration, lines.length);
    let consistency = 1.0;
    while (delims.size === 0 && consistency >= 0.9) {
      for (const [k, v] of modes) if (v[0] > 0 && v[1] > 0 && v[1] / total >= consistency) delims.set(k, v);
      consistency -= 0.01;
    }
    if (delims.size === 1) return [...delims.keys()][0];
    start = end;
    end += chunkLength;
  }
  if (!delims.size) return "";
  if (delims.size > 1) for (const d of [",", "\t", ";", " ", ":"]) if (delims.has(d)) return d;
  const items = [...delims.entries()].map(([k, v]) => [v, k]);
  items.sort((a, b) => py.cmpTuple(a[0], b[0]) || py.cmpStr(a[1], b[1]));
  return items[items.length - 1][1];
}

function sniff(line) {
  let d = guessQuoteAndDelimiter(line);
  if (!d) d = guessDelimiter(line);
  if (!d) throw new Error("Could not determine delimiter");
  return d;
}

/** Split text the way a file opened with newline="" is read line by line (\r, \n and \r\n end a line). */
function physicalLines(text) {
  const out = [];
  let i = 0;
  while (i < text.length) {
    let j = i;
    while (j < text.length && text[j] !== "\n" && text[j] !== "\r") j++;
    if (j < text.length) j += text[j] === "\r" && text[j + 1] === "\n" ? 2 : 1;
    out.push(text.slice(i, j));
    i = j;
  }
  return out;
}

/** Python csv.reader (QUOTE_MINIMAL, quotechar '"', doublequote, strict). Returns [{fields, line}]. */
function csvRecords(lines, delim, firstLineNo) {
  const recs = [];
  let state = "START_RECORD", fields = [], field = "";
  const save = () => { fields.push(field); field = ""; };
  let lineNo = firstLineNo - 1;
  for (const line of lines) {
    lineNo++;
    for (const c of line) {
      const nl = c === "\n" || c === "\r";
      switch (state) {
        case "START_RECORD":
          if (nl) { state = "EAT_CRNL"; break; }
          state = "START_FIELD";
        // falls through
        case "START_FIELD":
          if (nl) { save(); state = "EAT_CRNL"; }
          else if (c === '"') state = "IN_QUOTED_FIELD";
          else if (c === delim) save();
          else { field += c; state = "IN_FIELD"; }
          break;
        case "IN_FIELD":
          if (nl) { save(); state = "EAT_CRNL"; }
          else if (c === delim) { save(); state = "START_FIELD"; }
          else field += c;
          break;
        case "IN_QUOTED_FIELD":
          if (c === '"') state = "QUOTE_IN_QUOTED_FIELD";
          else field += c;
          break;
        case "QUOTE_IN_QUOTED_FIELD":
          if (c === '"') { field += c; state = "IN_QUOTED_FIELD"; }
          else if (c === delim) { save(); state = "START_FIELD"; }
          else if (nl) { save(); state = "EAT_CRNL"; }
          else throw new Error(`'${delim}' expected after '"'`);
          break;
        case "EAT_CRNL":
          if (!nl) throw new Error("new-line character seen in unquoted field - do you need to open the file with newline=''?");
          break;
      }
    }
    // end of the physical line (Python's EOL marker)
    if (state === "IN_QUOTED_FIELD") continue;
    if (state === "START_FIELD" || state === "IN_FIELD" || state === "QUOTE_IN_QUOTED_FIELD") save();
    state = "START_RECORD";
    recs.push({ fields, line: lineNo });
    fields = [];
  }
  if (state === "IN_QUOTED_FIELD" || field) throw new Error("unexpected end of data");
  return recs;
}

export function readCsvText(text) {
  const lines = physicalLines(text);
  if (!lines.length) throw new Error("No columns to parse from file");
  const delim = sniff(lines[0]);
  const recs = [...csvRecords([lines[0]], delim, 1), ...csvRecords(lines.slice(1), delim, 2)];
  // skip_blank_lines=True
  const kept = recs.filter((r) => r.fields.length > 1 || (r.fields.length === 1 && py.strip(r.fields[0])));
  if (!kept.length) throw new Error("No columns to parse from file");
  const ncols = kept[0].fields.length;
  for (const r of kept) {
    if (r.fields.length > ncols) throw new Error(`Expected ${ncols} fields in line ${r.line}, saw ${r.fields.length}`);
  }
  return kept.map((r) => r.fields.map((v) => py.strip(v)));
}

function decodeCsv(bytes) {
  try {
    let s = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(bytes);
    if (s.charCodeAt(0) === 0xfeff) s = s.slice(1); // utf-8-sig
    return s;
  } catch {
    let s = ""; // latin-1: every byte is one code point
    for (let i = 0; i < bytes.length; i += 8192) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192));
    return s;
  }
}

// ------------------------------------------------------------------ public: open a workbook
/**
 * Returns {name, ext, sheetNames, defaultSheet(), rows(sheet)} for core.load_sheet()/load_all_sheets().
 * XLSX: the SheetJS module.
 */
export function openBook(XLSX, bytes, fileName) {
  const ext = (fileName.match(/\.[^.]+$/) || [""])[0].toLowerCase();
  if ([".csv", ".tsv", ".txt"].includes(ext)) {
    const rows = readCsvText(decodeCsv(bytes));
    return { name: fileName, ext, sheetNames: ["(csv)"], defaultSheet: () => "(csv)", rows: () => rows.map((r) => [...r]) };
  }
  const wb = XLSX.read(bytes, { type: "array", cellNF: true, cellDates: false, cellText: true, cellFormula: false, dense: false });
  const date1904 = !!(wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
  const names = wb.SheetNames.slice();
  const isOpenXml = ext === ".xlsx" || ext === ".xlsm";
  const maxRow = (n) => {
    const ref = wb.Sheets[n] && wb.Sheets[n]["!ref"];
    return ref ? XLSX.utils.decode_range(ref).e.r + 1 : 0;
  };
  return {
    name: fileName,
    ext,
    sheetNames: names,
    // openpyxl: the sheet with the most rows (first one on a tie); pandas: the first sheet
    defaultSheet: () => (isOpenXml ? py.maxBy(names, maxRow) : names[0]),
    rows: (sheet) => {
      const ws = wb.Sheets[sheet];
      if (!ws || !ws["!ref"]) return [];
      const rg = XLSX.utils.decode_range(ws["!ref"]);
      const out = [];
      for (let r = 0; r <= rg.e.r; r++) { // openpyxl/pandas start at A1 whatever the used range says
        const row = [];
        for (let c = 0; c <= rg.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          row.push(isOpenXml ? xlsxCellStr(cell, date1904) : xlsCellStr(cell, date1904, XLSX));
        }
        out.push(row);
      }
      return out;
    },
  };
}
