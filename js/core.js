/*
 * core.js - port of core.py (extraction, column auto-detection, comparison) and meritto.py.
 * No browser or UI code here, so the same file runs in the page and in the Node test harness.
 * Function names follow the Python originals so the two can be read side by side.
 */
import * as py from "./py.js";

const { WS, W, B, END } = py;
const S = WS; // inside character classes
const RS = `[${WS}]`; // a Python "\s"

// --------------------------------------------------------------------------
// Constants
// --------------------------------------------------------------------------
export const FIELDS = [
  ["reg", "Registration No."],
  ["abc", "ABC ID"],
  ["name", "Student Name"],
  ["father", "Father's Name"],
  ["mother", "Mother's Name"],
  ["dob", "Date of Birth"],
];
export const FIELD_LABELS = Object.fromEntries(FIELDS);
export const COMPARE_FIELDS = ["name", "father", "mother", "dob", "abc"];
export const EXTRA_FIELDS = [["name_last", "Last name (optional)"]];
export const MAP_FIELDS = [...FIELDS, ...EXTRA_FIELDS];
export const MAP_LABELS = Object.fromEntries(MAP_FIELDS);
export const NAME_FIELDS = ["name", "father", "mother"];

export const MODE_NORMAL = "normal";
export const MODE_STRICT = "strict";
export const ALL_SHEETS = "(All sheets)";

export const ST_MATCH = "Match";
export const ST_CLOSE = "Close match";
export const ST_REVIEW = "Review";
export const ST_MISMATCH = "Mismatch";
export const ST_MISSING = "Missing";
export const ST_SKIPPED = "Not compared";

export const OV_VERIFIED = "Verified";
export const OV_MINOR = "Verified (minor differences)";
export const OV_REVIEW = "Needs review";
export const OV_MISMATCH = "Mismatch";
export const OV_NOTFOUND = "Not in Excel";
export const OV_NOREG = "No reg. no. in PDF";
export const OVERALL_ORDER = [OV_VERIFIED, OV_MINOR, OV_REVIEW, OV_MISMATCH, OV_NOTFOUND, OV_NOREG];

const TITLES = new Set(["MR", "MRS", "MS", "MISS", "DR", "SHRI", "SRI", "SMT", "KUMARI", "KUM", "LATE", "SHREE", "SHRIMATI"]);

export function Settings(o = {}) {
  return {
    day_first: true,
    match_threshold: 92.0,
    review_threshold: 75.0,
    ignore_titles: true,
    ocr: false,
    mode: MODE_NORMAL,
    ...o,
  };
}

export function col_letter(i) {
  let s = "";
  i += 1;
  while (i > 0) {
    const r = (i - 1) % 26;
    i = Math.floor((i - 1) / 26);
    s = String.fromCharCode(65 + r) + s;
  }
  return s;
}

export function col_index(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

// --------------------------------------------------------------------------
// Normalising and fuzzy scoring
// --------------------------------------------------------------------------
const pyStr = (s) => (s === null || s === undefined || s === "" ? "" : String(s)); // str(s or "")

function strip_accents(s) {
  // unicodedata.combining(c) != 0 for the marks that NFKD splits off Latin letters
  return s.normalize("NFKD").replace(/\p{M}/gu, "");
}

const RE_NOT_AZ_WS = new RegExp(`[^A-Z${S}]`, "gu");
export function clean_name(s, ignore_titles = true) {
  let t = strip_accents(pyStr(s)).toUpperCase();
  t = t.replace(RE_NOT_AZ_WS, " ");
  let words = py.split(t);
  if (ignore_titles) words = words.filter((w) => !TITLES.has(w));
  return words.join(" ");
}

const RE_SCI = /^\p{Nd}+(?:\.\p{Nd}+)?E\+\p{Nd}+$/u;
const RE_DOT0 = /^\p{Nd}+\.0+$/u;
export function norm_reg(s) {
  let t = py.strip(pyStr(s)).toUpperCase();
  if (RE_SCI.test(t)) t = decimalFixedIntPart(py.asciiDigits(t)); // Excel turned a long ID into 2.51E+11
  if (RE_DOT0.test(t)) t = t.split(".")[0];
  return t.replace(/[^A-Z0-9]/g, "");
}

/** format(Decimal("2.51612400699E+11"), "f").split(".")[0] - exact decimal shift, no floats. */
function decimalFixedIntPart(t) {
  const [mant, e] = t.split("E+");
  const [ip, fp = ""] = mant.split(".");
  const exp = Number(e);
  const digits = ip + fp;
  const point = ip.length + exp;
  const intPart = point >= digits.length ? digits + "0".repeat(point - digits.length) : digits.slice(0, point);
  // Decimal keeps leading zeros of the coefficient only as "0"
  return intPart.replace(/^0+(?=\d)/, "");
}

export function norm_abc(s) {
  return norm_reg(s).replace(/\P{Nd}/gu, "");
}

// rapidfuzz fuzz.ratio = (1 - indel_distance / (len1 + len2)) * 100
function lcs_len(a, b) {
  const A = Array.from(a), Bb = Array.from(b);
  if (!A.length || !Bb.length) return 0;
  let prev = new Uint32Array(Bb.length + 1), cur = new Uint32Array(Bb.length + 1);
  for (let i = 1; i <= A.length; i++) {
    for (let j = 1; j <= Bb.length; j++) {
      cur[j] = A[i - 1] === Bb[j - 1] ? prev[j - 1] + 1 : Math.max(prev[j], cur[j - 1]);
    }
    [prev, cur] = [cur, prev];
  }
  return prev[Bb.length];
}

export function _ratio(a, b) {
  const L = py.len(a) + py.len(b);
  const dist = L - 2 * lcs_len(a, b);
  const norm_dist = L ? dist / L : 0;
  return (1.0 - norm_dist) * 100;
}

export function _token_sort(a, b) {
  const srt = (s) => py.split(s).sort(py.cmpStr).join(" ");
  return _ratio(srt(a), srt(b));
}

export function edit_distance(a, b) {
  const A = Array.from(a), Bb = Array.from(b);
  let prev = Array.from({ length: Bb.length + 1 }, (_, i) => i);
  for (let i = 1; i <= A.length; i++) {
    const cur = [i];
    for (let j = 1; j <= Bb.length; j++) {
      cur.push(Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (A[i - 1] !== Bb[j - 1] ? 1 : 0)));
    }
    prev = cur;
  }
  return prev[prev.length - 1];
}

function _initials_match(ta, tb) {
  if (ta.length !== tb.length) return false;
  const pool = [...tb];
  let used_initial = false;
  const order = [...ta].sort((x, y) => py.len(y) - py.len(x)); // stable, like sorted(..., reverse=True)
  for (const t of order) {
    const at = pool.indexOf(t);
    if (at >= 0) { pool.splice(at, 1); continue; }
    const hit = pool.findIndex((p) => (py.len(t) === 1 && p.startsWith(t)) || (py.len(p) === 1 && t.startsWith(p)));
    if (hit < 0) return false;
    pool.splice(hit, 1);
    used_initial = true;
  }
  return used_initial;
}

const RE_WS_RUN = new RegExp(`${RS}+`, "gu");
const squash = (raw) => py.strip(String(raw)).replace(RE_WS_RUN, " ");

export function name_similarity(a_raw, b_raw, ignore_titles = true) {
  const a = clean_name(a_raw, ignore_titles), b = clean_name(b_raw, ignore_titles);
  if (a === b) {
    const same = squash(a_raw) === squash(b_raw);
    return [100.0, same ? "Exact" : "Same apart from case, spacing or punctuation"];
  }
  if (!a || !b) return [0.0, "Empty after cleaning"];
  const ta = a.split(" "), tb = b.split(" ");
  if (a.replaceAll(" ", "") === b.replaceAll(" ", "")) return [98.0, "Spacing differs"];
  const sa = [...ta].sort(py.cmpStr).join("\u0000"), sb = [...tb].sort(py.cmpStr).join("\u0000");
  if (sa === sb) return [97.0, "Same words, different order"];
  if (_initials_match(ta, tb)) return [95.0, "Initials vs full name"];
  const score = Math.max(_ratio(a, b), _token_sort(a, b));
  return [py.pyRound(score, 1), `${py.fmtFixed(score, 0)}% similar`];
}

// --------------------------------------------------------------------------
// Dates  (a date is {y, m, d})
// --------------------------------------------------------------------------
const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
const isLeap = (y) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
const DIM = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
export function daysIn(y, m) {
  return m === 2 && isLeap(y) ? 29 : DIM[m - 1];
}

function _ymd(y, m, d) {
  if (!(y >= 1 && y <= 9999 && m >= 1 && m <= 12 && d >= 1 && d <= daysIn(y, m))) return null;
  return { y, m, d };
}

function _year4(y) {
  y = py.pyInt(y);
  return y >= 100 ? y : (y <= 30 ? 2000 + y : 1900 + y);
}

/** proleptic Gregorian ordinal, like date.toordinal() */
export function toOrdinal(y, m, d) {
  const yy = y - 1;
  let n = yy * 365 + Math.floor(yy / 4) - Math.floor(yy / 100) + Math.floor(yy / 400);
  for (let i = 1; i < m; i++) n += daysIn(y, i);
  return n + d;
}
export function fromOrdinal(n) {
  let y = Math.floor(n / 365.2425) + 1;
  while (toOrdinal(y, 1, 1) > n) y--;
  while (toOrdinal(y + 1, 1, 1) <= n) y++;
  let m = 1;
  let rest = n - toOrdinal(y, 1, 1) + 1;
  while (rest > daysIn(y, m)) { rest -= daysIn(y, m); m++; }
  return { y, m, d: rest };
}

const D = "\\p{Nd}";
const RE_D1 = new RegExp(`^(${D}{4})[\\-/.](${D}{1,2})[\\-/.](${D}{1,2})(?:[ T][^\\n]*)?${END}`, "u");
const RE_D2 = new RegExp(`^(${D}{1,2})[\\-/.${S}](${D}{1,2})[\\-/.${S}](${D}{2,4})(?:${RS}[^\\n]*)?${END}`, "u");
const RE_D3 = new RegExp(`^(${D}{1,2})(?:st|nd|rd|th)?[${S}\\-/.]*([A-Za-z]{3,})[${S}\\-/.,]*(${D}{2,4})${END}`, "u");
const RE_D4 = new RegExp(`^([A-Za-z]{3,})[${S}\\-.]*(${D}{1,2})(?:st|nd|rd|th)?,?[${S}\\-.]*(${D}{2,4})${END}`, "u");
const RE_D5 = new RegExp(`^(${D}{5})(?:\\.${D}+)?${END}`, "u");

export function parse_date(value, day_first = true) {
  if (value === null || value === undefined) return null;
  if (typeof value === "object" && "y" in value) return value;
  const s = py.strip(String(value));
  if (!s) return null;
  let m = RE_D1.exec(s);
  if (m) return _ymd(py.pyInt(m[1]), py.pyInt(m[2]), py.pyInt(m[3]));
  m = RE_D2.exec(s);
  if (m) {
    const a = py.pyInt(m[1]), b = py.pyInt(m[2]), y = _year4(m[3]);
    let d, mo;
    if (a > 12) [d, mo] = [a, b];
    else if (b > 12) [mo, d] = [a, b];
    else if (day_first) [d, mo] = [a, b];
    else [mo, d] = [a, b];
    return _ymd(y, mo, d);
  }
  m = RE_D3.exec(s);
  if (m && MONTHS[m[2].slice(0, 3).toLowerCase()]) return _ymd(_year4(m[3]), MONTHS[m[2].slice(0, 3).toLowerCase()], py.pyInt(m[1]));
  m = RE_D4.exec(s);
  if (m && MONTHS[m[1].slice(0, 3).toLowerCase()]) return _ymd(_year4(m[3]), MONTHS[m[1].slice(0, 3).toLowerCase()], py.pyInt(m[2]));
  m = RE_D5.exec(s);
  if (m) {
    const n = py.pyInt(m[1]);
    if (n > 10000 && n < 80000) return fromOrdinal(toOrdinal(1899, 12, 30) + n);
  }
  return null;
}

export function fmt_date(d) {
  if (!d) return "";
  const p = (n) => String(n).padStart(2, "0");
  return `${p(d.d)}-${p(d.m)}-${d.y}`; // strftime %Y does not pad years below 1000 on Linux
}
const sameDate = (a, b) => a.y === b.y && a.m === b.m && a.d === b.d;

// --------------------------------------------------------------------------
// PDF field extraction (geometry comes from pdf.js, see pdfread.js)
// --------------------------------------------------------------------------
export function PdfRecord(file, path, page, pages, fields = {}, source = "text", unreadable = false, note = "") {
  return { file, path, page, pages, fields, source, unreadable, note };
}

const RE_LBL_NOT_AZ = new RegExp(`[^a-z${S}]`, "gu");
const RE_REGLBL = new RegExp(`regist|enrol|^reg(?: no| number)?${END}`, "u");
const RE_ABCLBL = new RegExp(`${B}abc${B}|${B}apaar${B}`, "u");
const RE_NAME_EXCL = /guardian|husband|institut|college|school|course|programme|program|exam|centre|center|university|parent|subject|paper/u;
export function label_field(label) {
  const l = py.strip(String(label).toLowerCase().replace(RE_LBL_NOT_AZ, " ").replace(RE_WS_RUN, " "));
  if (!l) return null;
  if (l.includes("mother")) return "mother";
  if (l.includes("father")) return "father";
  if (l.includes("birth") || l === "dob" || l === "d o b") return "dob";
  if (RE_REGLBL.test(l)) return "reg";
  if (RE_ABCLBL.test(l)) return "abc";
  if (l.includes("name") && !RE_NAME_EXCL.test(l)) return "name";
  return null;
}

export const _LABEL_RE = new RegExp(`^([^\\n]{2,60}?)${RS}*[:：]${RS}*([^\\n]+)${END}`, "u");
const RE_COLON_START = /^[:：]/u;
const RE_COLON_END = new RegExp(`[:：]${RS}*${END}`, "u");

/** words: {x0, x1, top, bottom, text} -> lines, each a list of segment strings (see core._lines_from_words). */
export function _lines_from_words(words, scale = 1.0) {
  words = words.filter((w) => py.strip(String(w.text)));
  words = [...words].sort((a, b) => (a.top - b.top) || (a.x0 - b.x0));
  const lines = [];
  for (const w of words) {
    const h = Math.max(w.bottom - w.top, 1);
    const ln = lines.find((l) => Math.abs(l.top - w.top) < Math.max(2.5 * scale, h * 0.35));
    if (ln) ln.w.push(w);
    else lines.push({ top: w.top, w: [w] });
  }
  const split_lines = [];
  for (const ln of [...lines].sort((a, b) => a.top - b.top)) {
    const ws = [...ln.w].sort((a, b) => a.x0 - b.x0);
    const segs = [];
    let cur = [];
    for (const w of ws) {
      const h = Math.max(w.bottom - w.top, 1);
      if (cur.length && w.x0 - cur[cur.length - 1].x1 > Math.max(18 * scale, h * 2.2)) {
        segs.push(cur);
        cur = [];
      }
      cur.push(w);
    }
    segs.push(cur);
    split_lines.push(segs);
  }
  // Where a second column of labels starts (x seen on 2+ lines after a clear gap); split a long value
  // glued to that column's label, but only when the rest really looks like "Label : value".
  const keys = [];
  for (const segs of split_lines) for (const sg of segs.slice(1)) if (sg.length) keys.push(py.pyRoundInt(sg[0].x0 / (3 * scale)));
  const anchors = [];
  for (const [k, n] of py.counter(keys)) if (n >= 2) anchors.push(k * 3 * scale);
  const text = (ws) => ws.map((w) => w.text).join(" ");
  const out = [];
  for (const segs0 of split_lines) {
    const fixed = [];
    segs0.forEach((sg, i) => {
      const after = i + 1 < segs0.length && RE_COLON_START.test(text(segs0[i + 1])) ? text(segs0[i + 1]) : "";
      let cut = null;
      for (let j = 1; j < sg.length; j++) {
        if (anchors.some((a) => Math.abs(sg[j].x0 - a) <= 3 * scale) && _LABEL_RE.test(py.strip(`${text(sg.slice(j))} ${after}`))) {
          cut = j;
          break;
        }
      }
      if (cut) fixed.push(sg.slice(0, cut), sg.slice(cut));
      else fixed.push(sg);
    });
    const segs = fixed.map((sg) => py.strip(sg.map((w) => w.text).join(" ")));
    const merged = [];
    for (const s of segs) {
      if (!s) continue;
      if (merged.length && (RE_COLON_START.test(s) || RE_COLON_END.test(merged[merged.length - 1]))) merged[merged.length - 1] += " " + s;
      else merged.push(s);
    }
    out.push(merged);
  }
  return out;
}

const RE_CV_REG = /[A-Za-z0-9/\-]{4,}/u;
const RE_CV_DOB = new RegExp(`${D}{1,4}[\\-/.${S}][${W}]{1,9}[\\-/.${S},]*${D}{2,4}`, "u");
const RE_CV_ABC = new RegExp(`${D}[${D}${S}\\-]{6,}${D}`, "u");
const RE_NOT_D = /\P{Nd}/gu;
function _clean_value(key, v) {
  v = py.strip(v);
  if (key === "reg") { const m = RE_CV_REG.exec(v); return m ? m[0] : v; }
  if (key === "dob") { const m = RE_CV_DOB.exec(v); return m ? m[0] : v; }
  if (key === "abc") { const m = RE_CV_ABC.exec(v); return m ? m[0].replace(RE_NOT_D, "") : v; }
  return py.stripChars(v, " .,:;");
}

const RE_4D = new RegExp(`${D}{4,}`, "u");
const RE_8D_FULL = new RegExp(`^${D}{8,}$`, "u");
const RE_FT_REG = new RegExp(`${B}(?:registration|regn?\\.?|enrol+ment)${RS}*(?:no\\.?|number)?${RS}*[:\\-]?${RS}*([A-Z0-9/\\-]{0,6}${D}{4,}[A-Z0-9/\\-]*)`, "iu");
const RE_FT_ABC = new RegExp(`${B}(?:abc|apaar)${RS}*(?:id|no\\.?)?${RS}*[:\\-]?${RS}*(${D}[${D} ]{6,}${D})`, "iu");
const RE_FT_DOB = new RegExp(`(?:date of birth|d\\.?${RS}?o\\.?${RS}?b\\.?)${RS}*[:\\-]?${RS}*(${D}{1,2}[/\\-.]${D}{1,2}[/\\-.]${D}{2,4})`, "iu");

export function extract_fields(lines) {
  const found = {};
  for (const segs of lines) {
    segs.forEach((seg, i) => {
      let key = null, val = null;
      const m = _LABEL_RE.exec(seg);
      if (m) { key = label_field(m[1]); val = m[2]; }
      else if (i + 1 < segs.length && py.split(seg).length <= 5) { // "Label | value" without a colon
        const k = label_field(seg);
        if (k && label_field(segs[i + 1]) === null) { key = k; val = segs[i + 1]; }
      }
      if (key && !(key in found) && py.strip(val)) {
        const cleaned = _clean_value(key, val);
        if (key === "reg" && !RE_4D.test(cleaned)) return; // OCR noise; full-text fallback may still find it
        if (key === "abc" && !RE_8D_FULL.test(cleaned)) return; // "ABC ID : -" / "N/A": no ABC ID
        found[key] = cleaned;
      }
    });
  }
  const full = lines.map((s) => s.join("  ")).join("\n");
  if (!("reg" in found)) { const m = RE_FT_REG.exec(full); if (m) found.reg = m[1]; }
  if (!("abc" in found)) { const m = RE_FT_ABC.exec(full); if (m) found.abc = m[1].replace(RE_NOT_D, ""); }
  if (!("dob" in found)) { const m = RE_FT_DOB.exec(full); if (m) found.dob = m[1]; }
  return found;
}

// --------------------------------------------------------------------------
// Excel / CSV grids
// --------------------------------------------------------------------------
export function SheetData(path, sheet, sheet_names, grid, ncols, origin = null, note = "") {
  return { path, sheet, sheet_names, grid, ncols, origin, note };
}

export function row_label(sd, r) {
  if (sd.origin && r < sd.origin.length && sd.origin[r]) return `${sd.origin[r][0]}!${sd.origin[r][1]}`;
  return String(r + 1);
}

/** rows already hold Python-equivalent cell strings (see excel.js) */
export function _finish_grid(rows) {
  const grid = rows.map((r) => [...r]);
  while (grid.length && !grid[grid.length - 1].some((c) => c)) grid.pop();
  let ncols = 0;
  for (const r of grid) r.forEach((c, i) => { if (c) ncols = Math.max(ncols, i + 1); });
  for (const r of grid) {
    while (r.length < ncols) r.push("");
    r.length = ncols;
  }
  return [grid, ncols];
}

/** book: {name, sheetNames, defaultSheet(names), rows(sheet)} from excel.js */
export function load_sheet(book, sheet = null) {
  const names = book.sheetNames;
  if (!names.includes(sheet)) sheet = book.defaultSheet();
  const [grid, ncols] = _finish_grid(book.rows(sheet));
  return SheetData(book.name, sheet, names, grid, ncols);
}

export function load_all_sheets(book) {
  const names = book.sheetNames;
  const header = [], where = new Map(), stacked = [], skipped = [];
  for (const n of names) {
    const sd = load_sheet(book, n);
    const hr = detect_header_row(sd.grid);
    const kinds = new Set();
    if (hr >= 0) {
      for (const c of sd.grid[hr]) {
        if (Math.max(...FIELDS.map(([f]) => header_score(f, c))) >= 0.6) kinds.add(py.maxBy(FIELDS, (f) => header_score(f[0], c))[0]);
      }
    }
    if (kinds.size < 2) { skipped.push(n); continue; }
    const cols = [], seen = new Set();
    for (const h of sd.grid[hr]) {
      let key = py.strip(h.replace(RE_WS_RUN, " ")).toLowerCase();
      if (!key) { cols.push(null); continue; }
      const base = key;
      let i = 2;
      while (seen.has(key)) { key = `${base} (${i})`; i += 1; }
      seen.add(key);
      if (!where.has(key)) {
        where.set(key, header.length);
        header.push(key === base ? py.strip(h) : `${py.strip(h)} (${i - 1})`);
      }
      cols.push(where.get(key));
    }
    for (const r of data_rows(sd.grid, hr)) {
      const vals = new Map();
      sd.grid[r].forEach((v, c) => { if (cols[c] !== null && v) vals.set(cols[c], v); });
      stacked.push([vals, n, r + 1]);
    }
  }
  if (!stacked.length) throw new Error("No sheet has a recognisable header row (registration no., name, ...).");
  const width = header.length + 1;
  const grid = [[...header, "Sheet"]];
  const origin = [null];
  for (const [vals, n, r] of stacked) {
    const row = new Array(width).fill("");
    for (const [c, v] of vals) row[c] = v;
    row[width - 1] = n;
    grid.push(row);
    origin.push([n, r]);
  }
  let note = `Combined ${names.length - skipped.length} sheet(s), ${stacked.length} rows.`;
  if (skipped.length) note += " Skipped (no header row found): " + skipped.join(", ");
  return SheetData(book.name, ALL_SHEETS, names, grid, width, origin, note);
}

// --------------------------------------------------------------------------
// Column auto-detection
// --------------------------------------------------------------------------
const HDR = {
  reg: [[/registration|regn|reg\.? ?no|reg ?number|^reg$|regd/u, 1], [/enrol+ment|enrol+ no/u, 0.9], [/roll/u, 0.55], [/univ.*(no|id)|student id|^id$/u, 0.45]],
  name: [[/name of (the )?(student|candidate)|(student|candidate|learner)'?s? ?name/u, 1], [/^(full ?)?name$/u, 0.9], [/^(student|candidate)$/u, 0.85],
    [/^first ?name$|given ?name/u, 0.8], [/name/u, 0.45]],
  name_last: [[/^(last|sur) ?name$|^surname$/u, 1], [/last ?name|surname|family name/u, 0.8]],
  abc: [[new RegExp(`${B}abc${B}|${B}apaar${B}`, "u"), 1]],
  father: [[/father/u, 1], [/f\.? ?name|fname/u, 0.6], [/guardian/u, 0.5]],
  mother: [[/mother/u, 1], [/m\.? ?name|mname/u, 0.6]],
  dob: [[new RegExp(`birth|${B}dob${B}|d\\.o\\.b`, "u"), 1], [/^b\.? ?date$|bdate/u, 0.8], [/date/u, 0.3]],
};
// Python "$" also matches before a final newline; header text is stripped, so plain "$" is equivalent here.
const RE_UNDERSCORE_WS = new RegExp(`[_${S}]+`, "gu");
const RE_NAME_HDR_EXCL = /father|mother|guardian|husband|course|programme|school|college|institut|subject|last ?name|surname|middle|family name|regional/u;
const RE_DOB_HDR_EXCL = /admission|issue|exam|publi|join/u;

export function header_score(key, header) {
  const t = py.strip(pyStr(header).toLowerCase().replace(RE_UNDERSCORE_WS, " "));
  if (!t) return 0.0;
  if (key === "name" && RE_NAME_HDR_EXCL.test(t)) return 0.0;
  if (key === "dob" && RE_DOB_HDR_EXCL.test(t)) return 0.0;
  let best = 0.0;
  for (const [pat, w] of HDR[key]) if (pat.test(t) && w > best) best = w;
  return best;
}

const anyField = (c) => Math.max(...FIELDS.map(([f]) => header_score(f, c)));
export function detect_header_row(grid) {
  let best = -1, best_n = 0;
  for (let r = 0; r < Math.min(20, grid.length); r++) {
    const n = grid[r].filter((c) => anyField(c) >= 0.6).length;
    if (n > best_n) { best = r; best_n = n; }
  }
  return best;
}

export function data_rows(grid, header_row) {
  const out = [];
  for (let r = header_row + 1; r < grid.length; r++) if (grid[r].some((c) => c)) out.push(r);
  return out;
}

export function ColumnGuess(col = null, confidence = "none", score = 0.0, reason = "") {
  return { col, confidence, score, reason };
}

const RE_5D = new RegExp(`${D}{5,}`, "u");
const RE_NAMELIKE = new RegExp(`^[A-Za-z][A-Za-z.${S}'\\-]{2,}${END}`, "u");

export function auto_detect(grid, header_row, pdf_records = null, settings = null, keep = null) {
  settings = settings || Settings();
  pdf_records = pdf_records || [];
  keep = Object.fromEntries(Object.entries(keep || {}).filter(([, v]) => v !== null && v !== undefined));
  const ncols = grid.length ? grid[0].length : 0;
  const rows = data_rows(grid, header_row);
  const sample = rows.slice(0, 400);
  const hdr = (c) => (header_row >= 0 && header_row < grid.length ? grid[header_row][c] : "");
  const vals = (c, rs) => rs.map((r) => grid[r][c]).filter((x) => x);
  const pdf_regs = pdf_records.filter((p) => p.fields.reg).map((p) => norm_reg(p.fields.reg));
  const pdf_abcs = pdf_records.filter((p) => p.fields.abc).map((p) => norm_abc(p.fields.abc));
  const cands = [];

  for (let c = 0; c < ncols; c++) { // ABC ID: header or PDF values only
    let h = header_score("abc", hdr(c));
    let found = 0;
    if (pdf_abcs.length) {
      const have = new Set(vals(c, rows).map(norm_abc));
      found = pdf_abcs.filter((x) => have.has(x)).length;
    }
    const why = [...(h >= 0.5 ? [`header "${hdr(c)}"`] : []), ...(found ? [`${found} of ${pdf_abcs.length} PDF ABC IDs found`] : [])];
    cands.push(["abc", c, Math.min(1.0, 0.8 * h + 0.7 * (pdf_abcs.length ? found / pdf_abcs.length : 0)), why]);
    h = header_score("name_last", hdr(c));
    if (h) cands.push(["name_last", c, 0.6 * h, [`header "${hdr(c)}"`]]);
  }

  for (let c = 0; c < ncols; c++) {
    const h = header_score("reg", hdr(c));
    const v = vals(c, sample);
    const like = v.length ? v.filter((x) => RE_5D.test(x.replaceAll(" ", "")) && parse_date(x) === null).length / v.length : 0;
    const uniq = v.length ? new Set(v.map(norm_reg)).size / v.length : 0;
    let p = 0.0, found = 0;
    if (pdf_regs.length) {
      const have = new Set(vals(c, rows).map(norm_reg));
      found = pdf_regs.filter((x) => have.has(x)).length;
      p = found / pdf_regs.length;
    }
    const why = [];
    if (h >= 0.5) why.push(`header "${hdr(c)}"`);
    if (found) why.push(`${found} of ${pdf_regs.length} PDF reg. numbers found`);
    else if (like > 0.7) why.push("values look like ID numbers");
    cands.push(["reg", c, Math.min(1.0, 0.55 * h + 0.2 * like * uniq + 0.7 * p + (h >= 0.5 && like > 0.7 ? 0.15 : 0)), why]);
  }

  let reg_col = keep.reg ?? null;
  if (reg_col === null) {
    const best = py.maxBy(cands.filter((x) => x[0] === "reg"), (x) => x[2]);
    reg_col = best && best[2] >= 0.3 ? best[1] : null;
  }

  let pairs = [];
  if (reg_col !== null && pdf_records.length) {
    const idx = new Map();
    for (const r of rows) { const k = norm_reg(grid[r][reg_col]); if (!idx.has(k)) idx.set(k, r); }
    pairs = pdf_records.filter((p) => p.fields.reg && idx.has(norm_reg(p.fields.reg))).map((p) => [p, idx.get(norm_reg(p.fields.reg))]);
  }

  for (const key of ["name", "father", "mother", "dob"]) {
    for (let c = 0; c < ncols; c++) {
      const h = header_score(key, hdr(c));
      const v = vals(c, sample);
      let like;
      if (key === "dob") like = v.length ? v.filter((x) => parse_date(x, settings.day_first)).length / v.length : 0;
      else like = v.length ? v.filter((x) => RE_NAMELIKE.test(x)).length / v.length : 0;
      let hits = 0, tot = 0;
      for (const [rec, r] of pairs) {
        const pv = rec.fields[key];
        if (!pv) continue;
        tot += 1;
        if (key === "dob") {
          const a = parse_date(pv, settings.day_first), b = parse_date(grid[r][c], settings.day_first);
          hits += a && b && sameDate(a, b) ? 1 : 0;
        } else {
          hits += name_similarity(pv, grid[r][c], settings.ignore_titles)[0] >= 85 ? 1 : 0;
        }
      }
      const pf = tot ? hits / tot : 0;
      const why = [];
      if (h >= 0.5) why.push(`header "${hdr(c)}"`);
      if (hits) why.push(`${hits} of ${tot} PDF values match`);
      else if (key === "dob" && like > 0.7) why.push(`${py.fmtPercent0(like)} of values are dates`);
      const s = key === "dob" ? 0.55 * h + 0.3 * like + 0.6 * pf : 0.6 * h + 0.1 * like + 0.6 * pf;
      cands.push([key, c, Math.min(1.0, s), why]);
    }
  }

  const result = Object.fromEntries(MAP_FIELDS.map(([k]) => [k, ColumnGuess()]));
  const used = new Set(Object.values(keep));
  for (const [k, c] of Object.entries(keep)) result[k] = ColumnGuess(c, "manual", 1.0, "Set manually");
  for (const [key, c, s, why] of [...cands].sort((a, b) => b[2] - a[2])) {
    if (result[key].col !== null || used.has(c) || s < 0.3) continue;
    result[key] = ColumnGuess(c, s >= 0.75 ? "high" : s >= 0.5 ? "medium" : "low", s, why.join(", "));
    used.add(c);
  }
  return result;
}

// --------------------------------------------------------------------------
// Comparison
// --------------------------------------------------------------------------
export function FieldResult(field, pdf_value = "", excel_value = "", status = ST_SKIPPED, score = null, note = "") {
  return { field, pdf_value, excel_value, status, score, note };
}
export function RecordResult(pdf, overall, excel_row = null, fields = {}, note = "") {
  return { pdf, overall, excel_row, fields, note };
}
export function RunResult(records, unmatched_excel) {
  return { records, unmatched_excel };
}
export function counts(result) {
  const c = Object.fromEntries(OVERALL_ORDER.map((k) => [k, 0]));
  for (const r of result.records) c[r.overall] = (c[r.overall] || 0) + 1;
  return c;
}

export function strict_similarity(a_raw, b_raw, ignore_titles = true) {
  const a = clean_name(a_raw, ignore_titles), b = clean_name(b_raw, ignore_titles);
  if (a === b) {
    const same = squash(a_raw) === squash(b_raw);
    return [100.0, same ? "Exact" : "Same apart from case, spacing or punctuation"];
  }
  if (!a || !b) return [0.0, "Empty after cleaning"];
  if (a.replaceAll(" ", "") === b.replaceAll(" ", "")) return [0.0, "Strict: words are split differently"];
  if ([...a.split(" ")].sort(py.cmpStr).join("\u0000") === [...b.split(" ")].sort(py.cmpStr).join("\u0000")) return [0.0, "Strict: same words, different order"];
  return [0.0, "Strict: text differs"];
}

export function compare_field(key, pdf_val, excel_val, s) {
  const p = py.strip(pyStr(pdf_val)), e = py.strip(pyStr(excel_val));
  const fr = FieldResult(key, p, e);
  if (key === "abc") {
    const pa = norm_abc(p), ea = norm_abc(e);
    fr.pdf_value = pa || p;
    fr.excel_value = ea || e;
    if (!pa) [fr.status, fr.note] = [ST_SKIPPED, "No ABC ID on the marksheet"];
    else if (!ea) [fr.status, fr.note] = [ST_MISSING, "Empty in Excel"];
    else if (pa === ea) [fr.status, fr.score, fr.note] = [ST_MATCH, 100.0, "Same ABC ID"];
    else [fr.status, fr.score, fr.note] = [ST_MISMATCH, 0.0, "Different ABC ID"];
    return fr;
  }
  let pd_ = null, ed = null;
  if (key === "dob") {
    pd_ = parse_date(p, s.day_first);
    ed = parse_date(e, s.day_first);
    fr.pdf_value = fmt_date(pd_) || p;
    fr.excel_value = fmt_date(ed) || e;
  }
  if (!p && !e) [fr.status, fr.note] = [ST_MISSING, "Empty in both"];
  else if (!p) [fr.status, fr.note] = [ST_MISSING, "Not found in PDF"];
  else if (!e) [fr.status, fr.note] = [ST_MISSING, "Empty in Excel"];
  else if (key === "dob") {
    if (pd_ && ed) {
      if (sameDate(pd_, ed)) [fr.status, fr.score, fr.note] = [ST_MATCH, 100.0, "Same date"];
      else if (pd_.y === ed.y && pd_.m === ed.d && pd_.d === ed.m) [fr.status, fr.score, fr.note] = [ST_REVIEW, 0.0, "Day and month look swapped"];
      else [fr.status, fr.score, fr.note] = [ST_MISMATCH, 0.0, "Different dates"];
    } else {
      const same = norm_reg(p) === norm_reg(e);
      [fr.status, fr.score] = same ? [ST_MATCH, 100.0] : [ST_MISMATCH, 0.0];
      fr.note = same ? "Same text" : "Could not read one of the dates";
    }
  } else {
    const sim = s.mode === MODE_STRICT ? strict_similarity : name_similarity;
    const [score, note] = sim(p, e, s.ignore_titles);
    fr.score = score;
    fr.note = note;
    if (score >= 100) fr.status = ST_MATCH;
    else if (score >= s.match_threshold) fr.status = ST_CLOSE;
    else if (score >= s.review_threshold) fr.status = ST_REVIEW;
    else fr.status = ST_MISMATCH;
  }
  return fr;
}

function _overall(statuses) {
  statuses = statuses.filter((x) => x !== ST_SKIPPED);
  if (!statuses.length) return OV_REVIEW;
  if (statuses.includes(ST_MISMATCH)) return OV_MISMATCH;
  if (statuses.includes(ST_REVIEW) || statuses.includes(ST_MISSING)) return OV_REVIEW;
  if (statuses.includes(ST_CLOSE)) return OV_MINOR;
  return OV_VERIFIED;
}

export function compare(grid, header_row, mapping, pdf_records, settings = null, origin = null) {
  settings = settings || Settings();
  const rc = mapping.reg ?? null;
  if (rc === null) throw new Error("Map the registration number column first; it is used to pair rows.");
  const cell = (r, c) => (c !== null && c !== undefined && c < grid[r].length ? grid[r][c] : "");
  const hasO = (r) => origin && r < origin.length && origin[r];
  const xrow = (r) => (hasO(r) ? `${origin[r][0]}!${origin[r][1]}` : r + 1);
  const where = (r) => (hasO(r) ? `sheet ${origin[r][0]} row ${origin[r][1]}` : `row ${r + 1}`);
  const last = mapping.name_last ?? null;
  const excel_value = (r, key) => {
    let v = cell(r, mapping[key] ?? null);
    if (key === "name" && last !== null && cell(r, last)) v = py.strip(`${v} ${cell(r, last)}`);
    return v;
  };

  const rows = data_rows(grid, header_row);
  const index = new Map(), abc_index = new Map();
  const ac = mapping.abc ?? null;
  for (const r of rows) {
    const k = norm_reg(cell(r, rc));
    if (k) { if (!index.has(k)) index.set(k, []); index.get(k).push(r); }
    const a = ac !== null ? norm_abc(cell(r, ac)) : "";
    if (a) { if (!abc_index.has(a)) abc_index.set(a, []); abc_index.get(a).push(r); }
  }
  const used = new Set(), results = [];
  for (const p of pdf_records) {
    const k = norm_reg(p.fields.reg);
    let paired_by = "reg", hits;
    if (!k) {
      const a = norm_abc(p.fields.abc);
      if (!(a && abc_index.has(a))) {
        results.push(RecordResult(p, OV_NOREG, null, {}, p.note || "Registration number not found in the PDF"));
        continue;
      }
      hits = abc_index.get(a);
      paired_by = "abc";
    } else hits = index.get(k) || [];
    if (!hits.length) {
      const near = [];
      for (const x of index.keys()) if (Math.abs(py.len(x) - py.len(k)) <= 2) near.push([edit_distance(k, x), x]);
      const close = near.filter((n) => n[0] <= 1).sort(py.cmpTuple);
      let note = "Registration number not found in Excel";
      if (close.length) {
        const r0 = index.get(close[0][1])[0];
        note += `. Closest in Excel: ${cell(r0, rc)} (${where(r0)})`;
      }
      results.push(RecordResult(p, OV_NOTFOUND, null, {}, note));
      continue;
    }
    const r = hits[0];
    used.add(r);
    const frs = paired_by === "reg"
      ? { reg: FieldResult("reg", pyStr(p.fields.reg ?? ""), cell(r, rc), ST_MATCH, 100.0, "Paired on this") }
      : { reg: FieldResult("reg", "", cell(r, rc), ST_SKIPPED, null, "Not on the marksheet; paired by ABC ID") };
    for (const key of COMPARE_FIELDS) {
      const c = mapping[key] ?? null;
      if (c === null) frs[key] = FieldResult(key, pyStr(p.fields[key] ?? ""), "", ST_SKIPPED, null, "Column not mapped");
      else frs[key] = compare_field(key, p.fields[key], excel_value(r, key), settings);
    }
    const notes = [];
    if (paired_by === "abc") notes.push("No reg. no. in the PDF; paired by ABC ID");
    if (hits.length > 1) {
      const label = paired_by === "abc" ? "ABC ID" : "Reg. no.";
      notes.push(`${label} appears in Excel at ${hits.map(where).join(", ")}; first used`);
    }
    results.push(RecordResult(p, _overall(Object.entries(frs).filter(([k2]) => k2 !== "reg").map(([, f]) => f.status)), xrow(r), frs, notes.join("; ")));
  }
  const unmatched = rows.filter((r) => !used.has(r) && cell(r, rc)).map((r) => ({ row: xrow(r), reg: cell(r, rc), name: excel_value(r, "name") }));
  return RunResult(results, unmatched);
}

// ==========================================================================
// Meritto vs Marksheet (meritto.py)
// ==========================================================================
export const M_CORRECT = "Correct";
export const M_MINOR = "Correct (minor differences)";
export const M_REVIEW = "Needs review";
export const M_OLD = "Still wrong (same as old value)";
export const M_WRONG = "Mismatch";
export const M_NOFIELD = "Field not in PDF";
export const M_NOPDF = "PDF not found";
export const M_MARKS = "Marks row (not checked)";
export const M_SKIP = "Not checked";
export const M_ORDER = [M_CORRECT, M_MINOR, M_REVIEW, M_OLD, M_WRONG, M_NOFIELD, M_NOPDF, M_MARKS, M_SKIP];
export const M_OK = [M_CORRECT, M_MINOR];
export const M_PROBLEM = [M_REVIEW, M_OLD, M_WRONG, M_NOFIELD, M_NOPDF];

export const ROLES = [
  ["enroll", "Enrollment / Reg. no."],
  ["expected", "Expected data (as per Meritto)"],
  ["claimed", "Old value in marksheet (optional)"],
  ["student", "Student name (optional)"],
  ["remark", "Remarks (optional)"],
  ["abc", "ABC ID (optional)"],
];
export const ROLE_LABELS = Object.fromEntries(ROLES);
export const CHECKABLE = ["name", "father", "mother", "dob", "abc", "reg"];

const _EXPECTED_HDR = new RegExp(`meritto|merito|${B}tr${B}|expected|correct|should be|data as per|as per record`, "iu");
const _CLAIMED_HDR = new RegExp(`${B}in (?:the )?mark ?sheet|mark ?sheet (?:data|value)|printed|wrong|incorrect|old value|current`, "iu");
const _REMARK_HDR = /remark|comment|issue|discrepanc|observation/iu;
const _LINE = new RegExp(`^${RS}*([^:：\\n]{2,60}?)${RS}*[:：]${RS}*([^\\n]*?)${RS}*${END}`, "u");
const _MARKS = new RegExp(`${B}marks?${B}|${B}total${B}|${B}grade${B}|${B}sgpa${B}|${B}cgpa${B}`, "iu");

export function parse_cell(text) {
  const out = [];
  for (let line of pyStr(text).split(/[\r\n]+/u)) {
    line = py.strip(line);
    if (!line) continue;
    const m = _LINE.exec(line);
    if (m) out.push([py.strip(m[1]), label_field(m[1]), py.strip(m[2])]);
    else out.push(["", null, line]);
  }
  return out;
}

export function Check(o) {
  return {
    excel_row: 0, enrollment: "", student: "", remark: "", label: "", field: null, expected: "", claimed: "",
    kind: "field", abc: "", pdf_file: "", pdf_path: "", pdf_page: null, pdf_value: "", status: "", score: null,
    note: "", parse_note: "", force_pdf_path: null, force_pdf_value: null, manual: [], ...o,
  };
}
export function field_label(chk) {
  return FIELD_LABELS[chk.field] || chk.label || "-";
}

function _role_score(role, h) {
  const t = py.strip(pyStr(h).replace(RE_WS_RUN, " "));
  if (!t) return 0.0;
  if (role === "enroll") return header_score("reg", t);
  if (role === "claimed") return _CLAIMED_HDR.test(t) ? 1.0 : 0.0;
  if (role === "expected") return _CLAIMED_HDR.test(t) ? 0.0 : (_EXPECTED_HDR.test(t) ? 1.0 : 0.0);
  if (role === "student") return header_score("name", t);
  if (role === "remark") return _REMARK_HDR.test(t) ? 1.0 : 0.0;
  if (role === "abc") return header_score("abc", t);
  return 0.0;
}

export function m_detect_header_row(grid) {
  let best = -1, best_n = 1;
  for (let r = 0; r < Math.min(30, grid.length); r++) {
    const n = ROLES.filter(([role]) => grid[r].some((c) => _role_score(role, c) >= 0.5)).length;
    if (n > best_n) { best = r; best_n = n; }
  }
  return best;
}

const _looks_labelled = (v) => parse_cell(v).some(([, f]) => f);
const RE_6D_FULL = new RegExp(`^${D}{6,}$`, "u");

export function detect_columns(grid, header_row, pdf_records = null, keep = null) {
  keep = Object.fromEntries(Object.entries(keep || {}).filter(([, v]) => v !== null && v !== undefined));
  const ncols = grid.length ? grid[0].length : 0;
  const hdr = (c) => (header_row >= 0 && header_row < grid.length ? grid[header_row][c] : "");
  const rows = [];
  for (let r = header_row + 1; r < grid.length; r++) if (grid[r].some((x) => x)) rows.push(r);
  const vals = (c) => rows.map((r) => grid[r][c]).filter((x) => x);
  const pdf_regs = new Set((pdf_records || []).filter((p) => p.fields.reg).map((p) => norm_reg(p.fields.reg)));
  const cands = [];
  for (let c = 0; c < ncols; c++) {
    const v = vals(c);
    const labelled = v.length ? v.filter(_looks_labelled).length / v.length : 0;
    const idlike = v.length ? v.filter((x) => RE_6D_FULL.test(norm_reg(x))).length / v.length : 0;
    const hits = pdf_regs.size ? v.filter((x) => pdf_regs.has(norm_reg(x))).length : 0;
    for (const [role] of ROLES) {
      const h = _role_score(role, hdr(c));
      const why = h >= 0.5 ? [`header "${hdr(c)}"`] : [];
      let s;
      if (role === "enroll") {
        s = 0.6 * h + 0.3 * idlike + (pdf_regs.size ? 0.5 * Math.min(1, hits / pdf_regs.size) : 0);
        if (hits) why.push(`${hits} PDF reg. numbers found`);
        else if (idlike > 0.7) why.push("values look like ID numbers");
      } else if (role === "expected" || role === "claimed") {
        s = 0.7 * h + 0.3 * labelled + (role === "expected" ? -0.01 * c : 0.01 * c * (h > 0 ? 1 : 0));
        if (labelled > 0.3) why.push(`${py.fmtPercent0(labelled)} of cells look like 'Label : value'`);
      } else s = 0.8 * h;
      cands.push([role, c, s, why]);
    }
  }
  const result = Object.fromEntries(ROLES.map(([k]) => [k, ColumnGuess()]));
  const used = new Set(Object.values(keep));
  for (const [k, c] of Object.entries(keep)) result[k] = ColumnGuess(c, "manual", 1.0, "Set manually");
  for (const [role, c, s, why] of [...cands].sort((a, b) => b[2] - a[2])) {
    if (result[role].col !== null || used.has(c) || s < 0.3) continue;
    result[role] = ColumnGuess(c, s >= 0.7 ? "high" : s >= 0.5 ? "medium" : "low", s, why.join(", "));
    used.add(c);
  }
  return result;
}

export function build_checks(grid, header_row, cols, last_row = null, row_label_fn = null) {
  if ((cols.enroll ?? null) === null || (cols.expected ?? null) === null) {
    throw new Error("Choose the enrollment column and the expected-data column first.");
  }
  const cell = (r, role) => py.strip((cols[role] ?? null) !== null && cols[role] < grid[r].length ? grid[r][cols[role]] : "");
  const row_no = row_label_fn || ((r) => r + 1);
  const end = !last_row ? grid.length : Math.min(grid.length, last_row);
  const checks = [];
  for (let r = header_row + 1; r < end; r++) {
    const enr = cell(r, "enroll"), exp = cell(r, "expected");
    if (!enr && !exp) continue; // blank row or section title
    const base = { excel_row: row_no(r), enrollment: enr, student: cell(r, "student"), remark: cell(r, "remark"), abc: cell(r, "abc") };
    if (!exp) { checks.push(Check({ ...base, kind: "empty", parse_note: "Expected-data cell is empty" })); continue; }
    const claimed = parse_cell(cell(r, "claimed"));
    const claimed_by_field = {};
    for (const [, f, v] of claimed) if (f) claimed_by_field[f] = v;
    const entries = parse_cell(exp);
    for (const [label, fld0, value] of entries) {
      let fld = fld0;
      const chk = Check({ ...base, label, field: fld, expected: value });
      if (fld === null && !label) {
        if (_MARKS.test(value)) chk.kind = "marks";
        else {
          fld = chk.remark ? label_field(chk.remark) : null;
          if (CHECKABLE.includes(fld)) { chk.field = fld; chk.parse_note = "Field taken from the Remarks column (no label in the cell)"; }
          else { chk.kind = "unsupported"; chk.parse_note = "No 'Label : value' found and Remarks do not name a field"; }
        }
      } else if (!CHECKABLE.includes(fld)) {
        chk.kind = "unsupported";
        chk.parse_note = `"${label}" is not a field read from marksheets`;
      }
      if (chk.kind === "field") {
        if (!chk.expected) { chk.kind = "empty"; chk.parse_note = "No value after the label"; }
        chk.claimed = claimed_by_field[chk.field] ?? "";
        if (!chk.claimed && entries.length === 1 && claimed.length === 1) chk.claimed = claimed[0][2];
      }
      checks.push(chk);
    }
  }
  _mark_duplicates(checks);
  return checks;
}

function _mark_duplicates(checks) {
  const seen = new Map();
  for (const c of checks) {
    if (c.kind !== "field") continue;
    const k = `${norm_reg(c.enrollment)}\u0000${c.field}`;
    if (!seen.has(k)) seen.set(k, []);
    seen.get(k).push(c);
  }
  for (const group of seen.values()) {
    if (group.length > 1) {
      for (const c of group) {
        const others = group.filter((o) => o !== c).map((o) => String(o.excel_row));
        c.parse_note = [c.parse_note, `Same student and field also on row(s) ${others.join(", ")}`].filter((x) => x).join("; ");
      }
    }
  }
}

export function pdf_index(pdf_records) {
  const by_reg = new Map(), by_abc = new Map();
  for (const p of pdf_records) {
    if (p.fields.reg) { const k = norm_reg(p.fields.reg); if (!by_reg.has(k)) by_reg.set(k, []); by_reg.get(k).push(p); }
    if (p.fields.abc) { const k = norm_abc(p.fields.abc); if (!by_abc.has(k)) by_abc.set(k, []); by_abc.get(k).push(p); }
  }
  return [by_reg, by_abc];
}

function _same(field_key, a, b, s) {
  return compare_field(field_key, a, b, { ...s, mode: MODE_STRICT }).status === ST_MATCH;
}

export function evaluate(chk, pdf_records, s, index = null) {
  chk.pdf_file = chk.pdf_path = chk.pdf_value = "";
  chk.pdf_page = null; chk.score = null; chk.note = "";
  if (chk.kind === "marks") { chk.status = M_MARKS; chk.note = "Marks are not printed as 'Label : value'; check by hand"; return chk; }
  if (chk.kind !== "field" || !chk.field || !chk.expected) { chk.status = M_SKIP; chk.note = chk.parse_note || "No expected value to check"; return chk; }
  const [by_reg, by_abc] = index || pdf_index(pdf_records);
  let recs, how;
  if (chk.force_pdf_path) { recs = pdf_records.filter((p) => p.path === chk.force_pdf_path); how = "chosen by hand"; }
  else {
    recs = by_reg.get(norm_reg(chk.enrollment)) || [];
    how = "";
    if (!recs.length && chk.abc && by_abc.has(norm_abc(chk.abc))) { recs = by_abc.get(norm_abc(chk.abc)); how = "matched by ABC ID"; }
  }
  if (!recs.length) {
    chk.status = M_NOPDF;
    const k = norm_reg(chk.enrollment);
    const near = [];
    if (k) for (const x of by_reg.keys()) if (Math.abs(py.len(x) - py.len(k)) <= 1) near.push([edit_distance(k, x), x]);
    near.sort(py.cmpTuple);
    chk.note = !k ? "No enrollment number on this row" : "No uploaded PDF has this enrollment number";
    if (near.length && near[0][0] <= 1) chk.note += `. Closest PDF: ${by_reg.get(near[0][1])[0].file} (${near[0][1]})`;
    return chk;
  }
  const rec = recs.find((p) => p.fields[chk.field]) || recs[0];
  chk.pdf_file = rec.file; chk.pdf_path = rec.path; chk.pdf_page = rec.page;
  chk.pdf_value = chk.force_pdf_value !== null && chk.force_pdf_value !== undefined ? chk.force_pdf_value : pyStr(rec.fields[chk.field] ?? "");
  const notes = [how, chk.parse_note].filter((x) => x);
  if (rec.unreadable) notes.push(rec.note);
  if (!chk.pdf_value) {
    chk.status = M_NOFIELD;
    chk.note = [`${field_label(chk)} not found in ${rec.file}`, ...notes].join("; ");
    return chk;
  }
  const fr = compare_field(chk.field, chk.pdf_value, chk.expected, s);
  chk.score = fr.score;
  if (fr.status === ST_MATCH) chk.status = M_CORRECT;
  else if (chk.claimed && _same(chk.field, chk.pdf_value, chk.claimed, s)) chk.status = M_OLD; // wins over close match
  else if (fr.status === ST_CLOSE) chk.status = M_MINOR;
  else if (fr.status === ST_REVIEW) chk.status = M_REVIEW;
  else chk.status = M_WRONG;
  if (chk.status !== M_CORRECT) notes.unshift(`PDF has "${fr.pdf_value}", expected "${fr.excel_value}" (${fr.note})`);
  else notes.unshift(fr.note);
  chk.note = notes.join("; ");
  return chk;
}

export function run_checks(checks, pdf_records, s) {
  const idx = pdf_index(pdf_records);
  for (const c of checks) evaluate(c, pdf_records, s, idx);
  return checks;
}

export function m_counts(checks) {
  const c = Object.fromEntries(M_ORDER.map((k) => [k, 0]));
  for (const x of checks) c[x.status] = (c[x.status] || 0) + 1;
  return c;
}
