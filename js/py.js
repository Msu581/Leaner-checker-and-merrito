/*
 * py.js - small helpers that reproduce Python behaviour the verifier depends on.
 *
 * JavaScript and Python differ in places that change results: what counts as whitespace,
 * what \d / \w / \b match, how round() breaks ties, how floats are printed, how strings sort.
 * Every function here mirrors the Python built-in named in its comment.
 */

// Characters Python's str.isspace() / str.split() / re "\s" treat as whitespace.
export const WS = "\\t\\n\\x0b\\x0c\\r\\x1c-\\x1f \\x85\\xa0\\u1680\\u2000-\\u200a\\u2028\\u2029\\u202f\\u205f\\u3000";
const WS_CLASS = new RegExp(`[${WS}]`, "u");
const WS_RUN = new RegExp(`[${WS}]+`, "u");
const WS_LEAD = new RegExp(`^[${WS}]+`, "u");
const WS_TRAIL = new RegExp(`[${WS}]+$`, "u");

// Python re (str patterns): \w = Unicode letters, numbers and underscore; \b is the boundary between \w and \W.
export const W = "\\p{L}\\p{N}_";
export const B = `(?:(?<=[${W}])(?![${W}])|(?<![${W}])(?=[${W}]))`;
// Python "$" (no MULTILINE) matches at the end or just before a final newline.
export const END = "(?=\\n?$)";

/** str.strip() */
export function strip(s) {
  return String(s).replace(WS_LEAD, "").replace(WS_TRAIL, "");
}

/** str.strip(chars) */
export function stripChars(s, chars) {
  let a = 0, b = s.length;
  while (a < b && chars.includes(s[a])) a++;
  while (b > a && chars.includes(s[b - 1])) b--;
  return s.slice(a, b);
}

/** str.split() with no argument */
export function split(s) {
  return String(s).split(WS_RUN).filter((x) => x !== "");
}

export function isSpace(ch) {
  return WS_CLASS.test(ch);
}

/** len(str): code points, not UTF-16 units */
export function len(s) {
  let n = 0;
  for (const _ of String(s)) n++;
  return n;
}

/** Python string ordering (by code point). */
export function cmpStr(a, b) {
  if (a === b) return 0;
  const A = Array.from(a), Bb = Array.from(b);
  const n = Math.min(A.length, Bb.length);
  for (let i = 0; i < n; i++) {
    if (A[i] !== Bb[i]) return A[i].codePointAt(0) - Bb[i].codePointAt(0);
  }
  return A.length - Bb.length;
}

/** Python tuple ordering for tuples of numbers/strings. */
export function cmpTuple(a, b) {
  for (let i = 0; i < Math.min(a.length, b.length); i++) {
    const x = a[i], y = b[i];
    const c = typeof x === "string" ? cmpStr(x, y) : (x < y ? -1 : x > y ? 1 : 0);
    if (c) return c;
  }
  return a.length - b.length;
}

/** int(str) for a run of Unicode decimal digits (Python accepts e.g. Devanagari digits). */
export function pyInt(s) {
  let n = 0;
  for (const ch of String(s)) n = n * 10 + digitValue(ch);
  return n;
}

const ND = /\p{Nd}/u;
export function digitValue(ch) {
  const cp = ch.codePointAt(0);
  if (cp >= 48 && cp <= 57) return cp - 48;
  // Unicode decimal digits come in runs of ten starting at zero.
  let start = cp;
  while (start > 0 && ND.test(String.fromCodePoint(start - 1))) start--;
  return (cp - start) % 10;
}

/** Replace Unicode decimal digits with ASCII ones. */
export function asciiDigits(s) {
  return String(s).replace(/\p{Nd}/gu, (ch) => String(digitValue(ch)));
}

// ------------------------------------------------------------------ exact float formatting
/** Exact decimal expansion of a finite double: {neg, digits, exp} meaning 0.digits x 10^exp ... see below.
 *  Returned as integer digit string `int` and number of digits after the point `frac`. */
function exactDecimal(v) {
  const buf = new DataView(new ArrayBuffer(8));
  buf.setFloat64(0, v);
  const hi = buf.getUint32(0), lo = buf.getUint32(4);
  const neg = (hi >>> 31) === 1;
  const e = (hi >>> 20) & 0x7ff;
  let m = (BigInt(hi & 0xfffff) << 32n) | BigInt(lo);
  let e2;
  if (e === 0) e2 = -1074;
  else { m |= 1n << 52n; e2 = e - 1075; }
  if (m === 0n) return { neg, digits: "0", frac: 0 };
  if (e2 >= 0) return { neg, digits: (m << BigInt(e2)).toString(), frac: 0 };
  const k = -e2;
  return { neg, digits: (m * 5n ** BigInt(k)).toString(), frac: k };
}

/** Round an exact decimal (digits with `frac` decimals) to `nd` decimals, ties to even. Returns {neg, intPart, fracPart}. */
function roundHalfEven(v, nd) {
  const { neg, digits, frac } = exactDecimal(v);
  let d = digits.padStart(frac + 1, "0");
  let intLen = d.length - frac;
  if (frac <= nd) {
    return { neg, intPart: d.slice(0, intLen), fracPart: d.slice(intLen).padEnd(nd, "0") };
  }
  const keep = d.slice(0, intLen + nd);
  const rest = d.slice(intLen + nd);
  let up = false;
  if (rest[0] > "5") up = true;
  else if (rest[0] === "5") {
    if (/[1-9]/.test(rest.slice(1))) up = true;
    else up = (Number(keep[keep.length - 1] || "0") % 2) === 1; // exact tie: round to even
  }
  let n = BigInt(keep || "0") + (up ? 1n : 0n);
  let s = n.toString().padStart(nd + 1, "0");
  return { neg, intPart: s.slice(0, s.length - nd), fracPart: s.slice(s.length - nd) };
}

/** round(x, nd) for floats */
export function pyRound(x, nd) {
  if (!Number.isFinite(x) || x === 0) return x;
  const r = roundHalfEven(x, nd);
  const v = Number(`${r.intPart}${nd ? "." + r.fracPart : ""}`);
  return r.neg ? -v : v;
}

/** round(x) with no ndigits (ties to even). */
export function pyRoundInt(x) {
  const f = Math.floor(x);
  const diff = x - f;
  if (diff < 0.5) return f;
  if (diff > 0.5) return f + 1;
  return f % 2 === 0 ? f : f + 1;
}

/** f"{x:.{nd}f}" */
export function fmtFixed(x, nd) {
  if (Number.isNaN(x)) return "nan";
  if (!Number.isFinite(x)) return x > 0 ? "inf" : "-inf";
  const r = roundHalfEven(x, nd);
  return `${r.neg ? "-" : ""}${r.intPart}${nd ? "." + r.fracPart : ""}`;
}

/** f"{x:.0%}" */
export function fmtPercent0(x) {
  return fmtFixed(x * 100, 0) + "%";
}

/** format(v, ".15g") */
export function fmtG15(v) {
  if (Number.isNaN(v)) return "nan";
  if (!Number.isFinite(v)) return v > 0 ? "inf" : "-inf";
  if (v === 0) return Object.is(v, -0) ? "-0" : "0";
  const P = 15;
  const { neg, digits, frac } = exactDecimal(v);
  const sig = digits.replace(/^0+/, "");
  // value = sig x 10^-frac, so the leading digit sits at 10^(len-1-frac)
  const lead = sig.length - 1 - frac;
  // round to P significant digits, ties to even
  let mant;
  let exp = lead;
  if (sig.length <= P) mant = sig;
  else {
    const keep = sig.slice(0, P), rest = sig.slice(P);
    let up = rest[0] > "5" || (rest[0] === "5" && (/[1-9]/.test(rest.slice(1)) || Number(keep[P - 1]) % 2 === 1));
    let n = BigInt(keep) + (up ? 1n : 0n);
    mant = n.toString();
    if (mant.length > P) { mant = mant.slice(0, P); exp += 1; }
  }
  mant = mant.replace(/0+$/, "") || "0";
  let out;
  if (exp < -4 || exp >= P) {
    const m = mant.length > 1 ? `${mant[0]}.${mant.slice(1)}` : mant;
    const es = Math.abs(exp) < 10 ? `0${Math.abs(exp)}` : String(Math.abs(exp));
    out = `${m}e${exp < 0 ? "-" : "+"}${es}`;
  } else if (exp >= 0) {
    const intp = mant.length > exp + 1 ? mant.slice(0, exp + 1) : mant.padEnd(exp + 1, "0");
    const fr = mant.slice(exp + 1);
    out = fr ? `${intp}.${fr}` : intp;
  } else {
    out = `0.${"0".repeat(-exp - 1)}${mant}`;
  }
  return (neg ? "-" : "") + out;
}

/** repr(float) / str(float) */
export function floatRepr(x) {
  if (Number.isNaN(x)) return "nan";
  if (!Number.isFinite(x)) return x > 0 ? "inf" : "-inf";
  if (x === 0) return Object.is(x, -0) ? "-0.0" : "0.0";
  const [m, e] = Math.abs(x).toExponential().split("e"); // shortest round-trip digits
  const digits = m.replace(".", "");
  const exp = Number(e);
  let out;
  if (exp < -4 || exp >= 16) {
    const mm = digits.length > 1 ? `${digits[0]}.${digits.slice(1)}` : digits;
    out = `${mm}e${exp < 0 ? "-" : "+"}${String(Math.abs(exp)).padStart(2, "0")}`;
  } else if (exp >= 0) {
    const intp = digits.length > exp + 1 ? digits.slice(0, exp + 1) : digits.padEnd(exp + 1, "0");
    out = `${intp}.${digits.slice(exp + 1) || "0"}`;
  } else {
    out = `0.${"0".repeat(-exp - 1)}${digits}`;
  }
  return (x < 0 ? "-" : "") + out;
}

/** str(int(v)) for an integral double (exact, like Python's arbitrary-precision int). */
export function intStr(v) {
  return BigInt(v).toString();
}

/** Python max(iterable, key=...) - first maximum wins. */
export function maxBy(arr, key) {
  let best, bestK, first = true;
  for (const x of arr) {
    const k = key(x);
    if (first || k > bestK) { best = x; bestK = k; first = false; }
  }
  return best;
}

/** Python Counter-like: Map preserving first-seen order. */
export function counter(keys) {
  const m = new Map();
  for (const k of keys) m.set(k, (m.get(k) || 0) + 1);
  return m;
}

/** Local date/time strings like Python's datetime.now().strftime(...) */
export function nowStamp(fmt) {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  const map = { Y: String(d.getFullYear()), m: p(d.getMonth() + 1), d: p(d.getDate()), H: p(d.getHours()), M: p(d.getMinutes()), S: p(d.getSeconds()) };
  return fmt.replace(/%([YmdHMS])/g, (_, k) => map[k]);
}
