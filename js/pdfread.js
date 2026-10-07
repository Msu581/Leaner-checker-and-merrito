/*
 * pdfread.js - read_pdf() from core.py, with pdf.js in place of pdfplumber.
 *
 * pdfplumber's extract_words(x_tolerance=1.5, y_tolerance=2) builds words from single characters:
 * a new word starts at a space, or where the gap to the previous character is more than 1.5 pt.
 * Each character box spans [baseline + descent*size, baseline + descent*size + size] (pdfminer).
 * pdf.js returns text runs instead, so the runs are split back into characters (widths shared
 * evenly within a run) and words are rebuilt with the same rules. The words then go through the
 * unchanged _lines_from_words() / extract_fields() logic.
 */
import { _lines_from_words, extract_fields, PdfRecord } from "./core.js";
import * as py from "./py.js";

const X_TOL = 1.5;
const Y_TOL = 2;
// pdfplumber's expand_ligatures=True
const LIGATURES = { "\ufb00": "ff", "\ufb03": "ffi", "\ufb04": "ffl", "\ufb01": "fi", "\ufb02": "fl", "\ufb06": "st", "\ufb05": "st" };

/** pdf.js text items -> pdfplumber-like words {text, x0, x1, top, bottom}. */
export function wordsFromTextContent(content) {
  const chars = [];
  for (const it of content.items) {
    if (typeof it.str !== "string" || !it.str.length) continue;
    const t = it.transform;
    const size = Math.hypot(t[2], t[3]) || Math.abs(t[3]) || 1;
    const style = content.styles[it.fontName] || {};
    const descent = Number.isFinite(style.descent) ? style.descent : 0; // unknown font metrics: pdfminer also uses 0
    const y0 = t[5] + descent * size; // device-space bottom of the glyph box
    const top = -(y0 + size), bottom = -y0; // pdfplumber measures down from the page top
    const glyphs = Array.from(it.str);
    const w = (it.width || 0) / glyphs.length;
    glyphs.forEach((g, i) => {
      chars.push({ text: g, x0: t[4] + i * w, x1: t[4] + (i + 1) * w, top, bottom });
    });
  }
  // cluster into lines by top (pdfplumber cluster_objects with y_tolerance), then left to right
  const order = [...chars].sort((a, b) => a.top - b.top);
  const lines = [];
  let cur = null, lastTop = null;
  for (const c of order) {
    if (cur && c.top - lastTop <= Y_TOL) cur.push(c);
    else { cur = [c]; lines.push(cur); }
    lastTop = c.top;
  }
  const words = [];
  for (const line of lines) {
    line.sort((a, b) => a.x0 - b.x0);
    let wd = null, prev = null;
    const flush = () => { if (wd) words.push(wd); wd = null; prev = null; };
    for (const c of line) {
      if (py.isSpace(c.text)) { flush(); continue; }
      if (prev && (c.x0 > prev.x1 + X_TOL || c.x1 < prev.x0 - X_TOL || Math.abs(c.top - prev.top) > Y_TOL)) flush();
      const txt = LIGATURES[c.text] || c.text;
      if (!wd) wd = { text: txt, x0: c.x0, x1: c.x1, top: c.top, bottom: c.bottom };
      else {
        wd.text += txt;
        wd.x0 = Math.min(wd.x0, c.x0); wd.x1 = Math.max(wd.x1, c.x1);
        wd.top = Math.min(wd.top, c.top); wd.bottom = Math.max(wd.bottom, c.bottom);
      }
      prev = c;
    }
    flush();
  }
  return words;
}

/**
 * read_pdf(path, use_ocr=False): one PdfRecord per page that has student details.
 * pdfjs: the pdf.js module; data: Uint8Array; name: file name; path: a unique key for the file;
 * assets: {cMapUrl, standardFontDataUrl} folders shipped with pdf.js (needed for some fonts).
 * OCR is not available in the browser, so use_ocr is always False here.
 */
export async function read_pdf(pdfjs, data, name, path, assets = {}) {
  const doc = await pdfjs.getDocument({
    data, disableFontFace: true, useSystemFonts: false, isEvalSupported: false,
    cMapUrl: assets.cMapUrl, cMapPacked: true, standardFontDataUrl: assets.standardFontDataUrl,
  }).promise;
  const recs = [];
  try {
    const n = doc.numPages;
    for (let i = 1; i <= n; i++) {
      const page = await doc.getPage(i);
      const content = await page.getTextContent({ disableNormalization: true });
      const fields = extract_fields(_lines_from_words(wordsFromTextContent(content)));
      if (Object.keys(fields).length) recs.push(PdfRecord(name, path, i, n, fields, "text"));
      page.cleanup();
    }
    if (!recs.length) {
      recs.push(PdfRecord(name, path, 1, n, {}, "text", true, "No readable text found. This looks like a scanned PDF; enable OCR."));
    }
  } finally {
    await doc.destroy();
  }
  return recs;
}
