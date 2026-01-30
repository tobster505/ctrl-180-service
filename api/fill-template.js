/**
 * CTRL 180 Export Service · fill-template (V7 — AUTO-SELECT TEMPLATE FIRST + ignore q.tpl by default)
 *
 * ✅ Chart block is intentionally unchanged from your V5/V6.x (same fetch/embed + cx/cy/cw/ch overrides).
 * ✅ Every text area supports per-box coordinate overrides (x/y/w/h/size/maxLines/align + dx/dy).
 * ✅ Paragraph blocks + Question blocks are separate boxes (so you can move them independently).
 * ✅ WorkWith will ONLY render if Gen provides it; otherwise page 6 stays blank (no fallbacks, no invention).
 *
 * NEW in V7:
 * ✅ Template selection is payload-driven (dominant + second) and is the PRIMARY path.
 * ✅ q.tpl is ignored by default; only honoured when override=1 (or debug=1).
 * ✅ safeCombo whitelist prevents unexpected template loads.
 */

export const config = { runtime: "nodejs" };

import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

/* ───────────── utils ───────────── */
const S = (v, fb = "") => (v == null ? String(fb) : String(v));
const N = (v, fb = 0) => (Number.isFinite(+v) ? +v : fb);

const norm = (v, fb = "") =>
  String(v ?? fb)
    .normalize("NFKC")
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2010-\u2014]/g, "-")
    .replace(/\u2026/g, "...")
    .replace(/\u00A0/g, " ")
    .replace(/[•·]/g, "-")
    .replace(/\u2194/g, "<->").replace(/\u2192/g, "->").replace(/\u2190/g, "<-")
    .replace(/\u2191/g, "^").replace(/\u2193/g, "v").replace(/[\u2196-\u2199]/g, "->")
    .replace(/\u21A9/g, "<-").replace(/\u21AA/g, "->")
    .replace(/\u00D7/g, "x")
    .replace(/[\u200B-\u200D\u2060]/g, "")
    .replace(/[\uD800-\uDFFF]/g, "")
    .replace(/[\uE000-\uF8FF]/g, "")
    .replace(/\t/g, " ")
    .replace(/\r\n?/g, "\n")
    .replace(/[ \f\v]+/g, " ")
    .replace(/[ \t]+\n/g, "\n")
    .trim();

const isObj = (v) => v && typeof v === "object" && !Array.isArray(v);

const bullets = (arr = []) =>
  arr.map((s) => norm(s || "")).filter(Boolean).map((s) => `- ${s}`).join("\n");

const pageOrNull = (pages, idx0) => (pages[idx0] ?? null);

/* ───────── DEFAULT LAYOUT (BASELINE) ───────── */
const DEFAULT_LAYOUT = {
  pages: {
    p1: {
      name: { x: 60, y: 458, w: 500, h: 60, size: 30, align: "center", maxLines: 1 },
      date: { x: 230, y: 613, w: 500, h: 40, size: 25, align: "left", maxLines: 1 },
    },

    p2: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p3: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p4: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p5: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p6: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p7: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },
    p8: { hdrName: { x: 380, y: 51, w: 400, h: 24, size: 13, align: "left", maxLines: 1 } },

    // Page 3 — Exec Summary split into paragraph + questions (fallback to legacy block if needed)
    p3Text: {
      exec_summary_text: { x: 25, y: 360, w: 550, h: 170, size: 16, align: "left", maxLines: 10 },
      exec_summary_q:    { x: 25, y: 540, w: 550, h: 170, size: 16, align: "left", maxLines: 10 },
      summary:           { x: 25, y: 380, w: 550, h: 250, size: 16, align: "left", maxLines: 13 }, // legacy
    },

    // Page 4 — Overview split + chart target (chart embed block stays unchanged)
    p4Text: {
      ctrl_overview_text: { x: 25, y: 155, w: 190, h: 240, size: 15, align: "left", maxLines: 30 },
      ctrl_overview_q:    { x: 25, y: 575, w: 550, h: 420, size: 16, align: "left", maxLines: 23 },
      chart: { x: 250, y: 160, w: 320, h: 320 },
    },

    // Page 5 — Deepdive + Themes split (fallbacks kept)
    p5Text: {
      sequence_text: { x: 25, y: 140, w: 550, h: 150, size: 16, align: "left", maxLines: 8 },
      sequence_q:    { x: 25, y: 300, w: 550, h: 140, size: 16, align: "left", maxLines: 7 },

      theme_text:    { x: 25, y: 500, w: 550, h: 80,  size: 16, align: "left", maxLines: 4 },
      theme_q:       { x: 25, y: 590, w: 550, h: 120, size: 16, align: "left", maxLines: 6 },

      sequence: { x: 25, y: 140, w: 550, h: 240, size: 16, align: "left", maxLines: 13 }, // legacy
      theme:    { x: 25, y: 540, w: 550, h: 160, size: 16, align: "left", maxLines: 9 },  // legacy
    },

    // Page 6 — WorkWith split, BUT we will only render if Gen provides it
    p6WorkWith: {
      collabCol_text: { x: 30,  y: 300, w: 270, h: 170, size: 15, align: "left", maxLines: 6 },
      collabCol_q:    { x: 30,  y: 480, w: 270, h: 280, size: 15, align: "left", maxLines: 10 },

      collabLe_text:  { x: 320, y: 300, w: 260, h: 170, size: 15, align: "left", maxLines: 6 },
      collabLe_q:     { x: 320, y: 480, w: 260, h: 280, size: 15, align: "left", maxLines: 10 },

      // legacy combined (kept but NOT used in Gen-or-blank mode)
      collabCol: { x: 30,  y: 300, w: 270, h: 420, size: 15, align: "left", maxLines: 14 },
      collabLe:  { x: 320, y: 300, w: 260, h: 420, size: 15, align: "left", maxLines: 14 },
    },

    // Page 7 — Actions (from tips split into 3)
    p7Actions: {
      act1: { x: 50,  y: 320, w: 440, h: 95, size: 17, align: "left", maxLines: 5 },
      act2: { x: 100, y: 470, w: 440, h: 95, size: 17, align: "left", maxLines: 5 },
      act3: { x: 50,  y: 610, w: 440, h: 95, size: 17, align: "left", maxLines: 5 },
    },
  },
};

/* ───────────── payload helpers ───────────── */
function parseDataParam(b64ish) {
  if (!b64ish) return {};
  let s = String(b64ish);
  try { s = decodeURIComponent(s); } catch {}
  s = s.replace(/-/g, "+").replace(/_/g, "/");
  while (s.length % 4) s += "=";
  try { return JSON.parse(Buffer.from(s, "base64").toString("utf8")); }
  catch { return {}; }
}

async function readPayload(req) {
  const q = req.method === "POST" ? (req.body || {}) : (req.query || {});
  if (q.data) return parseDataParam(q.data);
  if (req.method === "POST" && !q.data) {
    try {
      return typeof req.json === "function" ? await req.json() : (req.body || {});
    } catch {}
  }
  return {};
}

/* ───────────── template auto-select helpers (V7) ───────────── */
function resolveStateKey(any) {
  const raw = S(any, "").trim();
  if (!raw) return null;

  // direct key
  const up = raw.toUpperCase();
  if (["C","T","R","L"].includes(up)) return up;

  // from label-ish text
  const low = raw.toLowerCase();
  if (low.includes("concealed")) return "C";
  if (low.includes("triggered")) return "T";
  if (low.includes("regulated")) return "R";
  if (low.includes("lead")) return "L";

  // first-char fallback (e.g. "Triggered (Developing)" -> T)
  const c = up.charAt(0);
  if (["C","T","R","L"].includes(c)) return c;

  return null;
}

function computeDomSecondFromPayload(src) {
  const domKey =
    resolveStateKey(src?.dominantKey) ||
    resolveStateKey(src?.domKey) ||
    resolveStateKey(src?.dominantLabel) ||
    resolveStateKey(src?.domLabel) ||
    resolveStateKey(src?.dominantState) ||
    resolveStateKey(src?.domState) ||
    resolveStateKey(src?.subject?.dominantKey) ||
    resolveStateKey(src?.subject?.domKey) ||
    null;

  const secondKey =
    resolveStateKey(src?.secondKey) ||
    resolveStateKey(src?.secKey) ||
    resolveStateKey(src?.secondLabel) ||
    resolveStateKey(src?.secondLabel) ||
    resolveStateKey(src?.secondState) ||
    resolveStateKey(src?.subject?.secondKey) ||
    resolveStateKey(src?.subject?.secKey) ||
    null;

  return {
    dominantKey: domKey,
    secondKey: secondKey,
    templateKey: domKey && secondKey ? `${domKey}${secondKey}` : ""
  };
}

/* ───────────── drawing helpers ───────────── */
function drawTextBox(page, font, text, spec = {}, opts = {}) {
  if (!page) return;
  const { x = 40, y = 40, w = 540, size = 12, lineGap = 3, color = rgb(0, 0, 0), align = "left", h } = spec;
  const lineHeight = Math.max(1, size) + lineGap;
  const maxLines = opts.maxLines ?? spec.maxLines ?? (h ? Math.max(1, Math.floor(h / lineHeight)) : 40);

  const hard = norm(text || "");
  if (!hard) return;

  const lines = hard.split(/\n/).map((s) => s.trim());
  const wrapped = [];
  const widthOf = (s) => font.widthOfTextAtSize(s, Math.max(1, size));

  const wrapLine = (ln) => {
    const words = ln.split(/\s+/);
    let cur = "";
    for (let i = 0; i < words.length; i++) {
      const nxt = cur ? `${cur} ${words[i]}` : words[i];
      if (widthOf(nxt) <= w || !cur) cur = nxt;
      else { wrapped.push(cur); cur = words[i]; }
    }
    wrapped.push(cur);
  };
  for (const ln of lines) wrapLine(ln);

  const out = wrapped.slice(0, maxLines);
  const pageH = page.getHeight();
  let yCursor = pageH - y;

  for (const ln of out) {
    let xDraw = x;
    const wLn = widthOf(ln);
    if (align === "center") xDraw = x + (w - wLn) / 2;
    else if (align === "right") xDraw = x + (w - wLn);

    page.drawText(ln, { x: xDraw, y: yCursor - size, size: Math.max(1, size), font, color });
    yCursor -= lineHeight;
  }
}

function drawOverlayBox(page, fonts, text, spec = {}) {
  if (!page) return;
  const fontReg = fonts.reg;
  const fontBold = fonts.bold;

  const { x = 40, y = 40, w = 540, size = 15, lineGap = 6, color = rgb(0, 0, 0), align = "left", maxLines = 120 } = spec;
  const hard = norm(text || "");
  if (!hard) return;

  const pageH = page.getHeight();
  let yCursor = pageH - y;
  let usedLines = 0;

  const pushLine = (ln, { font, fSize, indent = 0 }) => {
    if (usedLines >= maxLines) return;
    const widthOf = (s) => font.widthOfTextAtSize(s, fSize);

    const words = ln.split(/\s+/);
    let current = "";
    const segments = [];

    for (let i = 0; i < words.length; i++) {
      const next = current ? current + " " + words[i] : words[i];
      if (widthOf(next) <= (w - indent) || !current) current = next;
      else { segments.push(current); current = words[i]; }
    }
    if (current) segments.push(current);

    for (let s of segments) {
      if (usedLines >= maxLines) break;
      let xDraw = x + indent;
      const wLn = widthOf(s);
      if (align === "center") xDraw = x + (w - wLn) / 2;
      else if (align === "right") xDraw = x + (w - wLn);

      page.drawText(s, { x: xDraw, y: yCursor - fSize, size: fSize, font, color });
      yCursor -= (fSize + lineGap);
      usedLines++;
    }
  };

  const rawLines = hard.split(/\n+/);
  for (let raw of rawLines) {
    const line = raw.trim();
    if (!line) { yCursor -= (size + lineGap); continue; }

    const isHeading = /^what\b.*:\s*$/i.test(line);
    const isBullet  = /^-\s+/.test(line);

    if (isHeading) {
      yCursor -= (size + lineGap) * 2;
      pushLine(line.replace(/:\s*$/, ":"), { font: fontBold, fSize: size + 1, indent: 0 });
      yCursor -= (size + lineGap) * 0.5;
      continue;
    }

    if (isBullet) {
      pushLine(line.replace(/^-\s+/, "• "), { font: fontReg, fSize: size, indent: 10 });
      continue;
    }

    pushLine(line, { font: fontReg, fSize: size, indent: 0 });
  }
}

/* ───────────── template loader ───────────── */
async function loadTemplateBytesLocal(filename) {
  const fname = String(filename || "").trim();
  if (!fname.endsWith(".pdf")) throw new Error(`Invalid template filename: ${fname}`);

  const __file = fileURLToPath(import.meta.url);
  const __dir  = path.dirname(__file);

  const candidates = [
    path.join(__dir, "..", "..", "public", fname),
    path.join(__dir, "..", "public", fname),
    path.join(process.cwd(), "public", fname),
  ];

  let lastErr;
  for (const pth of candidates) {
    try { return await fs.readFile(pth); }
    catch (err) { lastErr = err; }
  }
  throw new Error(`Template not found in /public: ${fname} (${lastErr?.message || "no detail"})`);
}

/* ───────────── chart fetch/embed (UNCHANGED) ───────────── */
async function fetchBytes(url, timeoutMs = 9000) {
  if (!url) return { ok:false, reason:"no_url", url:"", contentType:"", bytes:null };
  const u = String(url).trim();
  if (!/^https?:\/\//i.test(u)) return { ok:false, reason:"not_http", url:u, contentType:"", bytes:null };

  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), timeoutMs);

  try {
    const r = await fetch(u, { signal: ctrl.signal });
    if (!r.ok) return { ok:false, reason:`http_${r.status}`, url:u, contentType:"", bytes:null };
    const ct = (r.headers.get("content-type") || "").toLowerCase();
    const ab = await r.arrayBuffer();
    return { ok:true, reason:"ok", url:u, contentType:ct, bytes:new Uint8Array(ab) };
  } catch (e) {
    return { ok:false, reason:String(e?.name || e?.message || "fetch_error"), url:u, contentType:"", bytes:null };
  } finally {
    clearTimeout(t);
  }
}

function looksPng(u, ct = "") {
  const s = String(u || "").toLowerCase();
  return ct.includes("png") || s.endsWith(".png") || s.includes(".png?");
}
function looksJpg(u, ct = "") {
  const s = String(u || "").toLowerCase();
  return ct.includes("jpeg") || ct.includes("jpg") || s.endsWith(".jpg") || s.endsWith(".jpeg") || s.includes(".jpg?") || s.includes(".jpeg?");
}

/* ───────────── split helpers ───────────── */
function splitTipsInto3(tips) {
  const t = norm(tips || "");
  if (!t) return ["", "", ""];
  const parts = t.split(/\n{2,}/).map(s => s.trim()).filter(Boolean);
  const out = [parts[0] || "", parts[1] || "", parts[2] || ""];
  if (parts.length <= 3) return out;
  out[2] = norm([out[2], ...parts.slice(3)].filter(Boolean).join("\n\n"));
  return out;
}

/* ───────────── layout override engine ───────────── */
function applyBoxOverrides(spec, q, id) {
  const out = { ...(spec || {}) };

  // delta shifts
  if (q[`${id}_dx`] != null) out.x = N(out.x, 0) + N(q[`${id}_dx`], 0);
  if (q[`${id}_dy`] != null) out.y = N(out.y, 0) + N(q[`${id}_dy`], 0);

  // absolute overrides
  if (q[`${id}_x`] != null) out.x = N(q[`${id}_x`], out.x);
  if (q[`${id}_y`] != null) out.y = N(q[`${id}_y`], out.y);
  if (q[`${id}_w`] != null) out.w = N(q[`${id}_w`], out.w);
  if (q[`${id}_h`] != null) out.h = N(q[`${id}_h`], out.h);
  if (q[`${id}_size`] != null) out.size = N(q[`${id}_size`], out.size);
  if (q[`${id}_lineGap`] != null) out.lineGap = N(q[`${id}_lineGap`], out.lineGap);
  if (q[`${id}_maxLines`] != null) out.maxLines = N(q[`${id}_maxLines`], out.maxLines);
  if (q[`${id}_align`] != null) out.align = String(q[`${id}_align`]);

  return out;
}

function buildLayoutWithOverrides(q) {
  const L = JSON.parse(JSON.stringify(DEFAULT_LAYOUT));
  const P = L.pages;

  P.p1.name = applyBoxOverrides(P.p1.name, q, "p1_name");
  P.p1.date = applyBoxOverrides(P.p1.date, q, "p1_date");

  for (let i = 2; i <= 8; i++) {
    const k = `p${i}`;
    P[k].hdrName = applyBoxOverrides(P[k].hdrName, q, `p${i}_hdrName`);
  }

  // p3
  P.p3Text.exec_summary_text = applyBoxOverrides(P.p3Text.exec_summary_text, q, "p3_exec_summary_text");
  P.p3Text.exec_summary_q    = applyBoxOverrides(P.p3Text.exec_summary_q,    q, "p3_exec_summary_q");
  P.p3Text.summary           = applyBoxOverrides(P.p3Text.summary,           q, "p3_summary");

  // p4
  P.p4Text.ctrl_overview_text = applyBoxOverrides(P.p4Text.ctrl_overview_text, q, "p4_ctrl_overview_text");
  P.p4Text.ctrl_overview_q    = applyBoxOverrides(P.p4Text.ctrl_overview_q,    q, "p4_ctrl_overview_q");
  // chart stays controlled by cx/cy/cw/ch (see chart block)

  // p5
  P.p5Text.sequence_text = applyBoxOverrides(P.p5Text.sequence_text, q, "p5_sequence_text");
  P.p5Text.sequence_q    = applyBoxOverrides(P.p5Text.sequence_q,    q, "p5_sequence_q");
  P.p5Text.theme_text    = applyBoxOverrides(P.p5Text.theme_text,    q, "p5_theme_text");
  P.p5Text.theme_q       = applyBoxOverrides(P.p5Text.theme_q,       q, "p5_theme_q");
  P.p5Text.sequence      = applyBoxOverrides(P.p5Text.sequence,      q, "p5_sequence");
  P.p5Text.theme         = applyBoxOverrides(P.p5Text.theme,         q, "p5_theme");

  // p6
  P.p6WorkWith.collabCol_text = applyBoxOverrides(P.p6WorkWith.collabCol_text, q, "p6_collabCol_text");
  P.p6WorkWith.collabCol_q    = applyBoxOverrides(P.p6WorkWith.collabCol_q,    q, "p6_collabCol_q");
  P.p6WorkWith.collabLe_text  = applyBoxOverrides(P.p6WorkWith.collabLe_text,  q, "p6_collabLe_text");
  P.p6WorkWith.collabLe_q     = applyBoxOverrides(P.p6WorkWith.collabLe_q,     q, "p6_collabLe_q");
  P.p6WorkWith.collabCol      = applyBoxOverrides(P.p6WorkWith.collabCol,      q, "p6_collabCol");
  P.p6WorkWith.collabLe       = applyBoxOverrides(P.p6WorkWith.collabLe,       q, "p6_collabLe");

  // p7
  P.p7Actions.act1 = applyBoxOverrides(P.p7Actions.act1, q, "p7_act1");
  P.p7Actions.act2 = applyBoxOverrides(P.p7Actions.act2, q, "p7_act2");
  P.p7Actions.act3 = applyBoxOverrides(P.p7Actions.act3, q, "p7_act3");

  return L;
}

/* ───────────── handler ───────────── */
export default async function handler(req, res) {
  let tpl = "";
  let usingFallback = false;
  let chartFetch = { ok:false, reason:"not_attempted", url:"", contentType:"", bytes:null };
  let payloadSize = 0;

  // V7 debug extras
  let tplDecision = null;
  let domSecond = null;

  try {
    const q = req.method === "POST" ? (req.body || {}) : (req.query || {});
    const debugMode = String(q.debug || "").trim() === "1";

    const FALLBACK_TPL = "CTRL_PoC_180_Assessment_Report_template_fallback.pdf";

    // Read payload FIRST so auto-select is primary
    const src = await readPayload(req);
    payloadSize = Buffer.byteLength(JSON.stringify(src || {}), "utf8");

    // Auto-select (primary)
    domSecond = computeDomSecondFromPayload(src || {});
    const validCombos = new Set(["CT","CL","CR","TC","TR","TL","RC","RT","RL","LC","LR","LT"]);
    const safeCombo = validCombos.has(domSecond.templateKey) ? domSecond.templateKey : "CT";

    const computedTpl = `CTRL_PoC_180_Assessment_Report_template_${safeCombo}.pdf`;

    // Ignore q.tpl by default; only allow in override mode
    const allowOverride = String(q.override || "").trim() === "1" || debugMode;
    const requestedTpl = S(q.tpl || "").trim().replace(/[^A-Za-z0-9._-]/g, "");

    tpl = computedTpl;
    if (allowOverride && requestedTpl) tpl = requestedTpl;

    usingFallback = !tpl;
    if (usingFallback) tpl = FALLBACK_TPL;

    tplDecision = {
      allowOverride,
      requestedTpl: requestedTpl || null,
      dominantKey: domSecond.dominantKey || null,
      secondKey: domSecond.secondKey || null,
      templateKey: domSecond.templateKey || null,
      safeCombo,
      computedTpl,
      finalTpl: tpl
    };

    const pdfBytes = await loadTemplateBytesLocal(tpl);

    // Prefer textV2, then text, then fields, then top-level
    const T = isObj(src?.textV2) ? src.textV2 : (isObj(src?.text) ? src.text : (isObj(src?.fields) ? src.fields : {}));
    const G = isObj(src?.gen) ? src.gen : {};

    // STRICT WorkWith from Gen (or blank)
    const wwColText = norm(G?.adapt_with_colleagues || "");
    const wwColQ    = bullets([G?.adapt_with_colleagues_q1, G?.adapt_with_colleagues_q2]);
    const wwLeText  = norm(G?.adapt_with_leaders || "");
    const wwLeQ     = bullets([G?.adapt_with_leaders_q1, G?.adapt_with_leaders_q2]);
    const hasWW = !!(wwColText || wwColQ || wwLeText || wwLeQ);
    const workWithBuilt = hasWW ? {
      collabCol_text: wwColText,
      collabCol_q:    wwColQ,
      collabLe_text:  wwLeText,
      collabLe_q:     wwLeQ
    } : null;

    const P = {
      name:    norm(src?.person?.fullName || src?.fullName || "Perspective Overlay"),
      // V8 (date-only): broaden date label sources to match user v12.3 payloads
      dateLbl: norm(
        src?.dateLbl ||
        src?.identity?.dateLabel ||
        src?.identity?.dateLbl ||
        src?.person?.dateLabel ||
        src?.person?.dateLbl ||
        src?.date ||
        src?.Date ||
        ""
      ),

      // Legacy blocks (still supported)
      summary:   norm(T?.summary   || src?.summary   || ""),
      sequence:  norm(T?.sequence  || src?.sequence  || ""),
      themepair: norm(T?.themepair || src?.themepair || ""),
      tips:      norm(T?.tips      || src?.tips      || ""),

      // Split blocks (prefer gen pack)
      exec_summary_text: norm(G?.exec_summary || ""),
      exec_summary_q:    bullets([G?.exec_summary_q1, G?.exec_summary_q2, G?.exec_summary_q3, G?.exec_summary_q4]),

      ctrl_overview_text: norm(G?.ctrl_overview || norm(T?.ctrl_overview || src?.ctrl_overview || "")),
      ctrl_overview_q:    bullets([G?.ctrl_overview_q1, G?.ctrl_overview_q2, G?.ctrl_overview_q3, G?.ctrl_overview_q4]) ||
                          norm(T?.ctrl_overviewQ || src?.ctrl_overviewQ || ""), // legacy (safe)

      sequence_text: norm(G?.ctrl_deepdive || ""),
      sequence_q:    bullets([G?.ctrl_deepdive_q1, G?.ctrl_deepdive_q2]),

      theme_text: norm(G?.themes || ""),
      theme_q:    bullets([G?.themes_q1, G?.themes_q2]),

      // WorkWith STRICT: Gen or blank only
      workWithBuilt
    };

    const hasText = !!(isObj(src?.text) || isObj(src?.textV2) || isObj(src?.gen));
    const chartUrl = norm(
      src?.chartUrl ||
      src?.spiderChartUrl ||
      src?.chart?.spiderUrl ||
      ""
    );

    // Layout (with overrides)
    const L = buildLayoutWithOverrides(q).pages;

    // DEBUG JSON response
    if (debugMode) {
      const [a1,a2,a3] = splitTipsInto3(P.tips);

      const built = isObj(P.workWithBuilt) ? P.workWithBuilt : null;
      const wwCol  = built ? (built.collabCol_text || "") : "";
      const wwColQ = built ? (built.collabCol_q || "")    : "";
      const wwLe   = built ? (built.collabLe_text || "")  : "";
      const wwLeQ  = built ? (built.collabLe_q || "")     : "";

      res.setHeader("Content-Type", "application/json; charset=utf-8");
      res.setHeader("X-CTRL-TPL", tpl);
      res.setHeader("X-CTRL-TPL-FALLBACK", usingFallback ? "1" : "0");
      res.setHeader("X-CTRL-DEBUG", "1");

      res.status(200).end(JSON.stringify({
        ok: true,
        debug: true,
        tpl,
        usingFallback,
        tplDecision,
        received: {
          person: src?.person || null,
          dateLbl: src?.dateLbl || null,
          chartUrl: chartUrl || null,
          hasText,
          fields: {
            // legacy
            summary_len: P.summary.length,
            sequence_len: P.sequence.length,
            themepair_len: P.themepair.length,
            tips_len: P.tips.length,

            // split
            exec_summary_text_len: P.exec_summary_text.length,
            exec_summary_q_len: P.exec_summary_q.length,
            ctrl_overview_text_len: P.ctrl_overview_text.length,
            ctrl_overview_q_len: P.ctrl_overview_q.length,
            sequence_text_len: P.sequence_text.length,
            sequence_q_len: P.sequence_q.length,
            theme_text_len: P.theme_text.length,
            theme_q_len: P.theme_q.length,
          },
          hasWorkWith: !!(wwCol || wwLe || wwColQ || wwLeQ),
          page6Preview: {
            collabCol_text_len: wwCol.length,
            collabCol_q_len: wwColQ.length,
            collabLe_text_len: wwLe.length,
            collabLe_q_len: wwLeQ.length
          },
          page7Preview: {
            act1_len: a1.length, act2_len: a2.length, act3_len: a3.length
          },
          layoutPreview: {
            p3_exec_summary_text: L.p3Text.exec_summary_text,
            p3_exec_summary_q: L.p3Text.exec_summary_q,
            p4_ctrl_overview_text: L.p4Text.ctrl_overview_text,
            p4_ctrl_overview_q: L.p4Text.ctrl_overview_q,
            p6_collabCol_text: L.p6WorkWith.collabCol_text,
            p6_collabCol_q: L.p6WorkWith.collabCol_q,
            p6_collabLe_text: L.p6WorkWith.collabLe_text,
            p6_collabLe_q: L.p6WorkWith.collabLe_q
          }
        }
      }, null, 2));
      return;
    }

    // Normal PDF render
    const pdfDoc   = await PDFDocument.load(pdfBytes);
    const font     = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
    const fonts    = { reg: font, bold: fontBold };

    const pages = pdfDoc.getPages();
    const p1 = pageOrNull(pages, 0);
    const p2 = pageOrNull(pages, 1);
    const p3 = pageOrNull(pages, 2);
    const p4 = pageOrNull(pages, 3);
    const p5 = pageOrNull(pages, 4);
    const p6 = pageOrNull(pages, 5);
    const p7 = pageOrNull(pages, 6);
    const p8 = pageOrNull(pages, 7);

    // p1: name & date
    if (p1 && P.name)    drawTextBox(p1, fontBold, P.name,    L.p1.name, { maxLines: L.p1.name.maxLines ?? 1 });
    if (p1 && P.dateLbl) drawTextBox(p1, font,     P.dateLbl, L.p1.date, { maxLines: L.p1.date.maxLines ?? 1 });

    // headers (p2–p8)
    const putHeader = (page, spec) => {
      if (!page || !P.name || !spec) return;
      drawTextBox(page, font, P.name, spec, { maxLines: spec.maxLines ?? 1 });
    };
    putHeader(p2, L.p2.hdrName);
    putHeader(p3, L.p3.hdrName);
    putHeader(p4, L.p4.hdrName);
    putHeader(p5, L.p5.hdrName);
    putHeader(p6, L.p6.hdrName);
    putHeader(p7, L.p7.hdrName);
    putHeader(p8, L.p8.hdrName);

    // ─────────────────────────────────────────────────────────────
    // chart on page 4  ✅ UNCHANGED BLOCK
    // ─────────────────────────────────────────────────────────────
    const CHART4 = { ...L.p4Text.chart };
    if (q.cx != null) CHART4.x = N(q.cx, CHART4.x);
    if (q.cy != null) CHART4.y = N(q.cy, CHART4.y);
    if (q.cw != null) CHART4.w = N(q.cw, CHART4.w);
    if (q.ch != null) CHART4.h = N(q.ch, CHART4.h);

    if (p4 && chartUrl) {
      chartFetch = await fetchBytes(chartUrl, 9000);
      if (chartFetch.ok && chartFetch.bytes && chartFetch.bytes.length) {
        let img = null;
        if (looksPng(chartFetch.url, chartFetch.contentType)) img = await pdfDoc.embedPng(chartFetch.bytes);
        else if (looksJpg(chartFetch.url, chartFetch.contentType)) img = await pdfDoc.embedJpg(chartFetch.bytes);

        if (img) {
          const { x, y, w, h } = CHART4;
          const pageH = p4.getHeight();
          const yBottom = pageH - y - h;
          p4.drawImage(img, { x, y: yBottom, width: w, height: h });
        }
      }
    }

    // p3: summary (prefer split; fallback to legacy block)
    if (p3) {
      const hasSplit = !!(P.exec_summary_text || P.exec_summary_q);
      if (hasSplit) {
        if (P.exec_summary_text) drawOverlayBox(p3, fonts, P.exec_summary_text, L.p3Text.exec_summary_text);
        if (P.exec_summary_q)    drawOverlayBox(p3, fonts, P.exec_summary_q,    L.p3Text.exec_summary_q);
      } else if (P.summary) {
        drawOverlayBox(p3, fonts, P.summary, L.p3Text.summary);
      }
    }

    // p4: overview (split)
    if (p4) {
      if (P.ctrl_overview_text) drawOverlayBox(p4, fonts, P.ctrl_overview_text, L.p4Text.ctrl_overview_text);
      if (P.ctrl_overview_q)    drawOverlayBox(p4, fonts, P.ctrl_overview_q,    L.p4Text.ctrl_overview_q);
    }

    // p5: deepdive + themes (prefer split; fallback to legacy)
    if (p5) {
      const hasSeqSplit = !!(P.sequence_text || P.sequence_q);
      if (hasSeqSplit) {
        if (P.sequence_text) drawOverlayBox(p5, fonts, P.sequence_text, L.p5Text.sequence_text);
        if (P.sequence_q)    drawOverlayBox(p5, fonts, P.sequence_q,    L.p5Text.sequence_q);
      } else if (P.sequence) {
        drawOverlayBox(p5, fonts, P.sequence, L.p5Text.sequence);
      }

      const hasThemeSplit = !!(P.theme_text || P.theme_q);
      if (hasThemeSplit) {
        if (P.theme_text) drawOverlayBox(p5, fonts, P.theme_text, L.p5Text.theme_text);
        if (P.theme_q)    drawOverlayBox(p5, fonts, P.theme_q,    L.p5Text.theme_q);
      } else if (P.themepair) {
        drawOverlayBox(p5, fonts, P.themepair, L.p5Text.theme);
      }
    }

    // p6: WorkWith (STRICT: Gen or blank only)
    if (p6) {
      const built = isObj(P.workWithBuilt) ? P.workWithBuilt : null;
      if (built) {
        const colText = norm(built.collabCol_text || "");
        const colQ    = norm(built.collabCol_q || "");
        const leText  = norm(built.collabLe_text  || "");
        const leQ     = norm(built.collabLe_q     || "");

        if (colText) drawOverlayBox(p6, fonts, colText, L.p6WorkWith.collabCol_text);
        if (colQ)    drawOverlayBox(p6, fonts, colQ,    L.p6WorkWith.collabCol_q);
        if (leText)  drawOverlayBox(p6, fonts, leText,  L.p6WorkWith.collabLe_text);
        if (leQ)     drawOverlayBox(p6, fonts, leQ,     L.p6WorkWith.collabLe_q);
      }
      // else: leave page 6 blank
    }

    // p7: actions (tips split into 3)
    if (p7 && P.tips) {
      const [act1, act2, act3] = splitTipsInto3(P.tips);
      if (act1) drawOverlayBox(p7, fonts, act1, L.p7Actions.act1);
      if (act2) drawOverlayBox(p7, fonts, act2, L.p7Actions.act2);
      if (act3) drawOverlayBox(p7, fonts, act3, L.p7Actions.act3);
    }

    const bytes = await pdfDoc.save();

    const outName = S(req.query?.out || `CTRL_180_${P.name || "Perspective"}_${P.dateLbl || ""}.pdf`)
      .replace(/[^\w.-]+/g, "_");

    // Response headers
    res.setHeader("X-CTRL-TPL", tpl);
    res.setHeader("X-CTRL-TPL-FALLBACK", usingFallback ? "1" : "0");
    res.setHeader("X-CTRL-PAYLOAD-SIZE", String(payloadSize || 0));
    res.setHeader("X-CTRL-CHART", chartUrl ? "1" : "0");
    res.setHeader("X-CTRL-CHART-TARGET", "p4");
    res.setHeader("X-CTRL-CHART-FETCH", chartFetch?.ok ? "ok" : (chartFetch?.reason || "no"));
    res.setHeader("X-CTRL-WORKWITH", (isObj(P.workWithBuilt) ? "1" : "0"));
    res.setHeader("X-CTRL-DEBUG", "0");

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${outName}"`);
    res.end(Buffer.from(bytes));
  } catch (err) {
    console.error("fill-template-180 error", err);

    try {
      res.setHeader("X-CTRL-TPL", tpl || "");
      res.setHeader("X-CTRL-TPL-FALLBACK", usingFallback ? "1" : "0");
      res.setHeader("X-CTRL-PAYLOAD-SIZE", String(payloadSize || 0));
      res.setHeader("X-CTRL-CHART-FETCH", chartFetch?.reason || "error");
      res.setHeader("X-CTRL-DEBUG", "0");
    } catch {}

    res.status(400).json({
      ok: false,
      error: `fill-template-180 error: ${err?.message || String(err)}`
    });
  }
}
