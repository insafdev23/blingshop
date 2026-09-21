import { useState, useEffect, useRef, useMemo } from "react";
import { read as xlsxRead, utils as xlsxUtils, writeFile as xlsxWriteFile } from "xlsx";
import JsBarcode from "jsbarcode";

// ── Brand Colors ──────────────────────────────────────────────────
const GOLD = "#C9952A";
const GOLD_LIGHT = "#F5E6C8";
const GOLD_DARK = "#8B6510";
const RC_BLUE = "#9F1239"; // RC Boutique brand accent — deep red
const RC_LIGHT = "#FCE4E8";
const RC_DARK = "#7F1237";
const CREAM = "#FFFBF2";
const WHITE = "#FFFFFF";
const GRAY = "#6B7280";
const LIGHT = "#FAFAFA";
const BORDER = "#E5E7EB";

const BUSINESSES = ["Blingshop", "RC Boutique"];
const PAYMENT_METHODS = ["Cash", "Card", "Bank Transfer", "Split"];
const TABS = ["Dashboard", "Inventory", "POS", "Sales History", "Reports", "Customers", "Settings"];
const SALESPERSON_TABS = ["Dashboard", "Inventory", "POS", "Sales History", "Customers"];
const CATEGORIES = ["All", "Clothing", "Jewelry", "Necklace", "Bracelet", "Anklet", "Chain", "Earring", "Scarf", "Ring", "Tops", "Dress", "Skirt", "Pant", "Belts", "Shirt", "Blouse", "Accessories", "Bags", "Shoes", "Shawls", "Other"];
// Each business only stocks its own kind of item, so its category filter is narrowed to just
// these — switching business drops any other active category filter back to "All" (see
// categoriesForBiz). "All" still shows the full master list above, unscoped.
const BLINGSHOP_CATEGORIES = ["Necklace", "Bracelet", "Anklet", "Chain", "Earring", "Scarf", "Ring", "Other"];
const RC_CATEGORIES = ["Tops", "Dress", "Skirt", "Pant", "Bags", "Shawls", "Other", "Belts", "Shirt", "Blouse"];
const categoriesForBiz = biz => {
  if (biz === "Blingshop") return ["All", ...BLINGSHOP_CATEGORIES];
  if (biz === "RC Boutique") return ["All", ...RC_CATEGORIES];
  return CATEGORIES;
};
const DELIVERY_METHODS = ["In Store", "Courier", "Flash Delivery"];
// Who the customer actually hands the delivery fee to: "Shop" means we collect it at checkout
// and owe it to the rider afterward (adds to the sale total); "Rider" means the customer pays
// the rider directly, so it never touches our cash and isn't added to the sale total — we just
// keep a record of the fee for reporting.
const DELIVERY_PAID_TO = [
  { id: "Shop", label: "Customer pays me (I pay the rider)" },
  { id: "Rider", label: "Customer pays the rider directly" },
];

const BIZ_STYLE = {
  Blingshop: { primary: GOLD_DARK, light: GOLD_LIGHT, badge: "#FEF3C7", badgeText: GOLD_DARK },
  "RC Boutique": { primary: RC_DARK, light: RC_LIGHT, badge: RC_LIGHT, badgeText: RC_DARK },
};

// ── API ───────────────────────────────────────────────────────────
const isLocal =
  window.location.hostname === "localhost" ||
  window.location.hostname === "127.0.0.1" ||
  /^192\.168\./.test(window.location.hostname) ||
  /^10\./.test(window.location.hostname);

const API =
  import.meta.env.VITE_API_URL ||
  (isLocal ? `https://${window.location.hostname}:3001` : "/api");

async function apiFetch(path, options = {}) {
  const res = await fetch(`${API}${path}`, {
    headers: { "Content-Type": "application/json" },
    ...options,
    body: options.body ? JSON.stringify(options.body) : undefined,
  });
  return res.json();
}

function generateId() { return "BS" + Date.now().toString().slice(-6) + Math.floor(Math.random() * 90 + 10); }

// Short synthesized beep (no audio asset needed) — confirms a barcode scan landed on a real
// product without needing to glance at the screen, useful mid-checkout with a hand-scanner.
function playBeep() {
  try {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    const osc = ctx.createOscillator(), gain = ctx.createGain();
    osc.connect(gain); gain.connect(ctx.destination);
    osc.type = "sine"; osc.frequency.value = 880; gain.gain.value = 0.2;
    osc.start(); osc.stop(ctx.currentTime + 0.12);
    osc.onended = () => ctx.close();
  } catch { }
}

// ── SKU Generator ─────────────────────────────────────────────────
const CATEGORY_CODES = {
  Clothing: "C", Jewelry: "J", Accessories: "A",
  Bags: "B", Shoes: "S", Shawls: "W", Other: "O",
  Necklace: "N", Bracelet: "L", Anklet: "K", Chain: "H", Earring: "E", Scarf: "F", Ring: "R",
  Tops: "T", Dress: "D", Skirt: "I", Pant: "P", Belts: "Y", Shirt: "U", Blouse: "M",
};

// Short-form SKU: 1-char business + 1-char category + 2-digit sequence, no separators
// (e.g. "BJ01") — kept as compact as possible since every character adds ~11 bars to the
// printed barcode. The sequence number self-extends past 99 rather than colliding.
function generateSKU(category, business, existingProducts) {
  const catCode = CATEGORY_CODES[category] || "O";
  const bizCode = business === "RC Boutique" ? "R" : "B";
  const prefix = `${bizCode}${catCode}`;
  const same = existingProducts.filter(p => p.sku && p.sku.startsWith(prefix) && /^\d+$/.test(p.sku.slice(prefix.length)));
  const nums = same.map(p => parseInt(p.sku.slice(prefix.length))).filter(n => !isNaN(n));
  const next = nums.length > 0 ? Math.max(...nums) + 1 : 1;
  return `${prefix}${String(next).padStart(2, "0")}`;
}

// ── Quick Source (Bulk Add) pricing ────────────────────────────────
// Both businesses' THB rates fluctuate, so each is entered once per restock and reused
// (see blingshop_thb_rate / rc_thb_rate settings) rather than being hardcoded. Blingshop's
// starting price is final — cost doubled (100% markup) plus a flat packing fee, no review
// needed. RC Boutique's starting price is break-even only (cost + packing + cargo), with the
// profit margin decided later by editing the price during review.
const BLINGSHOP_PACKING_LKR = 200;
const RC_PACKING_LKR = 210;
const RC_CARGO_LKR = 1000;
const PACKING_COST = { "Blingshop": BLINGSHOP_PACKING_LKR, "RC Boutique": RC_PACKING_LKR + RC_CARGO_LKR };

const quickAddCostLkr = (costThb, rate) => Math.round((+costThb || 0) * (+rate || 0));
const quickAddPrice = (business, costThb, rate) => {
  const costLkr = quickAddCostLkr(costThb, rate);
  return business === "RC Boutique" ? costLkr + RC_PACKING_LKR + RC_CARGO_LKR : costLkr * 2 + BLINGSHOP_PACKING_LKR;
};

// ── Code 128 Barcode ──────────────────────────────────────────────
// Renders via the jsbarcode library rather than a hand-rolled encoder — a wrong bar pattern
// or checksum is invisible on screen but makes the barcode completely unscannable in real life,
// so this is delegated to a well-tested implementation instead of maintained by hand.
// Code128B always costs 11 modules per symbol (start + one per character + checksum + stop) plus
// a 2-module tail, so this formula gives the true bar count without actually running the encoder —
// used purely for sizing/layout math.
function code128BarCount(text) {
  return (String(text).length + 3) * 11 + 2;
}

function renderBarcodeCanvas(text, moduleWidthPx, heightPx) {
  const canvas = document.createElement("canvas");
  const value = String(text || "").trim();
  if (!value) return canvas;
  try {
    JsBarcode(canvas, value, {
      format: "CODE128B",
      width: Math.max(1, moduleWidthPx),
      height: Math.max(1, heightPx),
      margin: 0,
      displayValue: false,
    });
  } catch {
    // Characters outside Code128B's range — leave the canvas blank rather than crashing the label preview.
  }
  return canvas;
}

function BarcodeCanvas({ value, width = 220, height = 56, price }) {
  const ref = useRef();
  useEffect(() => {
    if (!ref.current || !value) return;
    const canvas = ref.current;
    const DPR = Math.max(2, window.devicePixelRatio || 1);
    const hasPrice = price !== undefined && !isNaN(Number(price)) && Number(price) > 0;
    const W = width, H = height + 22 + (hasPrice ? 20 : 0);
    canvas.width = Math.round(W * DPR);
    canvas.height = Math.round(H * DPR);
    canvas.style.width = W + "px";
    canvas.style.height = H + "px";
    const ctx = canvas.getContext("2d");
    ctx.scale(DPR, DPR);
    ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, W, H);
    const barCount = code128BarCount(String(value));
    const moduleWidthPx = (W - 16) / barCount;
    const bcCanvas = renderBarcodeCanvas(value, moduleWidthPx, height);
    if (bcCanvas.width > 0) ctx.drawImage(bcCanvas, 8, 4, W - 16, height);
    ctx.fillStyle = "#000"; ctx.font = "11px monospace"; ctx.textAlign = "center";
    ctx.fillText(String(value), W / 2, height + 18);
    if (hasPrice) { ctx.font = "bold 14px Arial, sans-serif"; ctx.fillText(`LKR ${Number(price).toLocaleString()}`, W / 2, height + 38); }
  }, [value, width, height, price]);
  return <canvas ref={ref} style={{ display: "block", maxWidth: "100%" }} />;
}

// ── Label Print System ────────────────────────────────────────────
const LABEL_PRESETS = {
  standard: [
    { id: "50x30", label: "50×30mm — Standard", w: 50, h: 30 },
    { id: "50x25", label: "50×25mm", w: 50, h: 25 },
    { id: "40x20", label: "40×20mm — Small", w: 40, h: 20 },
    { id: "38x25", label: "38×25mm — Compact", w: 38, h: 25 },
    { id: "70x40", label: "70×40mm — Large", w: 70, h: 40 },
  ],
  jewelry: [
    { id: "70x12", label: "70×12mm Dumbbell", w: 70, h: 12 },
    { id: "60x10", label: "60×10mm Dumbbell", w: 60, h: 10 },
    { id: "80x15", label: "80×15mm Dumbbell", w: 80, h: 15 },
  ],
};
const DOTS_PER_MM = 8;          // Zebra 203 dpi ≈ 8 dots/mm
const PRINT_PX_PER_MM = 300 / 25.4; // canvas rendered at 300 dpi

// ── Label Layout Engine ──────────────────────────────────────────
// Shared by canvas rendering (Browser Print / preview) and ZPL generation (Network / ZPL Code),
// so both outputs match and stay within physical bounds instead of clipping or over-densifying.
const BARCODE_FILL_RATIO = { 1: 0.35, 2: 0.50, 3: 0.65 }; // Narrow/Normal/Wide — fraction of label width the barcode aims to fill
const BARCODE_FLOOR_MM = { 1: 0.25, 2: 0.33, 3: 0.42 }; // readable-minimum module width per preset, for small labels
const BARCODE_MAX_MODULE_MM = 0.5; // keeps bars from ballooning into an overly wide/stretched barcode on large labels
const BARCODE_MIN_MODULE_MM = 0.2; // hard floor — below this a scanner realistically can't read it

// Given how many bars a barcode needs and the label's full physical width, size the module
// (bar) width proportionally to the label so bigger stickers get a bigger barcode instead of
// staying pinned at a fixed mm size — but never let it shrink below the preset's readable floor,
// and never let it grow past a sensible max width so it doesn't stretch disproportionately wide.
// Only force-shrinks below the floor (with a warning) when the barcode would truly overflow
// the label even at that floor.
function computeBarcodeModule(barsLength, wMmFull, scale) {
  const fillRatio = BARCODE_FILL_RATIO[scale] || BARCODE_FILL_RATIO[2];
  const floorMm = BARCODE_FLOOR_MM[scale] || BARCODE_FLOOR_MM[2];
  // Code128 needs a blank "quiet zone" on each side (~10x the bar width) or scanners can't
  // find the barcode's start/stop edges — this must come out of the label width, not be
  // squeezed away, or the barcode looks fine on screen but fails to scan in real life.
  const quietZoneMm = Math.max(2, floorMm * 8);
  const ceilingMm = Math.max(1, wMmFull - quietZoneMm * 2);
  const idealMm = (ceilingMm * fillRatio) / barsLength;
  let moduleMm = Math.min(Math.max(idealMm, floorMm), BARCODE_MAX_MODULE_MM);
  let totalMm = barsLength * moduleMm;
  let warning = null;
  if (totalMm > ceilingMm) {
    const fitMm = ceilingMm / barsLength;
    if (fitMm < BARCODE_MIN_MODULE_MM) {
      warning = "Barcode is too dense to scan reliably at this label size. Use a larger label or a shorter SKU.";
      moduleMm = BARCODE_MIN_MODULE_MM;
    } else {
      warning = `Barcode had to shrink to ${fitMm.toFixed(2)}mm bars to fit this label — try Narrow or a larger label for best results.`;
      moduleMm = fitMm;
    }
    totalMm = barsLength * moduleMm;
  }
  return { moduleMm, totalMm, warning };
}

// Computes an mm-based stacking layout for the standard rectangular label: which fields are
// active, their font sizes (shrunk toward a readable floor if needed, never clipped), and the
// barcode's real module width. Consumed by both drawStandardLabel (canvas) and generateZPLStandard.
function computeStandardLayout(product, wMm, hMm, opts) {
  const { showBiz = false, showName = true, showBarcode = true, showSKU = true, showPrice = true, barcodeScale = 2 } = opts || {};
  const warnings = [];
  const marginMm = Math.max(1, wMm * 0.04);
  const gapMm = Math.max(0.4, hMm * 0.03);
  const availW = wMm - marginMm * 2;
  const availH = hMm - marginMm * 2;

  const items = [];
  if (showBiz) items.push({ type: "text", key: "biz", pref: 2.4, min: 1.6 });
  if (showName) items.push({ type: "text", key: "name", pref: 2.6, min: 1.7 });
  if (showBarcode) {
    const skuText = String(product.sku || "");
    const { moduleMm, totalMm, warning } = computeBarcodeModule(code128BarCount(skuText), wMm, barcodeScale);
    if (warning) warnings.push(warning);
    items.push({ type: "barcode", skuText, moduleMm, wMm: Math.min(totalMm, wMm), hMm: Math.max(hMm * 0.26, 3.5) });
  }
  if (showSKU) items.push({ type: "text", key: "sku", pref: 2.8, min: 1.9 });
  if (showPrice) items.push({ type: "text", key: "price", pref: 3.2, min: 2.0 });

  const gapsTotal = Math.max(0, items.length - 1) * gapMm;
  const fixedH = items.filter(i => i.type === "barcode").reduce((s, i) => s + i.hMm, 0);
  const textItems = items.filter(i => i.type === "text");

  const prefTotal = textItems.reduce((s, i) => s + i.pref, 0) + fixedH + gapsTotal;
  let useFloor = false, scale = 1;
  if (prefTotal > availH) {
    const minTotal = textItems.reduce((s, i) => s + i.min, 0) + fixedH + gapsTotal;
    if (minTotal > availH) {
      warnings.push(`Selected fields don't comfortably fit on a ${wMm}×${hMm}mm label — text shrunk to the minimum readable size. Try a larger label or fewer fields.`);
      useFloor = true;
    } else {
      const flexRoom = textItems.reduce((s, i) => s + (i.pref - i.min), 0);
      const available = availH - minTotal;
      scale = flexRoom > 0 ? Math.max(0, Math.min(1, available / flexRoom)) : 1;
    }
  }
  // textScale is a manual multiplier applied after the auto shrink-to-fit sizing above, so
  // "Normal" (1x) reproduces the exact fitted size as before, while Small/Large scale that
  // fitted result up or down rather than re-running the fit against scaled targets (which would
  // just claw Large back down to near-minimum on tight labels, defeating the point of the control).
  const textScale = opts?.textScale || 1;
  for (const i of textItems) i.fontMm = (useFloor ? i.min : i.min + (i.pref - i.min) * scale) * textScale;

  // Vertical alignment of the whole content stack within the label — when the active fields
  // don't fill the height, "top" (default browser behaviour before this) left everything
  // pinned to the top edge with the barcode looking stranded; "center"/"bottom" distribute
  // the leftover space instead.
  const contentH = items.reduce((s, i) => s + (i.hMm ?? i.fontMm), 0) + gapsTotal;
  const vAlign = opts?.vAlign || "center";
  const slack = Math.max(0, availH - contentH);
  const startYMm = marginMm + (vAlign === "top" ? 0 : vAlign === "bottom" ? slack : slack / 2);

  return { marginMm, gapMm, availW, availH, items, warnings, startYMm };
}

// Lighter-weight equivalent for the dumbbell jewelry label — barcode module width and text
// floors, so the front-panel barcode stays scannable and back-panel price never clips.
function computeJewelryLayout(product, wMm, hMm, opts) {
  const { showBarcode = true, barcodeScale = 2, textScale = 1 } = opts || {};
  const warnings = [];
  const FRONT_W_MM = wMm * 0.42;
  const BACK_W_MM = wMm * 0.42;
  let barcode = null;
  if (showBarcode) {
    const skuText = String(product.sku || "");
    const { moduleMm, totalMm, warning } = computeBarcodeModule(code128BarCount(skuText), FRONT_W_MM, barcodeScale);
    if (warning) warnings.push(warning);
    barcode = { skuText, moduleMm, totalMm: Math.min(totalMm, FRONT_W_MM) };
  }
  const skuFontMm = Math.max(hMm * 0.42, 1.7) * textScale;
  return {
    barcode,
    skuFontMm,
    nameFontMm: Math.max(hMm * 0.36, 1.7) * textScale,
    priceFontMm: skuFontMm, // matches SKU size, per request — was Math.max(hMm * 0.50, 2.2)
    bizFontMm: Math.max(hMm * 0.32 * 0.70, 1.5) * textScale,
    FRONT_W_MM, BACK_W_MM,
    warnings,
  };
}

function drawStandardLabel(ctx, product, W, H, wMm, hMm, opts) {
  const pxPerMm = W / wMm;
  const bizColor = (product.business || "Blingshop") === "RC Boutique" ? RC_DARK : GOLD_DARK;
  const layout = computeStandardLayout(product, wMm, hMm, opts);

  ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, W, H);

  let yMm = layout.startYMm;
  const cxPx = W / 2;
  layout.items.forEach((item, idx) => {
    if (item.type === "barcode") {
      const totalWpx = item.wMm * pxPerMm;
      const startXpx = (wMm / 2 - item.wMm / 2) * pxPerMm;
      const barcodeHpx = item.hMm * pxPerMm;
      const ypx = yMm * pxPerMm;
      const moduleWidthPx = item.moduleMm * pxPerMm;
      const bcCanvas = renderBarcodeCanvas(item.skuText, moduleWidthPx, barcodeHpx);
      if (bcCanvas.width > 0) ctx.drawImage(bcCanvas, startXpx, ypx, totalWpx, barcodeHpx);
    } else {
      const fontPx = item.fontMm * pxPerMm;
      const text = item.key === "biz" ? (product.business || "Blingshop").toUpperCase()
        : item.key === "name" ? ((product.name || "").length > 28 ? product.name.substring(0, 27) + "…" : (product.name || ""))
          : item.key === "sku" ? String(product.sku || "")
            : `LKR ${Number(product.price || 0).toLocaleString()}`;
      ctx.fillStyle = item.key === "biz" ? bizColor : "#111";
      ctx.font = item.key === "sku" ? `${fontPx}px monospace` : `bold ${fontPx}px Arial,sans-serif`;
      ctx.textAlign = "center"; ctx.textBaseline = "alphabetic";
      const baselinePx = (yMm + item.fontMm * 0.82) * pxPerMm;
      ctx.fillText(text, cxPx, baselinePx);
      if (item.key === "biz") {
        ctx.fillStyle = bizColor + "55";
        ctx.fillRect(layout.marginMm * pxPerMm, (yMm + item.fontMm + 0.3) * pxPerMm, layout.availW * pxPerMm, 1.5);
      }
    }
    yMm += item.hMm ?? item.fontMm;
    if (idx < layout.items.length - 1) yMm += layout.gapMm;
  });
}

function drawJewelryLabel(ctx, product, W, H, wMm, hMm, opts) {
  const { showBiz = false, showName = true, showBarcode = true, showSKU = true, showPrice = true } = opts || {};
  const pxPerMm = W / wMm;
  const bizColor = (product.business || "Blingshop") === "RC Boutique" ? RC_DARK : GOLD_DARK;
  const FRONT_W = W * 0.42, BACK_X = W * 0.58, BACK_W = W * 0.42;
  const layout = computeJewelryLayout(product, wMm, hMm, opts);

  ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, W, H);

  // Front panel — barcode + optional SKU (module width enforced with a readable minimum)
  if (showBarcode && layout.barcode) {
    const bH = showSKU ? H * 0.58 : H * 0.88;
    const totalBWpx = layout.barcode.totalMm * pxPerMm;
    const startX = (FRONT_W - totalBWpx) / 2;
    const moduleWidthPx = layout.barcode.moduleMm * pxPerMm;
    const bcCanvas = renderBarcodeCanvas(layout.barcode.skuText, moduleWidthPx, bH);
    if (bcCanvas.width > 0) ctx.drawImage(bcCanvas, startX, 1.5, totalBWpx, bH);
    if (showSKU) {
      const sFsPx = layout.skuFontMm * pxPerMm;
      ctx.fillStyle = "#111"; ctx.font = `${sFsPx}px monospace`;
      ctx.textAlign = "center"; ctx.textBaseline = "alphabetic";
      ctx.fillText(String(product.sku || ""), FRONT_W / 2, H * 0.9);
    }
  }

  // Fold zone intentionally left blank — this canvas is screenshotted verbatim for the
  // Browser/USB print path, so any shading/dashes drawn here print for real on the physical
  // tag. The Front/FOLD/Back caption below the on-screen preview already explains the zone.

  // Back panel
  let by = 2;
  if (showBiz) {
    const banH = H * 0.30;
    ctx.fillStyle = bizColor; ctx.fillRect(BACK_X, 0, BACK_W, banH);
    ctx.fillStyle = "#fff"; ctx.font = `bold ${layout.bizFontMm * pxPerMm}px Arial`;
    ctx.textAlign = "center"; ctx.textBaseline = "middle";
    ctx.fillText((product.business || "Blingshop").toUpperCase(), BACK_X + BACK_W / 2, banH / 2);
    by = banH + 2;
  }
  if (showName) {
    const nFsPx = layout.nameFontMm * pxPerMm;
    const nMax = Math.max(4, Math.floor(BACK_W / (nFsPx * 0.6)));
    const nTxt = (product.name || "").length > nMax ? product.name.substring(0, nMax - 1) + "…" : (product.name || "");
    ctx.fillStyle = "#111"; ctx.font = `${nFsPx}px Arial`;
    ctx.textAlign = "center"; ctx.textBaseline = "alphabetic";
    ctx.fillText(nTxt, BACK_X + BACK_W / 2, by + nFsPx * 1.2);
    by += nFsPx * 1.6;
  }
  if (showPrice) {
    const pFsPx = layout.priceFontMm * pxPerMm;
    ctx.fillStyle = "#111"; ctx.font = `bold ${pFsPx}px Arial`;
    ctx.textAlign = "center"; ctx.textBaseline = "alphabetic";
    // When price is the only thing on the back panel (biz banner and name both off), center it
    // in the full label height instead of pinning it to the bottom — otherwise it sits stranded
    // low with a lot of empty space above it. With biz/name shown above, keep the old
    // bottom-anchored behavior (clamped so descenders never fall off the edge) so price fills
    // the remaining space below them.
    const onlyPrice = !showBiz && !showName;
    const baselineY = onlyPrice ? H / 2 + pFsPx * 0.35 : Math.min(H - pFsPx * 0.22, H * 0.92);
    ctx.fillText(`LKR ${Number(product.price || 0).toLocaleString()}`, BACK_X + BACK_W / 2, baselineY);
  }
}

function drawLabelToCanvas(canvas, product, labelType, wMm, hMm, opts) {
  if (!canvas || !product) return;
  const dpi = opts?.printDPI || 300;
  const pxPerMm = dpi / 25.4;
  const W = Math.round(wMm * pxPerMm);
  const H = Math.round(hMm * pxPerMm);
  canvas.width = W;
  canvas.height = H;
  const maxDispW = 370;
  const dispW = Math.min(maxDispW, W);
  const dispH = Math.round(dispW * H / W);
  canvas.style.width = dispW + "px";
  canvas.style.height = dispH + "px";
  const ctx = canvas.getContext("2d");
  if (labelType === "jewelry") drawJewelryLabel(ctx, product, W, H, wMm, hMm, opts);
  else drawStandardLabel(ctx, product, W, H, wMm, hMm, opts);
}

function generateZPLStandard(product, wMm, hMm, qty = 1, opts) {
  const layout = computeStandardLayout(product, wMm, hMm, opts);
  const W = Math.round(wMm * DOTS_PER_MM), H = Math.round(hMm * DOTS_PER_MM);
  const mm2dots = mm => Math.round(mm * DOTS_PER_MM);
  const name = (product.name || "").substring(0, 30).toUpperCase();
  const biz = (product.business || "Blingshop").toUpperCase();
  const price = `LKR ${Number(product.price || 0).toLocaleString()}`;
  const sku = String(product.sku || "");
  const lines = ["^XA", `^PW${W}`, `^LL${H}`, "^CI28"];

  let yMm = layout.startYMm;
  layout.items.forEach((item, idx) => {
    if (item.type === "barcode") {
      const bcHdots = mm2dots(item.hMm);
      const startXdots = mm2dots(wMm / 2 - item.wMm / 2);
      const moduleDots = Math.max(1, Math.round(item.moduleMm * DOTS_PER_MM));
      lines.push(`^FO${startXdots},${mm2dots(yMm)}^BY${moduleDots},3,${bcHdots}^BCN,${bcHdots},N,N^FD${sku}^FS`);
    } else {
      const fsDots = mm2dots(item.fontMm);
      const text = item.key === "biz" ? biz : item.key === "name" ? name : item.key === "sku" ? sku : price;
      lines.push(`^FO${mm2dots(layout.marginMm)},${mm2dots(yMm)}^A0N,${fsDots},${fsDots}^FD${text}^FS`);
      if (item.key === "biz") lines.push(`^FO${mm2dots(layout.marginMm)},${mm2dots(yMm + item.fontMm + 0.3)}^GB${mm2dots(layout.availW)},1,1^FS`);
    }
    yMm += item.hMm ?? item.fontMm;
    if (idx < layout.items.length - 1) yMm += layout.gapMm;
  });

  lines.push(`^PQ${qty},0,1,Y`, "^XZ");
  return lines.join("\n");
}

function generateZPLJewelry(product, wMm, hMm, qty = 1, opts) {
  const { showBiz = false, showName = true, showBarcode = true, showSKU = true, showPrice = true } = opts || {};
  const layout = computeJewelryLayout(product, wMm, hMm, opts);
  const W = Math.round(wMm * DOTS_PER_MM), H = Math.round(hMm * DOTS_PER_MM);
  const mm2dots = mm => Math.round(mm * DOTS_PER_MM);
  const name = (product.name || "").substring(0, 14);
  const biz = (product.business || "Blingshop").toUpperCase();
  const price = `LKR ${Number(product.price || 0).toLocaleString()}`;
  const sku = String(product.sku || "");
  const BACK_X = Math.round(W * 0.58);
  const sFs = mm2dots(layout.skuFontMm);
  const nFs = mm2dots(layout.nameFontMm);
  const pFs = mm2dots(layout.priceFontMm);
  const bcH = Math.round(H * (showSKU ? 0.58 : 0.88));
  const lines = ["^XA", `^PW${W}`, `^LL${H}`, "^CI28"];
  if (showBarcode && layout.barcode) {
    const moduleDots = Math.max(1, Math.round(layout.barcode.moduleMm * DOTS_PER_MM));
    lines.push(`^FO2,2^BY${moduleDots},2.5,${bcH}^BCN,${bcH},N,N^FD${sku}^FS`);
    if (showSKU) lines.push(`^FO2,${bcH + 2}^A0N,${sFs},${sFs}^FD${sku}^FS`);
  }
  let by = 0;
  if (showBiz) {
    const banH = Math.round(H * 0.32);
    lines.push(`^FO${BACK_X},0^GB${W - BACK_X},${banH},${banH}^FS`);
    lines.push(`^FO${BACK_X + 2},${Math.round(banH * 0.05)}^A0N,${Math.round(banH * 0.70)},${Math.round(banH * 0.70)}^FR^FD${biz}^FS`);
    by = banH + 2;
  }
  if (showName) { lines.push(`^FO${BACK_X + 2},${by}^A0N,${nFs},${nFs}^FD${name}^FS`); by += nFs + 2; }
  if (showPrice) {
    // Center vertically when price is the only field on the back panel; otherwise anchor near
    // the bottom edge so it fills the space left below biz/name — mirrors the canvas renderer.
    const onlyPrice = !showBiz && !showName;
    const yTop = onlyPrice ? Math.max(2, Math.round(H / 2 - pFs / 2)) : Math.max(2, H - pFs - 2);
    lines.push(`^FO${BACK_X + 2},${yTop}^A0N,${pFs},${pFs}^FD${price}^FS`);
  }
  lines.push(`^PQ${qty},0,1,Y`, "^XZ");
  return lines.join("\n");
}

const RETURN_POLICY = {
  "Blingshop": "No exchange. No refund.",
  "RC Boutique": "Exchange possible within 3 days of purchase. No refund.",
};
const wrapText = (s, w) => {
  const out = []; let cur = "";
  for (const word of s.split(" ")) {
    const next = cur ? `${cur} ${word}` : word;
    if (next.length > w) { if (cur) out.push(cur); cur = word; } else cur = next;
  }
  if (cur) out.push(cur);
  return out;
};
const paymentSummary = sale => {
  if (sale.paymentMethod === "Split") return `Split (Card: LKR ${(+sale.cardAmount || 0).toLocaleString()}, Cash: LKR ${(+sale.cashAmount || 0).toLocaleString()})`;
  return sale.paymentMethod || "Cash";
};

function buildReceiptText(sale, headerBiz) {
  const W = 48; // 76mm paper at default ESC/POS font (Font A) fits ~48 chars/line
  const line = "-".repeat(W);
  const center = s => s.length >= W ? s.slice(0, W) : s.padStart(Math.floor((W + s.length) / 2)).padEnd(W);
  const row = (l, r) => l.padEnd(W - r.length) + r;
  const biz = headerBiz || sale.items[0]?.business || "Blingshop";
  const lines = [center(biz.toUpperCase())];
  lines.push(...wrapText(RETURN_POLICY[biz] || "", W).map(center));
  lines.push(line, `Date: ${new Date(sale.date).toLocaleString()}`, `Receipt: #${sale.id}`);
  if (sale.staffName) lines.push(`Served by: ${sale.staffName}`);
  if (sale.customerName) lines.push(`Customer: ${sale.customerName}`);
  lines.push(line);
  for (const i of sale.items) lines.push(row(`  ${(i.sku || i.name).substring(0, W - 14)} x${i.qty}`, `${(i.price * i.qty).toLocaleString()}`));
  lines.push(line);
  if (sale.discount > 0) { lines.push(row("Subtotal:", `LKR ${sale.subtotal.toLocaleString()}`)); lines.push(row("Discount:", `-LKR ${sale.discount.toLocaleString()}`)); }
  if (sale.deliveryMethod && sale.deliveryMethod !== "In Store") lines.push(row(`Delivery (${sale.deliveryMethod}):`, sale.deliveryPaidTo === "Shop" ? `LKR ${(+sale.deliveryFee || 0).toLocaleString()}` : "Paid to rider"));
  lines.push(row("TOTAL:", `LKR ${sale.total.toLocaleString()}`), line);
  lines.push(`Payment: ${paymentSummary(sale)}`, line);
  lines.push(center("Thank you for shopping!"), center("Come again soon."), "", "");
  return lines.join("\n");
}

async function printReceiptBluetooth(sale, headerBiz, onStatus) {
  if (!navigator.bluetooth) { onStatus("error:Web Bluetooth not supported. Use Chrome on Android or desktop."); return; }
  try {
    onStatus("info:Scanning for ABM P323B printer...");
    const device = await navigator.bluetooth.requestDevice({ filters: [{ services: ["000018f0-0000-1000-8000-00805f9b34fb"] }], optionalServices: ["000018f0-0000-1000-8000-00805f9b34fb", "e7810a71-73ae-499d-8c15-faa9aef0c3f2"] });
    onStatus("info:Connecting...");
    const server = await device.gatt.connect();
    let characteristic = null;
    for (const svc of ["000018f0-0000-1000-8000-00805f9b34fb", "e7810a71-73ae-499d-8c15-faa9aef0c3f2"]) {
      try { const service = await server.getPrimaryService(svc); for (const ch of ["00002af1-0000-1000-8000-00805f9b34fb", "bef8d6c9-9c21-4c9e-b632-bd58c1009f9f"]) { try { characteristic = await service.getCharacteristic(ch); break; } catch { } } if (characteristic) break; } catch { }
    }
    if (!characteristic) { onStatus("error:Could not find print characteristic. Try USB instead."); return; }
    // Let the connection settle before writing — writing immediately after connect is a common cause of "GATT operation failed" on these printers.
    await new Promise(r => setTimeout(r, 300));
    onStatus("info:Printing...");
    const encoder = new TextEncoder(), text = buildReceiptText(sale, headerBiz);
    const init = new Uint8Array([0x1B, 0x40]), cut = new Uint8Array([0x1D, 0x56, 0x41, 0x00]);
    const textBytes = encoder.encode(text);
    const full = new Uint8Array(init.length + textBytes.length + cut.length);
    full.set(init, 0); full.set(textBytes, init.length); full.set(cut, init.length + textBytes.length);
    // BLE's default negotiated MTU only allows ~20 bytes per write-without-response call — larger chunks (the old 512) are silently rejected as "GATT operation failed" on many printers.
    const CHUNK_SIZE = 20;
    const useWriteWithoutResponse = characteristic.properties?.writeWithoutResponse !== false;
    for (let i = 0; i < full.length; i += CHUNK_SIZE) {
      const chunk = full.slice(i, i + CHUNK_SIZE);
      for (let attempt = 1; ; attempt++) {
        try {
          if (useWriteWithoutResponse) await characteristic.writeValueWithoutResponse(chunk);
          else await characteristic.writeValue(chunk);
          break;
        } catch (writeErr) {
          if (attempt >= 3) throw writeErr;
          await new Promise(r => setTimeout(r, 200));
        }
      }
      await new Promise(r => setTimeout(r, 15));
    }
    onStatus("success:Receipt printed successfully.");
  } catch (err) {
    if (err.name === "NotFoundError") onStatus("info:No printer selected.");
    else if (/GATT operation failed/i.test(err.message)) onStatus("error:Connection dropped mid-print. Keep the printer close and powered on, then try again.");
    else onStatus(`error:${err.message}`);
  }
}

async function printZebraZPL(zpl, zebraIP, onStatus) {
  if (zebraIP) {
    try {
      onStatus("info:Sending to Zebra printer...");
      const res = await fetch(`${API}/print-zpl?ip=${zebraIP}`, { method: "POST", body: zpl });
      const data = await res.json();
      if (data.success) onStatus("success:Sent to Zebra ZD421T successfully."); else onStatus(`error:${data.error}`);
    } catch { onStatus("error:Could not reach proxy server. Make sure server.js is running."); }
  } else { onStatus("info:No IP configured. Go to Settings to add the Zebra printer IP."); }
}

// ── Excel Import ──────────────────────────────────────────────────
const FIELD_VARIANTS = {
  name: ["product name", "name", "item name", "product", "item"],
  category: ["category", "type", "product category", "item category"],
  price: ["price", "price (lkr)", "price(lkr)", "unit price", "selling price", "amount"],
  cost: ["cost", "cost (lkr)", "cost(lkr)", "cost price", "unit cost", "buying price"],
  stock: ["stock", "stock quantity", "quantity", "stock qty", "qty", "qty in stock"],
  sku: ["sku", "sku code", "barcode", "code"],
  image: ["image", "image url", "photo", "photo url", "picture"],
  business: ["business", "store", "shop", "business name", "store name"],
};
const norm = h => String(h || "").trim().toLowerCase();

function mapExcelRow(rawRow) {
  const result = { name: "", category: "", price: "", cost: "", stock: "", sku: "", image: "", business: "" };
  const keys = Object.keys(rawRow);
  for (const field of Object.keys(FIELD_VARIANTS)) {
    const matchKey = keys.find(k => FIELD_VARIANTS[field].includes(norm(k)));
    if (matchKey !== undefined) result[field] = String(rawRow[matchKey] ?? "").trim();
  }
  return result;
}

function validateExcelRow(row, idx) {
  const errors = [];
  if (!row.name) errors.push("Missing name");
  const priceNum = parseFloat(row.price);
  if (row.price === "" || isNaN(priceNum) || priceNum <= 0) errors.push("Invalid price");
  const costNum = parseFloat(row.cost);
  if (row.cost !== "" && (isNaN(costNum) || costNum < 0)) errors.push("Invalid cost");
  const stockNum = parseInt(row.stock);
  if (row.stock === "" || isNaN(stockNum) || stockNum < 0) errors.push("Invalid stock");
  let category = CATEGORIES.slice(1).find(c => c.toLowerCase() === row.category.toLowerCase());
  if (!category) category = "Other";
  let business = BUSINESSES.find(b => b.toLowerCase() === row.business.toLowerCase());
  if (!business) business = "Blingshop";
  return { rowNum: idx + 2, name: row.name, category, price: priceNum, cost: isNaN(costNum) ? 0 : costNum, stock: stockNum, sku: row.sku, image: row.image, business, errors, valid: errors.length === 0 };
}

// ── SVG Icons ─────────────────────────────────────────────────────
const Icons = {
  box: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 003 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16z" /></svg>,
  chart: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="18" y1="20" x2="18" y2="10" /><line x1="12" y1="20" x2="12" y2="4" /><line x1="6" y1="20" x2="6" y2="14" /></svg>,
  cart: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="9" cy="21" r="1" /><circle cx="20" cy="21" r="1" /><path d="M1 1h4l2.68 13.39a2 2 0 002 1.61h9.72a2 2 0 002-1.61L23 6H6" /></svg>,
  edit: <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7" /><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z" /></svg>,
  trash: <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="3 6 5 6 21 6" /><path d="M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6" /><path d="M10 11v6M14 11v6" /><path d="M9 6V4a1 1 0 011-1h4a1 1 0 011 1v2" /></svg>,
  barcode: <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M3 5v14M7 5v14M11 5v14M15 5v14M19 5v14M21 5v14" /></svg>,
  print: <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="6 9 6 2 18 2 18 9" /><path d="M6 18H4a2 2 0 01-2-2v-5a2 2 0 012-2h16a2 2 0 012 2v5a2 2 0 01-2 2h-2" /><rect x="6" y="14" width="12" height="8" /></svg>,
  camera: <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M23 19a2 2 0 01-2 2H3a2 2 0 01-2-2V8a2 2 0 012-2h4l2-3h6l2 3h4a2 2 0 012 2z" /><circle cx="12" cy="13" r="4" /></svg>,
  close: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" /></svg>,
  plus: <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" /></svg>,
  search: <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" /></svg>,
  check: <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12" /></svg>,
  wifi: <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M5 12.55a11 11 0 0114.08 0" /><path d="M1.42 9a16 16 0 0121.16 0" /><path d="M8.53 16.11a6 6 0 016.95 0" /><circle cx="12" cy="20" r="1" fill="currentColor" /></svg>,
  copy: <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="9" y="9" width="13" height="13" rx="2" /><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1" /></svg>,
  alert: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" /><line x1="12" y1="9" x2="12" y2="13" /><line x1="12" y1="17" x2="12.01" y2="17" /></svg>,
  refresh: <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="23 4 23 10 17 10" /><path d="M20.49 15a9 9 0 11-2.12-9.36L23 10" /></svg>,
  user: <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2" /><circle cx="12" cy="7" r="4" /></svg>,
  history: <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="1 4 1 10 7 10" /><path d="M3.51 15a9 9 0 102.13-9.36L1 10" /></svg>,
  settings: <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="3" /><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-2 2 2 2 0 01-2-2v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 01-2-2 2 2 0 012-2h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 012-2 2 2 0 012 2v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 012 2 2 2 0 01-2 2h-.09a1.65 1.65 0 00-1.51 1z" /></svg>,
  dashboard: <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="3" width="7" height="7" /><rect x="14" y="3" width="7" height="7" /><rect x="3" y="14" width="7" height="7" /><rect x="14" y="14" width="7" height="7" /></svg>,
  image: <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5"><rect x="3" y="3" width="18" height="18" rx="2" /><circle cx="8.5" cy="8.5" r="1.5" /><polyline points="21 15 16 10 5 21" /></svg>,
  upload: <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" y1="3" x2="12" y2="15" /></svg>,
  download: <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" y1="15" x2="12" y2="3" /></svg>,
  reports: <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="16" y1="13" x2="8" y2="13" /><line x1="16" y1="17" x2="8" y2="17" /><polyline points="10 9 9 9 8 9" /></svg>,
  store: <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z" /><polyline points="9 22 9 12 15 12 15 22" /></svg>,
  users: <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2" /><circle cx="9" cy="7" r="4" /><path d="M23 21v-2a4 4 0 00-3-3.87" /><path d="M16 3.13a4 4 0 010 7.75" /></svg>,
  cake: <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M20 21v-8a2 2 0 00-2-2H6a2 2 0 00-2 2v8" /><path d="M4 16s.5-1 2-1 2.5 2 4 2 2.5-2 4-2 2 1 2 1" /><path d="M2 21h20" /><path d="M7 8v2" /><path d="M12 8v2" /><path d="M17 8v2" /><path d="M7 4h.01" /><path d="M12 4h.01" /><path d="M17 4h.01" /></svg>,
  phone: <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M22 16.92v3a2 2 0 01-2.18 2 19.79 19.79 0 01-8.63-3.07A19.5 19.5 0 013.07 9.81a19.79 19.79 0 01-3.07-8.67A2 2 0 012 1h3a2 2 0 012 1.72c.127.96.361 1.903.7 2.81a2 2 0 01-.45 2.11L6.09 8.91a16 16 0 006 6l1.27-1.27a2 2 0 012.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0122 16.92z" /></svg>,
  mail: <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z" /><polyline points="22,6 12,13 2,6" /></svg>,
};

// ── Business Badge ────────────────────────────────────────────────
function BizBadge({ business, size = "sm" }) {
  const s = BIZ_STYLE[business] || BIZ_STYLE["Blingshop"];
  const fs = size === "sm" ? 10 : 12;
  return (
    <span style={{ fontSize: fs, fontWeight: 700, padding: "2px 8px", borderRadius: 10, background: s.badge, color: s.primary, whiteSpace: "nowrap" }}>
      {business}
    </span>
  );
}

// ── Business Filter Toggle ────────────────────────────────────────
function BizFilter({ value, onChange, includeAll = true }) {
  const opts = includeAll ? ["All", ...BUSINESSES] : BUSINESSES;
  return (
    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
      {opts.map(b => {
        const active = value === b;
        return (
          <button key={b} onClick={() => onChange(b)} style={{
            padding: "6px 14px", borderRadius: 20, border: `1.5px solid ${active ? (b === "Blingshop" ? GOLD : b === "RC Boutique" ? RC_BLUE : "#374151") : BORDER}`,
            background: active ? (b === "Blingshop" ? GOLD_DARK : b === "RC Boutique" ? RC_DARK : "#374151") : WHITE,
            color: active ? WHITE : GRAY, cursor: "pointer", fontSize: 12, fontWeight: active ? 700 : 400,
          }}>{b}</button>
        );
      })}
    </div>
  );
}

// ── Camera Scanner ────────────────────────────────────────────────
function CameraScanner({ onDetect, onClose }) {
  const videoRef = useRef(); const [status, setStatus] = useState("Starting camera..."); const [manualSku, setManualSku] = useState(""); const streamRef = useRef();
  useEffect(() => {
    let interval;
    (async () => {
      try {
        const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } });
        streamRef.current = stream;
        if (videoRef.current) { videoRef.current.srcObject = stream; await videoRef.current.play(); }
        setStatus("Point camera at barcode, or enter SKU manually below.");
        if ("BarcodeDetector" in window) {
          const detector = new window.BarcodeDetector({ formats: ["code_128", "ean_13", "ean_8", "qr_code", "code_39", "upc_a"] });
          interval = setInterval(async () => { if (!videoRef.current) return; try { const codes = await detector.detect(videoRef.current); if (codes.length > 0) { clearInterval(interval); onDetect(codes[0].rawValue); } } catch { } }, 500);
        } else { setStatus("Camera active. Barcode auto-detection unavailable — enter SKU below."); }
      } catch { setStatus("Camera access denied. Enter SKU manually."); }
    })();
    return () => { clearInterval(interval); streamRef.current?.getTracks().forEach(t => t.stop()); };
  }, []);
  const stop = () => { streamRef.current?.getTracks().forEach(t => t.stop()); onClose(); };
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.85)", zIndex: 2000, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: "#111", borderRadius: 12, overflow: "hidden", width: "100%", maxWidth: 420 }}>
        <div style={{ background: GOLD_DARK, padding: "13px 18px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ color: WHITE, fontWeight: 700, fontSize: 14 }}>SCAN BARCODE</span>
          <button onClick={stop} style={{ background: "none", border: "none", color: "rgba(255,255,255,0.7)", cursor: "pointer" }}>{Icons.close}</button>
        </div>
        <div style={{ position: "relative", background: "#000" }}>
          <video ref={videoRef} style={{ width: "100%", display: "block", maxHeight: 280, objectFit: "cover" }} muted playsInline />
          <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", pointerEvents: "none" }}>
            <div style={{ width: 200, height: 80, border: `2px solid ${GOLD}`, borderRadius: 4, boxShadow: "0 0 0 2000px rgba(0,0,0,0.45)" }} />
          </div>
        </div>
        <div style={{ padding: 16, background: "#1c1c1c" }}>
          <p style={{ color: "#aaa", fontSize: 12, margin: "0 0 12px", textAlign: "center" }}>{status}</p>
          <div style={{ display: "flex", gap: 8 }}>
            <input value={manualSku} onChange={e => setManualSku(e.target.value)} onKeyDown={e => e.key === "Enter" && manualSku && onDetect(manualSku)} placeholder="Enter SKU manually..." style={{ flex: 1, padding: "9px 12px", borderRadius: 6, border: `1px solid ${GOLD_DARK}`, background: "#2a2a2a", color: "#fff", fontSize: 13 }} />
            <button onClick={() => manualSku && onDetect(manualSku)} style={{ padding: "9px 16px", background: GOLD, color: WHITE, border: "none", borderRadius: 6, fontWeight: 700, cursor: "pointer" }}>Search</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── Modal ─────────────────────────────────────────────────────────
function Modal({ title, onClose, children, maxWidth = 500 }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: WHITE, borderRadius: 12, width: "100%", maxWidth, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "18px 22px", borderBottom: `1px solid ${GOLD_LIGHT}` }}>
          <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: "#1F2937" }}>{title}</h3>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: GRAY }}>{Icons.close}</button>
        </div>
        <div style={{ padding: "20px 22px" }}>{children}</div>
      </div>
    </div>
  );
}

function StatusBanner({ msg }) {
  if (!msg) return null;
  const [type, text] = msg.split(":");
  const s = { success: { bg: "#F0FDF4", border: "#86EFAC", color: "#166534" }, error: { bg: "#FEF2F2", border: "#FCA5A5", color: "#991B1B" }, info: { bg: CREAM, border: GOLD_LIGHT, color: GOLD_DARK } }[type] || { bg: CREAM, border: GOLD_LIGHT, color: GOLD_DARK };
  return <div style={{ background: s.bg, border: `1px solid ${s.border}`, borderRadius: 8, padding: "10px 14px", fontSize: 13, color: s.color, marginTop: 8 }}>{text}</div>;
}

const inp = { display: "block", width: "100%", padding: "9px 12px", border: `1px solid ${BORDER}`, borderRadius: 8, marginTop: 4, boxSizing: "border-box", fontSize: 14, background: WHITE, outline: "none", color: "#1F2937" };
const labelStyle = { fontSize: 12, color: GOLD_DARK, fontWeight: 600, display: "block", marginBottom: 2 };

// ── Product Form ──────────────────────────────────────────────────
function ProductForm({ initial, onSave, onCancel, existingProducts }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(() => initial || { name: "", category: "Clothing", business: "Blingshop", price: "", cost: "", stock: "", image: "", sku: generateSKU("Clothing", "Blingshop", []) });
  const [skuEdited, setSkuEdited] = useState(isEdit);
  const fileRef = useRef();
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const handleCatOrBiz = (cat, biz) => { set("category", cat); set("business", biz); if (!skuEdited) set("sku", generateSKU(cat, biz, existingProducts)); };
  const handleImage = e => { const f = e.target.files[0]; if (!f) return; const r = new FileReader(); r.onload = ev => set("image", ev.target.result); r.readAsDataURL(f); };
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div>
        <label style={labelStyle}>Business</label>
        <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
          {BUSINESSES.map(b => (
            <button key={b} onClick={() => handleCatOrBiz(form.category, b)} style={{ flex: 1, padding: "10px", borderRadius: 8, border: `2px solid ${form.business === b ? (b === "Blingshop" ? GOLD : RC_BLUE) : BORDER}`, background: form.business === b ? (b === "Blingshop" ? GOLD_LIGHT : RC_LIGHT) : WHITE, fontWeight: form.business === b ? 700 : 400, color: form.business === b ? (b === "Blingshop" ? GOLD_DARK : RC_DARK) : GRAY, cursor: "pointer", fontSize: 13 }}>{b}</button>
          ))}
        </div>
      </div>
      <div><label style={labelStyle}>Product Name</label><input value={form.name} onChange={e => set("name", e.target.value)} placeholder="Product Name" style={inp} /></div>
      <div><label style={labelStyle}>Category</label>
        <select value={form.category} onChange={e => handleCatOrBiz(e.target.value, form.business)} style={{ ...inp, marginTop: 4 }}>
          {CATEGORIES.slice(1).map(c => <option key={c}>{c}</option>)}
        </select>
      </div>
      <div>
        <label style={labelStyle}>SKU Code</label>
        <div style={{ position: "relative" }}>
          <input value={form.sku} onChange={e => { set("sku", e.target.value); setSkuEdited(true); }} placeholder="Auto-generated" style={{ ...inp, paddingRight: skuEdited ? 12 : 90 }} />
          {!skuEdited && <span style={{ position: "absolute", right: 10, top: "50%", transform: "translateY(-50%)", fontSize: 11, color: GOLD_DARK, fontWeight: 600, background: GOLD_LIGHT, padding: "2px 8px", borderRadius: 10 }}>Auto</span>}
        </div>
        {!skuEdited && <div style={{ fontSize: 11, color: GRAY, marginTop: 4 }}>Format: {form.business === "Blingshop" ? "B" : "R"}{CATEGORY_CODES[form.category] || "O"}01. Click to override.</div>}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
        <div><label style={labelStyle}>Cost (LKR)</label><input value={form.cost ?? ""} onChange={e => set("cost", e.target.value)} placeholder="0" type="number" style={inp} /></div>
        <div><label style={labelStyle}>Price (LKR)</label><input value={form.price} onChange={e => set("price", e.target.value)} placeholder="0" type="number" style={inp} /></div>
        <div><label style={labelStyle}>Stock Qty</label><input value={form.stock} onChange={e => set("stock", e.target.value)} placeholder="0" type="number" style={inp} /></div>
      </div>
      <div>
        <label style={labelStyle}>Product Image</label>
        <div style={{ marginTop: 6, display: "flex", gap: 10, alignItems: "center" }}>
          <button onClick={() => fileRef.current.click()} style={{ padding: "8px 16px", background: LIGHT, color: GOLD_DARK, border: `1px solid ${BORDER}`, borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>{Icons.camera} Upload Photo</button>
          {form.image && <img src={form.image} alt="" style={{ width: 52, height: 52, objectFit: "cover", borderRadius: 8, border: `1px solid ${BORDER}` }} />}
        </div>
        <input ref={fileRef} type="file" accept="image/*" style={{ display: "none" }} onChange={handleImage} />
        <input value={form.image && !form.image.startsWith("data:") ? form.image : ""} onChange={e => set("image", e.target.value)} placeholder="Or paste image URL" style={{ ...inp, marginTop: 8, fontSize: 12, color: GRAY }} />
      </div>
      {form.sku && (
        <div style={{ background: LIGHT, border: `1px solid ${BORDER}`, borderRadius: 8, padding: 12, textAlign: "center" }}>
          <div style={{ fontSize: 11, color: GRAY, marginBottom: 4 }}>Barcode Preview</div>
          <BizBadge business={form.business} />
          <div style={{ display: "inline-block", marginTop: 8 }}><BarcodeCanvas value={form.sku} width={200} height={48} price={form.price} /></div>
        </div>
      )}
      <div style={{ display: "flex", gap: 10, paddingTop: 4, borderTop: `1px solid ${GOLD_LIGHT}`, marginTop: 4 }}>
        <button onClick={() => onSave(form)} style={{ flex: 1, padding: "11px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 14 }}>Save Product</button>
        <button onClick={onCancel} style={{ padding: "11px 20px", background: LIGHT, color: GRAY, border: `1px solid #E5E7EB`, borderRadius: 8, cursor: "pointer", fontSize: 14 }}>Cancel</button>
      </div>
    </div>
  );
}

// ── Excel Import Modal ────────────────────────────────────────────
function ImportModal({ existingProducts, onImport, onClose }) {
  const [rows, setRows] = useState([]); const [fileName, setFileName] = useState(""); const [importing, setImporting] = useState(false); const [status, setStatus] = useState(""); const fileRef = useRef();
  const downloadTemplate = () => {
    const sample = [
      { "Business": "Blingshop", "Product Name": "Gold Necklace", "Category": "Jewelry", "Cost (LKR)": 3000, "Price (LKR)": 4500, "Stock Quantity": 10, "SKU": "", "Image URL": "" },
      { "Business": "Blingshop", "Product Name": "Pearl Earrings", "Category": "Jewelry", "Cost (LKR)": 1400, "Price (LKR)": 2200, "Stock Quantity": 15, "SKU": "", "Image URL": "" },
      { "Business": "RC Boutique", "Product Name": "Silk Scarf", "Category": "Shawls", "Cost (LKR)": 2200, "Price (LKR)": 3500, "Stock Quantity": 8, "SKU": "", "Image URL": "" },
      { "Business": "RC Boutique", "Product Name": "Leather Bag", "Category": "Bags", "Cost (LKR)": 6000, "Price (LKR)": 9500, "Stock Quantity": 5, "SKU": "", "Image URL": "" },
      { "Business": "RC Boutique", "Product Name": "Summer Dress", "Category": "Clothing", "Cost (LKR)": 4000, "Price (LKR)": 6500, "Stock Quantity": 12, "SKU": "", "Image URL": "" },
    ];
    try {
      const ws = xlsxUtils.json_to_sheet(sample);
      ws["!cols"] = [{ wch: 14 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 30 }];
      const wb = xlsxUtils.book_new();
      xlsxUtils.book_append_sheet(wb, ws, "Products");
      const info = xlsxUtils.aoa_to_sheet([
        ["BLINGSHOP + RC BOUTIQUE — PRODUCT IMPORT TEMPLATE"], [""],
        ["REQUIRED FIELDS:"], ["Business", "Blingshop or RC Boutique"], ["Product Name", "Any text"], ["Category", "Clothing, Jewelry, Accessories, Bags, Shoes, Shawls, or Other"], ["Price (LKR)", "Number only (e.g. 4500)"], ["Stock Quantity", "Whole number (e.g. 10)"], [""],
        ["OPTIONAL FIELDS:"], ["Cost (LKR)", "Number only — what you paid for the item. Leave blank to default to 0"], ["SKU", "Leave blank to auto-generate (e.g. BJ01 or RC01)"], ["Image URL", "Paste a direct image link"], [""],
        ["CATEGORY OPTIONS:"], ["Blingshop", "Jewelry, Accessories, Other"], ["RC Boutique", "Clothing, Bags, Shoes, Shawls, Other"],
      ]);
      xlsxUtils.book_append_sheet(wb, info, "Instructions");
      xlsxWriteFile(wb, "BlingshopRC_Import_Template.xlsx");
    } catch { alert("Export requires xlsx. Run: npm install xlsx"); }
  };
  const handleFile = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    setFileName(file.name); setStatus(""); setRows([]);
    try {
      const buf = await file.arrayBuffer();
      const wb = xlsxRead(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = xlsxUtils.sheet_to_json(ws, { defval: "" });
      if (raw.length === 0) { setStatus("error:No data rows found in this file."); return; }
      setRows(raw.map(mapExcelRow).map((r, i) => validateExcelRow(r, i)));
    } catch (err) { setStatus(`error:Could not read file. ${err.message}`); }
  };
  const validRows = rows.filter(r => r.valid);
  const handleImport = async () => { if (!validRows.length) return; setImporting(true); try { await onImport(validRows); } finally { setImporting(false); } };
  return (
    <Modal title="Import Products from Excel" onClose={onClose} maxWidth={680}>
      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
        <div style={{ background: LIGHT, border: `1px solid ${BORDER}`, borderRadius: 8, padding: 14 }}>
          <div style={{ fontWeight: 700, fontSize: 13, color: "#1F2937", marginBottom: 6 }}>Required columns in your Excel sheet</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4, fontSize: 12, color: GRAY, lineHeight: 1.8 }}>
            <div><strong style={{ color: "#1F2937" }}>Business</strong> — Blingshop or RC Boutique</div>
            <div><strong style={{ color: "#1F2937" }}>Product Name</strong> — any text</div>
            <div><strong style={{ color: "#1F2937" }}>Category</strong> — Clothing, Jewelry, Bags, Shoes, Shawls, etc.</div>
            <div><strong style={{ color: "#1F2937" }}>Price (LKR)</strong> — number only</div>
            <div><strong style={{ color: "#1F2937" }}>Stock Quantity</strong> — whole number</div>
            <div style={{ color: "#9CA3AF" }}>Cost (LKR), SKU, Image URL — optional</div>
          </div>
          <button onClick={downloadTemplate} style={{ marginTop: 10, display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", background: WHITE, color: GOLD_DARK, border: `1px solid ${BORDER}`, borderRadius: 7, cursor: "pointer", fontSize: 12, fontWeight: 600 }}>{Icons.download} Download Template (with Instructions sheet)</button>
        </div>
        <button onClick={() => fileRef.current.click()} style={{ width: "100%", padding: "16px", background: WHITE, border: `2px dashed ${BORDER}`, borderRadius: 10, cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 6, color: GOLD_DARK }}>
          {Icons.upload}<span style={{ fontSize: 13, fontWeight: 600 }}>{fileName || "Click to select Excel file (.xlsx, .xls)"}</span>
        </button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleFile} />
        <StatusBanner msg={status} />
        {rows.length > 0 && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <span style={{ fontSize: 13, fontWeight: 700, color: "#1F2937" }}>{rows.length} rows found</span>
              <span style={{ fontSize: 12, color: validRows.length === rows.length ? "#059669" : "#B45309" }}>{validRows.length} ready, {rows.length - validRows.length} with errors</span>
            </div>
            <div style={{ maxHeight: 260, overflowY: "auto", border: `1px solid ${BORDER}`, borderRadius: 8 }}>
              <table style={{ width: "100%", fontSize: 12, borderCollapse: "collapse" }}>
                <thead>
                  <tr style={{ background: LIGHT }}>
                    {["Row", "Business", "Name", "Category", "Cost", "Price", "Stock", "Status"].map(h => <th key={h} style={{ padding: "8px 10px", textAlign: "left", position: "sticky", top: 0, background: LIGHT, color: GRAY, fontWeight: 600 }}>{h}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, i) => (
                    <tr key={i} style={{ borderTop: `1px solid ${BORDER}`, background: r.valid ? WHITE : "#FEF2F2" }}>
                      <td style={{ padding: "7px 10px", color: GRAY }}>{r.rowNum}</td>
                      <td style={{ padding: "7px 10px" }}><BizBadge business={r.business} /></td>
                      <td style={{ padding: "7px 10px", color: "#1F2937" }}>{r.name || "—"}</td>
                      <td style={{ padding: "7px 10px", color: "#1F2937" }}>{r.category}</td>
                      <td style={{ padding: "7px 10px" }}>{isNaN(r.cost) ? "0" : r.cost.toLocaleString()}</td>
                      <td style={{ padding: "7px 10px" }}>{isNaN(r.price) ? "—" : r.price.toLocaleString()}</td>
                      <td style={{ padding: "7px 10px" }}>{isNaN(r.stock) ? "—" : r.stock}</td>
                      <td style={{ padding: "7px 10px", color: r.valid ? "#059669" : "#DC2626", fontSize: 11 }}>{r.valid ? "Ready" : r.errors.join(", ")}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
        <div style={{ display: "flex", gap: 10, paddingTop: 4, borderTop: `1px solid ${GOLD_LIGHT}` }}>
          <button onClick={handleImport} disabled={!validRows.length || importing} style={{ flex: 1, padding: "11px", background: validRows.length ? GOLD_DARK : "#D1D5DB", color: WHITE, border: "none", borderRadius: 8, cursor: validRows.length ? "pointer" : "not-allowed", fontWeight: 700, fontSize: 14 }}>
            {importing ? "Importing..." : `Import ${validRows.length} Product${validRows.length === 1 ? "" : "s"}`}
          </button>
          <button onClick={onClose} style={{ padding: "11px 20px", background: LIGHT, color: GRAY, border: `1px solid #E5E7EB`, borderRadius: 8, cursor: "pointer", fontSize: 14 }}>Cancel</button>
        </div>
      </div>
    </Modal>
  );
}

// ── Bulk Add (Quick Source) Modal ────────────────────────────────────
// Fast-capture flow for sourcing trips: photo + cost only. Name/SKU/price are generated
// automatically and land in a pending queue for later review before entering real inventory.
function BulkAddModal({ existingProducts, pendingProducts, onSave, onClose, blingshopThbRate, setBlingshopThbRate, rcThbRate, setRcThbRate }) {
  const [business, setBusiness] = useState("Blingshop");
  // Deliberately not reset between saves (see reset() below) — sourcing trips add many of the
  // same category back-to-back (e.g. a dozen necklaces), so re-picking it each time would be
  // exactly the back-and-forth this screen exists to avoid.
  const [category, setCategory] = useState("Other");
  const [image, setImage] = useState("");
  const [costThb, setCostThb] = useState("");
  const [qty, setQty] = useState("1");
  // Not reset between saves either — same reasoning as category: a batch of old stock with no
  // recoverable cost stays in that mode until the user explicitly turns it off.
  const [unknownCost, setUnknownCost] = useState(false);
  const [manualPrice, setManualPrice] = useState("");
  const isRC = business === "RC Boutique";
  const storedRate = isRC ? rcThbRate : blingshopThbRate;
  const [rateInput, setRateInput] = useState(storedRate || "");
  const [saving, setSaving] = useState(false);
  const [savedCount, setSavedCount] = useState(0);
  const cameraRef = useRef();
  const galleryRef = useRef();

  // Each business remembers its own rate — switching the toggle should show that
  // business's last-saved rate, not whatever was typed for the other one.
  useEffect(() => { setRateInput(storedRate || ""); }, [business]); // eslint-disable-line react-hooks/exhaustive-deps

  const handleBusiness = b => { setBusiness(b); if (!categoriesForBiz(b).includes(category)) setCategory("Other"); };

  const price = unknownCost ? (+manualPrice || 0) : quickAddPrice(business, costThb, rateInput);
  const canSave = image && qty !== "" && +qty > 0 && (
    unknownCost ? manualPrice !== "" && +manualPrice > 0 : costThb !== "" && +costThb > 0 && +rateInput > 0
  );

  const handleImage = e => {
    const f = e.target.files[0]; if (!f) return;
    const r = new FileReader();
    r.onload = ev => setImage(ev.target.result);
    r.readAsDataURL(f);
  };

  const reset = () => {
    setImage(""); setCostThb(""); setQty("1");
    if (cameraRef.current) cameraRef.current.value = "";
    if (galleryRef.current) galleryRef.current.value = "";
  };

  const handleSave = async (closeAfter) => {
    if (!canSave || saving) return;
    setSaving(true);
    try {
      if (!unknownCost && +rateInput !== +storedRate) (isRC ? setRcThbRate : setBlingshopThbRate)(rateInput);
      const allKnown = [...existingProducts, ...pendingProducts];
      const sku = generateSKU(category, business, allKnown);
      const item = {
        id: generateId(), business, category, sku, name: `${sku} ${category}`,
        costThb: unknownCost ? 0 : +costThb, cost: unknownCost ? 0 : quickAddCostLkr(costThb, rateInput),
        costUnknown: unknownCost, price, stock: +qty, image,
      };
      await onSave(item);
      setSavedCount(c => c + 1);
      if (closeAfter) onClose(); else reset();
    } finally { setSaving(false); }
  };

  // Closing (via the X, backdrop, etc.) shouldn't discard a filled-in entry the user forgot
  // to explicitly save — auto-save it first if the form is currently valid.
  const handleClose = () => { if (canSave && !saving) handleSave(true); else onClose(); };

  return (
    <Modal title="Bulk Add — Quick Source" onClose={handleClose}>
      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
        <div style={{ background: LIGHT, border: `1px solid ${BORDER}`, borderRadius: 8, padding: "10px 12px", fontSize: 12, color: GRAY }}>
          Snap a photo and enter cost — that's it. Name, SKU, and selling price are generated automatically. Review and finalize everything later in <strong>Review Pending</strong> before it joins the real inventory.
        </div>

        <div>
          <label style={labelStyle}>Business</label>
          <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
            {BUSINESSES.map(b => (
              <button key={b} onClick={() => handleBusiness(b)} style={{ flex: 1, padding: "10px", borderRadius: 8, border: `2px solid ${business === b ? (b === "Blingshop" ? GOLD : RC_BLUE) : BORDER}`, background: business === b ? (b === "Blingshop" ? GOLD_LIGHT : RC_LIGHT) : WHITE, fontWeight: business === b ? 700 : 400, color: business === b ? (b === "Blingshop" ? GOLD_DARK : RC_DARK) : GRAY, cursor: "pointer", fontSize: 13 }}>{b}</button>
            ))}
          </div>
        </div>

        <div>
          <label style={labelStyle}>Category</label>
          <select value={category} onChange={e => setCategory(e.target.value)} style={{ ...inp, marginTop: 6 }}>
            {categoriesForBiz(business).slice(1).map(c => <option key={c}>{c}</option>)}
          </select>
          <div style={{ fontSize: 11, color: GRAY, marginTop: 4 }}>Stays selected across saves — set it once per batch (e.g. all necklaces), then add items without re-picking it.</div>
        </div>

        {!unknownCost && (
          <div>
            <label style={labelStyle}>Exchange Rate (LKR per THB)</label>
            <input value={rateInput} onChange={e => setRateInput(e.target.value)} placeholder="e.g. 9.5" type="number" style={inp} />
            <div style={{ fontSize: 11, color: GRAY, marginTop: 4 }}>THB rate fluctuates — enter it once per restock. It's remembered per business until you change it here again.</div>
          </div>
        )}

        <label style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer", fontSize: 13, color: "#374151", background: LIGHT, border: `1px solid ${BORDER}`, borderRadius: 8, padding: "10px 12px" }}>
          <input type="checkbox" checked={unknownCost} onChange={e => setUnknownCost(e.target.checked)} style={{ width: 15, height: 15, accentColor: GOLD_DARK, cursor: "pointer", flexShrink: 0 }} />
          I don't know the cost for this item (old stock) — let me enter the selling price directly
        </label>

        <div>
          <label style={labelStyle}>Product Photo</label>
          {image ? (
            <div style={{ marginTop: 6, position: "relative" }}>
              <img src={image} alt="" style={{ width: "100%", maxHeight: 200, objectFit: "contain", borderRadius: 10, border: `1px solid ${BORDER}`, display: "block" }} />
              <button onClick={() => setImage("")} style={{ position: "absolute", top: 8, right: 8, width: 28, height: 28, borderRadius: "50%", border: "none", background: "rgba(0,0,0,0.6)", color: WHITE, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>{Icons.close}</button>
            </div>
          ) : (
            <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
              <button onClick={() => cameraRef.current.click()} style={{ flex: 1, padding: "20px 8px", background: WHITE, border: `2px dashed ${BORDER}`, borderRadius: 10, cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 6, color: GOLD_DARK }}>
                {Icons.camera}<span style={{ fontSize: 13, fontWeight: 600 }}>Take Photo</span>
              </button>
              <button onClick={() => galleryRef.current.click()} style={{ flex: 1, padding: "20px 8px", background: WHITE, border: `2px dashed ${BORDER}`, borderRadius: 10, cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 6, color: GOLD_DARK }}>
                {Icons.image}<span style={{ fontSize: 13, fontWeight: 600 }}>From Gallery</span>
              </button>
            </div>
          )}
          <input ref={cameraRef} type="file" accept="image/*" capture="environment" style={{ display: "none" }} onChange={handleImage} />
          <input ref={galleryRef} type="file" accept="image/*" style={{ display: "none" }} onChange={handleImage} />
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <div>
            <label style={labelStyle}>{unknownCost ? "Selling Price (LKR)" : "Cost (THB)"}</label>
            <input value={unknownCost ? manualPrice : costThb} onChange={e => (unknownCost ? setManualPrice : setCostThb)(e.target.value)} placeholder="0" type="number" style={inp} />
          </div>
          <div>
            <label style={labelStyle}>Quantity</label>
            <input value={qty} onChange={e => setQty(e.target.value)} placeholder="1" type="number" min="1" style={inp} />
          </div>
        </div>

        {!unknownCost && costThb !== "" && +costThb > 0 && +rateInput > 0 && (
          <div style={{ background: GOLD_LIGHT, border: `1px solid ${GOLD}`, borderRadius: 8, padding: "10px 12px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span style={{ fontSize: 12, color: GOLD_DARK, fontWeight: 600 }}>{isRC ? "Starting Price (break-even)" : "Selling Price (auto)"}</span>
              <span style={{ fontSize: 16, fontWeight: 800, color: GOLD_DARK }}>LKR {price.toLocaleString()}</span>
            </div>
            {isRC && <div style={{ fontSize: 11, color: GOLD_DARK, marginTop: 4 }}>No profit margin yet — cost + packing (LKR {RC_PACKING_LKR}) + cargo (LKR {RC_CARGO_LKR}). Set the profit by adjusting price in Review Pending.</div>}
          </div>
        )}

        {unknownCost && manualPrice !== "" && +manualPrice > 0 && (
          <div style={{ background: GOLD_LIGHT, border: `1px solid ${GOLD}`, borderRadius: 8, padding: "10px 12px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span style={{ fontSize: 12, color: GOLD_DARK, fontWeight: 600 }}>Selling Price (manual)</span>
              <span style={{ fontSize: 16, fontWeight: 800, color: GOLD_DARK }}>LKR {(+manualPrice).toLocaleString()}</span>
            </div>
            <div style={{ fontSize: 11, color: GOLD_DARK, marginTop: 4 }}>Cost marked unknown — won't count toward Stock Worth in Reports until a real cost is filled in later.</div>
          </div>
        )}

        {savedCount > 0 && <StatusBanner msg={`success:${savedCount} item${savedCount === 1 ? "" : "s"} added to the pending queue.`} />}

        <div style={{ display: "flex", gap: 10, paddingTop: 4, borderTop: `1px solid ${GOLD_LIGHT}` }}>
          <button disabled={!canSave || saving} onClick={() => handleSave(false)} style={{ flex: 1, padding: "11px", background: canSave ? LIGHT : "#F3F4F6", color: canSave ? GOLD_DARK : "#9CA3AF", border: `1px solid ${BORDER}`, borderRadius: 8, cursor: canSave ? "pointer" : "not-allowed", fontWeight: 700, fontSize: 14 }}>Save &amp; Add Another</button>
          <button disabled={!canSave || saving} onClick={() => handleSave(true)} style={{ flex: 1, padding: "11px", background: canSave ? GOLD_DARK : "#D1D5DB", color: WHITE, border: "none", borderRadius: 8, cursor: canSave ? "pointer" : "not-allowed", fontWeight: 700, fontSize: 14 }}>Save &amp; Close</button>
        </div>
      </div>
    </Modal>
  );
}

// ── Review Pending Modal ─────────────────────────────────────────────
// Owner-only queue where Quick Source captures get their name/price/category/stock finalized
// before being bulk-imported into the real product catalog.
function ReviewPendingModal({ pendingProducts, setPendingProducts, setProducts, onClose }) {
  const [edits, setEdits] = useState({});
  const [selected, setSelected] = useState({});
  const [busy, setBusy] = useState(false);

  const getVal = (item, field) => edits[item.id]?.[field] ?? item[field];
  const setVal = (id, field, val) => setEdits(e => ({ ...e, [id]: { ...e[id], [field]: val } }));

  // Bulk Add seeds name as "SKU Category" (see BulkAddModal). Keep it in sync when the category
  // is corrected here — but only while the name still matches that auto pattern, so a name the
  // user has already typed something custom into never gets silently overwritten.
  const handleCategoryChange = (item, newCategory) => {
    const autoName = `${item.sku} ${getVal(item, "category")}`;
    setEdits(e => {
      const next = { ...e[item.id], category: newCategory };
      if (getVal(item, "name") === autoName) next.name = `${item.sku} ${newCategory}`;
      return { ...e, [item.id]: next };
    });
  };

  const toggleAll = () => {
    const allSelected = pendingProducts.every(p => selected[p.id]);
    const next = {}; pendingProducts.forEach(p => next[p.id] = !allSelected);
    setSelected(next);
  };

  const approveOne = async (item) => {
    const merged = { ...item, ...edits[item.id] };
    const newProd = {
      id: generateId(), name: merged.name, category: merged.category, business: merged.business,
      price: +merged.price, cost: +merged.cost, stock: +merged.stock, sku: merged.sku, image: merged.image || "",
    };
    try { await apiFetch("/products", { method: "POST", body: newProd }); } catch { }
    try { await apiFetch(`/pending-products/${item.id}`, { method: "DELETE" }); } catch { }
    setProducts(ps => [...ps, newProd]);
    setPendingProducts(ps => ps.filter(p => p.id !== item.id));
    setEdits(e => { const n = { ...e }; delete n[item.id]; return n; });
  };

  const approveSelected = async () => {
    const chosen = pendingProducts.filter(p => selected[p.id]);
    if (!chosen.length) return;
    setBusy(true);
    try { for (const item of chosen) await approveOne(item); } finally { setBusy(false); }
  };

  const rejectOne = async (id) => {
    if (!confirm("Discard this pending item? This cannot be undone.")) return;
    try { await apiFetch(`/pending-products/${id}`, { method: "DELETE" }); } catch { }
    setPendingProducts(ps => ps.filter(p => p.id !== id));
  };

  return (
    <Modal title={`Review Pending (${pendingProducts.length})`} onClose={onClose} maxWidth={860}>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        {pendingProducts.length === 0 && <div style={{ textAlign: "center", color: "#9CA3AF", padding: "30px 0", fontSize: 14 }}>Nothing pending — Bulk Add items will show up here.</div>}
        {pendingProducts.length > 0 && (
          <>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: "#374151", cursor: "pointer" }}>
                <input type="checkbox" checked={pendingProducts.length > 0 && pendingProducts.every(p => selected[p.id])} onChange={toggleAll} style={{ width: 15, height: 15, accentColor: GOLD_DARK, cursor: "pointer" }} />
                Select all
              </label>
              <button disabled={busy || !Object.values(selected).some(Boolean)} onClick={approveSelected} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: Object.values(selected).some(Boolean) ? "#059669" : "#D1D5DB", color: WHITE, border: "none", borderRadius: 8, cursor: Object.values(selected).some(Boolean) ? "pointer" : "not-allowed", fontWeight: 700, fontSize: 13 }}>{Icons.check} Approve Selected</button>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 10, maxHeight: 480, overflowY: "auto" }}>
              {pendingProducts.map(item => (
                <div key={item.id} style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "flex-start", background: LIGHT, border: `1px solid ${BORDER}`, borderRadius: 10, padding: 10 }}>
                  <input type="checkbox" checked={!!selected[item.id]} onChange={() => setSelected(s => ({ ...s, [item.id]: !s[item.id] }))} style={{ width: 15, height: 15, marginTop: 10, flexShrink: 0, accentColor: GOLD_DARK, cursor: "pointer" }} />
                  <div style={{ width: 56, height: 56, borderRadius: 8, overflow: "hidden", background: WHITE, flexShrink: 0, border: `1px solid ${BORDER}` }}>
                    {item.image ? <img src={item.image} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : null}
                  </div>
                  {/* minmax(120px,1fr) + minWidth:0 on each cell keeps this from being forced wider than the
                      modal by an input's intrinsic min-content size — the exact reason this list would
                      silently overflow off-screen on a narrow phone/tablet instead of wrapping. */}
                  <div style={{ flex: "1 1 240px", minWidth: 0, display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(120px,1fr))", gap: 8 }}>
                    <div style={{ gridColumn: "1 / -1", display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                      <BizBadge business={item.business} />
                      <span style={{ fontSize: 11, color: GRAY }}>SKU: {item.sku}</span>
                      {item.costUnknown
                        ? <span style={{ fontSize: 11, fontWeight: 700, color: "#B45309", background: "#FEF3C7", padding: "1px 7px", borderRadius: 10 }}>Cost unknown — old stock</span>
                        : <span style={{ fontSize: 11, color: GRAY }}>Cost: THB {(+item.costThb).toLocaleString()} (LKR {(+item.cost).toLocaleString()})</span>}
                    </div>
                    <div style={{ minWidth: 0 }}>
                      <label style={{ fontSize: 10, color: GRAY }}>Name</label>
                      <input value={getVal(item, "name")} onChange={e => setVal(item.id, "name", e.target.value)} style={{ ...inp, marginTop: 2, fontSize: 12, padding: "6px 8px" }} />
                    </div>
                    <div style={{ minWidth: 0 }}>
                      <label style={{ fontSize: 10, color: GRAY }}>Category</label>
                      <select value={getVal(item, "category")} onChange={e => handleCategoryChange(item, e.target.value)} style={{ ...inp, marginTop: 2, fontSize: 12, padding: "6px 8px" }}>
                        {CATEGORIES.slice(1).map(c => <option key={c}>{c}</option>)}
                      </select>
                    </div>
                    <div style={{ minWidth: 0 }}>
                      <label style={{ fontSize: 10, color: GRAY }}>Price (LKR)</label>
                      <input value={getVal(item, "price")} onChange={e => setVal(item.id, "price", e.target.value)} type="number" style={{ ...inp, marginTop: 2, fontSize: 12, padding: "6px 8px" }} />
                    </div>
                    <div style={{ minWidth: 0 }}>
                      <label style={{ fontSize: 10, color: GRAY }}>Stock</label>
                      <input value={getVal(item, "stock")} onChange={e => setVal(item.id, "stock", e.target.value)} type="number" style={{ ...inp, marginTop: 2, fontSize: 12, padding: "6px 8px" }} />
                    </div>
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6, flexShrink: 0 }}>
                    <button disabled={busy} onClick={() => approveOne(item)} style={{ width: 30, height: 30, display: "flex", alignItems: "center", justifyContent: "center", background: "#F0FDF4", color: "#059669", border: "none", borderRadius: 6, cursor: "pointer" }}>{Icons.check}</button>
                    <button disabled={busy} onClick={() => rejectOne(item.id)} style={{ width: 30, height: 30, display: "flex", alignItems: "center", justifyContent: "center", background: "#FEF2F2", color: "#DC2626", border: "none", borderRadius: 6, cursor: "pointer" }}>{Icons.trash}</button>
                  </div>
                </div>
              ))}
            </div>
          </>
        )}
      </div>
    </Modal>
  );
}

// ── Label Print Modal ─────────────────────────────────────────────
function LabelPrintModal({ product, zebraIP, onClose }) {
  const [labelType, setLabelType] = useState("standard");
  const [presetId, setPresetId] = useState("50x30");
  const [wMm, setWMm] = useState(50);
  const [hMm, setHMm] = useState(30);
  const [qty, setQty] = useState(1);
  const [activeTab, setActiveTab] = useState("browser");
  const [status, setStatus] = useState("");
  const [opts, setOpts] = useState({
    showBiz: false, showName: true, showBarcode: true, showSKU: true, showPrice: true,
    barcodeScale: 2, printDPI: 300, vAlign: "center", textScale: 1,
  });
  const canvasRef = useRef();

  const setOpt = (key, val) => setOpts(o => ({ ...o, [key]: val }));

  const layoutWarnings = useMemo(() => (
    labelType === "jewelry"
      ? computeJewelryLayout(product, wMm, hMm, opts).warnings
      : computeStandardLayout(product, wMm, hMm, opts).warnings
  ), [product, labelType, wMm, hMm, opts]);

  useEffect(() => {
    drawLabelToCanvas(canvasRef.current, product, labelType, wMm, hMm, opts);
  }, [product, labelType, wMm, hMm, opts]);

  const handleType = (t) => {
    setLabelType(t);
    const first = LABEL_PRESETS[t][0];
    setPresetId(first.id); setWMm(first.w); setHMm(first.h);
  };

  const handlePreset = (id) => {
    setPresetId(id);
    const p = LABEL_PRESETS[labelType].find(x => x.id === id);
    if (p) { setWMm(p.w); setHMm(p.h); }
  };

  const zpl = labelType === "jewelry"
    ? generateZPLJewelry(product, wMm, hMm, qty, opts)
    : generateZPLStandard(product, wMm, hMm, qty, opts);

  const printBrowser = () => {
    const canvas = canvasRef.current; if (!canvas) return;
    const imgData = canvas.toDataURL("image/png", 1.0);
    const labels = Array(qty).fill(`<img src="${imgData}" style="width:${wMm}mm;height:${hMm}mm;display:block;page-break-after:always"/>`).join("");
    const win = window.open("", "_blank", "width=1,height=1,left=-1000,top=-1000");
    if (!win) { setStatus("error:Popup blocked — allow popups for this site."); return; }
    win.document.write(`<!DOCTYPE html><html><head><title>${product.sku || product.name || "Label"}</title><style>
      @page{margin:0;size:${wMm}mm ${hMm}mm}
      *{box-sizing:border-box;margin:0;padding:0}
      body{background:#fff}
      img{width:${wMm}mm;height:${hMm}mm;display:block}
    </style></head><body>${labels}</body></html>`);
    win.document.close();
    win.focus();
    setTimeout(() => { win.print(); win.close(); }, 350);
  };

  const printNetwork = () => printZebraZPL(zpl, zebraIP, setStatus);
  const copyZPL = () => { navigator.clipboard.writeText(zpl); setStatus("success:ZPL code copied to clipboard."); };

  const tabBtn = (id, label) => (
    <button key={id} onClick={() => setActiveTab(id)} style={{
      padding: "8px 14px", border: "none", cursor: "pointer", background: "none",
      fontWeight: activeTab === id ? 700 : 400, fontSize: 13,
      color: activeTab === id ? GOLD_DARK : GRAY,
      borderBottom: `2px solid ${activeTab === id ? GOLD : "transparent"}`,
    }}>{label}</button>
  );

  return (
    <Modal title="Print Barcode Label" onClose={onClose}>
      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>

        {/* Label type */}
        <div style={{ display: "flex", gap: 8 }}>
          {[["standard", "Standard Rectangle"], ["jewelry", "Jewelry Dumbbell"]].map(([t, l]) => (
            <button key={t} onClick={() => handleType(t)} style={{
              flex: 1, padding: "9px 6px", border: `2px solid ${labelType === t ? GOLD_DARK : BORDER}`,
              borderRadius: 8, background: labelType === t ? GOLD_DARK : WHITE,
              color: labelType === t ? WHITE : GRAY, fontWeight: 700, fontSize: 13, cursor: "pointer",
            }}>{l}</button>
          ))}
        </div>

        {labelType === "jewelry" && (
          <div style={{ background: "#FEF3C7", border: "1px solid #FCD34D", borderRadius: 8, padding: "8px 12px", fontSize: 12, color: "#92400E" }}>
            Dumbbell labels have two adhesive end panels with a non-adhesive fold zone in the middle. Used for rings, chains, and small jewellery tags. Requires a thermal transfer (ribbon) printer.
          </div>
        )}

        {/* Print fields */}
        <div style={{ background: LIGHT, border: `1px solid ${BORDER}`, borderRadius: 10, padding: "12px 14px" }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 10 }}>Print Fields</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px 16px" }}>
            {[
              ["showBiz", "Business Name"],
              ["showName", "Product Name"],
              ["showBarcode", "Barcode"],
              ["showSKU", "SKU / Code"],
              ["showPrice", "Price"],
            ].map(([key, label]) => (
              <label key={key} style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer", fontSize: 13, color: "#374151", userSelect: "none" }}>
                <input type="checkbox" checked={opts[key]} onChange={e => setOpt(key, e.target.checked)}
                  style={{ width: 15, height: 15, accentColor: GOLD_DARK, cursor: "pointer" }} />
                {label}
              </label>
            ))}
          </div>

          <div style={{ marginTop: 12, display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-start" }}>
            {/* Barcode bar width */}
            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: GRAY, marginBottom: 6 }}>Barcode Bar Width</div>
              <div style={{ display: "flex", gap: 6 }}>
                {[[1, "Narrow"], [2, "Normal"], [3, "Wide"]].map(([v, l]) => (
                  <button key={v} onClick={() => setOpt("barcodeScale", v)} style={{
                    padding: "4px 12px", borderRadius: 20, fontSize: 12, cursor: "pointer",
                    border: `1px solid ${opts.barcodeScale === v ? GOLD_DARK : BORDER}`,
                    background: opts.barcodeScale === v ? GOLD_DARK : WHITE,
                    color: opts.barcodeScale === v ? WHITE : GRAY,
                    fontWeight: opts.barcodeScale === v ? 700 : 400,
                  }}>{l}</button>
                ))}
              </div>
            </div>

            {/* Text size — scales name/SKU/price fonts after the auto shrink-to-fit sizing */}
            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: GRAY, marginBottom: 6 }}>Text Size</div>
              <div style={{ display: "flex", gap: 6 }}>
                {[[0.8, "Small"], [1, "Normal"], [1.25, "Large"]].map(([v, l]) => (
                  <button key={v} onClick={() => setOpt("textScale", v)} style={{
                    padding: "4px 12px", borderRadius: 20, fontSize: 12, cursor: "pointer",
                    border: `1px solid ${opts.textScale === v ? GOLD_DARK : BORDER}`,
                    background: opts.textScale === v ? GOLD_DARK : WHITE,
                    color: opts.textScale === v ? WHITE : GRAY,
                    fontWeight: opts.textScale === v ? 700 : 400,
                  }}>{l}</button>
                ))}
              </div>
            </div>

            {/* Vertical alignment of the content stack (incl. barcode) within the label */}
            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: GRAY, marginBottom: 6 }}>Vertical Alignment</div>
              <div style={{ display: "flex", gap: 6 }}>
                {[["top", "Top"], ["center", "Center"], ["bottom", "Bottom"]].map(([v, l]) => (
                  <button key={v} onClick={() => setOpt("vAlign", v)} style={{
                    padding: "4px 12px", borderRadius: 20, fontSize: 12, cursor: "pointer",
                    border: `1px solid ${opts.vAlign === v ? GOLD_DARK : BORDER}`,
                    background: opts.vAlign === v ? GOLD_DARK : WHITE,
                    color: opts.vAlign === v ? WHITE : GRAY,
                    fontWeight: opts.vAlign === v ? 700 : 400,
                  }}>{l}</button>
                ))}
              </div>
            </div>

            {/* Print DPI */}
            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: GRAY, marginBottom: 6 }}>Print Resolution</div>
              <div style={{ display: "flex", gap: 6 }}>
                {[[200, "200 dpi"], [300, "300 dpi"], [600, "600 dpi"]].map(([v, l]) => (
                  <button key={v} onClick={() => setOpt("printDPI", v)} style={{
                    padding: "4px 12px", borderRadius: 20, fontSize: 12, cursor: "pointer",
                    border: `1px solid ${opts.printDPI === v ? GOLD_DARK : BORDER}`,
                    background: opts.printDPI === v ? GOLD_DARK : WHITE,
                    color: opts.printDPI === v ? WHITE : GRAY,
                    fontWeight: opts.printDPI === v ? 700 : 400,
                  }}>{l}</button>
                ))}
              </div>
              <div style={{ fontSize: 10, color: "#9CA3AF", marginTop: 4 }}>Higher DPI = sharper barcodes, bigger file</div>
            </div>
          </div>
        </div>

        {/* Size presets + custom */}
        <div>
          <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 6 }}>Label Size</div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 10 }}>
            {LABEL_PRESETS[labelType].map(p => (
              <button key={p.id} onClick={() => handlePreset(p.id)} style={{
                padding: "5px 11px", borderRadius: 20, fontSize: 12, cursor: "pointer",
                border: `1px solid ${presetId === p.id ? GOLD_DARK : BORDER}`,
                background: presetId === p.id ? GOLD_DARK : WHITE,
                color: presetId === p.id ? WHITE : GRAY,
                fontWeight: presetId === p.id ? 700 : 400,
              }}>{p.label}</button>
            ))}
          </div>
          <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
            {[["W (mm)", wMm, v => { setWMm(v); setPresetId("custom"); }],
            ["H (mm)", hMm, v => { setHMm(v); setPresetId("custom"); }]].map(([lbl, val, set]) => (
              <div key={lbl}>
                <div style={{ fontSize: 11, color: GRAY, marginBottom: 3 }}>{lbl}</div>
                <input type="number" value={val} min="10" max="200"
                  onChange={e => set(+e.target.value)}
                  style={{ ...inp, marginTop: 0, width: 68, padding: "6px 8px", fontSize: 13 }} />
              </div>
            ))}
            <div>
              <div style={{ fontSize: 11, color: GRAY, marginBottom: 3 }}>Quantity</div>
              <div style={{ display: "flex", alignItems: "center" }}>
                <button onClick={() => setQty(q => Math.max(1, q - 1))} style={{ width: 28, height: 34, border: `1px solid ${BORDER}`, borderRadius: "6px 0 0 6px", background: LIGHT, cursor: "pointer", fontWeight: 700, fontSize: 16 }}>−</button>
                <input type="number" value={qty} min="1" max="99" onChange={e => setQty(Math.max(1, Math.min(99, +e.target.value)))}
                  style={{ width: 40, height: 34, textAlign: "center", border: `1px solid ${BORDER}`, borderLeft: "none", borderRight: "none", fontSize: 14, fontWeight: 700, background: WHITE, outline: "none" }} />
                <button onClick={() => setQty(q => Math.min(99, q + 1))} style={{ width: 28, height: 34, border: `1px solid ${BORDER}`, borderRadius: "0 6px 6px 0", background: LIGHT, cursor: "pointer", fontWeight: 700, fontSize: 16 }}>+</button>
              </div>
            </div>
          </div>
        </div>

        {/* Live preview */}
        <div style={{ background: "#F8F8F8", borderRadius: 10, padding: 12, border: `1px solid ${BORDER}`, textAlign: "center" }}>
          <div style={{ fontSize: 11, color: GRAY, marginBottom: 8, fontWeight: 600, textTransform: "uppercase", letterSpacing: 1 }}>
            Preview — {wMm}×{hMm}mm
            {labelType === "jewelry" && <span style={{ marginLeft: 8, color: "#92400E", background: "#FEF3C7", padding: "1px 6px", borderRadius: 10, textTransform: "none", letterSpacing: 0 }}>Dumbbell</span>}
          </div>
          <div style={{ display: "inline-block", border: "1px solid #ddd", borderRadius: 4, overflow: "hidden", boxShadow: "0 1px 4px rgba(0,0,0,0.1)" }}>
            <canvas ref={canvasRef} />
          </div>
          {labelType === "jewelry" && (
            <div style={{ fontSize: 11, color: GRAY, marginTop: 6 }}>◀ Front (barcode) &nbsp;|&nbsp; FOLD &nbsp;|&nbsp; Back (name &amp; price) ▶</div>
          )}
        </div>

        {layoutWarnings.length > 0 && (
          <div style={{ background: "#FEF3C7", border: "1px solid #FCD34D", borderRadius: 8, padding: "8px 12px" }}>
            {layoutWarnings.map((w, i) => <div key={i} style={{ fontSize: 12, color: "#92400E" }}>⚠ {w}</div>)}
          </div>
        )}

        {/* Print methods */}
        <div>
          <div style={{ display: "flex", flexWrap: "wrap", borderBottom: `1px solid ${BORDER}` }}>
            {tabBtn("browser", "Browser Print")}
            {tabBtn("network", "Network (Zebra)")}
            {tabBtn("zpl", "ZPL Code")}
          </div>
          <div style={{ paddingTop: 12 }}>
            {activeTab === "browser" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                <p style={{ margin: 0, fontSize: 13, color: GRAY }}>
                  Opens a print dialog sized to the exact label in mm. Works with any USB, network, or cloud printer.{qty > 1 ? ` Prints ${qty} copies.` : ""}
                </p>
                <button onClick={printBrowser} style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, padding: "11px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 14 }}>
                  {Icons.print} Print {qty > 1 ? `${qty} Labels` : "Label"}
                </button>
              </div>
            )}
            {activeTab === "network" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                <p style={{ margin: 0, fontSize: 13, color: GRAY }}>
                  {zebraIP ? `Sends ZPL to Zebra ZD421T at ${zebraIP}.` : "No printer IP — go to Settings."}{qty > 1 ? ` Prints ${qty} copies via ^PQ command.` : ""}
                </p>
                <button onClick={printNetwork} style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, padding: "11px", background: "#2563EB", color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 14 }}>
                  {Icons.wifi} Send to Zebra ZD421T
                </button>
              </div>
            )}
            {activeTab === "zpl" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                <p style={{ margin: 0, fontSize: 13, color: GRAY }}>Paste into Zebra Setup Utility or ZebraDesigner to test.</p>
                <textarea readOnly value={zpl} style={{ width: "100%", height: 140, fontSize: 11, fontFamily: "monospace", background: "#1a1a1a", color: "#4ADE80", padding: 12, borderRadius: 8, border: "none", boxSizing: "border-box", resize: "vertical" }} />
                <button onClick={copyZPL} style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, padding: "11px", background: "#059669", color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 14 }}>
                  {Icons.copy} Copy ZPL Code
                </button>
              </div>
            )}
          </div>
        </div>

        <StatusBanner msg={status} />
      </div>
    </Modal>
  );
}

// ── Bulk Label Print Modal ─────────────────────────────────────────
// Same label design controls as LabelPrintModal, applied across many products at once —
// quantity per product is fixed to that product's current stock count rather than a manual
// stepper, so restocked items each get exactly as many labels as they need.
function BulkLabelPrintModal({ products, zebraIP, onClose }) {
  const [labelType, setLabelType] = useState("standard");
  const [presetId, setPresetId] = useState("50x30");
  const [wMm, setWMm] = useState(50);
  const [hMm, setHMm] = useState(30);
  const [activeTab, setActiveTab] = useState("browser");
  const [status, setStatus] = useState("");
  const [opts, setOpts] = useState({
    showBiz: false, showName: true, showBarcode: true, showSKU: true, showPrice: true,
    barcodeScale: 2, printDPI: 300, vAlign: "center", textScale: 1,
  });
  const canvasRef = useRef();

  const setOpt = (key, val) => setOpts(o => ({ ...o, [key]: val }));

  const withStock = products.filter(p => (+p.stock || 0) > 0);
  const outOfStockCount = products.length - withStock.length;
  const totalLabels = withStock.reduce((s, p) => s + (+p.stock || 0), 0);
  const previewProduct = withStock[0] || products[0];

  const layoutWarnings = useMemo(() => (
    previewProduct ? (labelType === "jewelry"
      ? computeJewelryLayout(previewProduct, wMm, hMm, opts).warnings
      : computeStandardLayout(previewProduct, wMm, hMm, opts).warnings) : []
  ), [previewProduct, labelType, wMm, hMm, opts]);

  useEffect(() => {
    if (previewProduct) drawLabelToCanvas(canvasRef.current, previewProduct, labelType, wMm, hMm, opts);
  }, [previewProduct, labelType, wMm, hMm, opts]);

  const handleType = (t) => {
    setLabelType(t);
    const first = LABEL_PRESETS[t][0];
    setPresetId(first.id); setWMm(first.w); setHMm(first.h);
  };

  const handlePreset = (id) => {
    setPresetId(id);
    const p = LABEL_PRESETS[labelType].find(x => x.id === id);
    if (p) { setWMm(p.w); setHMm(p.h); }
  };

  const genZpl = (p, qty) => labelType === "jewelry"
    ? generateZPLJewelry(p, wMm, hMm, qty, opts)
    : generateZPLStandard(p, wMm, hMm, qty, opts);

  const zpl = useMemo(() => withStock.map(p => genZpl(p, +p.stock)).join("\n"), [withStock, labelType, wMm, hMm, opts]); // eslint-disable-line react-hooks/exhaustive-deps

  const printBrowser = () => {
    if (!withStock.length) { setStatus("error:None of the selected products have stock — nothing to print."); return; }
    const imgTags = withStock.map(p => {
      const canvas = document.createElement("canvas");
      drawLabelToCanvas(canvas, p, labelType, wMm, hMm, opts);
      const imgData = canvas.toDataURL("image/png", 1.0);
      return Array(+p.stock).fill(`<img src="${imgData}" style="width:${wMm}mm;height:${hMm}mm;display:block;page-break-after:always"/>`).join("");
    }).join("");
    const win = window.open("", "_blank", "width=1,height=1,left=-1000,top=-1000");
    if (!win) { setStatus("error:Popup blocked — allow popups for this site."); return; }
    win.document.write(`<!DOCTYPE html><html><head><title>Bulk Labels (${withStock.length} products)</title><style>
      @page{margin:0;size:${wMm}mm ${hMm}mm}
      *{box-sizing:border-box;margin:0;padding:0}
      body{background:#fff}
      img{width:${wMm}mm;height:${hMm}mm;display:block}
    </style></head><body>${imgTags}</body></html>`);
    win.document.close();
    win.focus();
    setTimeout(() => { win.print(); win.close(); }, 350);
  };

  const printNetwork = () => {
    if (!withStock.length) { setStatus("error:None of the selected products have stock — nothing to print."); return; }
    printZebraZPL(zpl, zebraIP, setStatus);
  };
  const copyZPL = () => { navigator.clipboard.writeText(zpl); setStatus("success:ZPL code copied to clipboard."); };

  const tabBtn = (id, label) => (
    <button key={id} onClick={() => setActiveTab(id)} style={{
      padding: "8px 14px", border: "none", cursor: "pointer", background: "none",
      fontWeight: activeTab === id ? 700 : 400, fontSize: 13,
      color: activeTab === id ? GOLD_DARK : GRAY,
      borderBottom: `2px solid ${activeTab === id ? GOLD : "transparent"}`,
    }}>{label}</button>
  );

  return (
    <Modal title={`Bulk Print Labels (${products.length} products)`} onClose={onClose}>
      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>

        <div style={{ background: GOLD_LIGHT, border: `1px solid ${GOLD}`, borderRadius: 8, padding: "10px 12px", fontSize: 13, color: GOLD_DARK }}>
          <strong>{totalLabels.toLocaleString()} label{totalLabels === 1 ? "" : "s"}</strong> total — one per unit in stock, across {withStock.length} product{withStock.length === 1 ? "" : "s"}.
          {outOfStockCount > 0 && <div style={{ marginTop: 4, fontSize: 12 }}>{outOfStockCount} selected product{outOfStockCount === 1 ? "" : "s"} {outOfStockCount === 1 ? "has" : "have"} 0 stock and will be skipped.</div>}
        </div>

        {/* Label type */}
        <div style={{ display: "flex", gap: 8 }}>
          {[["standard", "Standard Rectangle"], ["jewelry", "Jewelry Dumbbell"]].map(([t, l]) => (
            <button key={t} onClick={() => handleType(t)} style={{
              flex: 1, padding: "9px 6px", border: `2px solid ${labelType === t ? GOLD_DARK : BORDER}`,
              borderRadius: 8, background: labelType === t ? GOLD_DARK : WHITE,
              color: labelType === t ? WHITE : GRAY, fontWeight: 700, fontSize: 13, cursor: "pointer",
            }}>{l}</button>
          ))}
        </div>

        {labelType === "jewelry" && (
          <div style={{ background: "#FEF3C7", border: "1px solid #FCD34D", borderRadius: 8, padding: "8px 12px", fontSize: 12, color: "#92400E" }}>
            Dumbbell labels have two adhesive end panels with a non-adhesive fold zone in the middle. Used for rings, chains, and small jewellery tags. Requires a thermal transfer (ribbon) printer.
          </div>
        )}

        {/* Print fields */}
        <div style={{ background: LIGHT, border: `1px solid ${BORDER}`, borderRadius: 10, padding: "12px 14px" }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 10 }}>Print Fields</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px 16px" }}>
            {[
              ["showBiz", "Business Name"],
              ["showName", "Product Name"],
              ["showBarcode", "Barcode"],
              ["showSKU", "SKU / Code"],
              ["showPrice", "Price"],
            ].map(([key, label]) => (
              <label key={key} style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer", fontSize: 13, color: "#374151", userSelect: "none" }}>
                <input type="checkbox" checked={opts[key]} onChange={e => setOpt(key, e.target.checked)}
                  style={{ width: 15, height: 15, accentColor: GOLD_DARK, cursor: "pointer" }} />
                {label}
              </label>
            ))}
          </div>

          <div style={{ marginTop: 12, display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-start" }}>
            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: GRAY, marginBottom: 6 }}>Barcode Bar Width</div>
              <div style={{ display: "flex", gap: 6 }}>
                {[[1, "Narrow"], [2, "Normal"], [3, "Wide"]].map(([v, l]) => (
                  <button key={v} onClick={() => setOpt("barcodeScale", v)} style={{
                    padding: "4px 12px", borderRadius: 20, fontSize: 12, cursor: "pointer",
                    border: `1px solid ${opts.barcodeScale === v ? GOLD_DARK : BORDER}`,
                    background: opts.barcodeScale === v ? GOLD_DARK : WHITE,
                    color: opts.barcodeScale === v ? WHITE : GRAY,
                    fontWeight: opts.barcodeScale === v ? 700 : 400,
                  }}>{l}</button>
                ))}
              </div>
            </div>

            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: GRAY, marginBottom: 6 }}>Text Size</div>
              <div style={{ display: "flex", gap: 6 }}>
                {[[0.8, "Small"], [1, "Normal"], [1.25, "Large"]].map(([v, l]) => (
                  <button key={v} onClick={() => setOpt("textScale", v)} style={{
                    padding: "4px 12px", borderRadius: 20, fontSize: 12, cursor: "pointer",
                    border: `1px solid ${opts.textScale === v ? GOLD_DARK : BORDER}`,
                    background: opts.textScale === v ? GOLD_DARK : WHITE,
                    color: opts.textScale === v ? WHITE : GRAY,
                    fontWeight: opts.textScale === v ? 700 : 400,
                  }}>{l}</button>
                ))}
              </div>
            </div>

            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: GRAY, marginBottom: 6 }}>Vertical Alignment</div>
              <div style={{ display: "flex", gap: 6 }}>
                {[["top", "Top"], ["center", "Center"], ["bottom", "Bottom"]].map(([v, l]) => (
                  <button key={v} onClick={() => setOpt("vAlign", v)} style={{
                    padding: "4px 12px", borderRadius: 20, fontSize: 12, cursor: "pointer",
                    border: `1px solid ${opts.vAlign === v ? GOLD_DARK : BORDER}`,
                    background: opts.vAlign === v ? GOLD_DARK : WHITE,
                    color: opts.vAlign === v ? WHITE : GRAY,
                    fontWeight: opts.vAlign === v ? 700 : 400,
                  }}>{l}</button>
                ))}
              </div>
            </div>

            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: GRAY, marginBottom: 6 }}>Print Resolution</div>
              <div style={{ display: "flex", gap: 6 }}>
                {[[200, "200 dpi"], [300, "300 dpi"], [600, "600 dpi"]].map(([v, l]) => (
                  <button key={v} onClick={() => setOpt("printDPI", v)} style={{
                    padding: "4px 12px", borderRadius: 20, fontSize: 12, cursor: "pointer",
                    border: `1px solid ${opts.printDPI === v ? GOLD_DARK : BORDER}`,
                    background: opts.printDPI === v ? GOLD_DARK : WHITE,
                    color: opts.printDPI === v ? WHITE : GRAY,
                    fontWeight: opts.printDPI === v ? 700 : 400,
                  }}>{l}</button>
                ))}
              </div>
            </div>
          </div>
        </div>

        {/* Size presets + custom */}
        <div>
          <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 6 }}>Label Size</div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 10 }}>
            {LABEL_PRESETS[labelType].map(p => (
              <button key={p.id} onClick={() => handlePreset(p.id)} style={{
                padding: "5px 11px", borderRadius: 20, fontSize: 12, cursor: "pointer",
                border: `1px solid ${presetId === p.id ? GOLD_DARK : BORDER}`,
                background: presetId === p.id ? GOLD_DARK : WHITE,
                color: presetId === p.id ? WHITE : GRAY,
                fontWeight: presetId === p.id ? 700 : 400,
              }}>{p.label}</button>
            ))}
          </div>
          <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
            {[["W (mm)", wMm, v => { setWMm(v); setPresetId("custom"); }],
            ["H (mm)", hMm, v => { setHMm(v); setPresetId("custom"); }]].map(([lbl, val, set]) => (
              <div key={lbl}>
                <div style={{ fontSize: 11, color: GRAY, marginBottom: 3 }}>{lbl}</div>
                <input type="number" value={val} min="10" max="200"
                  onChange={e => set(+e.target.value)}
                  style={{ ...inp, marginTop: 0, width: 68, padding: "6px 8px", fontSize: 13 }} />
              </div>
            ))}
          </div>
        </div>

        {/* Live preview — representative of the first product; same design applies to all */}
        <div style={{ background: "#F8F8F8", borderRadius: 10, padding: 12, border: `1px solid ${BORDER}`, textAlign: "center" }}>
          <div style={{ fontSize: 11, color: GRAY, marginBottom: 8, fontWeight: 600, textTransform: "uppercase", letterSpacing: 1 }}>
            Preview — {wMm}×{hMm}mm
            {labelType === "jewelry" && <span style={{ marginLeft: 8, color: "#92400E", background: "#FEF3C7", padding: "1px 6px", borderRadius: 10, textTransform: "none", letterSpacing: 0 }}>Dumbbell</span>}
          </div>
          <div style={{ display: "inline-block", border: "1px solid #ddd", borderRadius: 4, overflow: "hidden", boxShadow: "0 1px 4px rgba(0,0,0,0.1)" }}>
            <canvas ref={canvasRef} />
          </div>
          <div style={{ fontSize: 11, color: GRAY, marginTop: 6 }}>
            Showing {previewProduct?.name} ({previewProduct?.sku}) — this design applies to all {products.length} selected products.
          </div>
        </div>

        {layoutWarnings.length > 0 && (
          <div style={{ background: "#FEF3C7", border: "1px solid #FCD34D", borderRadius: 8, padding: "8px 12px" }}>
            {layoutWarnings.map((w, i) => <div key={i} style={{ fontSize: 12, color: "#92400E" }}>⚠ {w}</div>)}
          </div>
        )}

        {/* Print methods */}
        <div>
          <div style={{ display: "flex", flexWrap: "wrap", borderBottom: `1px solid ${BORDER}` }}>
            {tabBtn("browser", "Browser Print")}
            {tabBtn("network", "Network (Zebra)")}
            {tabBtn("zpl", "ZPL Code")}
          </div>
          <div style={{ paddingTop: 12 }}>
            {activeTab === "browser" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                <p style={{ margin: 0, fontSize: 13, color: GRAY }}>
                  Opens one print dialog sized to the exact label in mm, with all {totalLabels.toLocaleString()} labels queued back to back. Works with any USB, network, or cloud printer.
                </p>
                <button onClick={printBrowser} style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, padding: "11px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 14 }}>
                  {Icons.print} Print {totalLabels.toLocaleString()} Labels
                </button>
              </div>
            )}
            {activeTab === "network" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                <p style={{ margin: 0, fontSize: 13, color: GRAY }}>
                  {zebraIP ? `Sends one combined ZPL job to Zebra ZD421T at ${zebraIP} — ${totalLabels.toLocaleString()} labels across ${withStock.length} products.` : "No printer IP — go to Settings."}
                </p>
                <button onClick={printNetwork} style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, padding: "11px", background: "#2563EB", color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 14 }}>
                  {Icons.wifi} Send to Zebra ZD421T
                </button>
              </div>
            )}
            {activeTab === "zpl" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                <p style={{ margin: 0, fontSize: 13, color: GRAY }}>Combined ZPL for all {withStock.length} products — paste into Zebra Setup Utility or ZebraDesigner.</p>
                <textarea readOnly value={zpl} style={{ width: "100%", height: 140, fontSize: 11, fontFamily: "monospace", background: "#1a1a1a", color: "#4ADE80", padding: 12, borderRadius: 8, border: "none", boxSizing: "border-box", resize: "vertical" }} />
                <button onClick={copyZPL} style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, padding: "11px", background: "#059669", color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 14 }}>
                  {Icons.copy} Copy ZPL Code
                </button>
              </div>
            )}
          </div>
        </div>

        <StatusBanner msg={status} />
      </div>
    </Modal>
  );
}

// ── Receipt Print Modal ───────────────────────────────────────────
function ReceiptPrintModal({ sale, onClose }) {
  const [btStatus, setBtStatus] = useState(""); const [activeTab, setActiveTab] = useState("bt");
  const [headerBiz, setHeaderBiz] = useState(sale.items[0]?.business || "Blingshop");
  const printBrowser = () => {
    const rows = sale.items.map(i => `<tr><td>${i.sku || i.name}</td><td align="center">${i.qty}</td><td align="right">LKR ${(i.price * i.qty).toLocaleString()}</td></tr>`).join("");
    const win = window.open("", "_blank", "width=1,height=1,left=-1000,top=-1000");
    if (!win) return;
    win.document.write(`<html><head><title>Receipt ${sale.id}</title><style>@page{size:80mm auto;margin:4mm}body{font-family:monospace;width:72mm;font-size:12px}h2,p{text-align:center;margin:2px 0}hr{border:none;border-top:1px dashed #999;margin:5px 0}table{width:100%}th{font-size:10px;text-align:left}td{font-size:11px;padding:1px 0}.total{font-size:14px;font-weight:bold;text-align:right}</style></head><body>
      <h2>${headerBiz.toUpperCase()}</h2>
      <p style="font-size:10px;font-weight:bold">RETURN POLICY</p>
      <p style="font-size:10px">${RETURN_POLICY[headerBiz] || ""}</p><hr/>
      <p style="font-size:10px">${new Date(sale.date).toLocaleString()}</p><p style="font-size:10px">Receipt #${sale.id}</p>
      ${sale.staffName ? `<p style="font-size:10px">Served by: ${sale.staffName}</p>` : ""}
      ${sale.customerName ? `<p style="font-size:10px">Customer: ${sale.customerName}</p>` : ""}<hr/>
      <table><tr><th>Item</th><th align="center">Qty</th><th align="right">Amt</th></tr>${rows}</table><hr/>
      ${sale.discount > 0 ? `<p style="font-size:11px;text-align:right">Subtotal: LKR ${sale.subtotal.toLocaleString()}</p><p style="font-size:11px;text-align:right">Discount: -LKR ${sale.discount.toLocaleString()}</p>` : ""}
      ${sale.deliveryMethod && sale.deliveryMethod !== "In Store" ? `<p style="font-size:11px;text-align:right">Delivery (${sale.deliveryMethod}): ${sale.deliveryPaidTo === "Shop" ? `LKR ${(+sale.deliveryFee || 0).toLocaleString()}` : "Paid to rider"}</p>` : ""}
      <p class="total">Total: LKR ${sale.total.toLocaleString()}</p><hr/>
      <p style="font-size:10px;text-align:center">Payment: ${paymentSummary(sale)}</p><hr/>
      <p style="font-size:10px;text-align:center">Thank you for shopping!</p>
    </body></html>`);
    win.document.close();
    win.focus();
    setTimeout(() => { win.print(); win.close(); }, 350);
  };
  const tabBtn = id => ({ padding: "8px 16px", border: "none", cursor: "pointer", fontWeight: activeTab === id ? 700 : 400, fontSize: 13, background: activeTab === id ? WHITE : LIGHT, color: activeTab === id ? GOLD_DARK : GRAY, borderBottom: activeTab === id ? `2px solid ${GOLD}` : "2px solid transparent" });
  return (
    <Modal title="Print Receipt — ABM P323B" onClose={onClose}>
      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
        <div>
          <label style={labelStyle}>Receipt Heading</label>
          <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
            {BUSINESSES.map(b => (
              <button key={b} onClick={() => setHeaderBiz(b)} style={{ flex: 1, padding: "10px", borderRadius: 8, border: `2px solid ${headerBiz === b ? (b === "Blingshop" ? GOLD : RC_BLUE) : BORDER}`, background: headerBiz === b ? (b === "Blingshop" ? GOLD_LIGHT : RC_LIGHT) : WHITE, fontWeight: headerBiz === b ? 700 : 400, color: headerBiz === b ? (b === "Blingshop" ? GOLD_DARK : RC_DARK) : GRAY, cursor: "pointer", fontSize: 13 }}>{b}</button>
            ))}
          </div>
          <div style={{ fontSize: 11, color: GRAY, marginTop: 4 }}>{RETURN_POLICY[headerBiz]}</div>
        </div>
        <div style={{ borderBottom: `1px solid ${BORDER}`, display: "flex", flexWrap: "wrap" }}>
          <button style={tabBtn("bt")} onClick={() => setActiveTab("bt")}>Bluetooth</button>
          <button style={tabBtn("usb")} onClick={() => setActiveTab("usb")}>USB / Browser</button>
        </div>
        {activeTab === "bt" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <div style={{ background: LIGHT, borderRadius: 8, padding: 14, border: `1px solid ${BORDER}` }}>
              <div style={{ fontWeight: 700, color: "#1F2937", fontSize: 13, marginBottom: 8 }}>Bluetooth Direct Print</div>
              <ol style={{ margin: 0, paddingLeft: 18, fontSize: 13, color: GRAY, lineHeight: 2 }}>
                <li>Ensure ABM P323B is powered on</li><li>Enable Bluetooth on this device</li>
                <li>Click <strong>Connect &amp; Print</strong> below</li><li>Select <strong>ABM-P323B</strong> from the list</li>
              </ol>
              <div style={{ fontSize: 11, color: "#9CA3AF", marginTop: 8, borderTop: `1px solid ${BORDER}`, paddingTop: 8 }}>Requires Chrome on Android or Chrome desktop.</div>
            </div>
            <button onClick={() => printReceiptBluetooth(sale, headerBiz, setBtStatus)} style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, padding: "12px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 14 }}>{Icons.wifi} Connect &amp; Print via Bluetooth</button>
            <StatusBanner msg={btStatus} />
          </div>
        )}
        {activeTab === "usb" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <div style={{ background: LIGHT, borderRadius: 8, padding: 14, border: `1px solid ${BORDER}` }}>
              <ol style={{ margin: 0, paddingLeft: 18, fontSize: 13, color: GRAY, lineHeight: 2 }}>
                <li>Connect ABM P323B via USB</li><li>Set as default printer in Windows</li><li>Click Print Receipt below</li>
              </ol>
            </div>
            <button onClick={printBrowser} style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, padding: "12px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 14 }}>{Icons.print} Print Receipt</button>
          </div>
        )}
      </div>
    </Modal>
  );
}

// ── Stat Card ─────────────────────────────────────────────────────
function StatCard({ label, value, sub, accent }) {
  return (
    <div style={{ background: WHITE, borderRadius: 12, padding: 16, boxShadow: "0 1px 4px rgba(0,0,0,0.06)", borderTop: `3px solid ${accent}` }}>
      <div style={{ fontSize: 11, color: GRAY, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 6 }}>{label}</div>
      <div style={{ fontSize: 20, fontWeight: 800, color: "#1F2937" }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: GRAY, marginTop: 3 }}>{sub}</div>}
    </div>
  );
}

// ── Dashboard ─────────────────────────────────────────────────────
function Dashboard({ products, sales, currentStaff }) {
  const [bizView, setBizView] = useState("All");
  const today = new Date().toDateString();

  const filterSales = (s) => bizView === "All" ? s : s.filter(sale => sale.items.some(i => (i.business || "Blingshop") === bizView));
  const filterProducts = (p) => bizView === "All" ? p : p.filter(x => (x.business || "Blingshop") === bizView);

  const fSales = filterSales(sales);
  const fProducts = filterProducts(products);
  const tSales = filterSales(sales.filter(s => new Date(s.date).toDateString() === today));

  const calcRev = (saleList) => bizView === "All"
    ? saleList.reduce((s, x) => s + parseFloat(x.total || 0), 0)
    : saleList.reduce((s, sale) => s + sale.items.filter(i => (i.business || "Blingshop") === bizView).reduce((a, i) => a + parseFloat(i.price || 0) * parseFloat(i.qty || 0), 0), 0);

  const tRev = calcRev(tSales);
  const totRev = calcRev(fSales);
  const outOf = fProducts.filter(p => p.stock === 0);

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div>
          <h2 style={{ color: "#1F2937", margin: "0 0 2px", fontSize: 20, fontWeight: 800 }}>Dashboard</h2>
          <p style={{ margin: 0, color: GRAY, fontSize: 13 }}>{new Date().toLocaleDateString("en-US", { weekday: "long", year: "numeric", month: "long", day: "numeric" })}</p>
        </div>
      </div>
      <div style={{ marginBottom: 16 }}><BizFilter value={bizView} onChange={setBizView} /></div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 20 }}>
        <StatCard label="Today's Sales" value={tSales.length} sub="transactions" accent={GOLD} />
        <StatCard label="Today's Revenue" value={`LKR ${tRev.toLocaleString()}`} sub={bizView === "All" ? "Combined" : bizView} accent={GOLD_DARK} />
        {currentStaff?.role === "Owner" && <StatCard label="Total Revenue" value={`LKR ${totRev.toLocaleString()}`} sub="All time" accent="#5C4A1E" />}
        <StatCard label="Products" value={fProducts.length} sub={bizView === "All" ? "Both businesses" : bizView} accent="#7C6A2A" />
      </div>
      {outOf.length > 0 && (
        <div style={{ background: WHITE, border: "1px solid #FCA5A5", borderLeft: "4px solid #DC2626", borderRadius: 10, padding: 16 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12, color: "#991B1B" }}>{Icons.alert}<span style={{ fontWeight: 700, fontSize: 14 }}>Out of Stock</span></div>
          {outOf.map(p => (
            <div key={p.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: "1px solid #FEE2E2", fontSize: 13 }}>
              <div><span style={{ color: "#374151" }}>{p.name}</span><span style={{ marginLeft: 8 }}><BizBadge business={p.business || "Blingshop"} /></span></div>
              <span style={{ color: "#DC2626", fontWeight: 700, fontSize: 12, background: "#FEE2E2", padding: "2px 8px", borderRadius: 10 }}>Out of Stock</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ── Inventory ─────────────────────────────────────────────────────
function Inventory({ products, setProducts, zebraIP, currentStaff, blingshopThbRate, setBlingshopThbRate, rcThbRate, setRcThbRate }) {
  const [modal, setModal] = useState(null); const [editProd, setEditProd] = useState(null); const [barcodeProduct, setBarcodeProduct] = useState(null);
  const [catFilter, setCatFilter] = useState("All"); const [bizFilter, setBizFilter] = useState("All"); const [search, setSearch] = useState(""); const [showImport, setShowImport] = useState(false);
  const handleBizFilter = b => { setBizFilter(b); if (!categoriesForBiz(b).includes(catFilter)) setCatFilter("All"); };
  const [showBulkAdd, setShowBulkAdd] = useState(false); const [showReviewPending, setShowReviewPending] = useState(false);
  const [pendingProducts, setPendingProducts] = useState([]);
  const [selectedIds, setSelectedIds] = useState({}); const [showBulkPrint, setShowBulkPrint] = useState(false);

  useEffect(() => {
    (async () => { try { setPendingProducts(await apiFetch("/pending-products")); } catch { } })();
  }, []);

  const handleBulkAddSave = async (item) => {
    const withStaff = { ...item, staffId: currentStaff?.id || null, staffName: currentStaff?.name || null };
    try { await apiFetch("/pending-products", { method: "POST", body: withStaff }); } catch { }
    setPendingProducts(ps => [withStaff, ...ps]);
  };

  const filtered = products.filter(p =>
    (bizFilter === "All" || (p.business || "Blingshop") === bizFilter) &&
    (catFilter === "All" || p.category === catFilter) &&
    (p.name.toLowerCase().includes(search.toLowerCase()) || (p.sku || "").toLowerCase().includes(search.toLowerCase()))
  );
  const selectedProducts = filtered.filter(p => selectedIds[p.id]);
  const selectedCount = selectedProducts.length;
  const toggleSelectAll = () => {
    const allSelected = filtered.length > 0 && filtered.every(p => selectedIds[p.id]);
    setSelectedIds(s => {
      const next = { ...s };
      filtered.forEach(p => { next[p.id] = !allSelected; });
      return next;
    });
  };

  const handleSave = async (form) => {
    if (editProd) {
      const updated = { ...editProd, ...form, price: +form.price, cost: +(form.cost || 0), stock: +form.stock };
      try { await apiFetch(`/products/${editProd.id}`, { method: "PUT", body: updated }); } catch { }
      setProducts(ps => ps.map(p => p.id === editProd.id ? updated : p));
    } else {
      const id = generateId();
      const newProd = { ...form, id, sku: form.sku || generateSKU(form.category, form.business, products), price: +form.price, cost: +(form.cost || 0), stock: +form.stock };
      try { await apiFetch("/products", { method: "POST", body: newProd }); } catch { }
      setProducts(ps => [...ps, newProd]);
    }
    setModal(null); setEditProd(null);
  };

  const handleDelete = async (id) => {
    if (!confirm("Delete this product?")) return;
    try { await apiFetch(`/products/${id}`, { method: "DELETE" }); } catch { }
    setProducts(ps => ps.filter(p => p.id !== id));
  };

  const handleImportConfirm = async (importRows) => {
    const newProducts = []; const runningList = [...products];
    for (const row of importRows) {
      const id = generateId();
      const sku = row.sku?.trim() ? row.sku.trim() : generateSKU(row.category, row.business, runningList);
      const newProd = { id, name: row.name, category: row.category, business: row.business, price: row.price, cost: row.cost || 0, stock: row.stock, sku, image: row.image || "" };
      try { await apiFetch("/products", { method: "POST", body: newProd }); } catch { }
      newProducts.push(newProd); runningList.push(newProd);
    }
    setProducts(ps => [...ps, ...newProducts]); setShowImport(false);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div><h2 style={{ color: "#1F2937", margin: "0 0 2px", fontSize: 20, fontWeight: 800 }}>Inventory</h2><p style={{ margin: 0, color: GRAY, fontSize: 13 }}>{products.length} products total</p></div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {currentStaff?.role === "Owner" && pendingProducts.length > 0 && (
            <button onClick={() => setShowReviewPending(true)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: "#FEF3C7", color: "#92400E", border: "1px solid #FCD34D", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 13 }}>Review Pending ({pendingProducts.length})</button>
          )}
          {selectedCount > 0 && (
            <button onClick={() => setShowBulkPrint(true)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: "#F0FDF4", color: "#059669", border: "1px solid #86EFAC", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 13 }}>{Icons.barcode} Print Labels ({selectedCount})</button>
          )}
          <button onClick={() => setShowBulkAdd(true)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: WHITE, color: GOLD_DARK, border: `1px solid ${BORDER}`, borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 13 }}>{Icons.camera} Bulk Add</button>
          <button onClick={() => setShowImport(true)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: WHITE, color: GOLD_DARK, border: `1px solid ${BORDER}`, borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 13 }}>{Icons.upload} Import Excel</button>
          <button onClick={() => { setEditProd(null); setModal("add"); }} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 13 }}>{Icons.plus} Add Product</button>
        </div>
      </div>
      <div style={{ marginBottom: 10 }}><BizFilter value={bizFilter} onChange={handleBizFilter} /></div>
      <div style={{ position: "relative", marginBottom: 10 }}>
        <div style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", color: GRAY }}>{Icons.search}</div>
        <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search by name or SKU..." style={{ ...inp, paddingLeft: 38, marginTop: 0 }} />
      </div>
      <div style={{ display: "flex", gap: 6, marginBottom: 14, flexWrap: "wrap" }}>
        {categoriesForBiz(bizFilter).map(c => (<button key={c} onClick={() => setCatFilter(c)} style={{ padding: "5px 14px", borderRadius: 20, border: `1px solid ${catFilter === c ? GOLD : BORDER}`, background: catFilter === c ? GOLD_DARK : WHITE, color: catFilter === c ? WHITE : GRAY, cursor: "pointer", fontSize: 12, fontWeight: catFilter === c ? 700 : 400 }}>{c}</button>))}
      </div>
      {filtered.length > 0 && (
        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: "#374151", cursor: "pointer", marginBottom: 8 }}>
          <input type="checkbox" checked={filtered.every(p => selectedIds[p.id])} onChange={toggleSelectAll} style={{ width: 15, height: 15, accentColor: GOLD_DARK, cursor: "pointer" }} />
          Select all {catFilter !== "All" ? `${catFilter} ` : ""}({filtered.length})
        </label>
      )}
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {filtered.map(p => (
          <div key={p.id} style={{ background: WHITE, borderRadius: 10, padding: 12, display: "flex", gap: 12, alignItems: "center", boxShadow: "0 1px 4px rgba(0,0,0,0.06)", border: `1px solid ${BORDER}` }}>
            <input type="checkbox" checked={!!selectedIds[p.id]} onChange={() => setSelectedIds(s => ({ ...s, [p.id]: !s[p.id] }))} style={{ width: 15, height: 15, flexShrink: 0, accentColor: GOLD_DARK, cursor: "pointer" }} />
            <div style={{ width: 56, height: 56, borderRadius: 8, overflow: "hidden", background: LIGHT, flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center", border: `1px solid ${BORDER}`, color: GRAY }}>
              {p.image ? <img src={p.image} alt={p.name} style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : Icons.image}
            </div>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 2 }}>
                <span style={{ fontWeight: 700, fontSize: 14, color: "#1F2937" }}>{p.name}</span>
                <BizBadge business={p.business || "Blingshop"} />
              </div>
              <div style={{ fontSize: 12, color: GRAY }}>{p.category} &nbsp;&middot;&nbsp; SKU: {p.sku}</div>
              <div style={{ display: "flex", gap: 10, marginTop: 4, alignItems: "center" }}>
                <span style={{ color: GOLD_DARK, fontWeight: 800, fontSize: 14 }}>LKR {(+p.price).toLocaleString()}</span>
                <span style={{ fontSize: 12, fontWeight: 700, padding: "2px 8px", borderRadius: 10, color: p.stock === 0 ? "#DC2626" : p.stock <= 3 ? "#B45309" : "#059669", background: p.stock === 0 ? "#FEE2E2" : p.stock <= 3 ? "#FEF3C7" : "#F0FDF4" }}>
                  {p.stock === 0 ? "Out of Stock" : `${p.stock} in stock`}
                </span>
              </div>
            </div>
            <div style={{ display: "flex", gap: 6 }}>
              {[{ icon: Icons.edit, bg: "#F3F4F6", color: "#374151", fn: () => { setEditProd(p); setModal("add"); } }, { icon: Icons.barcode, bg: "#F0FDF4", color: "#059669", fn: () => setBarcodeProduct(p) }, { icon: Icons.trash, bg: "#FEF2F2", color: "#DC2626", fn: () => handleDelete(p.id) }].map((btn, i) => (
                <button key={i} onClick={btn.fn} style={{ width: 32, height: 32, display: "flex", alignItems: "center", justifyContent: "center", background: btn.bg, color: btn.color, border: "none", borderRadius: 6, cursor: "pointer" }}>{btn.icon}</button>
              ))}
            </div>
          </div>
        ))}
        {filtered.length === 0 && <div style={{ textAlign: "center", color: "#9CA3AF", padding: "40px 0", fontSize: 14 }}>No products found</div>}
      </div>
      {modal === "add" && <Modal title={editProd ? "Edit Product" : "Add New Product"} onClose={() => { setModal(null); setEditProd(null); }}><ProductForm initial={editProd} onSave={handleSave} onCancel={() => { setModal(null); setEditProd(null); }} existingProducts={products} /></Modal>}
      {barcodeProduct && <LabelPrintModal product={barcodeProduct} zebraIP={zebraIP} onClose={() => setBarcodeProduct(null)} />}
      {showImport && <ImportModal existingProducts={products} onImport={handleImportConfirm} onClose={() => setShowImport(false)} />}
      {showBulkAdd && <BulkAddModal existingProducts={products} pendingProducts={pendingProducts} onSave={handleBulkAddSave} onClose={() => setShowBulkAdd(false)} blingshopThbRate={blingshopThbRate} setBlingshopThbRate={setBlingshopThbRate} rcThbRate={rcThbRate} setRcThbRate={setRcThbRate} />}
      {showReviewPending && <ReviewPendingModal pendingProducts={pendingProducts} setPendingProducts={setPendingProducts} setProducts={setProducts} onClose={() => setShowReviewPending(false)} />}
      {showBulkPrint && <BulkLabelPrintModal products={selectedProducts} zebraIP={zebraIP} onClose={() => setShowBulkPrint(false)} />}
    </div>
  );
}

// ── CRM Helpers ───────────────────────────────────────────────────
function getBirthdayStatus(birthday) {
  if (!birthday) return null;
  const today = new Date();
  const bday = new Date(birthday);
  const thisYear = new Date(today.getFullYear(), bday.getMonth(), bday.getDate());
  const diff = Math.ceil((thisYear - today) / (1000 * 60 * 60 * 24));
  if (diff === 0) return { label: "Birthday today!", color: "#DC2626", bg: "#FEE2E2" };
  if (diff > 0 && diff <= 7) return { label: `Birthday in ${diff} day${diff === 1 ? "" : "s"}`, color: "#D97706", bg: "#FEF3C7" };
  if (diff < 0 && diff >= -3) return { label: "Birthday was recent", color: "#059669", bg: "#F0FDF4" };
  return null;
}

function getCustomerStats(customerId, sales) {
  const custSales = sales.filter(s => s.customerId === customerId);
  const totalSpend = custSales.reduce((s, x) => s + x.total, 0);
  const lastVisit = custSales.length ? new Date(Math.max(...custSales.map(s => new Date(s.date)))) : null;
  const daysSince = lastVisit ? Math.floor((new Date() - lastVisit) / (1000 * 60 * 60 * 24)) : null;
  return { totalSpend, lastVisit, totalSales: custSales.length, daysSince };
}

function CustomerForm({ initial, onSave, onCancel }) {
  const [form, setForm] = useState(initial || { name: "", phone: "", email: "", birthday: "", notes: "", preferredBiz: "Both" });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div><label style={labelStyle}>Full Name *</label><input value={form.name} onChange={e => set("name", e.target.value)} placeholder="Customer name" style={inp} /></div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <div><label style={labelStyle}>Phone</label><input value={form.phone} onChange={e => set("phone", e.target.value)} placeholder="+94 77..." style={inp} /></div>
        <div><label style={labelStyle}>Birthday</label><input type="date" value={form.birthday} onChange={e => set("birthday", e.target.value)} style={inp} /></div>
      </div>
      <div><label style={labelStyle}>Email</label><input value={form.email} onChange={e => set("email", e.target.value)} placeholder="email@example.com" style={inp} /></div>
      <div>
        <label style={labelStyle}>Usually shops at</label>
        <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
          {["Blingshop", "RC Boutique", "Both"].map(b => (
            <button key={b} type="button" onClick={() => set("preferredBiz", b)} style={{ flex: 1, padding: "9px", borderRadius: 8, border: `1.5px solid ${form.preferredBiz === b ? GOLD : BORDER}`, background: form.preferredBiz === b ? GOLD_LIGHT : WHITE, color: form.preferredBiz === b ? GOLD_DARK : GRAY, fontWeight: form.preferredBiz === b ? 700 : 400, cursor: "pointer", fontSize: 13 }}>{b}</button>
          ))}
        </div>
      </div>
      <div><label style={labelStyle}>Notes</label><textarea value={form.notes} onChange={e => set("notes", e.target.value)} placeholder="Any notes about this customer..." style={{ ...inp, resize: "vertical", minHeight: 72 }} /></div>
      <div style={{ display: "flex", gap: 10, paddingTop: 4, borderTop: `1px solid ${GOLD_LIGHT}`, marginTop: 4 }}>
        <button type="button" onClick={() => onSave(form)} style={{ flex: 1, padding: "11px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 14 }}>Save Customer</button>
        <button type="button" onClick={onCancel} style={{ padding: "11px 20px", background: LIGHT, color: GRAY, border: `1px solid #E5E7EB`, borderRadius: 8, cursor: "pointer", fontSize: 14 }}>Cancel</button>
      </div>
    </div>
  );
}

// ── POS ───────────────────────────────────────────────────────────
function POS({ products, setProducts, addSale, customers, setCustomers, exchangeCredit, clearExchangeCredit, currentStaff }) {
  const [cart, setCart] = useState([]);
  const [search, setSearch] = useState("");
  const [catFilter, setCatFilter] = useState("All");
  const [bizFilter, setBizFilter] = useState("All");
  const handleBizFilter = b => { setBizFilter(b); if (!categoriesForBiz(b).includes(catFilter)) setCatFilter("All"); };
  const [discount, setDiscount] = useState(0);
  const [deliveryMethod, setDeliveryMethod] = useState("In Store");
  const [deliveryFee, setDeliveryFee] = useState("");
  const [deliveryPaidTo, setDeliveryPaidTo] = useState("Shop");
  const [paymentMethod, setPaymentMethod] = useState("Cash");
  const [splitCard, setSplitCard] = useState("");
  const [splitCash, setSplitCash] = useState("");
  const [printSale, setPrintSale] = useState(null);
  const [scanning, setScanning] = useState(false);
  const [successSale, setSuccessSale] = useState(null);
  const [custSearch, setCustSearch] = useState("");
  const [selectedCustomer, setSelectedCustomer] = useState(null);
  const [showCustSearch, setShowCustSearch] = useState(false);
  const [quickAdd, setQuickAdd] = useState(null);
  const [isWide, setIsWide] = useState(window.innerWidth >= 900);

  useEffect(() => {
    const onResize = () => setIsWide(window.innerWidth >= 900);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  // Pre-select customer from exchange credit on mount
  useEffect(() => {
    if (exchangeCredit?.customerId) {
      const cust = customers.find(c => c.id === exchangeCredit.customerId);
      if (cust) setSelectedCustomer(cust);
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const filtered = products.filter(p =>
    (bizFilter === "All" || (p.business || "Blingshop") === bizFilter) &&
    (catFilter === "All" || p.category === catFilter) &&
    (p.name.toLowerCase().includes(search.toLowerCase()) || (p.sku || "").toLowerCase().includes(search.toLowerCase())) &&
    p.stock > 0
  );

  const addToCart = p => { setCart(c => { const ex = c.find(x => x.id === p.id); if (ex) { if (ex.qty >= p.stock) return c; return c.map(x => x.id === p.id ? { ...x, qty: x.qty + 1 } : x); } return [...c, { ...p, qty: 1, business: p.business || "Blingshop" }]; }); };
  const handleScan = sku => { setScanning(false); const p = products.find(x => x.sku === sku || x.id === sku); if (p && p.stock > 0) { playBeep(); addToCart(p); } else alert(`Product not found: ${sku}`); };
  const updateQty = (id, qty) => { if (qty < 1) { setCart(c => c.filter(x => x.id !== id)); return; } const p = products.find(x => x.id === id); if (qty > p.stock) return; setCart(c => c.map(x => x.id === id ? { ...x, qty } : x)); };
  const subtotal = cart.reduce((s, x) => s + x.price * x.qty, 0);
  const discountAmt = Math.round(subtotal * discount / 100);
  const creditAmt = Math.min(exchangeCredit?.amount || 0, Math.max(0, subtotal - discountAmt));
  // Only add the delivery fee to what we collect from the customer when it's paid to us — if
  // they pay the rider directly, that cash never touches the business, so it's tracked but not billed.
  const isDelivery = deliveryMethod !== "In Store";
  const deliveryFeeToShop = isDelivery && deliveryPaidTo === "Shop" ? (+deliveryFee || 0) : 0;
  const total = subtotal - discountAmt - creditAmt + deliveryFeeToShop;
  const cartByBiz = BUSINESSES.reduce((acc, b) => { const items = cart.filter(i => (i.business || "Blingshop") === b); if (items.length) acc[b] = items; return acc; }, {});
  const matches = customers.filter(c => c.name.toLowerCase().includes(custSearch.toLowerCase()) || (c.phone || "").includes(custSearch)).slice(0, 6);

  const saveQuickAdd = async () => {
    if (!quickAdd.name.trim()) return;
    const id = generateId();
    const newCust = { id, name: quickAdd.name.trim(), phone: quickAdd.phone.trim(), email: "", birthday: "", preferredBiz: "Both", notes: "", createdAt: new Date().toISOString() };
    try { await apiFetch("/customers", { method: "POST", body: newCust }); } catch { }
    setCustomers(cs => [...cs, newCust]);
    setSelectedCustomer(newCust); setQuickAdd(null); setCustSearch(""); setShowCustSearch(false);
  };

  const checkout = () => {
    if (!cart.length) return;
    let cashAmount = 0, cardAmount = 0, bankAmount = 0;
    if (paymentMethod === "Cash") cashAmount = total;
    else if (paymentMethod === "Card") cardAmount = total;
    else if (paymentMethod === "Bank Transfer") bankAmount = total;
    else if (paymentMethod === "Split") {
      cardAmount = +splitCard || 0; cashAmount = +splitCash || 0;
      if (Math.abs((cardAmount + cashAmount) - total) > 0.5) { alert(`Card + Cash must add up to the total (LKR ${total.toLocaleString()}).`); return; }
    }
    const sale = { id: generateId(), date: new Date().toISOString(), items: cart, subtotal, discount: discountAmt + creditAmt, total, customerId: selectedCustomer?.id || null, customerName: selectedCustomer?.name || null, staffId: currentStaff?.id || null, staffName: currentStaff?.name || null, paymentMethod, cashAmount, cardAmount, bankAmount, deliveryMethod, deliveryFee: isDelivery ? (+deliveryFee || 0) : 0, deliveryPaidTo: isDelivery ? deliveryPaidTo : null };
    setProducts(ps => ps.map(p => { const ci = cart.find(x => x.id === p.id); return ci ? { ...p, stock: p.stock - ci.qty } : p; }));
    addSale(sale); setSuccessSale(sale); setCart([]); setDiscount(0); setSelectedCustomer(null); setCustSearch("");
    setPaymentMethod("Cash"); setSplitCard(""); setSplitCash("");
    setDeliveryMethod("In Store"); setDeliveryFee(""); setDeliveryPaidTo("Shop");
    if (clearExchangeCredit) clearExchangeCredit();
  };

  const cartPanel = (
    <div style={{ background: WHITE, borderRadius: 12, border: `1px solid ${BORDER}`, overflow: "hidden", boxShadow: "0 2px 8px rgba(0,0,0,0.06)" }}>
      <div style={{ padding: "12px 16px", background: LIGHT, borderBottom: `1px solid ${BORDER}`, display: "flex", alignItems: "center", gap: 6, color: "#374151", fontWeight: 700, fontSize: 14 }}>
        {Icons.cart} Cart &nbsp;<span style={{ background: cart.length > 0 ? GOLD_DARK : "#D1D5DB", color: WHITE, borderRadius: 10, padding: "1px 8px", fontSize: 12 }}>{cart.length}</span>
      </div>
      {exchangeCredit && (
        <div style={{ padding: "8px 16px", background: "#F0FDF4", borderBottom: "1px solid #86EFAC", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontSize: 12, fontWeight: 700, color: "#166534" }}>Exchange Credit</div>
            {exchangeCredit.customerName && <div style={{ fontSize: 11, color: "#15803D" }}>{exchangeCredit.customerName}</div>}
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontWeight: 800, color: "#166534", fontSize: 14 }}>LKR {exchangeCredit.amount.toLocaleString()}</span>
            <button onClick={clearExchangeCredit} style={{ background: "none", border: "none", cursor: "pointer", color: GRAY, fontSize: 16, lineHeight: 1 }}>{Icons.close}</button>
          </div>
        </div>
      )}
      {cart.length === 0 ? (
        <div style={{ padding: "40px 20px", textAlign: "center", color: "#9CA3AF", fontSize: 13 }}>No items yet<br />Tap a product to add</div>
      ) : (
        <div style={{ padding: "12px 16px", maxHeight: isWide ? "calc(100vh - 180px)" : "none", overflowY: isWide ? "auto" : "visible" }}>
          {Object.entries(cartByBiz).map(([biz, items]) => (
            <div key={biz} style={{ marginBottom: 12 }}>
              <div style={{ marginBottom: 6 }}><BizBadge business={biz} size="md" /></div>
              {items.map(item => (
                <div key={item.id} style={{ display: "flex", alignItems: "center", gap: 10, paddingBottom: 8, marginBottom: 8, borderBottom: `1px solid ${GOLD_LIGHT}` }}>
                  <div style={{ flex: 1, fontSize: 13, color: "#1F2937", fontWeight: 500 }}>{item.name}</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 1 }}>
                    <button onClick={() => updateQty(item.id, item.qty - 1)} style={{ width: 28, height: 28, borderRadius: "6px 0 0 6px", border: `1px solid ${BORDER}`, background: LIGHT, cursor: "pointer", fontWeight: 700, fontSize: 16 }}>−</button>
                    <div style={{ width: 34, height: 28, display: "flex", alignItems: "center", justifyContent: "center", border: `1px solid ${BORDER}`, borderLeft: "none", borderRight: "none", fontSize: 13, fontWeight: 700 }}>{item.qty}</div>
                    <button onClick={() => updateQty(item.id, item.qty + 1)} style={{ width: 28, height: 28, borderRadius: "0 6px 6px 0", border: `1px solid ${BORDER}`, background: LIGHT, cursor: "pointer", fontWeight: 700, fontSize: 16 }}>+</button>
                  </div>
                  <div style={{ minWidth: 82, textAlign: "right", fontSize: 13, fontWeight: 700, color: GOLD_DARK }}>LKR {(item.price * item.qty).toLocaleString()}</div>
                </div>
              ))}
            </div>
          ))}

          <div style={{ marginBottom: 12, paddingBottom: 12, borderBottom: `1px solid ${GOLD_LIGHT}` }}>
            <div style={{ fontSize: 13, color: GRAY, fontWeight: 500, marginBottom: 6 }}>Customer (optional)</div>
            {selectedCustomer ? (
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: GOLD_LIGHT, borderRadius: 8, padding: "8px 12px", border: `1px solid ${GOLD}` }}>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 13, color: GOLD_DARK }}>{selectedCustomer.name}</div>
                  {selectedCustomer.phone && <div style={{ fontSize: 11, color: GRAY }}>{selectedCustomer.phone}</div>}
                </div>
                <button type="button" onClick={() => setSelectedCustomer(null)} style={{ background: "none", border: "none", cursor: "pointer", color: GRAY }}>{Icons.close}</button>
              </div>
            ) : (
              <div style={{ position: "relative" }}>
                <div style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: GRAY }}>{Icons.search}</div>
                <input value={custSearch} onChange={e => { setCustSearch(e.target.value); setShowCustSearch(true); setQuickAdd(null); }} onFocus={() => setShowCustSearch(true)} placeholder="Search customer name or phone..." style={{ ...inp, marginTop: 0, paddingLeft: 34, fontSize: 13 }} />
                {showCustSearch && custSearch && !quickAdd && (
                  <div style={{ position: "absolute", top: "100%", left: 0, right: 0, background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 8, boxShadow: "0 4px 12px rgba(0,0,0,0.1)", zIndex: 100, maxHeight: 220, overflowY: "auto" }}>
                    {matches.map(c => (
                      <div key={c.id} style={{ padding: "10px 14px", cursor: "pointer", borderBottom: `1px solid ${BORDER}`, fontSize: 13 }} onClick={() => { setSelectedCustomer(c); setCustSearch(""); setShowCustSearch(false); }}>
                        <div style={{ fontWeight: 600, color: "#1F2937" }}>{c.name}</div>
                        {c.phone && <div style={{ fontSize: 11, color: GRAY }}>{c.phone}</div>}
                      </div>
                    ))}
                    <div style={{ padding: "10px 14px" }}>
                      {matches.length === 0 && <div style={{ fontSize: 13, color: GRAY, marginBottom: 8 }}>No customer found for "{custSearch}"</div>}
                      <button type="button" onClick={() => setQuickAdd({ name: custSearch, phone: "" })} style={{ width: "100%", padding: "8px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 6, cursor: "pointer", fontSize: 12, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}>{Icons.plus} Add "{custSearch}" as new customer</button>
                    </div>
                  </div>
                )}
              </div>
            )}
            {quickAdd && (
              <div style={{ marginTop: 8, background: LIGHT, border: `1px solid ${BORDER}`, borderRadius: 8, padding: 12 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 8 }}>Add New Customer</div>
                <input value={quickAdd.name} onChange={e => setQuickAdd(q => ({ ...q, name: e.target.value }))} placeholder="Full name" style={{ ...inp, marginTop: 0, marginBottom: 8, fontSize: 13 }} />
                <input value={quickAdd.phone} onChange={e => setQuickAdd(q => ({ ...q, phone: e.target.value }))} placeholder="Phone (optional)" style={{ ...inp, marginTop: 0, marginBottom: 8, fontSize: 13 }} />
                <div style={{ display: "flex", gap: 8 }}>
                  <button type="button" onClick={saveQuickAdd} style={{ flex: 1, padding: "8px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 6, cursor: "pointer", fontSize: 12, fontWeight: 700 }}>Save &amp; Select</button>
                  <button type="button" onClick={() => setQuickAdd(null)} style={{ padding: "8px 14px", background: WHITE, color: GRAY, border: `1px solid ${BORDER}`, borderRadius: 6, cursor: "pointer", fontSize: 12 }}>Cancel</button>
                </div>
              </div>
            )}
          </div>

          <div style={{ marginBottom: 12, paddingBottom: 12, borderBottom: `1px solid ${GOLD_LIGHT}` }}>
            <div style={{ fontSize: 13, color: GRAY, fontWeight: 500, marginBottom: 6 }}>Delivery</div>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {DELIVERY_METHODS.map(m => (
                <button key={m} onClick={() => setDeliveryMethod(m)} style={{ flex: "1 1 auto", padding: "7px 10px", borderRadius: 8, border: `1.5px solid ${deliveryMethod === m ? GOLD_DARK : BORDER}`, background: deliveryMethod === m ? GOLD_LIGHT : WHITE, color: deliveryMethod === m ? GOLD_DARK : GRAY, fontWeight: deliveryMethod === m ? 700 : 400, fontSize: 12, cursor: "pointer" }}>{m}</button>
              ))}
            </div>
            {isDelivery && (
              <div style={{ marginTop: 8 }}>
                <label style={{ fontSize: 11, color: GRAY }}>Delivery Fee (LKR)</label>
                <input type="number" min="0" value={deliveryFee} onChange={e => setDeliveryFee(e.target.value)} placeholder="0" style={{ ...inp, marginTop: 2, fontSize: 13 }} />
                <div style={{ display: "flex", gap: 6, marginTop: 8, flexWrap: "wrap" }}>
                  {DELIVERY_PAID_TO.map(o => (
                    <button key={o.id} onClick={() => setDeliveryPaidTo(o.id)} style={{ flex: "1 1 auto", padding: "7px 10px", borderRadius: 8, border: `1.5px solid ${deliveryPaidTo === o.id ? GOLD_DARK : BORDER}`, background: deliveryPaidTo === o.id ? GOLD_LIGHT : WHITE, color: deliveryPaidTo === o.id ? GOLD_DARK : GRAY, fontWeight: deliveryPaidTo === o.id ? 700 : 400, fontSize: 11, cursor: "pointer" }}>{o.label}</button>
                  ))}
                </div>
                {deliveryPaidTo === "Shop" && +deliveryFee > 0 && (
                  <div style={{ fontSize: 11, color: "#166534", marginTop: 6 }}>Adds LKR {(+deliveryFee).toLocaleString()} to the total — remember to pay the rider afterward.</div>
                )}
                {deliveryPaidTo === "Rider" && +deliveryFee > 0 && (
                  <div style={{ fontSize: 11, color: GRAY, marginTop: 6 }}>Not added to the total — customer pays the rider LKR {(+deliveryFee).toLocaleString()} directly.</div>
                )}
              </div>
            )}
          </div>

          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
            <label style={{ fontSize: 13, color: GRAY, fontWeight: 500 }}>Discount</label>
            <div style={{ display: "flex", alignItems: "center", border: `1px solid ${BORDER}`, borderRadius: 8, overflow: "hidden" }}>
              <input type="number" min="0" max="100" value={discount} onChange={e => setDiscount(+e.target.value)} style={{ width: 54, padding: "6px 10px", border: "none", fontSize: 13, background: WHITE, outline: "none", textAlign: "center" }} />
              <span style={{ padding: "6px 10px", background: LIGHT, color: GRAY, fontSize: 13, borderLeft: `1px solid ${BORDER}` }}>%</span>
            </div>
            {discount > 0 && <span style={{ color: "#059669", fontSize: 13, fontWeight: 600 }}>− LKR {discountAmt.toLocaleString()}</span>}
          </div>
          {deliveryFeeToShop > 0 && (
            <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", color: "#166534", fontSize: 13 }}>
              <span style={{ fontWeight: 600 }}>Delivery Fee ({deliveryMethod})</span>
              <span style={{ fontWeight: 700 }}>+ LKR {deliveryFeeToShop.toLocaleString()}</span>
            </div>
          )}
          {creditAmt > 0 && (
            <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", color: "#166534", fontSize: 13 }}>
              <span style={{ fontWeight: 600 }}>Exchange Credit</span>
              <span style={{ fontWeight: 700 }}>− LKR {creditAmt.toLocaleString()}</span>
            </div>
          )}
          <div style={{ display: "flex", justifyContent: "space-between", padding: "12px 0", borderTop: `2px solid ${GOLD_LIGHT}` }}>
            <span style={{ fontWeight: 800, fontSize: 18, color: "#1F2937" }}>Total</span>
            <span style={{ fontWeight: 800, fontSize: 20, color: GOLD_DARK }}>LKR {total.toLocaleString()}</span>
          </div>
          <div style={{ marginBottom: 12 }}>
            <label style={{ fontSize: 13, color: GRAY, fontWeight: 500, display: "block", marginBottom: 6 }}>Payment Method</label>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {PAYMENT_METHODS.map(m => (
                <button key={m} onClick={() => setPaymentMethod(m)} style={{ flex: "1 1 auto", padding: "7px 10px", borderRadius: 8, border: `1.5px solid ${paymentMethod === m ? GOLD_DARK : BORDER}`, background: paymentMethod === m ? GOLD_LIGHT : WHITE, color: paymentMethod === m ? GOLD_DARK : GRAY, fontWeight: paymentMethod === m ? 700 : 400, fontSize: 12, cursor: "pointer" }}>{m === "Split" ? "Split (Card+Cash)" : m}</button>
              ))}
            </div>
            {paymentMethod === "Split" && (
              <>
                <div style={{ marginTop: 8, display: "flex", gap: 8 }}>
                  <div style={{ flex: 1 }}>
                    <label style={{ fontSize: 11, color: GRAY }}>Card Amount</label>
                    <input type="number" min="0" value={splitCard} onChange={e => setSplitCard(e.target.value)} placeholder="0" style={{ ...inp, marginTop: 2, fontSize: 13 }} />
                  </div>
                  <div style={{ flex: 1 }}>
                    <label style={{ fontSize: 11, color: GRAY }}>Cash Amount</label>
                    <input type="number" min="0" value={splitCash} onChange={e => setSplitCash(e.target.value)} placeholder="0" style={{ ...inp, marginTop: 2, fontSize: 13 }} />
                  </div>
                </div>
                <div style={{ fontSize: 11, marginTop: 4, fontWeight: 600, color: Math.abs((+splitCard || 0) + (+splitCash || 0) - total) < 0.5 ? "#059669" : "#DC2626" }}>
                  {Math.abs((+splitCard || 0) + (+splitCash || 0) - total) < 0.5 ? "Matches total ✓" : `Remaining: LKR ${(total - (+splitCard || 0) - (+splitCash || 0)).toLocaleString()}`}
                </div>
              </>
            )}
          </div>
          <button onClick={checkout} style={{ width: "100%", padding: "13px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 10, cursor: "pointer", fontWeight: 800, fontSize: 15, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>{Icons.check} Confirm &amp; Checkout</button>
        </div>
      )}
    </div>
  );

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <div><h2 style={{ color: "#1F2937", margin: "0 0 2px", fontSize: 20, fontWeight: 800 }}>Point of Sale</h2><p style={{ margin: 0, color: GRAY, fontSize: 13 }}>{filtered.length} items available</p></div>
        <button onClick={() => setScanning(true)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 13 }}>{Icons.camera} Scan</button>
      </div>

      <div style={{ display: isWide ? "grid" : "block", gridTemplateColumns: isWide ? "1fr 360px" : undefined, gap: isWide ? 16 : 0, alignItems: "start" }}>
        {/* Left: product browser */}
        <div>
          <div style={{ position: "relative", marginBottom: 10 }}>
            <div style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", color: GRAY }}>{Icons.search}</div>
            <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search or enter SKU..." style={{ ...inp, paddingLeft: 38, marginTop: 0 }} />
          </div>
          <BizFilter value={bizFilter} onChange={handleBizFilter} />
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", margin: "10px 0" }}>
            {categoriesForBiz(bizFilter).map(c => (<button key={c} onClick={() => setCatFilter(c)} style={{ padding: "5px 14px", borderRadius: 20, border: `1px solid ${catFilter === c ? GOLD : BORDER}`, background: catFilter === c ? GOLD_DARK : WHITE, color: catFilter === c ? WHITE : GRAY, cursor: "pointer", fontSize: 12, fontWeight: catFilter === c ? 700 : 400 }}>{c}</button>))}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: isWide ? "1fr 1fr 1fr" : "1fr 1fr", gap: 10 }}>
            {filtered.map(p => (
              <button key={p.id} onClick={() => addToCart(p)} style={{ background: WHITE, border: `1px solid ${BORDER}`, borderRadius: 10, padding: 10, cursor: "pointer", textAlign: "left", boxShadow: "0 1px 3px rgba(0,0,0,0.05)" }}
                onMouseEnter={e => { e.currentTarget.style.boxShadow = `0 0 0 2px ${GOLD}`; e.currentTarget.style.borderColor = GOLD; }}
                onMouseLeave={e => { e.currentTarget.style.boxShadow = "0 1px 3px rgba(0,0,0,0.05)"; e.currentTarget.style.borderColor = BORDER; }}>
                <div style={{ width: "100%", aspectRatio: "1/1", borderRadius: 7, overflow: "hidden", background: LIGHT, display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 6, border: `1px solid ${BORDER}`, color: GRAY }}>
                  {p.image ? <img src={p.image} alt={p.name} style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : Icons.image}
                </div>
                <div style={{ fontSize: 13, fontWeight: 700, color: "#1F2937", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{p.name}</div>
                <div style={{ marginTop: 3 }}><BizBadge business={p.business || "Blingshop"} /></div>
                <div style={{ fontSize: 13, color: GOLD_DARK, fontWeight: 800, marginTop: 4 }}>LKR {(+p.price).toLocaleString()}</div>
                <div style={{ fontSize: 11, color: GRAY, marginTop: 1 }}>Stock: {p.stock}</div>
              </button>
            ))}
          </div>
          {filtered.length === 0 && <div style={{ textAlign: "center", color: "#9CA3AF", padding: "40px 0", fontSize: 14 }}>No products found</div>}
        </div>

        {/* Right: cart — always visible on wide, only when has items on mobile */}
        {(isWide || cart.length > 0) && (
          <div style={{ position: isWide ? "sticky" : "relative", top: isWide ? 72 : "auto", marginTop: isWide ? 0 : 14 }}>
            {cartPanel}
          </div>
        )}
      </div>

      {scanning && <CameraScanner onDetect={handleScan} onClose={() => setScanning(false)} />}
      {successSale && (
        <Modal title="Sale Completed" onClose={() => setSuccessSale(null)}>
          <div style={{ textAlign: "center", padding: "10px 0" }}>
            <div style={{ width: 56, height: 56, borderRadius: "50%", background: "#F0FDF4", border: "2px solid #86EFAC", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 16px", color: "#059669" }}>{Icons.check}</div>
            <div style={{ fontSize: 13, color: GRAY, marginBottom: 4 }}>Receipt #{successSale.id}</div>
            {successSale.customerName && <div style={{ fontSize: 13, color: GOLD_DARK, fontWeight: 600, marginBottom: 4 }}>{successSale.customerName}</div>}
            <div style={{ fontSize: 26, fontWeight: 800, color: GOLD_DARK, marginBottom: 6 }}>LKR {successSale.total.toLocaleString()}</div>
            {Object.entries(successSale.items.reduce((acc, i) => { const b = i.business || "Blingshop"; acc[b] = (acc[b] || 0) + i.price * i.qty; return acc; }, {})).map(([b, v]) => (
              <div key={b} style={{ fontSize: 12, color: GRAY, marginBottom: 2 }}><BizBadge business={b} /> &nbsp;LKR {v.toLocaleString()}</div>
            ))}
            {successSale.deliveryMethod !== "In Store" && (
              <div style={{ marginTop: 10, background: successSale.deliveryPaidTo === "Shop" ? "#FEF3C7" : LIGHT, border: `1px solid ${successSale.deliveryPaidTo === "Shop" ? "#FCD34D" : BORDER}`, borderRadius: 8, padding: "8px 12px", fontSize: 12, color: successSale.deliveryPaidTo === "Shop" ? "#92400E" : GRAY }}>
                {successSale.deliveryMethod}
                {successSale.deliveryFee > 0 && ` — LKR ${(+successSale.deliveryFee).toLocaleString()} `}
                {successSale.deliveryFee > 0 && (successSale.deliveryPaidTo === "Shop" ? "collected — remember to pay the rider" : "paid to rider directly")}
              </div>
            )}
            <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
              <button onClick={() => { setSuccessSale(null); setPrintSale(successSale); }} style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 8, padding: "11px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 14 }}>{Icons.print} Print Receipt</button>
              <button onClick={() => setSuccessSale(null)} style={{ padding: "11px 20px", background: LIGHT, color: GRAY, border: "1px solid #E5E7EB", borderRadius: 8, cursor: "pointer" }}>Close</button>
            </div>
          </div>
        </Modal>
      )}
      {printSale && <ReceiptPrintModal sale={printSale} onClose={() => setPrintSale(null)} />}
    </div>
  );
}

// ── Refund Modal ──────────────────────────────────────────────────
function RefundModal({ sale, refunds, onClose, onRefund }) {
  const prevRefunds = (refunds || []).filter(r => r.sale_id === sale.id);

  const alreadyRefunded = {};
  for (const r of prevRefunds) {
    const ritems = Array.isArray(r.items) ? r.items : JSON.parse(r.items || "[]");
    for (const ri of ritems) {
      alreadyRefunded[ri.product_id] = (alreadyRefunded[ri.product_id] || 0) + ri.qty;
    }
  }

  const refundableItems = sale.items.map(item => {
    const pid = item.product_id || item.id;
    return { ...item, product_id: pid, maxQty: item.qty - (alreadyRefunded[pid] || 0) };
  }).filter(i => i.maxQty > 0);

  const [selected, setSelected] = useState(() => {
    const init = {};
    for (const item of refundableItems) init[item.product_id] = { ...item, refundQty: item.maxQty };
    return init;
  });
  const [notes, setNotes] = useState("");
  const [busy, setBusy] = useState(false);

  const toggle = pid => setSelected(s => {
    if (s[pid]) { const { [pid]: _, ...rest } = s; return rest; }
    const item = refundableItems.find(i => i.product_id === pid);
    return { ...s, [pid]: { ...item, refundQty: item.maxQty } };
  });

  const setQty = (pid, qty) => {
    const max = refundableItems.find(i => i.product_id === pid).maxQty;
    setSelected(s => ({ ...s, [pid]: { ...s[pid], refundQty: Math.max(1, Math.min(qty, max)) } }));
  };

  const refundAmt = Object.values(selected).reduce((s, i) => s + parseFloat(i.price) * i.refundQty, 0);
  const allSel = refundableItems.every(i => selected[i.product_id]);

  const doProcess = async (type) => {
    if (!Object.keys(selected).length) { alert("Select at least one item."); return; }
    setBusy(true);
    const items = Object.values(selected).map(i => ({
      product_id: i.product_id, name: i.name, qty: i.refundQty,
      price: parseFloat(i.price), business: i.business || "Blingshop",
    }));
    await onRefund({
      items, type, refundAmt, notes,
      customerId: sale.customerId, customerName: sale.customerName
    });
    setBusy(false);
    onClose();
  };

  return (
    <Modal title={`Refund / Exchange — #${sale.id}`} onClose={onClose}>
      <div style={{ fontSize: 13 }}>
        {prevRefunds.length > 0 && (
          <div style={{ background: "#FEF3C7", border: "1px solid #FCD34D", borderRadius: 8, padding: "10px 12px", marginBottom: 14 }}>
            <div style={{ fontWeight: 700, color: "#92400E", marginBottom: 4 }}>Previous Refunds</div>
            {prevRefunds.map(r => (
              <div key={r.id} style={{ fontSize: 12, color: "#78350F", marginBottom: 2 }}>
                {new Date(r.date).toLocaleDateString()} — LKR {parseFloat(r.refund_amount).toLocaleString()} ({r.type})
                {r.notes && <span style={{ color: "#A16207" }}> · {r.notes}</span>}
              </div>
            ))}
          </div>
        )}

        {refundableItems.length === 0 ? (
          <div style={{ textAlign: "center", color: GRAY, padding: "20px 0" }}>All items have already been fully refunded.</div>
        ) : (<>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <div style={{ fontWeight: 700, color: "#374151" }}>Select items to return</div>
            <button onClick={() => {
              if (allSel) { setSelected({}); return; }
              const all = {};
              for (const item of refundableItems) all[item.product_id] = { ...item, refundQty: item.maxQty };
              setSelected(all);
            }} style={{ fontSize: 12, color: GOLD_DARK, background: "none", border: "none", cursor: "pointer", fontWeight: 700 }}>
              {allSel ? "Deselect All" : "Select All"}
            </button>
          </div>

          {refundableItems.map(item => {
            const pid = item.product_id;
            const isChecked = !!selected[pid];
            return (
              <div key={pid} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 0", borderBottom: `1px solid ${BORDER}` }}>
                <input type="checkbox" checked={isChecked} onChange={() => toggle(pid)}
                  style={{ width: 16, height: 16, cursor: "pointer", accentColor: GOLD_DARK }} />
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 600, color: "#1F2937" }}>{item.name}</div>
                  <div style={{ display: "flex", gap: 6, marginTop: 2 }}>
                    <BizBadge business={item.business || "Blingshop"} />
                    <span style={{ fontSize: 11, color: GRAY }}>LKR {parseFloat(item.price).toLocaleString()} each · max {item.maxQty}</span>
                  </div>
                </div>
                {isChecked && (
                  <div style={{ display: "flex", alignItems: "center" }}>
                    <button onClick={() => setQty(pid, selected[pid].refundQty - 1)}
                      style={{ width: 26, height: 26, borderRadius: "6px 0 0 6px", border: `1px solid ${BORDER}`, background: LIGHT, cursor: "pointer", fontWeight: 700 }}>−</button>
                    <div style={{ width: 32, height: 26, display: "flex", alignItems: "center", justifyContent: "center", border: `1px solid ${BORDER}`, borderLeft: "none", borderRight: "none", fontSize: 13, fontWeight: 700 }}>{selected[pid].refundQty}</div>
                    <button onClick={() => setQty(pid, selected[pid].refundQty + 1)}
                      style={{ width: 26, height: 26, borderRadius: "0 6px 6px 0", border: `1px solid ${BORDER}`, background: LIGHT, cursor: "pointer", fontWeight: 700 }}>+</button>
                  </div>
                )}
                <div style={{ minWidth: 84, textAlign: "right", fontWeight: 700, fontSize: 13, color: isChecked ? "#DC2626" : GRAY }}>
                  {isChecked
                    ? `− LKR ${(parseFloat(item.price) * selected[pid].refundQty).toLocaleString()}`
                    : `LKR ${(parseFloat(item.price) * item.maxQty).toLocaleString()}`}
                </div>
              </div>
            );
          })}

          <div style={{ marginTop: 12 }}>
            <textarea value={notes} onChange={e => setNotes(e.target.value)} placeholder="Notes (optional)"
              rows={2} style={{
                width: "100%", padding: "8px 10px", border: `1px solid ${BORDER}`, borderRadius: 8,
                fontSize: 13, resize: "vertical", boxSizing: "border-box", fontFamily: "inherit"
              }} />
          </div>

          {refundAmt > 0 && (
            <div style={{
              background: "#FEF2F2", border: "1px solid #FCA5A5", borderRadius: 8, padding: "10px 14px",
              marginTop: 12, display: "flex", justifyContent: "space-between", alignItems: "center"
            }}>
              <span style={{ fontWeight: 700, color: "#991B1B" }}>Refund Amount</span>
              <span style={{ fontWeight: 800, fontSize: 18, color: "#DC2626" }}>LKR {refundAmt.toLocaleString()}</span>
            </div>
          )}

          <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
            <button onClick={() => doProcess("refund")} disabled={busy || !Object.keys(selected).length}
              style={{
                flex: 1, padding: "11px", background: "#DC2626", color: WHITE, border: "none", borderRadius: 8,
                cursor: "pointer", fontWeight: 700, fontSize: 14, opacity: (!Object.keys(selected).length || busy) ? 0.5 : 1
              }}>
              {busy ? "Processing…" : "Process Refund"}
            </button>
            <button onClick={() => doProcess("exchange")} disabled={busy || !Object.keys(selected).length}
              style={{
                flex: 1, padding: "11px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8,
                cursor: "pointer", fontWeight: 700, fontSize: 14, opacity: (!Object.keys(selected).length || busy) ? 0.5 : 1
              }}>
              {busy ? "Processing…" : "Exchange"}
            </button>
          </div>
          <div style={{ fontSize: 11, color: GRAY, marginTop: 8, textAlign: "center" }}>
            Refund returns items to stock &amp; records the return. Exchange opens POS with store credit applied.
          </div>
        </>)}
      </div>
    </Modal>
  );
}

// ── Sales History ─────────────────────────────────────────────────
function SalesHistory({ sales, refunds, addRefund, setProducts, currentStaff, onExchange }) {
  const [sel, setSel] = useState(null);
  const [printSale, setPrintSale] = useState(null);
  const [refundSale, setRefundSale] = useState(null);
  const [bizFilter, setBizFilter] = useState("All");

  const filtered = bizFilter === "All" ? sales : sales.filter(s => s.items.some(i => (i.business || "Blingshop") === bizFilter));

  const saleRefunds = id => (refunds || []).filter(r => r.sale_id === id);
  const isFullyRefunded = sale => {
    const sr = saleRefunds(sale.id);
    if (!sr.length) return false;
    return sale.items.every(item => {
      const pid = item.product_id || item.id;
      const refunded = sr.reduce((s, r) => {
        const ritems = Array.isArray(r.items) ? r.items : JSON.parse(r.items || "[]");
        return s + ritems.filter(i => i.product_id === pid).reduce((a, i) => a + i.qty, 0);
      }, 0);
      return refunded >= item.qty;
    });
  };

  const handleRefund = async ({ items, type, refundAmt, notes, customerId, customerName }) => {
    const refund = {
      id: generateId(), sale_id: refundSale.id, date: new Date().toISOString(),
      items, refund_amount: refundAmt, type, notes,
      staffId: currentStaff?.id || null, staffName: currentStaff?.name || null,
    };
    await addRefund(refund);
    if (type === "exchange") onExchange({ amount: refundAmt, customerId, customerName });
  };

  return (
    <div>
      <div style={{ marginBottom: 12 }}><h2 style={{ color: "#1F2937", margin: "0 0 2px", fontSize: 20, fontWeight: 800 }}>Sales History</h2><p style={{ margin: 0, color: GRAY, fontSize: 13 }}>{sales.length} transactions total</p></div>
      <div style={{ marginBottom: 14 }}><BizFilter value={bizFilter} onChange={setBizFilter} /></div>
      {filtered.length === 0 && <div style={{ textAlign: "center", color: "#9CA3AF", padding: "40px 0", fontSize: 14 }}>No transactions found</div>}
      {[...filtered].reverse().map(s => {
        const bizSet = [...new Set(s.items.map(i => i.business || "Blingshop"))];
        const sr = saleRefunds(s.id);
        const fullyRefunded = isFullyRefunded(s);
        return (
          <div key={s.id} style={{ background: WHITE, borderRadius: 10, padding: "12px 14px", marginBottom: 8, boxShadow: "0 1px 4px rgba(0,0,0,0.06)", border: `1px solid ${BORDER}`, cursor: "pointer" }} onClick={() => setSel(s)}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <div style={{ fontWeight: 700, color: "#1F2937", fontSize: 14 }}>#{s.id}</div>
                  {fullyRefunded && <span style={{ fontSize: 10, fontWeight: 700, background: "#FEE2E2", color: "#DC2626", padding: "2px 7px", borderRadius: 10 }}>REFUNDED</span>}
                  {!fullyRefunded && sr.length > 0 && <span style={{ fontSize: 10, fontWeight: 700, background: "#FEF3C7", color: "#92400E", padding: "2px 7px", borderRadius: 10 }}>PARTIAL REFUND</span>}
                </div>
                {s.customerName && <div style={{ fontSize: 12, color: GOLD_DARK, fontWeight: 600 }}>{s.customerName}</div>}
                <div style={{ fontSize: 12, color: GRAY, marginTop: 2 }}>{new Date(s.date).toLocaleString()}</div>
                <div style={{ display: "flex", gap: 4, marginTop: 4, flexWrap: "wrap" }}>
                  {bizSet.map(b => <BizBadge key={b} business={b} />)}
                  {s.deliveryMethod && s.deliveryMethod !== "In Store" && (
                    <span style={{ fontSize: 10, fontWeight: 700, padding: "2px 8px", borderRadius: 10, background: s.deliveryPaidTo === "Shop" ? "#FEF3C7" : "#F3F4F6", color: s.deliveryPaidTo === "Shop" ? "#92400E" : "#374151" }}>
                      {s.deliveryMethod}{s.deliveryPaidTo === "Rider" ? " · rider paid" : ""}
                    </span>
                  )}
                </div>
              </div>
              <div style={{ textAlign: "right" }}>
                <div style={{ fontWeight: 800, color: fullyRefunded ? "#DC2626" : GOLD_DARK, fontSize: 16, textDecoration: fullyRefunded ? "line-through" : "none" }}>LKR {parseFloat(s.total).toLocaleString()}</div>
                {s.discount > 0 && <div style={{ fontSize: 11, color: "#059669", marginTop: 2 }}>Saved LKR {parseFloat(s.discount).toLocaleString()}</div>}
                <div style={{ fontSize: 11, color: GRAY, marginTop: 2 }}>{s.items.length} item(s)</div>
                <div style={{ fontSize: 11, color: GRAY, marginTop: 2 }}>{s.paymentMethod || "Cash"}</div>
              </div>
            </div>
          </div>
        );
      })}

      {sel && (
        <Modal title={`Transaction #${sel.id}`} onClose={() => setSel(null)}>
          <div style={{ fontSize: 13 }}>
            <div style={{ color: GRAY, marginBottom: 6, fontSize: 12 }}>{new Date(sel.date).toLocaleString()}</div>
            {sel.customerName && <div style={{ color: GOLD_DARK, fontWeight: 700, marginBottom: 14, fontSize: 14 }}>{sel.customerName}</div>}
            {BUSINESSES.map(biz => {
              const bizItems = sel.items.filter(i => (i.business || "Blingshop") === biz);
              if (!bizItems.length) return null;
              return (
                <div key={biz} style={{ marginBottom: 12 }}>
                  <div style={{ marginBottom: 6 }}><BizBadge business={biz} size="md" /></div>
                  <div style={{ background: LIGHT, borderRadius: 8, padding: 10 }}>
                    {bizItems.map(i => (<div key={i.id || i.name} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderBottom: `1px solid ${BORDER}` }}><span>{i.name} &times; {i.qty}</span><span style={{ fontWeight: 700, color: GOLD_DARK }}>LKR {(parseFloat(i.price) * i.qty).toLocaleString()}</span></div>))}
                    <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", fontWeight: 700, fontSize: 13 }}>
                      <span>{biz} Subtotal</span>
                      <span style={{ color: biz === "Blingshop" ? GOLD_DARK : RC_DARK }}>LKR {bizItems.reduce((s, i) => s + parseFloat(i.price) * i.qty, 0).toLocaleString()}</span>
                    </div>
                  </div>
                </div>
              );
            })}
            {sel.discount > 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", color: "#059669" }}><span>Discount</span><span>− LKR {parseFloat(sel.discount).toLocaleString()}</span></div>}
            {sel.deliveryMethod && sel.deliveryMethod !== "In Store" && (
              <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", color: sel.deliveryPaidTo === "Shop" ? "#92400E" : GRAY }}>
                <span>Delivery ({sel.deliveryMethod})</span>
                <span style={{ fontWeight: 600 }}>{sel.deliveryPaidTo === "Shop" ? `+ LKR ${(+sel.deliveryFee || 0).toLocaleString()} — pay rider` : `LKR ${(+sel.deliveryFee || 0).toLocaleString()} paid to rider directly`}</span>
              </div>
            )}
            <div style={{ display: "flex", justifyContent: "space-between", padding: "10px 0 4px", fontWeight: 800, fontSize: 16, color: GOLD_DARK, borderTop: `2px solid ${GOLD_LIGHT}`, marginTop: 4 }}><span>Total</span><span>LKR {parseFloat(sel.total).toLocaleString()}</span></div>
            <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", fontSize: 12, color: GRAY }}><span>Payment</span><span>{paymentSummary(sel)}</span></div>

            {saleRefunds(sel.id).length > 0 && (
              <div style={{ marginTop: 12, background: "#FEF3C7", border: "1px solid #FCD34D", borderRadius: 8, padding: "10px 12px" }}>
                <div style={{ fontWeight: 700, color: "#92400E", marginBottom: 4, fontSize: 12 }}>Refund History</div>
                {saleRefunds(sel.id).map(r => (
                  <div key={r.id} style={{ fontSize: 12, color: "#78350F", marginBottom: 2 }}>
                    {new Date(r.date).toLocaleDateString()} — LKR {parseFloat(r.refund_amount).toLocaleString()} ({r.type})
                    {r.notes && <span style={{ color: "#A16207" }}> · {r.notes}</span>}
                  </div>
                ))}
              </div>
            )}

            <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
              <button onClick={() => { setSel(null); setPrintSale(sel); }} style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 8, padding: "11px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 14 }}>{Icons.print} Print Receipt</button>
              {!isFullyRefunded(sel) && (
                <button onClick={() => { setSel(null); setRefundSale(sel); }} style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 8, padding: "11px", background: "#DC2626", color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 14 }}>
                  Refund / Exchange
                </button>
              )}
            </div>
          </div>
        </Modal>
      )}

      {refundSale && (
        <RefundModal sale={refundSale} refunds={refunds} onClose={() => setRefundSale(null)} onRefund={handleRefund} />
      )}
      {printSale && <ReceiptPrintModal sale={printSale} onClose={() => setPrintSale(null)} />}
    </div>
  );
}

// ── Reports ───────────────────────────────────────────────────────
function Reports({ products, sales }) {
  const [bizView, setBizView] = useState("All");
  const [range, setRange] = useState("month");
  const [customFrom, setCustomFrom] = useState("");
  const [customTo, setCustomTo] = useState("");
  const [trendProduct, setTrendProduct] = useState("");
  const now = new Date();

  const rangeStart = () => {
    if (range === "today") { const d = new Date(now); d.setHours(0, 0, 0, 0); return d; }
    if (range === "week") { const d = new Date(now); d.setDate(d.getDate() - 6); d.setHours(0, 0, 0, 0); return d; }
    if (range === "month") { return new Date(now.getFullYear(), now.getMonth(), 1); }
    if (range === "year") { return new Date(now.getFullYear(), 0, 1); }
    if (range === "custom" && customFrom) return new Date(customFrom);
    return new Date(0);
  };
  const rangeEnd = () => {
    if (range === "custom" && customTo) { const d = new Date(customTo); d.setHours(23, 59, 59, 999); return d; }
    const d = new Date(now); d.setHours(23, 59, 59, 999); return d;
  };

  const periodSales = sales.filter(s => { const d = new Date(s.date); return d >= rangeStart() && d <= rangeEnd(); });

  const calcItemsRev = (items) => bizView === "All" ? items.reduce((s, i) => s + i.price * i.qty, 0) : items.filter(i => (i.business || "Blingshop") === bizView).reduce((s, i) => s + i.price * i.qty, 0);
  const filteredSales = bizView === "All" ? periodSales : periodSales.filter(s => s.items.some(i => (i.business || "Blingshop") === bizView));
  const totalRevenue = filteredSales.reduce((s, x) => s + calcItemsRev(x.items), 0);
  const totalSales = filteredSales.length;
  const totalDiscount = filteredSales.reduce((s, x) => s + parseFloat(x.discount || 0), 0);
  const avgOrder = totalSales > 0 ? totalRevenue / totalSales : 0;
  const totalItems = filteredSales.reduce((s, x) => s + x.items.filter(i => bizView === "All" || (i.business || "Blingshop") === bizView).reduce((a, i) => a + i.qty, 0), 0);

  const bizRevenue = {};
  for (const sale of periodSales) {
    for (const item of sale.items) {
      const b = (item.business || "Blingshop");
      bizRevenue[b] = (bizRevenue[b] || 0) + item.price * item.qty;
    }
  }

  const paymentTotals = filteredSales.reduce((acc, x) => {
    acc.Cash += parseFloat(x.cashAmount || 0);
    acc.Card += parseFloat(x.cardAmount || 0);
    acc["Bank Transfer"] += parseFloat(x.bankAmount || 0);
    return acc;
  }, { Cash: 0, Card: 0, "Bank Transfer": 0 });
  const paymentMax = Math.max(...Object.values(paymentTotals), 1);

  // Delivery fees the shop collected are cash we now owe out to riders (a liability, not
  // profit); fees customers pay riders directly never touch our cash — tracked separately so
  // the owner can see both what's payable and how much delivery cost customers in total.
  const deliverySales = filteredSales.filter(s => s.deliveryMethod && s.deliveryMethod !== "In Store");
  const deliveryByMethod = {};
  let owedToRiders = 0, paidDirectToRiders = 0;
  for (const s of deliverySales) {
    const m = s.deliveryMethod, fee = +s.deliveryFee || 0;
    if (!deliveryByMethod[m]) deliveryByMethod[m] = { count: 0, collected: 0, direct: 0 };
    deliveryByMethod[m].count += 1;
    if (s.deliveryPaidTo === "Shop") { deliveryByMethod[m].collected += fee; owedToRiders += fee; }
    else { deliveryByMethod[m].direct += fee; paidDirectToRiders += fee; }
  }

  const catRevenue = {};
  for (const sale of filteredSales) {
    for (const item of sale.items) {
      if (bizView !== "All" && (item.business || "Blingshop") !== bizView) continue;
      const prod = products.find(p => p.id === item.product_id || p.id === item.id);
      const cat = prod?.category || "Other";
      catRevenue[cat] = (catRevenue[cat] || 0) + item.price * item.qty;
    }
  }
  const catMax = Math.max(...Object.values(catRevenue), 1);

  const prodSales = {};
  for (const sale of filteredSales) {
    for (const item of sale.items) {
      if (bizView !== "All" && (item.business || "Blingshop") !== bizView) continue;
      if (!prodSales[item.name]) prodSales[item.name] = { name: item.name, business: item.business || "Blingshop", qty: 0, revenue: 0 };
      prodSales[item.name].qty += item.qty; prodSales[item.name].revenue += item.price * item.qty;
    }
  }
  const allProdSales = Object.values(prodSales).sort((a, b) => b.revenue - a.revenue);
  const topProducts = allProdSales.slice(0, 8);

  // Trend arrows on Top Selling Products compare each product's revenue against the
  // equivalent-length period immediately before the selected range (e.g. this month vs last
  // month) — a flat revenue total doesn't tell you whether a product is picking up or fading.
  const rangeMs = rangeEnd().getTime() - rangeStart().getTime();
  const prevEnd = new Date(rangeStart().getTime() - 1);
  const prevStart = new Date(prevEnd.getTime() - rangeMs);
  const prevPeriodSales = sales.filter(s => { const d = new Date(s.date); return d >= prevStart && d <= prevEnd; });
  const prevFilteredSales = bizView === "All" ? prevPeriodSales : prevPeriodSales.filter(s => s.items.some(i => (i.business || "Blingshop") === bizView));
  const prevProdRevenue = {};
  for (const sale of prevFilteredSales) {
    for (const item of sale.items) {
      if (bizView !== "All" && (item.business || "Blingshop") !== bizView) continue;
      prevProdRevenue[item.name] = (prevProdRevenue[item.name] || 0) + item.price * item.qty;
    }
  }

  const trendDays = range === "today" ? 1 : range === "week" ? 7 : 30;
  const trend = Array.from({ length: trendDays }, (_, i) => {
    const d = new Date(now); d.setDate(d.getDate() - (trendDays - 1 - i)); d.setHours(0, 0, 0, 0);
    const label = d.toLocaleDateString("en-US", { month: "short", day: "numeric" });
    const rev = filteredSales.filter(s => { const sd = new Date(s.date); sd.setHours(0, 0, 0, 0); return sd.getTime() === d.getTime(); }).reduce((s, x) => s + calcItemsRev(x.items), 0);
    return { label, value: rev };
  });
  const trendMax = Math.max(...trend.map(t => t.value), 1);

  // Per-product day-by-day breakdown, shown as its own chart below Top Selling Products.
  // Falls back to the current top product whenever the previously picked one drops out of the
  // filtered range (switched business/date range), same pattern as the category-filter reset.
  const selectedProductName = trendProduct && allProdSales.some(p => p.name === trendProduct) ? trendProduct : (allProdSales[0]?.name || "");
  const productTrend = selectedProductName ? Array.from({ length: trendDays }, (_, i) => {
    const d = new Date(now); d.setDate(d.getDate() - (trendDays - 1 - i)); d.setHours(0, 0, 0, 0);
    const label = d.toLocaleDateString("en-US", { month: "short", day: "numeric" });
    const dayItems = filteredSales
      .filter(s => { const sd = new Date(s.date); sd.setHours(0, 0, 0, 0); return sd.getTime() === d.getTime(); })
      .flatMap(s => s.items.filter(i => i.name === selectedProductName && (bizView === "All" || (i.business || "Blingshop") === bizView)));
    const units = dayItems.reduce((s, i) => s + i.qty, 0);
    const rev = dayItems.reduce((s, i) => s + i.price * i.qty, 0);
    return { label, units, rev };
  }) : [];
  const productTrendMax = Math.max(...productTrend.map(t => t.units), 1);

  const fProducts = bizView === "All" ? products : products.filter(p => (p.business || "Blingshop") === bizView);
  const stockValue = fProducts.reduce((s, p) => s + p.price * p.stock, 0);
  const stockWorth = fProducts.reduce((s, p) => s + (p.cost || 0) * p.stock, 0);
  const outOfStock = fProducts.filter(p => p.stock === 0).length;
  const lowStock = fProducts.filter(p => p.stock > 0 && p.stock <= 3).length;

  const exportReport = () => {
    const summary = [
      { Metric: "Business", Value: bizView }, { Metric: "Period", Value: range },
      { Metric: "Total Sales", Value: totalSales }, { Metric: "Total Revenue (LKR)", Value: totalRevenue },
      { Metric: "Items Sold", Value: totalItems }, { Metric: "Discounts Given (LKR)", Value: totalDiscount },
      { Metric: "Avg Order Value (LKR)", Value: Math.round(avgOrder) },
      { Metric: "Cash Collected (LKR)", Value: paymentTotals.Cash }, { Metric: "Card Collected (LKR)", Value: paymentTotals.Card },
      { Metric: "Bank Transfer Collected (LKR)", Value: paymentTotals["Bank Transfer"] },
      { Metric: "Delivery Fees Owed to Riders (LKR)", Value: owedToRiders }, { Metric: "Delivery Fees Paid Direct to Riders (LKR)", Value: paidDirectToRiders },
    ];
    const salesData = filteredSales.map(s => ({
      "Receipt ID": s.id, "Date": new Date(s.date).toLocaleString(), "Customer": s.customerName || "",
      "Blingshop (LKR)": s.items.filter(i => (i.business || "Blingshop") === "Blingshop").reduce((a, i) => a + i.price * i.qty, 0),
      "RC Boutique (LKR)": s.items.filter(i => (i.business || "Blingshop") === "RC Boutique").reduce((a, i) => a + i.price * i.qty, 0),
      "Discount": s.discount, "Total (LKR)": s.total, "Payment Method": s.paymentMethod || "Cash",
      "Cash (LKR)": parseFloat(s.cashAmount || 0), "Card (LKR)": parseFloat(s.cardAmount || 0), "Bank Transfer (LKR)": parseFloat(s.bankAmount || 0),
      "Delivery Method": s.deliveryMethod || "In Store", "Delivery Fee (LKR)": parseFloat(s.deliveryFee || 0),
      "Delivery Fee Paid To": s.deliveryMethod && s.deliveryMethod !== "In Store" ? (s.deliveryPaidTo === "Shop" ? "Shop (owed to rider)" : "Rider directly") : "",
    }));
    const topData = topProducts.map((p, i) => {
      const prevRev = prevProdRevenue[p.name] || 0;
      const changePct = prevRev > 0 ? Math.round(((p.revenue - prevRev) / prevRev) * 100) : null;
      return { "Rank": i + 1, "Product": p.name, "Business": p.business, "Units Sold": p.qty, "Revenue (LKR)": p.revenue, "Prev Period Revenue (LKR)": prevRev, "Change vs Prev Period (%)": changePct === null ? "New" : changePct };
    });
    const invData = fProducts.map(p => ({ "SKU": p.sku, "Product": p.name, "Business": p.business || "Blingshop", "Category": p.category, "Cost": p.cost || 0, "Price": p.price, "Stock": p.stock, "Stock Worth (Cost)": (p.cost || 0) * p.stock, "Stock Value": p.price * p.stock, "Status": p.stock === 0 ? "Out of Stock" : p.stock <= 3 ? "Low Stock" : "In Stock" }));
    try {
      const wb = xlsxUtils.book_new();
      xlsxUtils.book_append_sheet(wb, xlsxUtils.json_to_sheet(summary), "Summary");
      xlsxUtils.book_append_sheet(wb, xlsxUtils.json_to_sheet(salesData), "Sales");
      xlsxUtils.book_append_sheet(wb, xlsxUtils.json_to_sheet(topData), "Top Products");
      xlsxUtils.book_append_sheet(wb, xlsxUtils.json_to_sheet(invData), "Inventory");
      xlsxWriteFile(wb, `Report_${bizView.replace(" ", "_")}_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch { alert("Export requires xlsx. Run: npm install xlsx"); }
  };

  const rangeBtn = (id, lbl) => (<button key={id} onClick={() => setRange(id)} style={{ padding: "6px 14px", borderRadius: 20, border: `1px solid ${range === id ? GOLD : BORDER}`, background: range === id ? GOLD_DARK : WHITE, color: range === id ? WHITE : GRAY, cursor: "pointer", fontSize: 12, fontWeight: range === id ? 700 : 400 }}>{lbl}</button>);

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div><h2 style={{ color: "#1F2937", margin: "0 0 2px", fontSize: 20, fontWeight: 800 }}>Reports</h2><p style={{ margin: 0, color: GRAY, fontSize: 13 }}>Financial and inventory overview</p></div>
        <button onClick={exportReport} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 13 }}>{Icons.download} Export to Excel</button>
      </div>

      <div style={{ background: WHITE, borderRadius: 10, padding: 14, border: `1px solid ${BORDER}`, marginBottom: 14 }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 10 }}>View By Business</div>
        <BizFilter value={bizView} onChange={setBizView} />
        {bizView === "All" && Object.keys(bizRevenue).length > 0 && (
          <div style={{ display: "flex", gap: 12, marginTop: 12, flexWrap: "wrap" }}>
            {Object.entries(bizRevenue).map(([b, v]) => (
              <div key={b} style={{ flex: 1, minWidth: 120, background: b === "Blingshop" ? GOLD_LIGHT : RC_LIGHT, borderRadius: 8, padding: "10px 14px", border: `1px solid ${b === "Blingshop" ? GOLD : RC_BLUE}` }}>
                <div style={{ fontSize: 11, color: b === "Blingshop" ? GOLD_DARK : RC_DARK, fontWeight: 600, marginBottom: 2 }}>{b}</div>
                <div style={{ fontSize: 16, fontWeight: 800, color: b === "Blingshop" ? GOLD_DARK : RC_DARK }}>LKR {v.toLocaleString()}</div>
                <div style={{ fontSize: 11, color: GRAY }}>This period</div>
              </div>
            ))}
          </div>
        )}
      </div>

      <div style={{ background: WHITE, borderRadius: 10, padding: 14, border: `1px solid ${BORDER}`, marginBottom: 14 }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 10 }}>Date Range</div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {[["today", "Today"], ["week", "Last 7 Days"], ["month", "This Month"], ["year", "This Year"], ["custom", "Custom"]].map(([id, lbl]) => rangeBtn(id, lbl))}
        </div>
        {range === "custom" && (
          <div style={{ display: "flex", gap: 10, marginTop: 12, flexWrap: "wrap" }}>
            <div><label style={labelStyle}>From</label><input type="date" value={customFrom} onChange={e => setCustomFrom(e.target.value)} style={{ ...inp, marginTop: 4, width: "auto" }} /></div>
            <div><label style={labelStyle}>To</label><input type="date" value={customTo} onChange={e => setCustomTo(e.target.value)} style={{ ...inp, marginTop: 4, width: "auto" }} /></div>
          </div>
        )}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 14 }}>
        <StatCard label="Total Revenue" value={`LKR ${totalRevenue.toLocaleString()}`} sub={`${totalSales} transactions`} accent={GOLD_DARK} />
        <StatCard label="Items Sold" value={totalItems} sub={`Avg order: LKR ${Math.round(avgOrder).toLocaleString()}`} accent={GOLD} />
        <StatCard label="Discounts Given" value={`LKR ${totalDiscount.toLocaleString()}`} sub={`Across ${totalSales} sales`} accent="#2563EB" />
        <StatCard label="Stock Value (Selling)" value={`LKR ${stockValue.toLocaleString()}`} sub={`${fProducts.length} products`} accent="#059669" />
        <StatCard label="Total Stock Worth (Cost)" value={`LKR ${stockWorth.toLocaleString()}`} sub={`${fProducts.length} products`} accent="#7C3AED" />
      </div>

      {trendDays > 1 && (
        <div style={{ background: WHITE, borderRadius: 12, padding: 16, border: `1px solid ${BORDER}`, marginBottom: 14 }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: "#1F2937", marginBottom: 14 }}>Revenue Trend</div>
          <div style={{ display: "flex", alignItems: "flex-end", gap: 3, height: 100, overflowX: "auto" }}>
            {trend.map((t, i) => (
              <div key={i} style={{ display: "flex", flexDirection: "column", alignItems: "center", flex: 1, minWidth: 24 }}>
                <div title={`LKR ${t.value.toLocaleString()}`} style={{ width: "100%", background: t.value > 0 ? GOLD : GOLD_LIGHT, borderRadius: "4px 4px 0 0", height: `${Math.max(4, Math.round((t.value / trendMax) * 84))}px` }} />
                {trendDays <= 14 && <div style={{ fontSize: 9, color: GRAY, marginTop: 3, whiteSpace: "nowrap", transform: "rotate(-30deg)", transformOrigin: "top center" }}>{t.label}</div>}
              </div>
            ))}
          </div>
        </div>
      )}

      {Object.keys(catRevenue).length > 0 && (
        <div style={{ background: WHITE, borderRadius: 12, padding: 16, border: `1px solid ${BORDER}`, marginBottom: 14 }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: "#1F2937", marginBottom: 14 }}>Revenue by Category</div>
          {Object.entries(catRevenue).sort((a, b) => b[1] - a[1]).map(([cat, rev]) => (
            <div key={cat} style={{ marginBottom: 10 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 4 }}>
                <span style={{ fontWeight: 600, color: "#374151" }}>{cat}</span>
                <span style={{ color: GOLD_DARK, fontWeight: 700 }}>LKR {rev.toLocaleString()}</span>
              </div>
              <div style={{ background: GOLD_LIGHT, borderRadius: 6, height: 10, overflow: "hidden" }}>
                <div style={{ background: GOLD_DARK, height: "100%", width: `${Math.round((rev / catMax) * 100)}%`, borderRadius: 6 }} />
              </div>
            </div>
          ))}
        </div>
      )}

      {totalSales > 0 && (
        <div style={{ background: WHITE, borderRadius: 12, padding: 16, border: `1px solid ${BORDER}`, marginBottom: 14 }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: "#1F2937", marginBottom: 14 }}>Payment Methods</div>
          {Object.entries(paymentTotals).map(([method, amt]) => (
            <div key={method} style={{ marginBottom: 10 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 4 }}>
                <span style={{ fontWeight: 600, color: "#374151" }}>{method}</span>
                <span style={{ color: "#2563EB", fontWeight: 700 }}>LKR {amt.toLocaleString()}</span>
              </div>
              <div style={{ background: "#DBEAFE", borderRadius: 6, height: 10, overflow: "hidden" }}>
                <div style={{ background: "#2563EB", height: "100%", width: `${Math.round((amt / paymentMax) * 100)}%`, borderRadius: 6 }} />
              </div>
            </div>
          ))}
          <div style={{ fontSize: 11, color: GRAY, marginTop: 6 }}>Split payments are counted toward both Cash and Card for their respective portions.</div>
        </div>
      )}

      {deliverySales.length > 0 && (
        <div style={{ background: WHITE, borderRadius: 12, padding: 16, border: `1px solid ${BORDER}`, marginBottom: 14 }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: "#1F2937", marginBottom: 14 }}>Delivery</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 14 }}>
            <div style={{ background: "#FEF3C7", border: "1px solid #FCD34D", borderRadius: 8, padding: "10px 12px" }}>
              <div style={{ fontSize: 11, color: "#92400E", fontWeight: 600 }}>Owed to Riders</div>
              <div style={{ fontSize: 16, fontWeight: 800, color: "#92400E" }}>LKR {owedToRiders.toLocaleString()}</div>
              <div style={{ fontSize: 11, color: "#92400E" }}>Collected from customers, not yet paid out</div>
            </div>
            <div style={{ background: LIGHT, border: `1px solid ${BORDER}`, borderRadius: 8, padding: "10px 12px" }}>
              <div style={{ fontSize: 11, color: GRAY, fontWeight: 600 }}>Paid Direct to Riders</div>
              <div style={{ fontSize: 16, fontWeight: 800, color: "#374151" }}>LKR {paidDirectToRiders.toLocaleString()}</div>
              <div style={{ fontSize: 11, color: GRAY }}>Customers paid the rider themselves</div>
            </div>
          </div>
          {Object.entries(deliveryByMethod).map(([method, d]) => (
            <div key={method} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${BORDER}`, fontSize: 13 }}>
              <span style={{ fontWeight: 600, color: "#374151" }}>{method}</span>
              <span style={{ color: GRAY }}>{d.count} order{d.count === 1 ? "" : "s"}</span>
              <span style={{ color: "#92400E", fontWeight: 700 }}>LKR {d.collected.toLocaleString()} owed</span>
              <span style={{ color: GRAY }}>LKR {d.direct.toLocaleString()} direct</span>
            </div>
          ))}
        </div>
      )}

      {topProducts.length > 0 && (
        <div style={{ background: WHITE, borderRadius: 12, padding: 16, border: `1px solid ${BORDER}`, marginBottom: 14 }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: "#1F2937", marginBottom: 12 }}>Top Selling Products</div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", fontSize: 13, borderCollapse: "collapse" }}>
              <thead><tr style={{ borderBottom: `2px solid ${BORDER}` }}>
                {["#", "Product", "Business", "Units", "Revenue", "Trend"].map(h => <th key={h} style={{ padding: "6px 8px", textAlign: h === "Revenue" || h === "Units" || h === "Trend" ? "right" : "left", color: GRAY, fontWeight: 600, fontSize: 11 }}>{h}</th>)}
              </tr></thead>
              <tbody>
                {topProducts.map((p, i) => {
                  const prevRev = prevProdRevenue[p.name] || 0;
                  const change = prevRev > 0 ? ((p.revenue - prevRev) / prevRev) * 100 : null;
                  return (
                    <tr key={p.name} style={{ borderBottom: `1px solid ${GOLD_LIGHT}` }}>
                      <td style={{ padding: "9px 8px", color: GRAY, fontWeight: 700 }}>{i + 1}</td>
                      <td style={{ padding: "9px 8px", color: "#1F2937", fontWeight: 500 }}>{p.name}</td>
                      <td style={{ padding: "9px 8px" }}><BizBadge business={p.business} /></td>
                      <td style={{ padding: "9px 8px", textAlign: "right" }}>{p.qty}</td>
                      <td style={{ padding: "9px 8px", textAlign: "right", color: GOLD_DARK, fontWeight: 700 }}>LKR {p.revenue.toLocaleString()}</td>
                      <td style={{ padding: "9px 8px", textAlign: "right" }}>
                        {change === null
                          ? <span style={{ fontSize: 11, color: GRAY, fontWeight: 600 }}>New</span>
                          : <span style={{ fontSize: 11, fontWeight: 700, color: change > 0 ? "#059669" : change < 0 ? "#DC2626" : GRAY }}>{change > 0 ? "▲" : change < 0 ? "▼" : "–"} {Math.abs(change).toFixed(0)}%</span>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div style={{ fontSize: 11, color: GRAY, marginTop: 8 }}>Trend compares each product's revenue this period to the equivalent-length period right before it.</div>
        </div>
      )}

      {allProdSales.length > 0 && (
        <div style={{ background: WHITE, borderRadius: 12, padding: 16, border: `1px solid ${BORDER}`, marginBottom: 14 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
            <div style={{ fontWeight: 700, fontSize: 14, color: "#1F2937" }}>Product Sales Trend</div>
            <select value={selectedProductName} onChange={e => setTrendProduct(e.target.value)} style={{ ...inp, marginTop: 0, width: "auto", minWidth: 180, fontSize: 12, padding: "6px 10px" }}>
              {allProdSales.map(p => <option key={p.name} value={p.name}>{p.name}</option>)}
            </select>
          </div>
          <div style={{ display: "flex", alignItems: "flex-end", gap: 3, height: 100, overflowX: "auto" }}>
            {productTrend.map((t, i) => (
              <div key={i} style={{ display: "flex", flexDirection: "column", alignItems: "center", flex: 1, minWidth: 24 }}>
                <div title={`${t.units} sold — LKR ${t.rev.toLocaleString()}`} style={{ width: "100%", background: t.units > 0 ? RC_DARK : RC_LIGHT, borderRadius: "4px 4px 0 0", height: `${Math.max(4, Math.round((t.units / productTrendMax) * 84))}px` }} />
                {trendDays <= 14 && <div style={{ fontSize: 9, color: GRAY, marginTop: 3, whiteSpace: "nowrap", transform: "rotate(-30deg)", transformOrigin: "top center" }}>{t.label}</div>}
              </div>
            ))}
          </div>
          <div style={{ fontSize: 11, color: GRAY, marginTop: trendDays <= 14 ? 18 : 6 }}>Units sold per day for the selected product, over the current date range.</div>
        </div>
      )}

      <div style={{ background: WHITE, borderRadius: 12, padding: 16, border: `1px solid ${BORDER}`, marginBottom: 14 }}>
        <div style={{ fontWeight: 700, fontSize: 14, color: "#1F2937", marginBottom: 12 }}>Inventory Status</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 14 }}>
          {[["Total Products", fProducts.length, "#374151"], ["Out of Stock", outOfStock, "#DC2626"], ["Low (1–3)", lowStock, "#D97706"]].map(([lbl, val, col]) => (
            <div key={lbl} style={{ textAlign: "center", padding: "12px 8px", background: LIGHT, borderRadius: 10, border: `1px solid ${BORDER}` }}>
              <div style={{ fontSize: 22, fontWeight: 800, color: col }}>{val}</div>
              <div style={{ fontSize: 11, color: GRAY, marginTop: 3 }}>{lbl}</div>
            </div>
          ))}
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", fontSize: 12, borderCollapse: "collapse" }}>
            <thead><tr style={{ borderBottom: `2px solid ${BORDER}` }}>
              {["Product", "Business", "Category", "Cost", "Price", "Stock", "Stock Worth", "Stock Value"].map(h => <th key={h} style={{ padding: "6px 8px", textAlign: ["Cost", "Price", "Stock", "Stock Worth", "Stock Value"].includes(h) ? "right" : "left", color: GRAY, fontWeight: 600, fontSize: 11 }}>{h}</th>)}
            </tr></thead>
            <tbody>
              {[...fProducts].sort((a, b) => a.stock - b.stock).map(p => (
                <tr key={p.id} style={{ borderBottom: `1px solid ${GOLD_LIGHT}` }}>
                  <td style={{ padding: "8px", color: "#1F2937", fontWeight: 500 }}>{p.name}</td>
                  <td style={{ padding: "8px" }}><BizBadge business={p.business || "Blingshop"} /></td>
                  <td style={{ padding: "8px", color: GRAY }}>{p.category}</td>
                  <td style={{ padding: "8px", textAlign: "right" }}>LKR {Number(p.cost || 0).toLocaleString()}</td>
                  <td style={{ padding: "8px", textAlign: "right" }}>LKR {Number(p.price).toLocaleString()}</td>
                  <td style={{ padding: "8px", textAlign: "right" }}>
                    <span style={{ fontWeight: 700, fontSize: 12, padding: "2px 8px", borderRadius: 10, color: p.stock === 0 ? "#DC2626" : p.stock <= 3 ? "#B45309" : "#059669", background: p.stock === 0 ? "#FEE2E2" : p.stock <= 3 ? "#FEF3C7" : "#F0FDF4" }}>
                      {p.stock === 0 ? "Out of Stock" : p.stock}
                    </span>
                  </td>
                  <td style={{ padding: "8px", textAlign: "right", color: "#7C3AED", fontWeight: 600 }}>LKR {((p.cost || 0) * p.stock).toLocaleString()}</td>
                  <td style={{ padding: "8px", textAlign: "right", color: GOLD_DARK, fontWeight: 600 }}>LKR {(p.price * p.stock).toLocaleString()}</td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr style={{ borderTop: `2px solid ${BORDER}` }}>
                <td colSpan={6} style={{ padding: "10px 8px", fontWeight: 800, color: "#1F2937", fontSize: 13 }}>Totals</td>
                <td style={{ padding: "10px 8px", textAlign: "right", fontWeight: 800, color: "#7C3AED", fontSize: 13 }}>LKR {stockWorth.toLocaleString()}</td>
                <td style={{ padding: "10px 8px", textAlign: "right", fontWeight: 800, color: GOLD_DARK, fontSize: 13 }}>LKR {stockValue.toLocaleString()}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
      {filteredSales.length === 0 && <div style={{ textAlign: "center", color: "#9CA3AF", padding: "20px 0", fontSize: 14 }}>No sales data for the selected period and business.</div>}
    </div>
  );
}

// ── Customers Tab ─────────────────────────────────────────────────
function Customers({ customers, setCustomers, sales }) {
  const [search, setSearch] = useState("");
  const [modal, setModal] = useState(null);
  const [editCust, setEditCust] = useState(null);
  const [selected, setSelected] = useState(null);
  const [bizFilter, setBizFilter] = useState("All");

  const filtered = customers.filter(c => {
    const matchSearch = c.name.toLowerCase().includes(search.toLowerCase()) || (c.phone || "").includes(search);
    const matchBiz = bizFilter === "All" || c.preferredBiz === bizFilter || c.preferredBiz === "Both";
    return matchSearch && matchBiz;
  });

  const birthdays = customers.filter(c => c.birthday && getBirthdayStatus(c.birthday));

  const handleSave = async (form) => {
    if (!form.name.trim()) return;
    if (editCust) {
      const updated = { ...editCust, ...form };
      const res = await apiFetch(`/customers/${editCust.id}`, { method: "PUT", body: updated }).catch(e => ({ error: e.message }));
      if (res?.error) { alert("Failed to save customer: " + res.error); return; }
      setCustomers(cs => cs.map(c => c.id === editCust.id ? updated : c));
    } else {
      const id = generateId();
      const newC = { ...form, id, createdAt: new Date().toISOString() };
      const res = await apiFetch("/customers", { method: "POST", body: newC }).catch(e => ({ error: e.message }));
      if (res?.error) { alert("Failed to save customer: " + res.error); return; }
      setCustomers(cs => [...cs, newC]);
    }
    setModal(null); setEditCust(null);
  };

  const handleDelete = async (id) => {
    if (!confirm("Delete this customer?")) return;
    try { await apiFetch(`/customers/${id}`, { method: "DELETE" }); } catch { }
    setCustomers(cs => cs.filter(c => c.id !== id));
    if (selected?.id === id) setSelected(null);
  };

  const selectedStats = selected ? getCustomerStats(selected.id, sales) : null;
  const selectedSales = selected ? sales.filter(s => s.customerId === selected.id) : [];

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div>
          <h2 style={{ color: "#1F2937", margin: "0 0 2px", fontSize: 20, fontWeight: 800 }}>Customers</h2>
          <p style={{ margin: 0, color: GRAY, fontSize: 13 }}>{customers.length} customers registered</p>
        </div>
        <button onClick={() => { setEditCust(null); setModal("add"); }} style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 13 }}>{Icons.plus} Add Customer</button>
      </div>

      {birthdays.length > 0 && (
        <div style={{ background: WHITE, border: `1px solid #FCD34D`, borderLeft: `4px solid ${GOLD}`, borderRadius: 10, padding: 14, marginBottom: 14 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, color: GOLD_DARK, fontWeight: 700, fontSize: 14 }}>{Icons.cake} Upcoming Birthdays</div>
          {birthdays.map(c => {
            const status = getBirthdayStatus(c.birthday);
            return (
              <div key={c.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 0", borderBottom: `1px solid ${GOLD_LIGHT}`, fontSize: 13 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ color: "#1F2937", fontWeight: 600 }}>{c.name}</span>
                  {c.phone && <span style={{ color: GRAY, fontSize: 11 }}>{c.phone}</span>}
                </div>
                <span style={{ fontSize: 11, fontWeight: 700, padding: "2px 10px", borderRadius: 10, color: status.color, background: status.bg }}>{status.label}</span>
              </div>
            );
          })}
        </div>
      )}

      <div style={{ position: "relative", marginBottom: 10 }}>
        <div style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", color: GRAY }}>{Icons.search}</div>
        <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search by name or phone..." style={{ ...inp, paddingLeft: 38, marginTop: 0 }} />
      </div>
      <div style={{ marginBottom: 14 }}><BizFilter value={bizFilter} onChange={setBizFilter} /></div>

      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {filtered.map(c => {
          const stats = getCustomerStats(c.id, sales);
          const bday = getBirthdayStatus(c.birthday);
          const isOpen = selected?.id === c.id;
          return (
            <div key={c.id} style={{ background: WHITE, borderRadius: 10, border: `1px solid ${isOpen ? GOLD : BORDER}`, boxShadow: isOpen ? `0 0 0 1px ${GOLD_LIGHT}` : "0 1px 4px rgba(0,0,0,0.06)", overflow: "hidden" }}>
              <div style={{ padding: 12, display: "flex", gap: 12, alignItems: "center", cursor: "pointer" }} onClick={() => setSelected(isOpen ? null : c)}>
                <div style={{ width: 42, height: 42, borderRadius: "50%", background: GOLD_LIGHT, border: `2px solid ${GOLD}`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, color: GOLD_DARK, fontWeight: 800, fontSize: 16 }}>
                  {c.name.charAt(0).toUpperCase()}
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{ fontWeight: 700, fontSize: 14, color: "#1F2937" }}>{c.name}</span>
                    {bday && <span style={{ fontSize: 10, fontWeight: 700, padding: "1px 6px", borderRadius: 8, color: bday.color, background: bday.bg }}>{bday.label}</span>}
                  </div>
                  <div style={{ fontSize: 12, color: GRAY, marginTop: 2 }}>
                    {c.phone && <span style={{ marginRight: 10 }}>{c.phone}</span>}
                    {stats.totalSales > 0 && <span>{stats.totalSales} visit{stats.totalSales !== 1 ? "s" : ""} · LKR {stats.totalSpend.toLocaleString()}</span>}
                    {stats.totalSales === 0 && <span style={{ color: "#9CA3AF" }}>No purchases yet</span>}
                  </div>
                  {c.preferredBiz && c.preferredBiz !== "Both" && <div style={{ marginTop: 4 }}><BizBadge business={c.preferredBiz} /></div>}
                  {c.preferredBiz === "Both" && (
                    <div style={{ display: "flex", gap: 4, marginTop: 4 }}>
                      <BizBadge business="Blingshop" /><BizBadge business="RC Boutique" />
                    </div>
                  )}
                </div>
                <div style={{ display: "flex", gap: 6 }}>
                  <button type="button" onClick={e => { e.stopPropagation(); setEditCust(c); setModal("add"); }} style={{ width: 30, height: 30, display: "flex", alignItems: "center", justifyContent: "center", background: "#F3F4F6", color: "#374151", border: "none", borderRadius: 6, cursor: "pointer" }}>{Icons.edit}</button>
                  <button type="button" onClick={e => { e.stopPropagation(); handleDelete(c.id); }} style={{ width: 30, height: 30, display: "flex", alignItems: "center", justifyContent: "center", background: "#FEF2F2", color: "#DC2626", border: "none", borderRadius: 6, cursor: "pointer" }}>{Icons.trash}</button>
                </div>
              </div>

              {isOpen && (
                <div style={{ borderTop: `1px solid ${GOLD_LIGHT}`, padding: 14, background: CREAM }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 14 }}>
                    {[
                      ["Total Spend", `LKR ${stats.totalSpend.toLocaleString()}`],
                      ["Visits", stats.totalSales],
                      ["Last Visit", stats.daysSince !== null ? (stats.daysSince === 0 ? "Today" : `${stats.daysSince}d ago`) : "—"],
                    ].map(([lbl, val]) => (
                      <div key={lbl} style={{ background: WHITE, borderRadius: 8, padding: "10px 12px", border: `1px solid ${BORDER}` }}>
                        <div style={{ fontSize: 11, color: GRAY, marginBottom: 3 }}>{lbl}</div>
                        <div style={{ fontSize: 15, fontWeight: 800, color: GOLD_DARK }}>{val}</div>
                      </div>
                    ))}
                  </div>

                  <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 12 }}>
                    {c.email && <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: GRAY }}>{Icons.mail}<span>{c.email}</span></div>}
                    {c.birthday && <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: GRAY }}>{Icons.cake}<span>{new Date(c.birthday).toLocaleDateString("en-US", { month: "long", day: "numeric" })}</span></div>}
                    {c.phone && <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: GRAY }}>{Icons.phone}<span>{c.phone}</span></div>}
                  </div>

                  {c.notes && <div style={{ fontSize: 13, color: GRAY, background: WHITE, borderRadius: 8, padding: "10px 12px", border: `1px solid ${BORDER}`, marginBottom: 12 }}>{c.notes}</div>}

                  {selectedSales.length > 0 && (
                    <div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 8 }}>Purchase History</div>
                      <div style={{ maxHeight: 200, overflowY: "auto", display: "flex", flexDirection: "column", gap: 6 }}>
                        {[...selectedSales].reverse().map(s => {
                          const bizSet = [...new Set(s.items.map(i => i.business || "Blingshop"))];
                          return (
                            <div key={s.id} style={{ background: WHITE, borderRadius: 8, padding: "10px 12px", border: `1px solid ${BORDER}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                              <div>
                                <div style={{ fontSize: 13, fontWeight: 600, color: "#1F2937" }}>#{s.id}</div>
                                <div style={{ fontSize: 11, color: GRAY }}>{new Date(s.date).toLocaleDateString("en-US", { year: "numeric", month: "short", day: "numeric" })}</div>
                                <div style={{ display: "flex", gap: 4, marginTop: 4 }}>{bizSet.map(b => <BizBadge key={b} business={b} />)}</div>
                              </div>
                              <div style={{ textAlign: "right" }}>
                                <div style={{ fontWeight: 800, color: GOLD_DARK, fontSize: 14 }}>LKR {s.total.toLocaleString()}</div>
                                <div style={{ fontSize: 11, color: GRAY }}>{s.items.length} item{s.items.length !== 1 ? "s" : ""}</div>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}
                  {selectedSales.length === 0 && <div style={{ textAlign: "center", color: "#9CA3AF", padding: "12px 0", fontSize: 13 }}>No purchases linked to this customer yet.</div>}
                </div>
              )}
            </div>
          );
        })}
        {filtered.length === 0 && <div style={{ textAlign: "center", color: "#9CA3AF", padding: "40px 0", fontSize: 14 }}>No customers found</div>}
      </div>

      {modal === "add" && (
        <Modal title={editCust ? "Edit Customer" : "Add New Customer"} onClose={() => { setModal(null); setEditCust(null); }}>
          <CustomerForm initial={editCust} onSave={handleSave} onCancel={() => { setModal(null); setEditCust(null); }} />
        </Modal>
      )}
    </div>
  );
}

// ── Staff Form ────────────────────────────────────────────────────
function StaffForm({ initial, onSave, onCancel }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(initial || { name: "", username: "", password: "", role: "Salesperson", active: true });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div><label style={labelStyle}>Full Name</label><input value={form.name} onChange={e => set("name", e.target.value)} placeholder="Staff member name" style={inp} /></div>
      <div><label style={labelStyle}>Username</label><input value={form.username} onChange={e => set("username", e.target.value)} placeholder="username" style={inp} /></div>
      <div>
        <label style={labelStyle}>{isEdit ? "New Password (leave blank to keep current)" : "Password"}</label>
        <input type="password" value={form.password} onChange={e => set("password", e.target.value)} placeholder={isEdit ? "Leave blank to keep unchanged" : "Password"} style={inp} />
      </div>
      <div>
        <label style={labelStyle}>Role</label>
        <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
          {["Owner", "Salesperson"].map(r => (
            <button key={r} type="button" onClick={() => set("role", r)} style={{ flex: 1, padding: "9px", borderRadius: 8, border: `1.5px solid ${form.role === r ? GOLD : BORDER}`, background: form.role === r ? GOLD_LIGHT : WHITE, color: form.role === r ? GOLD_DARK : GRAY, fontWeight: form.role === r ? 700 : 400, cursor: "pointer", fontSize: 13 }}>{r}</button>
          ))}
        </div>
      </div>
      {isEdit && (
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button type="button" onClick={() => set("active", !form.active)} style={{ width: 44, height: 24, borderRadius: 12, border: "none", background: form.active ? GOLD : "#D1D5DB", position: "relative", cursor: "pointer" }}>
            <div style={{ width: 18, height: 18, borderRadius: "50%", background: WHITE, position: "absolute", top: 3, left: form.active ? 23 : 3, transition: "left 0.2s" }} />
          </button>
          <span style={{ fontSize: 13, color: "#374151" }}>{form.active ? "Active" : "Deactivated"}</span>
        </div>
      )}
      <div style={{ display: "flex", gap: 10, paddingTop: 4, borderTop: `1px solid ${GOLD_LIGHT}`, marginTop: 4 }}>
        <button type="button" onClick={() => onSave(form)} style={{ flex: 1, padding: "11px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 14 }}>Save Staff Member</button>
        <button type="button" onClick={onCancel} style={{ padding: "11px 20px", background: LIGHT, color: GRAY, border: `1px solid #E5E7EB`, borderRadius: 8, cursor: "pointer", fontSize: 14 }}>Cancel</button>
      </div>
    </div>
  );
}

// ── Staff Management (Settings section) ──────────────────────────
function StaffManagement({ staffList, setStaffList, currentStaff }) {
  const [modal, setModal] = useState(null);
  const [editStaff, setEditStaff] = useState(null);
  const [error, setError] = useState("");

  const handleSave = async (form) => {
    setError("");
    if (!form.name.trim() || !form.username.trim()) { setError("Name and username are required."); return; }
    if (editStaff) {
      const body = { name: form.name.trim(), username: form.username.trim(), role: form.role, active: form.active };
      if (form.password) body.password = form.password;
      const res = await apiFetch(`/staff/${editStaff.id}`, { method: "PUT", body }).catch(() => ({ success: false, error: "Backend not reachable" }));
      if (res?.error) { setError(res.error); return; }
      setStaffList(list => list.map(s => s.id === editStaff.id ? { ...s, ...body } : s));
    } else {
      if (!form.password) { setError("Password is required for a new account."); return; }
      const id = generateId();
      const res = await apiFetch("/staff", { method: "POST", body: { id, name: form.name.trim(), username: form.username.trim(), password: form.password, role: form.role } }).catch(() => ({ success: false, error: "Backend not reachable" }));
      if (res?.error) { setError(res.error); return; }
      setStaffList(list => [...list, { id, name: form.name.trim(), username: form.username.trim(), role: form.role, active: 1 }]);
    }
    setModal(null); setEditStaff(null);
  };

  const handleDelete = async (s) => {
    if (s.id === currentStaff.id) { alert("You cannot remove your own account while logged in."); return; }
    if (!confirm(`Remove ${s.name}'s account?`)) return;
    try { await apiFetch(`/staff/${s.id}`, { method: "DELETE" }); } catch { }
    setStaffList(list => list.filter(x => x.id !== s.id));
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <span style={{ fontSize: 13, color: GRAY }}>{staffList.length} staff account{staffList.length !== 1 ? "s" : ""}</span>
        <button onClick={() => { setEditStaff(null); setModal("add"); }} style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 12 }}>{Icons.plus} Add Staff</button>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {staffList.map(s => (
          <div key={s.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: LIGHT, borderRadius: 8, border: `1px solid ${BORDER}`, opacity: s.active === 0 || s.active === false ? 0.5 : 1 }}>
            <div style={{ width: 36, height: 36, borderRadius: "50%", background: GOLD_LIGHT, display: "flex", alignItems: "center", justifyContent: "center", color: GOLD_DARK, fontWeight: 800, fontSize: 14, flexShrink: 0 }}>{s.name.charAt(0).toUpperCase()}</div>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 700, fontSize: 13, color: "#1F2937" }}>{s.name} {s.id === currentStaff.id && <span style={{ fontSize: 11, color: GOLD_DARK }}>(You)</span>}</div>
              <div style={{ fontSize: 11, color: GRAY }}>@{s.username} · {s.role} {(s.active === 0 || s.active === false) && "· Deactivated"}</div>
            </div>
            <div style={{ display: "flex", gap: 6 }}>
              <button onClick={() => { setEditStaff({ ...s, password: "", active: s.active !== 0 && s.active !== false }); setModal("add"); }} style={{ width: 28, height: 28, display: "flex", alignItems: "center", justifyContent: "center", background: "#F3F4F6", color: "#374151", border: "none", borderRadius: 6, cursor: "pointer" }}>{Icons.edit}</button>
              <button onClick={() => handleDelete(s)} style={{ width: 28, height: 28, display: "flex", alignItems: "center", justifyContent: "center", background: "#FEF2F2", color: "#DC2626", border: "none", borderRadius: 6, cursor: "pointer" }}>{Icons.trash}</button>
            </div>
          </div>
        ))}
      </div>
      {modal === "add" && (
        <Modal title={editStaff ? "Edit Staff Member" : "Add Staff Member"} onClose={() => { setModal(null); setEditStaff(null); setError(""); }}>
          {error && <StatusBanner msg={`error:${error}`} />}
          <div style={{ marginTop: error ? 10 : 0 }}>
            <StaffForm initial={editStaff} onSave={handleSave} onCancel={() => { setModal(null); setEditStaff(null); setError(""); }} />
          </div>
        </Modal>
      )}
    </div>
  );
}

// ── Login Screen ──────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleLogin = async () => {
    if (!username.trim() || !password) { setError("Enter username and password."); return; }
    setLoading(true); setError("");
    try {
      const res = await apiFetch("/auth/login", { method: "POST", body: { username: username.trim(), password } });
      if (res.success) onLogin(res.staff);
      else setError(res.error || "Invalid username or password.");
    } catch {
      setError("Could not reach the server. Make sure the backend is running.");
    } finally { setLoading(false); }
  };

  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: `linear-gradient(135deg, ${GOLD_DARK}, ${GOLD})`, padding: 16 }}>
      <div style={{ background: WHITE, borderRadius: 16, padding: 32, width: "100%", maxWidth: 380, boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
        <div style={{ textAlign: "center", marginBottom: 24 }}>
          <div style={{ width: 56, height: 56, borderRadius: 14, background: GOLD_LIGHT, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}>
            <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke={GOLD_DARK} strokeWidth="2"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2" /></svg>
          </div>
          <div style={{ fontWeight: 800, fontSize: 18, color: "#1F2937" }}>Blingshop &amp; RC Boutique</div>
          <div style={{ fontSize: 12, color: GRAY, marginTop: 2 }}>Staff Sign In</div>
        </div>
        <div style={{ marginBottom: 14 }}>
          <label style={labelStyle}>Username</label>
          <input value={username} onChange={e => setUsername(e.target.value)} onKeyDown={e => e.key === "Enter" && handleLogin()} placeholder="Username" style={inp} autoFocus />
        </div>
        <div style={{ marginBottom: 6 }}>
          <label style={labelStyle}>Password</label>
          <input type="password" value={password} onChange={e => setPassword(e.target.value)} onKeyDown={e => e.key === "Enter" && handleLogin()} placeholder="Password" style={inp} />
        </div>
        <StatusBanner msg={error ? `error:${error}` : ""} />
        <button onClick={handleLogin} disabled={loading} style={{ width: "100%", marginTop: 16, padding: "12px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 10, cursor: loading ? "default" : "pointer", fontWeight: 800, fontSize: 15, opacity: loading ? 0.7 : 1 }}>
          {loading ? "Signing in..." : "Sign In"}
        </button>
      </div>
    </div>
  );
}

// ── Settings ──────────────────────────────────────────────────────
function Settings({ staffList, setStaffList, currentStaff, zebraIP, setZebraIP }) {
  const [ipInput, setIpInput] = useState(zebraIP);
  const s = { background: WHITE, borderRadius: 12, padding: 20, boxShadow: "0 1px 4px rgba(0,0,0,0.06)", border: `1px solid ${BORDER}`, marginBottom: 14 };
  const t = { margin: "0 0 14px", fontSize: 14, fontWeight: 800, color: "#1F2937", paddingBottom: 10, borderBottom: `1px solid ${BORDER}` };
  return (
    <div>
      <div style={{ marginBottom: 20 }}><h2 style={{ color: "#1F2937", margin: "0 0 2px", fontSize: 20, fontWeight: 800 }}>Settings</h2><p style={{ margin: 0, color: GRAY, fontSize: 13 }}>Configure your POS system</p></div>
      <div style={s}>
        <h4 style={t}>Staff Accounts</h4>
        <StaffManagement staffList={staffList} setStaffList={setStaffList} currentStaff={currentStaff} />
      </div>
      <div style={s}>
        <h4 style={t}>Businesses</h4>
        {BUSINESSES.map(b => {
          const bs = BIZ_STYLE[b];
          return (
            <div key={b} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 0", borderBottom: `1px solid ${GOLD_LIGHT}` }}>
              <div style={{ width: 10, height: 10, borderRadius: "50%", background: bs.primary, flexShrink: 0 }} />
              <div><div style={{ fontWeight: 700, color: "#1F2937", fontSize: 14 }}>{b}</div><div style={{ fontSize: 11, color: GRAY }}>{b === "Blingshop" ? "Jewelry & Accessories" : "Clothing, Bags, Shoes, Shawls"}</div></div>
            </div>
          );
        })}
      </div>
      <div style={s}>
        <h4 style={t}>Zebra ZD421T — Network Settings</h4>
        <p style={{ fontSize: 13, color: GRAY, margin: "0 0 12px" }}>Enter the printer IP for direct ZPL network printing.</p>
        <div style={{ display: "flex", gap: 8 }}>
          <input value={ipInput} onChange={e => setIpInput(e.target.value)} placeholder="e.g. 192.168.1.100" style={{ ...inp, marginTop: 0, flex: 1 }} />
          <button onClick={() => setZebraIP(ipInput)} style={{ padding: "9px 18px", background: GOLD_DARK, color: WHITE, border: "none", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 13 }}>Save</button>
        </div>
        {zebraIP && <div style={{ marginTop: 10, fontSize: 12, color: "#059669", fontWeight: 600 }}>Zebra configured at: {zebraIP}</div>}
        <p style={{ fontSize: 12, color: "#9CA3AF", marginTop: 10, marginBottom: 0 }}>Hold Feed button for 2 seconds on ZD421T to print its IP address.</p>
      </div>
      <div style={s}>
        <h4 style={t}>ABM P323B — Bluetooth</h4>
        <p style={{ fontSize: 13, color: GRAY, margin: 0 }}>No configuration required. Use the Bluetooth option in Print Receipt. Requires Chrome.</p>
      </div>
    </div>
  );
}

const TAB_ICONS = { Dashboard: Icons.dashboard, Inventory: Icons.box, POS: Icons.cart, "Sales History": Icons.history, Reports: Icons.reports, Customers: Icons.users, Settings: Icons.settings };

// ── App ───────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("POS");
  const [products, setProductsState] = useState([]);
  const [sales, setSalesState] = useState([]);
  const [customers, setCustomersState] = useState([]);
  const [staffList, setStaffListState] = useState([]);
  const [refunds, setRefundsState] = useState([]);
  const [currentStaff, setCurrentStaff] = useState(null);
  const [zebraIP, setZebraIPState] = useState("");
  const [blingshopThbRate, setBlingshopThbRateState] = useState("");
  const [rcThbRate, setRcThbRateState] = useState("");
  const [exchangeCredit, setExchangeCredit] = useState(null);
  const [booting, setBooting] = useState(true);
  const [loading, setLoading] = useState(false);

  // Restore a saved login session on first load
  useEffect(() => {
    const saved = localStorage.getItem("blingshop_staff");
    if (saved) {
      try { setCurrentStaff(JSON.parse(saved)); } catch { }
    }
    setBooting(false);
  }, []);

  // Load app data once a staff member is logged in
  useEffect(() => {
    if (!currentStaff) return;
    setLoading(true);
    Promise.all([apiFetch("/products"), apiFetch("/sales"), apiFetch("/customers"), apiFetch("/staff"), apiFetch("/settings"), apiFetch("/refunds")])
      .then(([prods, sls, custs, staff, settings, refs]) => {
        setProductsState(Array.isArray(prods) ? prods : []);
        setSalesState(Array.isArray(sls) ? sls : []);
        setCustomersState(Array.isArray(custs) ? custs : []);
        setStaffListState(Array.isArray(staff) ? staff : []);
        setRefundsState(Array.isArray(refs) ? refs : []);
        setZebraIPState(settings?.zebra_ip || "");
        setBlingshopThbRateState(settings?.blingshop_thb_rate || "");
        setRcThbRateState(settings?.rc_thb_rate || "");
        setLoading(false);
      })
      .catch(() => { console.warn("Backend not reachable — running in local preview mode."); setLoading(false); });
  }, [currentStaff]);

  const handleLogin = (staff) => {
    setCurrentStaff(staff);
    localStorage.setItem("blingshop_staff", JSON.stringify(staff));
  };

  const handleLogout = () => {
    if (!confirm("Log out?")) return;
    setCurrentStaff(null);
    localStorage.removeItem("blingshop_staff");
    setProductsState([]); setSalesState([]); setCustomersState([]); setStaffListState([]); setRefundsState([]);
    setExchangeCredit(null);
    setTab("POS");
  };

  const addSale = async (sale) => {
    const stamped = { ...sale, staffId: currentStaff?.id || null, staffName: currentStaff?.name || null };
    try { await apiFetch("/sales", { method: "POST", body: stamped }); } catch { }
    setSalesState(s => [...s, stamped]);
  };

  const addRefund = async (refund) => {
    try { await apiFetch("/refunds", { method: "POST", body: refund }); } catch { }
    setRefundsState(rs => [...rs, refund]);
    setProductsState(ps => ps.map(p => {
      const ri = refund.items.find(i => i.product_id === p.id);
      return ri ? { ...p, stock: p.stock + ri.qty } : p;
    }));
  };

  const setZebraIP = async (ip) => {
    setZebraIPState(ip);
    try { await apiFetch("/settings/zebra_ip", { method: "PUT", body: { value: ip } }); } catch { }
  };

  const setBlingshopThbRate = async (rate) => {
    setBlingshopThbRateState(rate);
    try { await apiFetch("/settings/blingshop_thb_rate", { method: "PUT", body: { value: rate } }); } catch { }
  };

  const setRcThbRate = async (rate) => {
    setRcThbRateState(rate);
    try { await apiFetch("/settings/rc_thb_rate", { method: "PUT", body: { value: rate } }); } catch { }
  };

  const visibleTabs = currentStaff?.role === "Owner" ? TABS : SALESPERSON_TABS;

  if (booting) {
    return <div style={{ minHeight: "100vh", background: WHITE }} />;
  }

  if (!currentStaff) {
    return <LoginScreen onLogin={handleLogin} />;
  }

  if (loading) return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: WHITE, fontFamily: "-apple-system, sans-serif" }}>
      <div style={{ textAlign: "center" }}>
        <div style={{ width: 48, height: 48, border: `4px solid ${GOLD_LIGHT}`, borderTop: `4px solid ${GOLD_DARK}`, borderRadius: "50%", animation: "spin 0.8s linear infinite", margin: "0 auto 16px" }} />
        <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
        <div style={{ color: GOLD_DARK, fontWeight: 700, fontSize: 16 }}>Loading POS...</div>
        <div style={{ color: GRAY, fontSize: 13, marginTop: 4 }}>Connecting to database</div>
      </div>
    </div>
  );

  return (
    <div style={{ background: WHITE, minHeight: "100vh", fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif" }}>
      <div style={{ background: GOLD_DARK, padding: "0 20px", height: 56, display: "flex", alignItems: "center", justifyContent: "space-between", boxShadow: "0 2px 8px rgba(0,0,0,0.15)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 32, height: 32, borderRadius: 8, background: "rgba(255,255,255,0.15)", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke={WHITE} strokeWidth="2"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2" /></svg>
          </div>
          <div>
            <div style={{ color: WHITE, fontWeight: 800, fontSize: 15, letterSpacing: 1 }}>Blingshop &amp; RC Boutique</div>
            <div style={{ color: "rgba(255,255,255,0.6)", fontSize: 10, letterSpacing: 2, textTransform: "uppercase" }}>Point of Sale</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, background: "rgba(255,255,255,0.12)", padding: "5px 12px", borderRadius: 6 }}>
            <span style={{ color: "rgba(255,255,255,0.7)" }}>{Icons.user}</span>
            <span style={{ color: WHITE, fontSize: 12, fontWeight: 600 }}>{currentStaff.name}</span>
            <span style={{ color: "rgba(255,255,255,0.55)", fontSize: 11 }}>· {currentStaff.role}</span>
          </div>
          <button onClick={handleLogout} style={{ background: "rgba(255,255,255,0.12)", border: "none", color: WHITE, padding: "7px 12px", borderRadius: 6, cursor: "pointer", fontSize: 12, fontWeight: 700 }}>Logout</button>
        </div>
      </div>
      <div style={{ background: WHITE, display: "flex", overflowX: "auto", borderBottom: "1px solid #E5E7EB", padding: "0 4px" }}>
        {visibleTabs.map(t => (<button key={t} onClick={() => setTab(t)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "14px 14px 12px", background: "none", border: "none", borderBottom: `2px solid ${tab === t ? GOLD : "transparent"}`, color: tab === t ? GOLD_DARK : GRAY, fontWeight: tab === t ? 700 : 400, cursor: "pointer", fontSize: 13, whiteSpace: "nowrap", flexShrink: 0 }}>{TAB_ICONS[t]}{t}</button>))}
      </div>
      <div style={{ padding: 16, maxWidth: tab === "POS" ? 1400 : 640, margin: "0 auto" }}>
        {tab === "Dashboard" && <Dashboard products={products} sales={sales} currentStaff={currentStaff} />}
        {tab === "Inventory" && <Inventory products={products} setProducts={setProductsState} zebraIP={zebraIP} currentStaff={currentStaff} blingshopThbRate={blingshopThbRate} setBlingshopThbRate={setBlingshopThbRate} rcThbRate={rcThbRate} setRcThbRate={setRcThbRate} />}
        {tab === "POS" && <POS products={products} setProducts={setProductsState} addSale={addSale} customers={customers} setCustomers={setCustomersState} exchangeCredit={exchangeCredit} clearExchangeCredit={() => setExchangeCredit(null)} currentStaff={currentStaff} />}
        {tab === "Sales History" && <SalesHistory sales={sales} refunds={refunds} addRefund={addRefund} setProducts={setProductsState} currentStaff={currentStaff} onExchange={credit => { setExchangeCredit(credit); setTab("POS"); }} />}
        {tab === "Reports" && currentStaff.role === "Owner" && <Reports products={products} sales={sales} />}
        {tab === "Customers" && <Customers customers={customers} setCustomers={setCustomersState} sales={sales} />}
        {tab === "Settings" && currentStaff.role === "Owner" && <Settings staffList={staffList} setStaffList={setStaffListState} currentStaff={currentStaff} zebraIP={zebraIP} setZebraIP={setZebraIP} />}
      </div>
    </div>
  );
}