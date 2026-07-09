import React, { useState, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import {
  Upload, FileSpreadsheet, AlertTriangle, Database, Download,
  Sparkles, Search, ChevronDown, Info, Package, Sun, Moon, Copy,
} from "lucide-react";

/* ------------------------------------------------------------------ *
 *  Transformation rules (derived from the SSH-27 workflow)
 * ------------------------------------------------------------------ */

const ENTITY_MAP = {
  FRCA: { solids: "CARREFOUR FRANCE",      prints: "FRANCE PRINTS" },
  FRCH: { solids: "CARREFOUR FRANCE FRCH", prints: "FRANCE PRINTS" },
  PFCA: { solids: "CARREFOUR FRANCE",      prints: "FRANCE PRINTS" },
  REC5: { solids: "CARREFOUR FRANCE",      prints: "FRANCE PRINTS" },
  MACA: { solids: "CARREFOUR MORROCO",     prints: "MORROCO PRINTS" },
  BECA: { solids: "CARREFOUR BELGIUM",     prints: "BELGIUM PRINTS" },
  ROCA: { solids: "CARREFOUR ROMANIA",     prints: "ROMANIA PRINTS" },
  ESCA: { solids: "CARREFOUR SPAIN",       prints: "SPAIN PRINTS" },
  POCA: { solids: "CARREFOUR POLAND",      prints: "POLAND PRINTS" },
  AECA: { solids: "CARREFOUR DUBAI",       prints: "DUBAI PRINTS" },
};

// Description prefix / keyword -> article (French Carrefour nomenclature)
const ARTICLE_RULES = [
  { re: /\bDRAP\s*HOUSSE\b|\bDH\d?\b/i, article: "FITTED SHEET",  print: false },
  { re: /\bDRAP\s*PLAT\b|\bDP\d?\b/i,   article: "FLAT SHEET",    print: false },
  { re: /\bTAIE|\bTO\b|\bTT\b/i,        article: "PILLOW CASE",   print: false },
  { re: /\bTRAVERSIN|\bBOLSTER/i,       article: "BOLSTER",       print: false },
  { re: /\bPARURE|\bHDC\b|\bDUVET|\bQUILT/i, article: "DUVET/QUILT COVER SET", print: true },
];

const PRINT_DESIGN_TOKENS = [
  "MICROSAND","MICROPALM","MICROPANSY","MICROLINE","MICROFLOR","THEA","OLIVE",
  "SAGARA","DEHLI","DIP DYE","BLEUET","NYC","CHRIS","DASH","SPRINGS","MICROFLORA",
];

const num = (v) => (v === null || v === undefined || v === "" || isNaN(+v) ? null : +v);
const clean = (s) => (s === null || s === undefined ? "" : String(s).trim());

function parsePrice(raw) {
  if (raw === null || raw === undefined) return null;
  const s = String(raw);
  const m = s.match(/(\d+(?:\.\d+)?)\s*$/); // trailing number after last underscore/space
  return m ? +m[1] : num(raw);
}
function parseSeason(basket) {
  const p = clean(basket).slice(0, 5).toUpperCase();
  if (p === "SSH27") return "SSH-27";
  if (p === "AWH26") return "AWH-26";
  if (/^\d?PV27/.test(clean(basket))) return "SSH-27"; // Spain PV variant
  return p || "";
}
function parseEntity(basket) {
  const b = clean(basket).toUpperCase();
  // standard: chars 6-9 (after 5-char season). PV variant: after the PV segment.
  let code = b.slice(5, 9);
  if (!ENTITY_MAP[code]) {
    const m = b.match(/(FRCA|FRCH|PFCA|REC5|MACA|BECA|ROCA|ESCA|POCA|AECA)/);
    if (m) code = m[1];
  }
  return code;
}
function deriveArticle(desc) {
  for (const r of ARTICLE_RULES) if (r.re.test(desc)) return r;
  return { article: "", print: null };
}
function looksPrinted(desc) {
  const up = desc.toUpperCase();
  return PRINT_DESIGN_TOKENS.some((t) => up.includes(t));
}
function parseSize(desc) {
  const m = String(desc).match(/(\d{2,3})\s*[xX]\s*(\d{2,3})/);
  return m ? `${m[1]}x${m[2]}` : "";
}
function deriveCategory(desc) {
  const up = String(desc).toUpperCase();
  if (/MICRO/.test(up)) return "MICROFIBER";
  if (/\bALG\b|COTON|COTTON|BIO/.test(up)) return "BIO COTTON";
  return "";
}
function addDays(d, n) {
  if (!(d instanceof Date) || isNaN(d)) return null;
  const r = new Date(d); r.setDate(r.getDate() + n); return r;
}
function fmtDate(d) {
  if (!(d instanceof Date) || isNaN(d)) return "";
  return d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}
function toDate(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === "number") { // Excel serial → JS date
    const d = new Date(Math.round((v - 25569) * 86400000));
    return isNaN(d) ? null : d;
  }
  if (typeof v === "string" && v.trim()) { const d = new Date(v); return isNaN(d) ? null : d; }
  return null;
}

const STATUS_META = {
  OKPOOL: { label: "OK to pool",   tone: "green",  note: "Validated — clear to commit" },
  PBPOOL: { label: "Pool issue",   tone: "amber",  note: "Pooling problem — resolve before booking" },
  PBMOQ:  { label: "Below MOQ",    tone: "red",    note: "Under minimum order qty — resolve before booking" },
};

/* ------------------------------------------------------------------ *
 *  Core: build enriched line from a raw basket row
 * ------------------------------------------------------------------ */
function buildLine(row, dbIndex) {
  const basket = clean(row.Basket);
  const code = clean(row.ProductCode);
  const desc = clean(row.Description);
  const season = parseSeason(basket);
  const entityCode = parseEntity(basket);
  const price = parsePrice(row.Price);
  const qty = num(row.Quantity);

  const dbHit = dbIndex ? dbIndex[code] : null;
  const isNew = !dbHit;

  // route: repeats follow their existing DB sheet (authoritative); new -> heuristic
  const art = deriveArticle(desc);
  let isPrint;
  if (dbHit) isPrint = dbHit.sheet === "prints";
  else isPrint = art.print === true || looksPrinted(desc);
  const sheet = isPrint ? "prints" : "solids";

  // Customer/programme belongs to THIS order line → take it from the basket entity code.
  // (A shared basic code can sell to several entities, so the DB owner is only a fallback.)
  const entity = ENTITY_MAP[entityCode];
  const routedName = entity ? (isPrint ? entity.prints : entity.solids) : (dbHit?.owner || "");
  const ownerVerify = !entity && !dbHit?.owner;

  // attributes: prefer DB (repeat), else derive
  const article  = dbHit?.article  || art.article;
  const sizeName = dbHit?.sizeName  || "";
  const size     = parseSize(desc)  || dbHit?.size || "";
  const category = dbHit?.category  || deriveCategory(desc);
  const quality  = dbHit?.quality   || deriveCategory(desc);
  const design   = dbHit?.design    || (isPrint ? PRINT_DESIGN_TOKENS.find((t)=>desc.toUpperCase().includes(t)) || "" : "MICROFIBER");

  const pcb = num(dbHit?.pcb); // basket has no PCB; repeats inherit it
  const len = num(row.Lenght), wid = num(row.Width), hei = num(row.Height);
  const netWtCtn = num(row.NetWeight), grossWtCtn = num(row.GrossWeight);
  const cartons = pcb && qty ? Math.round(qty / pcb) : null;
  const netWtPcs = pcb && netWtCtn ? +(netWtCtn / pcb).toFixed(4) : num(dbHit?.netWtPcs);
  const totalNetWt = cartons && netWtCtn ? +(netWtCtn * cartons).toFixed(2) : null;
  const totalGrossWt = cartons && grossWtCtn ? +(grossWtCtn * cartons).toFixed(2) : null;
  const cbm = len && wid && hei ? +((len * wid * hei) / 1e6).toFixed(6) : null;
  const netCbm = cbm && cartons ? +(cbm * cartons).toFixed(4) : null;
  const value = qty && price ? +(qty * price).toFixed(2) : null;

  const initFri = toDate(row.Fri_date);
  const initEtd = toDate(row.ETD) || toDate(row.Initial_ETD);
  const ytmShip = addDays(initFri, -15);
  const weekDate = addDays(initFri, -67);

  const status = clean(row.StatusCode).toUpperCase();
  const flowType = clean(row.FlowType);

  return {
    _sheet: sheet, isNew, status, ownerVerify,
    season, basket, entityCode, owner: routedName,
    code, description: desc, color: clean(row.Color),
    article, sizeName, size, category, quality, design,
    packing: clean(row.PackingComments),
    qty, price, value, pcb, cartons,
    netWtCtn, grossWtCtn, netWtPcs, totalNetWt, totalGrossWt,
    len, wid, hei, cbm, netCbm,
    initFri, initEtd, ytmShip, weekDate, flowType,
    missingPcb: !pcb,
  };
}

/* ------------------------------------------------------------------ *
 *  DB index builder (from an uploaded current database workbook)
 * ------------------------------------------------------------------ */
// Natural key for a physical order line: code + colour + basket.
const dupKey = (code, color, basket) =>
  `${clean(code).toUpperCase()}|${clean(color).toUpperCase()}|${clean(basket).toUpperCase()}`;

function buildDbIndex(wb) {
  const codes = {};
  const keys = new Set();
  const g = (r, ks) => { for (const k of ks) if (r[k] !== undefined && r[k] !== null && r[k] !== "") return r[k]; return ""; };
  if (wb.Sheets["Sheet1"]) {
    XLSX.utils.sheet_to_json(wb.Sheets["Sheet1"], { defval: "" }).forEach((r) => {
      const c = clean(g(r, ["PRODUCT CODE"])); if (!c) return;
      keys.add(dupKey(c, g(r, ["COLOR 1"]), g(r, ["BASKET"])));
      codes[c] = codes[c] || {
        sheet: "solids", owner: clean(g(r, ["CUSTOMER"])),
        article: clean(g(r, ["ARTICLE"])), sizeName: clean(g(r, ["SIZE NAME"])),
        size: clean(g(r, ["SIZE"])), category: clean(g(r, ["CATEGORY"])),
        quality: clean(g(r, ["QUALITY"])), design: clean(g(r, ["DESIGN"])),
        pcb: g(r, ["PCB"]), netWtPcs: g(r, ["NET WT/PCS"]),
      };
    });
  }
  if (wb.Sheets["Sheet2"]) {
    XLSX.utils.sheet_to_json(wb.Sheets["Sheet2"], { defval: "" }).forEach((r) => {
      const c = clean(g(r, ["PRODUCT CODE"])); if (!c) return;
      keys.add(dupKey(c, g(r, ["color"]), g(r, ["Basket"])));
      if (codes[c]) return;
      codes[c] = {
        sheet: "prints", owner: clean(g(r, ["PROGRAM"])),
        article: clean(g(r, ["ARTICLE"])), sizeName: clean(g(r, ["SIZE NAME"])),
        size: clean(g(r, ["SIZE"])), category: clean(g(r, ["QUALITY"])),
        quality: clean(g(r, ["QUALITY"])), design: clean(g(r, ["DESIGN"])),
        pcb: g(r, ["PCB"]), netWtPcs: "",
      };
    });
  }
  return { codes, keys };
}

/* ------------------------------------------------------------------ *
 *  Export: DB-column-ordered sheets
 * ------------------------------------------------------------------ */
const SOLID_COLS = ["CUSTOMER","CATEGORY","QUALITY","SEASON","YTM","BASKET","PSS","SHIPPING MARKS","TAG CARDS","ORDER TYPE","PRODUCT CODE","DESCRIPTION","DESIGN","PACKING COMMENTS","COLOR 1","COLOR CODE","ARTICLE","SIZE NAME","SIZE","ORDER  QTY","CUTSIZE","MTR","CANCELLED QTY","PCB","NO. OF CARTONS","NET WT/PCS","NET WT/CTN","GROSS WT/CTN","TOTAL NET WT.","TOTAL GROSS WT.","L","W","H","CBM","Net CBM","PRICE","TOTAL VALUE","WEEK DATE","# OF DAYS","YTM SHIP DATES\n(INTERNAL FRI)","INITIAL FRI","INITIAL ETD","Remarks"];
const PRINT_COLS = ["PROGRAM","QUALITY","Season","YTM#","SHIPPING MARK","BRAND NAME","PSS","Basket","FlowType","Packaging status","PRODUCT CODE","Description","DETAIL SIZE DESCRIPTION","NEW/REPEAT","MASTER CONTRACT","SAM","DESIGN","widths","color","ARTICLE","SIZE NAME","SIZE","Quantity","cancelled qty","PCB","ORDER CARTONS","NET WT/PCS","NET WT/CTN","GrossWeight","TOTAL NET WT.","TOTAL GROSS WT.","L","W","H","CBM","TOTAL VOLUME","TAG CARDS/yellow tags","LINE PRODUCT CODE (CARTON STICKERS)","PRICE","TOTAL VALUE","Week Date","# of Days","YTM SHIP DATES","INITIAL FRI","ETD","COMMENTS","MERGED","ETD REMAKRS","GREIGE FILE"];

const D = (d) => (d instanceof Date && !isNaN(d) ? d : "");
const B = (v) => (v === null || v === undefined ? "" : v);

function solidRow(L) {
  return {
    CUSTOMER: L.owner, CATEGORY: L.category, QUALITY: L.quality, SEASON: L.season,
    YTM: "", BASKET: L.basket, PSS: "", "SHIPPING MARKS": "", "TAG CARDS": "",
    "ORDER TYPE": L.flowType, "PRODUCT CODE": L.code, DESCRIPTION: L.description,
    DESIGN: L.design, "PACKING COMMENTS": L.packing, "COLOR 1": L.color, "COLOR CODE": "",
    ARTICLE: L.article, "SIZE NAME": L.sizeName, SIZE: L.size, "ORDER  QTY": B(L.qty),
    CUTSIZE: "", MTR: "", "CANCELLED QTY": "", PCB: B(L.pcb), "NO. OF CARTONS": B(L.cartons),
    "NET WT/PCS": B(L.netWtPcs), "NET WT/CTN": B(L.netWtCtn), "GROSS WT/CTN": B(L.grossWtCtn),
    "TOTAL NET WT.": B(L.totalNetWt), "TOTAL GROSS WT.": B(L.totalGrossWt),
    L: B(L.len), W: B(L.wid), H: B(L.hei), CBM: B(L.cbm), "Net CBM": B(L.netCbm),
    PRICE: B(L.price), "TOTAL VALUE": B(L.value), "WEEK DATE": D(L.weekDate), "# OF DAYS": 52,
    "YTM SHIP DATES\n(INTERNAL FRI)": D(L.ytmShip), "INITIAL FRI": D(L.initFri),
    "INITIAL ETD": D(L.initEtd), Remarks: L.isNew ? "NEW — await PSS" : (STATUS_META[L.status]?.tone !== "green" ? STATUS_META[L.status]?.label || "" : ""),
  };
}
function printRow(L) {
  return {
    PROGRAM: L.owner, QUALITY: L.quality, Season: L.season, "YTM#": "", "SHIPPING MARK": "",
    "BRAND NAME": "", PSS: "", Basket: L.basket, FlowType: L.flowType, "Packaging status": "",
    "PRODUCT CODE": L.code, Description: L.description, "DETAIL SIZE DESCRIPTION": "",
    "NEW/REPEAT": L.isNew ? "NEW" : "REPEAT", "MASTER CONTRACT": "", SAM: "", DESIGN: L.design,
    widths: "", color: L.color, ARTICLE: L.article, "SIZE NAME": L.sizeName, SIZE: L.size,
    Quantity: B(L.qty), "cancelled qty": "", PCB: B(L.pcb), "ORDER CARTONS": B(L.cartons),
    "NET WT/PCS": B(L.netWtPcs), "NET WT/CTN": B(L.netWtCtn), GrossWeight: B(L.grossWtCtn),
    "TOTAL NET WT.": B(L.totalNetWt), "TOTAL GROSS WT.": B(L.totalGrossWt),
    L: B(L.len), W: B(L.wid), H: B(L.hei), CBM: B(L.cbm), "TOTAL VOLUME": B(L.netCbm),
    "TAG CARDS/yellow tags": "", "LINE PRODUCT CODE (CARTON STICKERS)": "",
    PRICE: B(L.price), "TOTAL VALUE": B(L.value), "Week Date": D(L.weekDate), "# of Days": 52,
    "YTM SHIP DATES": D(L.ytmShip), "INITIAL FRI": D(L.initFri), ETD: D(L.initEtd),
    COMMENTS: L.isNew ? "NEW — await PSS" : "", MERGED: "", "ETD REMAKRS": "", "GREIGE FILE": "",
  };
}
// Append plain row-objects into an existing ExcelJS worksheet, immediately
// after the LAST ROW THAT ACTUALLY HAS DATA — inheriting that row's formatting
// (font, fill, borders, number formats). Merchandising sheets often carry a
// band of pre-formatted but empty rows below the data (e.g. Sheet1 data ends at
// row 1083 while the sheet range runs to 1367); ExcelJS's addRow() would append
// after that band and leave a blank gap, so we find the real last row ourselves.
function appendRowsToSheet(ws, rowObjects) {
  if (!ws || !rowObjects.length) return;

  // Map header text -> column index from row 1.
  const headerRow = ws.getRow(1);
  const headers = [];              // col -> header text
  const colByHeader = {};
  headerRow.eachCell({ includeEmpty: true }, (cell, col) => {
    const h = String(cell.value ?? "").trim();
    headers[col] = h;
    if (h) colByHeader[h] = col;
  });
  const lastCol = headers.length - 1;

  // Columns that reliably mark a row as real data (ignores blank, styled-only rows).
  const anchorCols = ["PRODUCT CODE", "Basket", "BASKET", "PROGRAM", "CUSTOMER"]
    .map((h) => colByHeader[h])
    .filter(Boolean);
  if (!anchorCols.length) anchorCols.push(1);

  const scanTo = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  let lastDataRow = 1;
  for (let r = 2; r <= scanTo; r++) {
    const row = ws.getRow(r);
    const hasData = anchorCols.some((c) => {
      const v = row.getCell(c).value;
      return v !== null && v !== undefined && String(v).trim() !== "";
    });
    if (hasData) lastDataRow = r;
  }

  // Copy formatting from the last populated row (fall back to header only if empty).
  const styleRow = ws.getRow(lastDataRow >= 2 ? lastDataRow : 1);

  rowObjects.forEach((obj, i) => {
    const row = ws.getRow(lastDataRow + 1 + i);
    for (let col = 1; col <= lastCol; col++) {
      const h = headers[col];
      const cell = row.getCell(col);
      if (lastDataRow >= 2) cell.style = styleRow.getCell(col).style; // inherit formatting
      cell.value = h && obj[h] !== undefined && obj[h] !== "" ? obj[h] : null;
    }
  });
}

// Preferred path: append the new/repeat lines directly into the database
// workbook the user uploaded, so its fonts, colours, column widths, and
// number formats carry straight through.
async function exportIntoDatabase(lines, includeDupes, dbFile) {
  const usable = includeDupes ? lines : lines.filter((l) => !l.isDup);
  const solids = usable.filter((l) => l._sheet === "solids").map(solidRow);
  const prints = usable.filter((l) => l._sheet === "prints").map(printRow);

  const buf = await dbFile.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);

  appendRowsToSheet(wb.getWorksheet("Sheet1"), solids);
  appendRowsToSheet(wb.getWorksheet("Sheet2"), prints);

  const out = await wb.xlsx.writeBuffer();
  const blob = new Blob([out], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = dbFile.name.replace(/\.xlsx?$/i, "") + "_updated.xlsx";
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}

// Fallback when no database has been uploaded: nothing to inherit formatting
// from, so build a plain new workbook the same way as before.
function downloadWorkbook(lines, includeDupes) {
  const usable = includeDupes ? lines : lines.filter((l) => !l.isDup);
  const wb = XLSX.utils.book_new();
  const solids = usable.filter((l) => l._sheet === "solids").map(solidRow);
  const prints = usable.filter((l) => l._sheet === "prints").map(printRow);
  const ws1 = XLSX.utils.json_to_sheet(solids.length ? solids : [{}], { header: SOLID_COLS });
  const ws2 = XLSX.utils.json_to_sheet(prints.length ? prints : [{}], { header: PRINT_COLS });
  XLSX.utils.book_append_sheet(wb, ws1, "SOLIDS");
  XLSX.utils.book_append_sheet(wb, ws2, "PRINTS");
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([out], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = "basket_ready_for_database.xlsx";
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}

/* ------------------------------------------------------------------ *
 *  UI
 * ------------------------------------------------------------------ */
const CSS = `
.bic{
  --ink:#131A26; --ink2:#3A4658; --muted:#727E90; --line:#E3E7EE; --line2:#EDF0F5;
  --paper:#EBEEF3; --paper2:#F5F7FA; --card:#FFFFFF; --navy:#1F3864; --teal:#2E6C6E;
  --green:#1F7A54; --amber:#B0741A; --red:#B4341F; --violet:#6B4E9E; --steel:#59697F;
  --b-new-bg:#F0E9FA; --b-new-fg:#5E4390; --b-rep-bg:#E6F1EA; --b-rep-fg:#1B6E4B;
  --b-sol-bg:#E9EDF3; --b-sol-fg:#4C5C73; --b-pri-bg:#E1F0EE; --b-pri-fg:#276460;
  --b-dup-bg:#FBE7E1; --b-dup-fg:#A82F1C;
  --bn-warn-bg:#FCF5E8; --bn-warn-bd:#EAD9B0; --bn-warn-fg:#6B4E12;
  --bn-info-bg:#ECF2FA; --bn-info-bd:#CFDDF0; --bn-info-fg:#264566;
  --thead:#F6F8FB; --hover:#F7F9FC; --isnew-bg:#FBFAFF; --isdup-bg:#FCF5F2;
  --chip-on:#EAF1FA; --ok-tint:#F1FAF8;
  --grad:linear-gradient(120deg,var(--navy),var(--teal));
  --shadow-sm:0 1px 2px rgba(20,32,54,.05), 0 1px 3px rgba(20,32,54,.04);
  --shadow:0 2px 4px rgba(20,32,54,.05), 0 6px 16px rgba(20,32,54,.07);
  --shadow-lg:0 10px 34px rgba(20,32,54,.12);
  --ring:0 0 0 3px rgba(31,56,100,.22);
  --mono:"JetBrains Mono",ui-monospace,"SF Mono",Menlo,Consolas,monospace;
  --sans:"Inter",system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;
}
.bic.dark{
  --ink:#E7EBF1; --ink2:#AEB9C7; --muted:#7A8698; --line:#2A3341; --line2:#212934;
  --paper:#0C1017; --paper2:#0F141C; --card:#161D28; --navy:#7A9AD8; --teal:#54BAB0;
  --green:#46B482; --amber:#D6A441; --red:#E0684F; --violet:#A98BE0; --steel:#8496AE;
  --b-new-bg:rgba(169,139,224,.18); --b-new-fg:#C3AEEF; --b-rep-bg:rgba(70,180,130,.18); --b-rep-fg:#6FD0A2;
  --b-sol-bg:rgba(132,150,174,.20); --b-sol-fg:#AEBCD0; --b-pri-bg:rgba(84,186,176,.20); --b-pri-fg:#79D6CB;
  --b-dup-bg:rgba(224,104,79,.20); --b-dup-fg:#EE977F;
  --bn-warn-bg:#241E13; --bn-warn-bd:#463A1E; --bn-warn-fg:#E6CC92;
  --bn-info-bg:#151F2D; --bn-info-bd:#274061; --bn-info-fg:#ADC6EB;
  --thead:#1A212C; --hover:#1A212C; --isnew-bg:#191726; --isdup-bg:#241A18;
  --chip-on:#1B2740; --ok-tint:#12201E;
  --shadow-sm:0 1px 2px rgba(0,0,0,.30);
  --shadow:0 2px 6px rgba(0,0,0,.34), 0 10px 22px rgba(0,0,0,.30);
  --shadow-lg:0 14px 40px rgba(0,0,0,.48);
  --ring:0 0 0 3px rgba(122,154,216,.30);
}
*{box-sizing:border-box}
.bic{font-family:var(--sans);color:var(--ink);background:linear-gradient(180deg,var(--paper2),var(--paper));min-height:100vh;font-size:14px;line-height:1.5;letter-spacing:-.005em;-webkit-font-smoothing:antialiased;transition:background .25s,color .25s}
.bic-wrap{max-width:1200px;margin:0 auto;padding:0 22px 72px}

/* signature: loom warp threads, navy→teal */
.warp{height:5px;background:
  repeating-linear-gradient(90deg, rgba(255,255,255,.5) 0 1px, transparent 1px 5px),
  var(--grad);
  opacity:.95}

.hdr{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;padding:26px 0 20px;margin-bottom:2px;border-bottom:1px solid var(--line);flex-wrap:wrap}
.brand{display:flex;flex-direction:column;gap:8px}
.brand .kick{align-self:flex-start;font-family:var(--mono);font-size:10.5px;letter-spacing:.18em;text-transform:uppercase;color:var(--b-pri-fg);font-weight:600;background:var(--b-pri-bg);padding:4px 9px;border-radius:6px}
.brand h1{margin:0;font-size:30px;line-height:1.05;letter-spacing:-.03em;font-weight:800}
.brand p{margin:3px 0 0;color:var(--muted);font-size:13.5px;max-width:56ch;line-height:1.5}
.hdr-r{display:flex;align-items:center;gap:12px}
.hdr .season-flag{font-family:var(--mono);font-size:12px}
.themebtn{width:40px;height:40px;border-radius:11px;border:1px solid var(--line);background:var(--card);color:var(--ink2);cursor:pointer;display:grid;place-items:center;box-shadow:var(--shadow-sm);transition:transform .15s,border-color .15s,color .15s,box-shadow .15s}
.themebtn:hover{border-color:var(--navy);color:var(--navy);transform:translateY(-1px);box-shadow:var(--shadow)}
.themebtn:active{transform:translateY(0)}

.drops{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:22px}
@media(max-width:720px){.drops{grid-template-columns:1fr}}
.drop{position:relative;background:var(--card);border:1.5px dashed var(--line);border-radius:14px;padding:20px;display:flex;gap:15px;align-items:center;cursor:pointer;box-shadow:var(--shadow-sm);transition:transform .16s,border-color .16s,box-shadow .16s,background .16s;overflow:hidden}
.drop:hover{border-color:var(--navy);background:var(--hover);transform:translateY(-2px);box-shadow:var(--shadow)}
.drop.ok{border-style:solid;border-color:transparent;background:var(--ok-tint);box-shadow:var(--shadow-sm),inset 3px 0 0 var(--teal)}
.drop .ic{width:46px;height:46px;border-radius:12px;display:grid;place-items:center;background:var(--grad);color:#fff;flex:none;box-shadow:var(--shadow-sm);transition:transform .16s}
.drop.opt .ic{background:linear-gradient(120deg,var(--steel),#7C8CA4)}
.drop.ok .ic{background:linear-gradient(120deg,var(--teal),#3E8F84)}
.drop:hover .ic{transform:scale(1.05) rotate(-2deg)}
.drop .t{font-weight:680;font-size:14.5px;letter-spacing:-.01em}
.drop .s{color:var(--muted);font-size:12.5px;margin-top:3px}
.drop .s b{color:var(--ink2);font-family:var(--mono);font-weight:600}

.banner{margin-top:16px;border-radius:12px;padding:13px 15px;display:flex;gap:11px;align-items:flex-start;font-size:13px;box-shadow:var(--shadow-sm);animation:rise .25s ease both}
.banner.warn{background:var(--bn-warn-bg);border:1px solid var(--bn-warn-bd);color:var(--bn-warn-fg)}
.banner.info{background:var(--bn-info-bg);border:1px solid var(--bn-info-bd);color:var(--bn-info-fg)}
@keyframes rise{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:none}}

.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-top:20px}
.kpi{position:relative;background:var(--card);border:1px solid var(--line);border-radius:13px;padding:14px 15px 13px;box-shadow:var(--shadow-sm);transition:transform .16s,box-shadow .16s;overflow:hidden}
.kpi::before{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--grad);opacity:0;transition:opacity .16s}
.kpi:hover{transform:translateY(-2px);box-shadow:var(--shadow)}
.kpi:hover::before{opacity:1}
.kpi .l{font-size:10.5px;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);font-weight:700}
.kpi .v{font-family:var(--mono);font-size:25px;font-weight:700;margin-top:5px;letter-spacing:-.03em;font-variant-numeric:tabular-nums;line-height:1.1}
.kpi .sub{font-size:11px;color:var(--muted);margin-top:2px;font-family:var(--mono)}
.kpi.accent .v{color:var(--navy)}

.controls{display:flex;gap:10px;align-items:center;margin:22px 0 12px;flex-wrap:wrap}
.tabs{display:inline-flex;background:var(--card);border:1px solid var(--line);border-radius:11px;padding:3px;box-shadow:var(--shadow-sm)}
.tab{border:0;background:transparent;padding:8px 15px;border-radius:8px;font-size:13px;font-weight:640;color:var(--muted);cursor:pointer;font-family:var(--sans);transition:color .15s,background .15s}
.tab:hover{color:var(--ink)}
.tab.on{background:var(--grad);color:#fff;box-shadow:var(--shadow-sm)}
.tab .n{font-family:var(--mono);opacity:.85;margin-left:7px;font-size:12px;font-variant-numeric:tabular-nums}
.chip{border:1px solid var(--line);background:var(--card);border-radius:22px;padding:7px 13px;font-size:12px;font-weight:640;cursor:pointer;color:var(--ink2);display:inline-flex;gap:6px;align-items:center;box-shadow:var(--shadow-sm);transition:transform .14s,border-color .14s,color .14s,background .14s}
.chip:hover{transform:translateY(-1px);color:var(--ink);border-color:var(--ink2)}
.chip.on{border-color:var(--navy);color:var(--navy);background:var(--chip-on)}
.chip .dot{width:8px;height:8px;border-radius:50%}
.search{margin-left:auto;position:relative}
.search input{border:1px solid var(--line);border-radius:11px;padding:9px 12px 9px 34px;font-size:13px;width:230px;font-family:var(--sans);background:var(--card);color:var(--ink);box-shadow:var(--shadow-sm);transition:border-color .15s,box-shadow .15s}
.search input:focus{outline:none;border-color:var(--navy);box-shadow:var(--ring)}
.search input::placeholder{color:var(--muted)}
.search svg{position:absolute;left:10px;top:10px;color:var(--muted);pointer-events:none}
.btn{border:0;border-radius:11px;padding:10px 17px;font-size:13px;font-weight:680;cursor:pointer;display:inline-flex;gap:8px;align-items:center;font-family:var(--sans);transition:transform .14s,box-shadow .14s,filter .14s}
.btn.pri{background:var(--grad);color:#fff;box-shadow:var(--shadow)}
.btn.pri:hover:not(:disabled){transform:translateY(-1px);filter:brightness(1.05);box-shadow:var(--shadow-lg)}
.btn.pri:active:not(:disabled){transform:translateY(0)}
.btn.pri:disabled{background:var(--line);color:var(--muted);cursor:not-allowed;box-shadow:none}
button:focus-visible,input:focus-visible,summary:focus-visible{outline:none;box-shadow:var(--ring)}

.tablewrap{background:var(--card);border:1px solid var(--line);border-radius:14px;overflow:hidden;margin-top:8px;box-shadow:var(--shadow)}
.scroll{overflow-x:auto}
.scroll::-webkit-scrollbar{height:10px}
.scroll::-webkit-scrollbar-thumb{background:var(--line);border-radius:10px;border:3px solid var(--card)}
.scroll::-webkit-scrollbar-thumb:hover{background:var(--muted)}
table{border-collapse:separate;border-spacing:0;width:100%;font-size:12.5px}
thead th{position:sticky;top:0;z-index:2;background:var(--thead);text-align:left;padding:11px 11px;font-size:10px;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);font-weight:700;border-bottom:1px solid var(--line);white-space:nowrap;backdrop-filter:blur(4px)}
tbody td{padding:9px 11px;border-bottom:1px solid var(--line2);white-space:nowrap;font-variant-numeric:tabular-nums}
tbody tr{border-left:3px solid transparent;transition:background .12s}
tbody tr:hover{background:var(--hover)}
tbody tr:last-child td{border-bottom:0}
tr.solids{border-left-color:var(--steel)}
tr.prints{border-left-color:var(--teal)}
tr.isnew{background:var(--isnew-bg)}
tr.isdup{background:var(--isdup-bg)}
tr.isdup td:not(.keepcol){opacity:.5}
tr.isdup .dupcode{text-decoration:line-through;text-decoration-color:var(--red)}
.m{font-family:var(--mono);font-variant-numeric:tabular-nums}
.rt{color:var(--muted);font-size:11px}
.badge{font-family:var(--mono);font-size:10px;font-weight:700;padding:2px 7px;border-radius:6px;letter-spacing:.03em}
.b-new{background:var(--b-new-bg);color:var(--b-new-fg)}
.b-rep{background:var(--b-rep-bg);color:var(--b-rep-fg)}
.b-sol{background:var(--b-sol-bg);color:var(--b-sol-fg)}
.b-pri{background:var(--b-pri-bg);color:var(--b-pri-fg)}
.b-dup{background:var(--b-dup-bg);color:var(--b-dup-fg)}
.sdot{width:9px;height:9px;border-radius:50%;display:inline-block;vertical-align:middle;box-shadow:0 0 0 3px color-mix(in srgb, currentColor 14%, transparent)}
.s-green{background:var(--green)}.s-amber{background:var(--amber)}.s-red{background:var(--red)}
.owner{color:var(--ink2)}
.verify{color:var(--amber);font-size:10px;font-family:var(--mono)}
.pcbwarn{color:var(--amber)}

.pss{margin-top:20px;background:var(--card);border:1px solid var(--line);border-radius:14px;padding:18px 20px;box-shadow:var(--shadow-sm)}
.pss h3{margin:0 0 4px;font-size:15.5px;font-weight:750;letter-spacing:-.01em;display:flex;gap:9px;align-items:center}
.pss p{margin:0 0 14px;color:var(--muted);font-size:12.5px;line-height:1.5}
.pss .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(212px,1fr));gap:10px}
.pss .card2{border:1px solid var(--line);border-left:3px solid var(--violet);border-radius:10px;padding:11px 13px;transition:transform .14s,box-shadow .14s}
.pss .card2:hover{transform:translateY(-1px);box-shadow:var(--shadow-sm)}
.pss .card2 .c{font-family:var(--mono);font-weight:700;font-size:13px}
.pss .card2 .d{color:var(--muted);font-size:11.5px;margin-top:3px;line-height:1.4}

.empty{text-align:center;padding:72px 20px;color:var(--muted)}
.empty .ic{width:62px;height:62px;border-radius:16px;background:var(--grad);display:grid;place-items:center;margin:0 auto 16px;color:#fff;box-shadow:var(--shadow)}
.rules{margin-top:18px}
.rules summary{cursor:pointer;font-size:12.5px;color:var(--ink2);font-weight:640;display:flex;gap:7px;align-items:center;user-select:none;padding:4px 2px;border-radius:8px}
.rules summary:hover{color:var(--ink)}
.rules .body{margin-top:10px;font-size:12.5px;color:var(--ink2);background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px 18px;line-height:1.6;box-shadow:var(--shadow-sm)}
.rules code{font-family:var(--mono);background:var(--line2);color:var(--ink);padding:1px 5px;border-radius:5px;font-size:11.5px}
.rules ul{margin:8px 0 0;padding-left:18px}.rules li{margin:4px 0}
.foot{margin-top:22px;color:var(--muted);font-size:11.5px;text-align:center}

@media (prefers-reduced-motion: reduce){
  *{animation-duration:.001ms !important;transition-duration:.001ms !important}
}
`;

const fmtN = (n) => (n === null || n === undefined || n === "" ? "—" : Number(n).toLocaleString("en-US"));
const fmtM = (n) => (n === null || n === undefined || n === "" ? "—" : "$" + Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }));

export default function App() {
  const [lines, setLines] = useState([]);
  const [basketName, setBasketName] = useState("");
  const [dbName, setDbName] = useState("");
  const [dbIndex, setDbIndex] = useState(null);
  const [dbKeys, setDbKeys] = useState(null);
  const [dbFile, setDbFile] = useState(null);
  const [exporting, setExporting] = useState(false);
  const [rawRows, setRawRows] = useState([]);
  const [view, setView] = useState("all");
  const [statusFilter, setStatusFilter] = useState(null);
  const [newOnly, setNewOnly] = useState(false);
  const [dupOnly, setDupOnly] = useState(false);
  const [includeDupes, setIncludeDupes] = useState(false);
  const [dark, setDark] = useState(false);
  const [q, setQ] = useState("");
  const [err, setErr] = useState("");
  const basketRef = useRef(), dbRef = useRef();

  const reprocess = useCallback((rows, idx, keys) => {
    const built = rows.filter((r) => clean(r.ProductCode)).map((r) => buildLine(r, idx));
    // duplicate pass: against the database (already exists) then within the file (repeated line)
    const seen = new Set();
    built.forEach((l) => {
      const k = dupKey(l.code, l.color, l.basket);
      if (keys && keys.has(k)) l.dupType = "db";
      else if (seen.has(k)) l.dupType = "basket";
      else l.dupType = null;
      l.isDup = !!l.dupType;
      seen.add(k);
    });
    setLines(built);
  }, []);

  const readBasket = async (file) => {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ws = wb.Sheets["ExportDetail"] || wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
      if (!rows.length || rows[0].ProductCode === undefined) {
        setErr("That file doesn't look like a basket export — expected an ExportDetail sheet with a ProductCode column.");
        return;
      }
      setErr(""); setBasketName(file.name); setRawRows(rows); reprocess(rows, dbIndex, dbKeys);
    } catch (e) { setErr("Couldn't read that file: " + e.message); }
  };
  const readDb = async (file) => {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const { codes, keys } = buildDbIndex(wb);
      setDbIndex(codes); setDbKeys(keys); setDbName(file.name); setDbFile(file);
      if (rawRows.length) reprocess(rawRows, codes, keys);
    } catch (e) { setErr("Couldn't read the database file: " + e.message); }
  };

  const seasons = useMemo(() => [...new Set(lines.map((l) => l.season).filter(Boolean))], [lines]);
  const counts = useMemo(() => {
    const c = { solids: 0, prints: 0, isNew: 0, repeat: 0, OKPOOL: 0, PBPOOL: 0, PBMOQ: 0, value: 0, missingPcb: 0, dup: 0, dupDb: 0, dupBasket: 0 };
    lines.forEach((l) => {
      c[l._sheet]++; l.isNew ? c.isNew++ : c.repeat++;
      if (c[l.status] !== undefined) c[l.status]++;
      c.value += l.value || 0; if (l.missingPcb) c.missingPcb++;
      if (l.isDup) { c.dup++; l.dupType === "db" ? c.dupDb++ : c.dupBasket++; }
    });
    return c;
  }, [lines]);

  const exportCount = useMemo(() => lines.filter((l) => includeDupes || !l.isDup).length, [lines, includeDupes]);

  const filtered = useMemo(() => lines.filter((l) => {
    if (view !== "all" && l._sheet !== view) return false;
    if (statusFilter && l.status !== statusFilter) return false;
    if (newOnly && !l.isNew) return false;
    if (dupOnly && !l.isDup) return false;
    if (q) {
      const s = (l.code + " " + l.description + " " + l.basket + " " + l.owner + " " + l.color).toLowerCase();
      if (!s.includes(q.toLowerCase())) return false;
    }
    return true;
  }), [lines, view, statusFilter, newOnly, dupOnly, q]);

  const newCodes = useMemo(() => {
    const seen = new Set(); const out = [];
    lines.filter((l) => l.isNew).forEach((l) => { if (!seen.has(l.code)) { seen.add(l.code); out.push(l); } });
    return out;
  }, [lines]);

  const hasData = lines.length > 0;

  return (
    <div className={"bic" + (dark ? " dark" : "")}>
      <style>{CSS}</style>
      <div className="warp" />
      <div className="bic-wrap">
        <div className="hdr">
          <div className="brand">
            <span className="kick">Merchandising · Basket Intake</span>
            <h1>Basket → Database Console</h1>
            <p>Turn a Carrefour <code style={{fontFamily:"var(--mono)",fontSize:12}}>ExportDetail</code> into database-ready rows: parsed, routed, enriched from your current book, and flagged for the fields that still need a PSS.</p>
          </div>
          <div className="hdr-r">
            {seasons.length > 0 && (
              <div className="season-flag">
                {seasons.map((s) => (
                  <span key={s} className="badge b-sol" style={{ marginLeft: 6 }}>{s}</span>
                ))}
              </div>
            )}
            <button className="themebtn" onClick={() => setDark(!dark)} title={dark ? "Switch to light" : "Switch to dark"} aria-label="Toggle theme">
              {dark ? <Sun size={17} /> : <Moon size={17} />}
            </button>
          </div>
        </div>

        {/* Uploads */}
        <div className="drops">
          <div className={"drop" + (basketName ? " ok" : "")} onClick={() => basketRef.current.click()}>
            <div className="ic"><Upload size={20} /></div>
            <div>
              <div className="t">Basket export {basketName && "· loaded"}</div>
              <div className="s">{basketName ? <>Reading <b>{basketName}</b> — <b>{lines.length}</b> lines</> : <>Drop or choose the monthly <b>ExportDetail</b> .xlsx · required</>}</div>
            </div>
            <input ref={basketRef} type="file" accept=".xlsx,.xls" hidden onChange={(e) => e.target.files[0] && readBasket(e.target.files[0])} />
          </div>
          <div className={"drop opt" + (dbName ? " ok" : "")} onClick={() => dbRef.current.click()}>
            <div className="ic"><Database size={20} /></div>
            <div>
              <div className="t">Current database {dbName && "· loaded"}</div>
              <div className="s">{dbName ? <>Matching against <b>{Object.keys(dbIndex||{}).length}</b> known codes</> : <>Optional — enables repeat-code auto-fill (PCB, weights, article)</>}</div>
            </div>
            <input ref={dbRef} type="file" accept=".xlsx,.xls" hidden onChange={(e) => e.target.files[0] && readDb(e.target.files[0])} />
          </div>
        </div>

        {err && <div className="banner warn"><AlertTriangle size={16} style={{ flex: "none", marginTop: 1 }} /><div>{err}</div></div>}

        {hasData && seasons.length > 1 && (
          <div className="banner warn">
            <AlertTriangle size={16} style={{ flex: "none", marginTop: 1 }} />
            <div><b>Mixed seasons in one basket.</b> This file spans {seasons.join(" + ")}. Split off any off-season lines (e.g. AWH-26) into their own database before committing — they don't belong in the SSH-27 book.</div>
          </div>
        )}
        {hasData && !dbName && (
          <div className="banner info">
            <Info size={16} style={{ flex: "none", marginTop: 1 }} />
            <div>No database loaded, so every code shows as <b>NEW</b> and cartons/piece-weights can't be filled. Add your current workbook to auto-enrich the repeat codes.</div>
          </div>
        )}
        {hasData && counts.dup > 0 && (
          <div className="banner warn">
            <Copy size={16} style={{ flex: "none", marginTop: 1 }} />
            <div>
              <b>{counts.dup} duplicate {counts.dup === 1 ? "line" : "lines"} held back.</b>{" "}
              {counts.dupDb > 0 && <>{counts.dupDb} already {counts.dupDb === 1 ? "exists" : "exist"} in your database</>}
              {counts.dupDb > 0 && counts.dupBasket > 0 && ", "}
              {counts.dupBasket > 0 && <>{counts.dupBasket} repeated within this file</>}. They're excluded from the export so you never double-append — override with “Include duplicates” if you meant to add them.
            </div>
          </div>
        )}

        {hasData && (
          <>
            <div className="kpis">
              <div className="kpi accent"><div className="l">Lines</div><div className="v">{fmtN(lines.length)}</div><div className="sub">{[...new Set(lines.map(l=>l.basket))].length} baskets</div></div>
              <div className="kpi"><div className="l">Solids</div><div className="v" style={{color:"var(--steel)"}}>{fmtN(counts.solids)}</div><div className="sub">→ Sheet1</div></div>
              <div className="kpi"><div className="l">Prints</div><div className="v" style={{color:"var(--teal)"}}>{fmtN(counts.prints)}</div><div className="sub">→ Sheet2</div></div>
              <div className="kpi"><div className="l">New codes</div><div className="v" style={{color:"var(--violet)"}}>{fmtN(newCodes.length)}</div><div className="sub">{fmtN(counts.repeat)} repeat lines</div></div>
              <div className="kpi"><div className="l">Clean to book</div><div className="v" style={{color:"var(--green)"}}>{fmtN(counts.OKPOOL)}</div><div className="sub">{fmtN(counts.PBPOOL+counts.PBMOQ)} flagged</div></div>
              <div className="kpi"><div className="l">Duplicates</div><div className="v" style={{color:counts.dup?"var(--red)":"var(--muted)"}}>{fmtN(counts.dup)}</div><div className="sub">held from export</div></div>
              <div className="kpi"><div className="l">Order value</div><div className="v">{counts.value?("$"+Math.round(counts.value).toLocaleString("en-US")):"—"}</div><div className="sub">gross</div></div>
            </div>

            <div className="controls">
              <div className="tabs">
                {[["all","All"],["solids","Solids"],["prints","Prints"]].map(([k,lab])=>(
                  <button key={k} className={"tab"+(view===k?" on":"")} onClick={()=>setView(k)}>{lab}<span className="n">{k==="all"?lines.length:counts[k]}</span></button>
                ))}
              </div>
              {["OKPOOL","PBPOOL","PBMOQ"].map((s)=>(
                <button key={s} className={"chip"+(statusFilter===s?" on":"")} onClick={()=>setStatusFilter(statusFilter===s?null:s)}>
                  <span className={"dot s-"+STATUS_META[s].tone} style={{background:`var(--${STATUS_META[s].tone})`}} />{STATUS_META[s].label}<span className="m rt">{counts[s]}</span>
                </button>
              ))}
              <button className={"chip"+(newOnly?" on":"")} onClick={()=>setNewOnly(!newOnly)}><Sparkles size={12}/>New only</button>
              {counts.dup > 0 && (
                <button className={"chip"+(dupOnly?" on":"")} onClick={()=>setDupOnly(!dupOnly)}><Copy size={12}/>Duplicates<span className="m rt">{counts.dup}</span></button>
              )}
              <div className="search"><Search size={15}/><input placeholder="Code, design, colour…" value={q} onChange={(e)=>setQ(e.target.value)} /></div>
              {counts.dup > 0 && (
                <button className={"chip"+(includeDupes?" on":"")} onClick={()=>setIncludeDupes(!includeDupes)} title="Include duplicate lines in the exported file">
                  {includeDupes ? "Including dupes" : "Include dupes"}
                </button>
              )}
              <button
                className="btn pri"
                disabled={!hasData || exporting}
                title={dbFile ? `Appends directly into ${dbName}, keeping its formatting.` : "No database loaded — exporting as a new plain file."}
                onClick={async () => {
                  setExporting(true); setErr("");
                  try {
                    if (dbFile) await exportIntoDatabase(lines, includeDupes, dbFile);
                    else downloadWorkbook(lines, includeDupes);
                  } catch (e) {
                    setErr("Export failed: " + e.message);
                  } finally {
                    setExporting(false);
                  }
                }}
              >
                <Download size={15}/>
                {exporting ? "Exporting…" : dbFile ? `Append ${fmtN(exportCount)} rows to database` : `Export ${fmtN(exportCount)} rows`}
              </button>
            </div>

            <div className="tablewrap"><div className="scroll">
              <table>
                <thead><tr>
                  <th>St</th><th>Route</th><th>Code</th><th>Description</th><th>Colour</th>
                  <th>Owner</th><th>Article</th><th>Size</th>
                  <th style={{textAlign:"right"}}>Qty</th><th style={{textAlign:"right"}}>Price</th><th style={{textAlign:"right"}}>Value</th>
                  <th style={{textAlign:"right"}}>PCB</th><th style={{textAlign:"right"}}>Ctns</th>
                  <th>ETD</th><th>FRI</th>
                </tr></thead>
                <tbody>
                  {filtered.map((l,i)=>(
                    <tr key={i} className={l._sheet+(l.isNew?" isnew":"")+(l.isDup?" isdup":"")}>
                      <td className="keepcol" title={STATUS_META[l.status]?.note||l.status}><span className={"sdot s-"+(STATUS_META[l.status]?.tone||"amber")} style={{background:`var(--${STATUS_META[l.status]?.tone||"amber"})`}} /></td>
                      <td className="keepcol"><span className={"badge "+(l._sheet==="prints"?"b-pri":"b-sol")}>{l._sheet==="prints"?"PRINT":"SOLID"}</span></td>
                      <td className="m keepcol">
                        <span className={l.isDup?"dupcode":""}>{l.code}</span>{" "}
                        {l.isDup
                          ? <span className="badge b-dup" style={{marginLeft:4}} title={l.dupType==="db"?"Already in the database — won't be re-added":"Repeated within this file — only the first is kept"}>DUP</span>
                          : (l.isNew?<span className="badge b-new" style={{marginLeft:4}}>NEW</span>:<span className="badge b-rep" style={{marginLeft:4}}>REP</span>)}
                      </td>
                      <td>{l.description}</td>
                      <td>{l.color}</td>
                      <td className="owner">{l.owner||<span className="verify">?verify</span>} {l.ownerVerify&&<span className="verify" title="entity code not recognised">·chk</span>}</td>
                      <td>{l.article||<span className="rt">—</span>}</td>
                      <td className="m">{l.size||"—"}</td>
                      <td className="m" style={{textAlign:"right"}}>{fmtN(l.qty)}</td>
                      <td className="m" style={{textAlign:"right"}}>{l.price?l.price.toFixed(2):"—"}</td>
                      <td className="m" style={{textAlign:"right"}}>{fmtM(l.value)}</td>
                      <td className={"m"+(l.missingPcb?" pcbwarn":"")} style={{textAlign:"right"}} title={l.missingPcb?"No PCB — needs packing spec / PSS":""}>{l.pcb||"·"}</td>
                      <td className="m" style={{textAlign:"right"}}>{l.cartons||"·"}</td>
                      <td className="m rt">{fmtDate(l.initEtd)}</td>
                      <td className="m rt">{fmtDate(l.initFri)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div></div>
            {filtered.length===0 && <div className="foot">No lines match these filters.</div>}

            {newCodes.length>0 && (
              <div className="pss">
                <h3><Package size={17} color="var(--violet)"/> {newCodes.length} new codes awaiting PSS</h3>
                <p>Not in your current database — these need the PSS and packing spec before PCB, cartons and piece-weights can be completed. Production can still be sequenced by ETD in the meantime.</p>
                <div className="grid">
                  {newCodes.map((l)=>(
                    <div className="card2" key={l.code}>
                      <div className="c">{l.code}</div>
                      <div className="d">{l.description}</div>
                      <div className="d m" style={{marginTop:3}}>{l._sheet==="prints"?"PRINT":"SOLID"} · {l.owner||"?"} · ETD {fmtDate(l.initEtd)}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <details className="rules">
              <summary><ChevronDown size={14}/> How each field is decided</summary>
              <div className="body">
                <b>Straight from the basket:</b> code, description, colour, quantity, L/W/H, per-carton net &amp; gross weight, basket, flow type, FRI &amp; ETD dates.
                <ul>
                  <li><b>Price</b> — trailing number of the coded string (<code>CFS_FOB_PKPQM_USD_6.4900</code> → <code>6.49</code>).</li>
                  <li><b>Season / Owner</b> — from the basket code: <code>SSH27</code>+<code>FRCA</code> → SSH-27 · France. Repeats take their owner from your database (authoritative); new codes use the entity map and are marked <span className="verify">·chk</span> if unrecognised.</li>
                  <li><b>Route</b> — repeats follow their existing sheet; new codes route by article/design (PARURE + a print name → Prints, else Solids).</li>
                  <li><b>Article · Size · Design · Category</b> — pulled from the database for repeats, else derived from the description.</li>
                  <li><b>PCB</b> — not in the basket. Inherited from the database for repeats; <span className="pcbwarn">blank for new codes</span> until the packing spec arrives (this blocks cartons, piece-weights and totals).</li>
                  <li><b>Dates</b> — INITIAL FRI = basket FRI · YTM SHIP = FRI − 15d · WEEK DATE = FRI − 67d · # of days = 52.</li>
                  <li><b>Left blank for the second pass:</b> PSS, YTM lot #, shipping marks, tag cards, colour code, validation date.</li>
                  <li><b>Export</b> — with a database loaded, rows are appended straight into that workbook (its fonts, fills, column widths and number formats carry through); without one, a plain new workbook is produced instead.</li>
                </ul>
              </div>
            </details>
          </>
        )}

        {!hasData && (
          <div className="empty">
            <div className="ic"><FileSpreadsheet size={26}/></div>
            <div style={{fontWeight:640,color:"var(--ink2)",fontSize:15}}>Load a basket export to begin</div>
            <div style={{marginTop:4}}>Add your current database too and repeat codes fill themselves in.</div>
          </div>
        )}
      </div>
    </div>
  );
}
