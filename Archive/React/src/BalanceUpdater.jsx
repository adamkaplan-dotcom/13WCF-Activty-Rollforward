import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

/* ═══════════════════════════════════════════════════════════════
   ZIP UTILITIES — parse, decompress, replace files, rebuild
   Uses browser-native DecompressionStream / CompressionStream
   ═══════════════════════════════════════════════════════════════ */
const _te = new TextEncoder();
const _td = new TextDecoder();

/* CRC32 table */
const _crcT = new Uint32Array(256);
for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); _crcT[n] = c; }
function _crc32(d) { let c = 0xFFFFFFFF; for (let i = 0; i < d.length; i++) c = _crcT[(c ^ d[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }

/* inflate/deflate via browser streams */
async function _streamXform(data, xf) {
  const w = xf.writable.getWriter(); w.write(data); w.close();
  const r = xf.readable.getReader(); const ch = [];
  while (true) { const { done, value } = await r.read(); if (done) break; ch.push(value); }
  const tot = ch.reduce((a, c) => a + c.length, 0); const out = new Uint8Array(tot);
  let o = 0; for (const c of ch) { out.set(c, o); o += c.length; } return out;
}
const _inflate = (d) => _streamXform(d, new DecompressionStream("deflate-raw"));
const _deflate = (d) => _streamXform(d, new CompressionStream("deflate-raw"));

/* Parse ZIP → array of { name, method, compSz, uncompSz, localOff, cdOff, cdLen, ... } */
function parseZipEntries(buf) {
  const u8 = new Uint8Array(buf);
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let eocd = -1;
  for (let i = u8.length - 22; i >= 0; i--) { if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; } }
  if (eocd < 0) return { entries: [], u8, dv, eocdOff: -1 };
  const cdOff = dv.getUint32(eocd + 16, true);
  const num = dv.getUint16(eocd + 10, true);
  const entries = []; let off = cdOff;
  for (let i = 0; i < num; i++) {
    if (dv.getUint32(off, true) !== 0x02014b50) break;
    const method = dv.getUint16(off + 10, true);
    const compSz = dv.getUint32(off + 20, true);
    const uncompSz = dv.getUint32(off + 24, true);
    const nL = dv.getUint16(off + 28, true);
    const eL = dv.getUint16(off + 30, true);
    const cL = dv.getUint16(off + 32, true);
    const localOff = dv.getUint32(off + 42, true);
    const name = _td.decode(u8.subarray(off + 46, off + 46 + nL));
    entries.push({ name, method, compSz, uncompSz, localOff, cdOff: off, cdLen: 46 + nL + eL + cL });
    off += 46 + nL + eL + cL;
  }
  return { entries, u8, dv, eocdOff: eocd };
}

/* Read a file's decompressed content from the ZIP */
async function zipReadFile(entry, u8, dv) {
  const lnL = dv.getUint16(entry.localOff + 26, true);
  const leL = dv.getUint16(entry.localOff + 28, true);
  const dataOff = entry.localOff + 30 + lnL + leL;
  const comp = u8.subarray(dataOff, dataOff + entry.compSz);
  if (entry.method === 0) return comp;
  if (entry.method === 8) return await _inflate(comp);
  throw new Error("Unsupported compression: " + entry.method);
}

/* Rebuild ZIP with replacements */
async function rebuildZip(parsed, replacements) {
  const { entries, u8, dv, eocdOff } = parsed;
  const patchMap = {};
  for (const [name, rawBytes] of Object.entries(replacements)) {
    const comp = await _deflate(rawBytes);
    patchMap[name] = { newComp: comp, newCrc: _crc32(rawBytes), newUncompSz: rawBytes.length, newCompSz: comp.length };
  }
  const parts = []; const newOffsets = {};
  for (const e of entries) {
    newOffsets[e.name] = parts.reduce((a, p) => a + p.length, 0);
    const lnL = dv.getUint16(e.localOff + 26, true);
    const leL = dv.getUint16(e.localOff + 28, true);
    const hdrEnd = e.localOff + 30 + lnL + leL;
    if (patchMap[e.name]) {
      const pm = patchMap[e.name];
      const hdr = new Uint8Array(30 + lnL + leL);
      hdr.set(u8.subarray(e.localOff, hdrEnd));
      const hv = new DataView(hdr.buffer);
      hv.setUint16(8, 8, true);
      hv.setUint32(14, pm.newCrc, true);
      hv.setUint32(18, pm.newCompSz, true);
      hv.setUint32(22, pm.newUncompSz, true);
      parts.push(hdr); parts.push(pm.newComp);
    } else {
      parts.push(u8.subarray(e.localOff, hdrEnd + e.compSz));
    }
  }
  const cdStart = parts.reduce((a, p) => a + p.length, 0);
  for (const e of entries) {
    const cd = new Uint8Array(e.cdLen);
    cd.set(u8.subarray(e.cdOff, e.cdOff + e.cdLen));
    const cv = new DataView(cd.buffer);
    cv.setUint32(42, newOffsets[e.name], true);
    if (patchMap[e.name]) {
      const pm = patchMap[e.name];
      cv.setUint16(10, 8, true);
      cv.setUint32(16, pm.newCrc, true);
      cv.setUint32(20, pm.newCompSz, true);
      cv.setUint32(24, pm.newUncompSz, true);
    }
    parts.push(cd);
  }
  const cdEnd = parts.reduce((a, p) => a + p.length, 0);
  const eocdBuf = new Uint8Array(22);
  eocdBuf.set(u8.subarray(eocdOff, eocdOff + 22));
  const ev = new DataView(eocdBuf.buffer);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, cdEnd - cdStart, true);
  ev.setUint32(16, cdStart, true);
  parts.push(eocdBuf);
  const total = parts.reduce((a, p) => a + p.length, 0);
  const result = new Uint8Array(total);
  let wo = 0; for (const p of parts) { result.set(p, wo); wo += p.length; }
  return result;
}

/* ═══════════════════════════════════════════════════════════════
   EXTRACT cell style maps from original XLSX
   ═══════════════════════════════════════════════════════════════ */
async function extractCellStyleMaps(origBuf) {
  const parsed = parseZipEntries(origBuf);
  const { entries, u8, dv } = parsed;
  const maps = {};
  for (const e of entries) {
    if (!e.name.startsWith("xl/worksheets/sheet") || !e.name.endsWith(".xml")) continue;
    const raw = await zipReadFile(e, u8, dv);
    const text = _td.decode(raw);
    const map = {};
    const re = /<c\s+r="([A-Z]+[0-9]+)"([^>]*?)(?:\s*\/>|>)/g;
    let m;
    while ((m = re.exec(text)) !== null) {
      const sMatch = m[2].match(/\bs="(\d+)"/);
      if (sMatch) map[m[1]] = sMatch[1];
    }
    if (Object.keys(map).length > 0) maps[e.name] = map;
  }
  return maps;
}

/* ═══════════════════════════════════════════════════════════════
   PATCH OUTPUT XLSX
   Transplants styles.xml + theme from original (formatting)
   Restores cell style indices in worksheets
   ═══════════════════════════════════════════════════════════════ */
async function patchXlsxOutput(origBuf, outBuf, cellStyleMaps) {
  const origParsed = parseZipEntries(origBuf);
  const outParsed  = parseZipEntries(outBuf);
  const replacements = {};

  /* ── Transplant styles.xml and theme from original ───────────── */
  for (const e of origParsed.entries) {
    if (e.name === "xl/styles.xml" || e.name === "xl/theme/theme1.xml") {
      const raw = await zipReadFile(e, origParsed.u8, origParsed.dv);
      replacements[e.name] = raw;
    }
  }

  /* ── Patch worksheets: restore s= style indices ────────────── */
  for (const e of outParsed.entries) {
    if (!e.name.startsWith("xl/worksheets/sheet") || !e.name.endsWith(".xml")) continue;
    const raw = await zipReadFile(e, outParsed.u8, outParsed.dv);
    let xml = _td.decode(raw);
    let changed = false;

    const origMap = cellStyleMaps[e.name];
    if (origMap && Object.keys(origMap).length > 0) {
      xml = xml.replace(/<c\s+r="([A-Z]+[0-9]+)"([^>]*?)(\/?>)/g, (match, addr, rest, close) => {
        const origS = origMap[addr];
        if (!origS) {
          const cleaned = rest.replace(/\s*s="\d+"/, "");
          return `<c r="${addr}"${cleaned}${close}`;
        }
        if (/\bs="\d+"/.test(rest)) {
          return `<c r="${addr}"${rest.replace(/\bs="\d+"/, `s="${origS}"`)}${close}`;
        }
        return `<c r="${addr}" s="${origS}"${rest}${close}`;
      });
      changed = true;
    }

    if (changed) replacements[e.name] = _te.encode(xml);
  }

  return await rebuildZip(outParsed, replacements);
}

/* ═══════════════════════════════════════════════════════════════
   HELPER FUNCTIONS — SheetJS utilities
   ═══════════════════════════════════════════════════════════════ */
const addr = (r, c) => XLSX.utils.encode_cell({ r, c });
const getCell = (ws, r, c) => ws[addr(r, c)];
const getVal = (ws, r, c) => { const cell = getCell(ws, r, c); return cell ? cell.v : undefined; };
const getNum = (ws, r, c) => { const v = getVal(ws, r, c); return typeof v === "number" ? v : null; };
const putNum = (ws, r, c, val) => { ws[addr(r, c)] = { t: "n", v: val }; };
const putStr = (ws, r, c, val) => { ws[addr(r, c)] = { t: "s", v: val }; };
const delCell = (ws, r, c) => { delete ws[addr(r, c)]; };
const endRC = (ws) => {
  const ref = ws["!ref"];
  if (!ref) return { r: 0, c: 0 };
  const range = XLSX.utils.decode_range(ref);
  return { r: range.e.r, c: range.e.c };
};
const setRange = (ws, maxR, maxC) => {
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
};

/* ═══════════════════════════════════════════════════════════════
   PROCESS BALANCE UPDATER

   1. Read column K (Current Week) from USDx Balances tab (K10 header)
   2. Find last date in row 14 of Beginning Balances tab
   3. Paste column K data into the next column
   4. Copy formula from row 10 to generate new date header
   5. Copy formulas for rows 12-20 and 148-end from previous column
   ═══════════════════════════════════════════════════════════════ */
async function processBalanceUpdater(weeklyBuf, rollforwardBuf, log) {
  log("═".repeat(60));
  log("STEP 1: Reading Column K from USDx Balances");
  log("═".repeat(60));

  /* ── 1A: Read Weekly Balances file (data_only mode) ──────────── */
  log("Parsing Weekly Balances file with SheetJS...");
  const wbWeekly = XLSX.read(weeklyBuf, { type: "array", cellFormula: false, cellStyles: true });
  log("✓ Weekly file parsed");

  const wsUSDx = wbWeekly.Sheets["USDx Balances"];
  if (!wsUSDx) throw new Error("Sheet 'USDx Balances' not found in Weekly Balances file");
  log("✓ Found 'USDx Balances' sheet");

  const balanceCol = 10;  // Column K (0-indexed)
  const refCol = 1;       // Column B (0-indexed)
  const headerRow = 9;    // Row 10 (0-indexed)
  const dataStartRow = 12; // Row 13 (0-indexed)

  const balanceHeader = getVal(wsUSDx, headerRow, balanceCol);
  log(`Column K header (K10): ${balanceHeader}`);

  /* Extract Current Week balances with reference numbers */
  const balances = {};
  const usdxEnd = endRC(wsUSDx);
  for (let r = dataStartRow; r <= usdxEnd.r; r++) {
    const refNum = getVal(wsUSDx, r, refCol);
    const balance = getNum(wsUSDx, r, balanceCol);

    if (refNum != null) {
      const refStr = String(refNum).trim();
      // Skip section headers
      const skipKeywords = ['bullish', 'coindesk', 'fiat', 'balance', 'total'];
      if (!skipKeywords.some(kw => refStr.toLowerCase().includes(kw))) {
        balances[refStr] = balance != null ? balance : 0.0;
      }
    }
  }
  log(`Extracted ${Object.keys(balances).length} Current Week balances from column K`);

  log("");
  log("═".repeat(60));
  log("STEP 2: Opening Activity Rollforward file");
  log("═".repeat(60));

  /* ── 2A: Read Rollforward file (preserve formulas) ────────────── */
  log("Parsing Activity Rollforward file with SheetJS...");
  const wbRollforward = XLSX.read(rollforwardBuf, { type: "array", cellFormula: true, cellStyles: true });
  log("✓ Rollforward file parsed");

  const wsBB = wbRollforward.Sheets["Beginning Balances"];
  if (!wsBB) throw new Error("Sheet 'Beginning Balances' not found in Activity Rollforward file");

  log("✓ Found 'Beginning Balances' sheet");

  log("");
  log("═".repeat(60));
  log("STEP 3: Finding last date in row 14");
  log("═".repeat(60));

  /* ── 3A: Find last date in row 14 ──────────────────────────────── */
  const dateRow = 13;  // Row 14 (0-indexed)
  const bbEnd = endRC(wsBB);
  let lastColWithDate = -1;

  for (let c = 0; c <= bbEnd.c; c++) {
    const cellVal = getVal(wsBB, dateRow, c);
    if (cellVal != null) lastColWithDate = c;
  }

  if (lastColWithDate < 0) {
    throw new Error("No dates found in row 14 of Beginning Balances");
  }

  const lastColLetter = XLSX.utils.encode_col(lastColWithDate);
  log(`Last date found in column ${lastColLetter} (row 14)`);

  /* New column is one to the right */
  const newCol = lastColWithDate + 1;
  const newColLetter = XLSX.utils.encode_col(newCol);
  log(`Will paste into column ${newColLetter}`);

  log("");
  log("═".repeat(60));
  log("STEP 4: Copying formula from row 10 for date header");
  log("═".repeat(60));

  /* ── 4A: Copy formula from row 10 (previous column) ────────────── */
  const formulaRow = 9;  // Row 10 (0-indexed)
  const sourceFormulaCell = getCell(wsBB, formulaRow, lastColWithDate);

  if (sourceFormulaCell && sourceFormulaCell.f) {
    // Parse and adjust formula reference (shift column right by 1)
    const adjustedFormula = sourceFormulaCell.f.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
      const colNum = XLSX.utils.decode_col(col);
      const newColNum = colNum + 1;
      const newColRef = XLSX.utils.encode_col(newColNum);
      return `${newColRef}${row}`;
    });

    wsBB[addr(formulaRow, newCol)] = { t: sourceFormulaCell.t, f: adjustedFormula };
    log(`Copied and adjusted formula from ${lastColLetter}10 to ${newColLetter}10`);
    log(`Formula: ${adjustedFormula}`);
  } else {
    log(`No formula found in row 10, column ${lastColLetter}`);
  }

  log("");
  log("═".repeat(60));
  log("STEP 5: Pasting Current Week balances");
  log("═".repeat(60));

  /* ── 5A: Match and paste balances by Column B ──────────────────── */
  const targetRefCol = 1;  // Column B (0-indexed)
  const targetDataStart = 14;  // Row 15 (0-indexed)
  let matched = 0;
  let notMatched = 0;

  for (let r = targetDataStart; r <= bbEnd.r; r++) {
    const refNum = getVal(wsBB, r, targetRefCol);

    if (refNum != null) {
      const refStr = String(refNum).trim();

      if (balances[refStr] !== undefined) {
        putNum(wsBB, r, newCol, balances[refStr]);
        matched++;
      } else {
        putNum(wsBB, r, newCol, 0.0);
        notMatched++;
      }
    }
  }

  log(`✓ Matched and pasted: ${matched} records`);
  if (notMatched > 0) {
    log(`⚠ Not found in source: ${notMatched} records (set to 0.0)`);
  }

  log("");
  log("═".repeat(60));
  log("STEP 6: Copying formulas from previous week");
  log("═".repeat(60));

  /* ── 6A: Copy formulas/cells for rows 12-20 and 148-end ────────── */
  let formulasCopied = 0;
  log(`Sheet has ${bbEnd.r + 1} rows, ${bbEnd.c + 1} columns`);

  // Copy rows 12-20 (0-indexed: 11-19)
  log("Copying formulas for rows 12-20...");
  for (let r = 11; r <= 19; r++) {
    const sourceCell = getCell(wsBB, r, lastColWithDate);
    if (sourceCell) {
      if (sourceCell.f) {
        // Adjust formula references
        const adjustedFormula = sourceCell.f.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
          const colNum = XLSX.utils.decode_col(col);
          const newColNum = colNum + 1;
          const newColRef = XLSX.utils.encode_col(newColNum);
          return `${newColRef}${row}`;
        });
        wsBB[addr(r, newCol)] = { ...sourceCell, f: adjustedFormula };
        formulasCopied++;
      } else {
        wsBB[addr(r, newCol)] = { ...sourceCell };
      }
    }
  }

  log(`✓ Copied cells for rows 12-20 from column ${lastColLetter}`);

  // Copy rows 148 to last row (0-indexed: 147-end)
  const endRow = Math.min(bbEnd.r, 10000); // Safety limit: max 10k rows
  log(`Copying formulas for rows 148-${endRow + 1}...`);
  for (let r = 147; r <= endRow; r++) {
    const sourceCell = getCell(wsBB, r, lastColWithDate);
    if (sourceCell) {
      if (sourceCell.f) {
        // Adjust formula references
        const adjustedFormula = sourceCell.f.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
          const colNum = XLSX.utils.decode_col(col);
          const newColNum = colNum + 1;
          const newColRef = XLSX.utils.encode_col(newColNum);
          return `${newColRef}${row}`;
        });
        wsBB[addr(r, newCol)] = { ...sourceCell, f: adjustedFormula };
        formulasCopied++;
      } else {
        wsBB[addr(r, newCol)] = { ...sourceCell };
      }
    }
  }

  log(`✓ Copied cells for rows 148-${endRow + 1} from column ${lastColLetter}`);
  log(`✓ Total formulas copied: ${formulasCopied}`);

  /* Update sheet range */
  setRange(wsBB, bbEnd.r, newCol);

  log("");
  log("═".repeat(60));
  log("SUMMARY");
  log("═".repeat(60));
  log(`Column pasted to: ${newColLetter}`);
  log(`Current Week balances matched: ${matched}`);
  log(`Formulas copied: ${formulasCopied}`);
  log(`✓ Process completed successfully!`);
  log("═".repeat(60));

  return {
    workbook: wbRollforward,
    stats: {
      column: newColLetter,
      totalBalances: Object.keys(balances).length,
      balancesMatched: matched,
      formulasCopied: formulasCopied
    }
  };
}

/* ═══════════════════════════════════════════════════════════════
   REACT COMPONENT — Balance Updater UI
   ═══════════════════════════════════════════════════════════════ */
export default function BalanceUpdater() {
  const [weeklyFile, setWeeklyFile] = useState(null);
  const [rollforwardFile, setRollforwardFile] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [logs, setLogs] = useState([]);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);
  const [weeklyDragOver, setWeeklyDragOver] = useState(false);
  const [rollforwardDragOver, setRollforwardDragOver] = useState(false);

  const weeklyInputRef = useRef(null);
  const rollforwardInputRef = useRef(null);

  const log = useCallback((msg) => {
    setLogs(prev => [...prev, msg]);
  }, []);

  const handleWeeklyFile = useCallback((e) => {
    const file = e.target.files[0];
    if (file) setWeeklyFile(file);
  }, []);

  const handleRollforwardFile = useCallback((e) => {
    const file = e.target.files[0];
    if (file) setRollforwardFile(file);
  }, []);

  /* ── Drag and Drop Handlers ────────────────────────────────────── */
  const handleWeeklyDrop = useCallback((e) => {
    e.preventDefault();
    setWeeklyDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
      setWeeklyFile(file);
    }
  }, []);

  const handleRollforwardDrop = useCallback((e) => {
    e.preventDefault();
    setRollforwardDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
      setRollforwardFile(file);
    }
  }, []);

  const handleWeeklyDragOver = useCallback((e) => {
    e.preventDefault();
    setWeeklyDragOver(true);
  }, []);

  const handleRollforwardDragOver = useCallback((e) => {
    e.preventDefault();
    setRollforwardDragOver(true);
  }, []);

  const handleWeeklyDragLeave = useCallback(() => {
    setWeeklyDragOver(false);
  }, []);

  const handleRollforwardDragLeave = useCallback(() => {
    setRollforwardDragOver(false);
  }, []);

  const processFiles = useCallback(async () => {
    if (!weeklyFile || !rollforwardFile) return;

    setProcessing(true);
    setLogs([]);
    setError(null);
    setResult(null);

    try {
      /* Read files as ArrayBuffers */
      log("Reading file: " + weeklyFile.name + " (" + (weeklyFile.size / 1024 / 1024).toFixed(2) + " MB)");
      const weeklyBuf = await weeklyFile.arrayBuffer();
      log("✓ Weekly file loaded into memory");

      log("Reading file: " + rollforwardFile.name + " (" + (rollforwardFile.size / 1024 / 1024).toFixed(2) + " MB)");
      const rollforwardBuf = await rollforwardFile.arrayBuffer();
      log("✓ Rollforward file loaded into memory");

      /* Extract cell styles from original rollforward file */
      log("");
      log("Extracting cell styles from Activity Rollforward file...");
      const cellStyleMaps = await extractCellStyleMaps(rollforwardBuf);
      log(`✓ Extracted styles for ${Object.keys(cellStyleMaps).length} worksheets`);

      /* Process balance updater logic */
      log("");
      const { workbook, stats } = await processBalanceUpdater(weeklyBuf, rollforwardBuf, log);
      log("✓ Balance processing complete");

      /* Write workbook to buffer using SheetJS */
      log("");
      log("Generating output XLSX...");
      const outBuf = XLSX.write(workbook, { type: "array", bookType: "xlsx", cellStyles: true });
      log(`✓ Generated XLSX (${(outBuf.byteLength / 1024 / 1024).toFixed(2)} MB)`);

      /* Patch output to restore formatting */
      log("");
      log("Restoring original formatting...");
      const finalBuf = await patchXlsxOutput(rollforwardBuf, outBuf, cellStyleMaps);
      log(`✓ Formatting restored (final size: ${(finalBuf.byteLength / 1024 / 1024).toFixed(2)} MB)`);

      /* Create download blob */
      const timestamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, -5);
      const filename = `Activity_Rollforward_Updated_${timestamp}.xlsx`;
      const blob = new Blob([finalBuf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);

      setResult({ url, filename, stats });
      log("");
      log(`✓ File ready for download: ${filename}`);

    } catch (err) {
      setError(err.message);
      log("");
      log(`✗ ERROR: ${err.message}`);
      log(`✗ Stack trace: ${err.stack}`);
      console.error("Full error details:", err);
    } finally {
      setProcessing(false);
    }
  }, [weeklyFile, rollforwardFile, log]);

  return (
    <div style={styles.container}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:opsz,wght@9..40,400;9..40,500;9..40,600;9..40,700&display=swap" rel="stylesheet" />
      <div style={styles.header} />

      {/* Header */}
      <div style={styles.topBar}>
        <div style={styles.topBarInner}>
          <div style={styles.logo}>
            <span style={styles.logoText}>B</span>
          </div>
          <span style={styles.brandName}>bullish</span>
          <div style={styles.divider} />
          <div>
            <div style={styles.title}>Balance Updater</div>
            <div style={styles.subtitle}>Update Activity Rollforward with Current Week balances from USDx tab</div>
          </div>
        </div>
      </div>

      <div style={styles.card}>
        <div style={styles.uploadSection}>
          <div
            style={{
              ...styles.uploadBox,
              ...(weeklyFile && styles.uploadBoxActive),
              ...(weeklyDragOver && styles.uploadBoxDragOver)
            }}
            onClick={() => weeklyInputRef.current?.click()}
            onDrop={handleWeeklyDrop}
            onDragOver={handleWeeklyDragOver}
            onDragEnter={handleWeeklyDragOver}
            onDragLeave={handleWeeklyDragLeave}
          >
            <input
              ref={weeklyInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={handleWeeklyFile}
              style={styles.fileInput}
            />
            <div style={styles.uploadTitle}>Weekly Balances File</div>
            <div style={styles.uploadSubtitle}>
              {weeklyFile ? "✓ " + weeklyFile.name : "Drop or click to upload"}
            </div>
          </div>

          <div
            style={{
              ...styles.uploadBox,
              ...(rollforwardFile && styles.uploadBoxActive),
              ...(rollforwardDragOver && styles.uploadBoxDragOver)
            }}
            onClick={() => rollforwardInputRef.current?.click()}
            onDrop={handleRollforwardDrop}
            onDragOver={handleRollforwardDragOver}
            onDragEnter={handleRollforwardDragOver}
            onDragLeave={handleRollforwardDragLeave}
          >
            <input
              ref={rollforwardInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={handleRollforwardFile}
              style={styles.fileInput}
            />
            <div style={styles.uploadTitle}>Activity Rollforward File</div>
            <div style={styles.uploadSubtitle}>
              {rollforwardFile ? "✓ " + rollforwardFile.name : "Drop or click to upload"}
            </div>
          </div>
        </div>

        <button
          onClick={processFiles}
          disabled={!weeklyFile || !rollforwardFile || processing}
          style={{
            ...styles.processButton,
            ...((!weeklyFile || !rollforwardFile || processing) && styles.processButtonDisabled)
          }}
        >
          {processing ? "Processing..." : "Process Files"}
        </button>

        {error && (
          <div style={styles.errorBox}>
            <h3>❌ Error</h3>
            <p>{error}</p>
          </div>
        )}

        {result && (
          <div style={styles.successBox}>
            <h3 style={styles.successTitle}>✓ Processing Complete!</h3>
            <div style={styles.statsGrid}>
              <div style={styles.statCard}>
                <div style={styles.statValue}>Column {result.stats.column}</div>
                <div style={styles.statLabel}>Data Pasted To</div>
              </div>
              <div style={styles.statCard}>
                <div style={styles.statValue}>{result.stats.totalBalances}</div>
                <div style={styles.statLabel}>Total Balances</div>
              </div>
              <div style={styles.statCard}>
                <div style={styles.statValue}>{result.stats.balancesMatched}</div>
                <div style={styles.statLabel}>Matched</div>
              </div>
              <div style={styles.statCard}>
                <div style={styles.statValue}>{result.stats.formulasCopied}</div>
                <div style={styles.statLabel}>Formulas Copied</div>
              </div>
            </div>
            <a
              href={result.url}
              download={result.filename}
              style={styles.downloadButton}
            >
              ⬇️ Download Updated File
            </a>
          </div>
        )}

        <div style={styles.logSection}>
          <h3 style={styles.logTitle}>Processing Log</h3>
          <div style={styles.logOutput}>
            {logs.length > 0 ? (
              logs.map((log, i) => (
                <div key={i}>{log}</div>
              ))
            ) : (
              <div style={{ color: C.mut }}>
                Upload files and click "Process Files" to begin...
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
   COLOR PALETTE — Matches 13WCF Rollforward Tool
   ═══════════════════════════════════════════════════════════════ */
const C = {
  bg: "#000", card: "#111", bdr: "#222",
  green: "#23D389", blue: "#2870F0", yellow: "#FDBD5B",
  purple: "#A456EA", red: "#D04E59",
  white: "#FFF", text: "#E8E8E8", mid: "#B0B0B0",
  mut: "#777", dim: "#555", div: "#1E1E1E",
  grad: "linear-gradient(135deg,#18E589,#2870E6)",
  gs: "rgba(35,211,137,0.12)",
};

/* ═══════════════════════════════════════════════════════════════
   STYLES
   ═══════════════════════════════════════════════════════════════ */
const styles = {
  container: {
    minHeight: "100vh",
    background: C.bg,
    fontFamily: "'DM Sans',sans-serif",
    color: C.text
  },
  header: {
    background: C.grad,
    height: "3px"
  },
  topBar: {
    background: C.card,
    borderBottom: `1px solid ${C.bdr}`,
    padding: "16px 32px"
  },
  topBarInner: {
    display: "flex",
    alignItems: "center",
    gap: "14px",
    maxWidth: "960px",
    margin: "0 auto"
  },
  logo: {
    width: "26px",
    height: "26px",
    borderRadius: "6px",
    background: C.grad,
    display: "flex",
    alignItems: "center",
    justifyContent: "center"
  },
  logoText: {
    color: "#fff",
    fontSize: "13px",
    fontWeight: "700"
  },
  brandName: {
    fontSize: "15px",
    fontWeight: "700",
    color: C.white,
    letterSpacing: "-0.3px"
  },
  divider: {
    width: "1px",
    height: "22px",
    background: C.bdr
  },
  card: {
    maxWidth: "960px",
    margin: "32px auto",
    background: C.card,
    borderRadius: "8px",
    border: `1px solid ${C.bdr}`,
    padding: "32px"
  },
  title: {
    fontSize: "18px",
    fontWeight: "700",
    marginBottom: "8px",
    color: C.white
  },
  subtitle: {
    fontSize: "12px",
    color: C.mut,
    marginBottom: "24px"
  },
  uploadSection: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "16px",
    marginBottom: "24px"
  },
  uploadBox: {
    borderWidth: "1px",
    borderStyle: "solid",
    borderColor: C.bdr,
    borderRadius: "6px",
    padding: "20px",
    background: C.bg,
    cursor: "pointer",
    transition: "all 0.15s"
  },
  uploadBoxActive: {
    background: C.gs,
    borderColor: C.green
  },
  uploadBoxDragOver: {
    background: C.gs,
    borderColor: C.blue,
    borderWidth: "2px",
    borderStyle: "dashed"
  },
  uploadTitle: {
    fontSize: "12px",
    fontWeight: "600",
    color: C.text,
    marginBottom: "8px"
  },
  uploadSubtitle: {
    fontSize: "10px",
    color: C.mut
  },
  fileInput: {
    display: "none"
  },
  uploadButton: {
    padding: "10px 20px",
    background: C.blue,
    color: "white",
    border: "none",
    borderRadius: "6px",
    cursor: "pointer",
    fontSize: "12px",
    fontWeight: "600",
    transition: "all 0.2s",
    marginTop: "12px"
  },
  processButton: {
    width: "100%",
    padding: "14px",
    background: C.green,
    color: C.bg,
    border: "none",
    borderRadius: "6px",
    fontSize: "14px",
    fontWeight: "700",
    cursor: "pointer",
    marginBottom: "20px",
    transition: "all 0.2s"
  },
  processButtonDisabled: {
    background: C.dim,
    cursor: "not-allowed",
    color: C.mut
  },
  errorBox: {
    background: C.card,
    border: `1px solid ${C.red}`,
    borderRadius: "6px",
    padding: "16px",
    marginBottom: "16px",
    color: C.red
  },
  successBox: {
    background: C.card,
    border: `1px solid ${C.green}`,
    borderRadius: "6px",
    padding: "20px",
    marginBottom: "20px"
  },
  successTitle: {
    fontSize: "14px",
    fontWeight: "600",
    color: C.green,
    marginBottom: "16px"
  },
  statsGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(4, 1fr)",
    gap: "12px",
    margin: "16px 0"
  },
  statCard: {
    background: C.bg,
    border: `1px solid ${C.bdr}`,
    padding: "14px",
    borderRadius: "6px",
    textAlign: "center"
  },
  statValue: {
    fontSize: "20px",
    fontWeight: "700",
    color: C.white,
    marginBottom: "4px"
  },
  statLabel: {
    fontSize: "10px",
    color: C.mut,
    textTransform: "uppercase",
    letterSpacing: "0.5px"
  },
  downloadButton: {
    display: "inline-block",
    padding: "12px 24px",
    background: C.blue,
    color: "white",
    textDecoration: "none",
    borderRadius: "6px",
    fontWeight: "600",
    marginTop: "10px",
    fontSize: "13px"
  },
  logSection: {
    marginTop: "24px"
  },
  logTitle: {
    fontSize: "13px",
    fontWeight: "600",
    marginBottom: "10px",
    color: C.white
  },
  logOutput: {
    background: C.bg,
    border: `1px solid ${C.bdr}`,
    color: C.text,
    padding: "16px",
    borderRadius: "6px",
    fontFamily: "monospace",
    fontSize: "11px",
    maxHeight: "400px",
    overflowY: "auto",
    lineHeight: "1.6"
  }
};
