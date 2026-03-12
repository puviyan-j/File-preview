/**
 * Parser Web Worker
 *
 * Parses XLSX/XLS/ODS (via SheetJS) and CSV/TSV files inside a Web Worker.
 * Extracts cell styles directly from XLSX XML using JSZip (SheetJS community
 * edition does not export styles).
 * Emits row chunks progressively so the main thread can render immediately.
 */

import * as XLSX from 'xlsx';

import JSZip from 'jszip';
import { DOMParser } from 'xmldom';
import { SCALE_FACTOR } from '../engine/workbookModel.js';

// ─── Message Handler ──────────────────────────────────────────
self.onmessage = function (e) {
  const { type, buffer, fileName } = e.data;
  if (type === 'PARSE') {
    parseFile(buffer, fileName);
  }
};

// ─── Format Detection ─────────────────────────────────────────
function detectFormat(fileName) {
  const ext = (fileName || '').split('.').pop().toLowerCase();
  if (ext === 'csv') return 'csv';
  if (ext === 'tsv') return 'tsv';
  if (ext === 'ods') return 'ods';
  if (ext === 'xls') return 'xls';
  return 'xlsx';
}

// ─── Main Parse Entry ─────────────────────────────────────────
async function parseFile(buffer, fileName) {
  self.postMessage({ type: 'PARSING_START', fileName });
  const format = detectFormat(fileName);

  if (format === 'csv' || format === 'tsv') {
    parseDelimited(buffer, format === 'tsv' ? '\t' : ',', fileName);
  } else {
    await parseWorkbook(buffer, fileName, format);
  }
}

// ─── XLSX Style Extraction via JSZip ──────────────────────────

async function extractXlsxStyles(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const stylesXml = await zip.file('xl/styles.xml')?.async('string');
    if (!stylesXml) return null;

    const parser = new DOMParser();
    const doc = parser.parseFromString(stylesXml, 'text/xml');

    const getNodes = (p, t) => Array.from(p.getElementsByTagNameNS('*', t));
    const getFirstNode = (p, t) => p.getElementsByTagNameNS('*', t)[0] || null;

    // --- 1. Theme Colors ---
    const themeXml = await zip.file('xl/theme/theme1.xml')?.async('string');
    const themeColors = new Array(12);
    if (themeXml) {
      const themeDoc = parser.parseFromString(themeXml, 'text/xml');
      const clrScheme = getFirstNode(themeDoc, 'clrScheme');
      if (clrScheme) {
        const schemeMap = {
          lt1: 0, dk1: 1, lt2: 2, dk2: 3,
          accent1: 4, accent2: 5, accent3: 6, accent4: 7,
          accent5: 8, accent6: 9, hlink: 10, folHlink: 11
        };
        const children = Array.from(clrScheme.childNodes).filter(n => n.nodeType === 1);
        for (const child of children) {
          const name = child.localName || child.tagName.split(':').pop();
          const idx = schemeMap[name];
          if (idx !== undefined) {
            const sysClr = getFirstNode(child, 'sysClr');
            if (sysClr) themeColors[idx] = sysClr.getAttribute('lastClr') || (idx % 2 === 0 ? 'FFFFFF' : '000000');
            else {
              const srgbClr = getFirstNode(child, 'srgbClr');
              if (srgbClr) themeColors[idx] = srgbClr.getAttribute('val');
            }
          }
        }
      }
    }
    // Fallback defaults
    const defaults = ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646'];
    for (let i = 0; i < 10; i++) if (!themeColors[i]) themeColors[i] = defaults[i];

    function resolveColor(el) {
      if (!el) return null;
      let hex = el.getAttribute('rgb');
      if (hex && hex.length > 6) hex = hex.slice(-6); // strip alpha
      if (!hex) {
        const theme = el.getAttribute('theme');
        if (theme !== null) {
          const tIdx = parseInt(theme);
          if (tIdx >= 0 && tIdx < themeColors.length) hex = themeColors[tIdx];
        }
      }
      if (!hex || hex.length < 6) return null;

      const tint = parseFloat(el.getAttribute('tint') || '0');
      let r = parseInt(hex.slice(0, 2), 16);
      let g = parseInt(hex.slice(2, 4), 16);
      let b = parseInt(hex.slice(4, 6), 16);
      if (tint !== 0) {
        const applyTint = (v, tn) => Math.round(tn > 0 ? v * (1 - tn) + 255 * tn : v * (1 + tn));
        r = applyTint(r, tint);
        g = applyTint(g, tint);
        b = applyTint(b, tint);
      }
      const toHex = (x) => Math.max(0, Math.min(255, x)).toString(16).padStart(2, '0');
      return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
    }

    // --- 2. Fonts ---
    const fonts = getNodes(doc, 'font').map(el => {
      const f = {};
      const sz = getFirstNode(el, 'sz');
      if (sz) f.fontSize = parseFloat(sz.getAttribute('val')) || 11;
      const nm = getFirstNode(el, 'name');
      if (nm) f.fontFamily = nm.getAttribute('val');
      if (getFirstNode(el, 'b')) f.fontWeight = 'bold';
      if (getFirstNode(el, 'i')) f.fontStyle = 'italic';
      const uNode = getFirstNode(el, 'u');
      if (uNode && uNode.getAttribute('val') !== 'none') f.textDecoration = 'underline';
      const cNode = getFirstNode(el, 'color');
      if (cNode) {
        const c = resolveColor(cNode);
        if (c) f.color = c;
      }
      return f;
    });

    // --- 3. Fills ---
    const fills = getNodes(doc, 'fill').map(el => {
      const pf = getFirstNode(el, 'patternFill');
      if (!pf) return {};
      const patternType = pf.getAttribute('patternType') || '';
      // 'none' means no fill. For 'solid' or named patterns, use fgColor.
      if (patternType === 'none') return {};
      const fgColorEl = getFirstNode(pf, 'fgColor');
      const background = resolveColor(fgColorEl);
      if (!background) return {};
      return { backgroundColor: background };
    });

    // --- 4. Borders ---
    const borders = getNodes(doc, 'border').map(el => {
      const b = {};
      const sides = ['left', 'right', 'top', 'bottom', 'diagonal'];
      for (const side of sides) {
        const sideEl = getFirstNode(el, side);
        if (!sideEl) continue;
        const style = sideEl.getAttribute('style');
        if (!style || style === 'none') continue;
        const colorEl = getFirstNode(sideEl, 'color');
        const color = resolveColor(colorEl) || '#000000';
        b[`border${side.charAt(0).toUpperCase() + side.slice(1)}`] = { style, color };
      }
      return b;
    });

    // --- 5. Number Formats (built-in + custom) ---
    const numFmtMap = {};
    // Built-in Excel number format IDs
    const builtinFmts = {
      0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
      9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
      14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
      18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
      22: 'm/d/yy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
      39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
      45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
    };
    Object.assign(numFmtMap, builtinFmts);
    const numFmtsEl = getFirstNode(doc, 'numFmts');
    if (numFmtsEl) {
      for (const nf of getNodes(numFmtsEl, 'numFmt')) {
        const id = parseInt(nf.getAttribute('numFmtId') || '0');
        const fmt = nf.getAttribute('formatCode') || '';
        numFmtMap[id] = fmt;
      }
    }

    // --- 6. Cell XFs (the primary style lookup table) ---
    const cellXfsEl = getFirstNode(doc, 'cellXfs');
    const xfNodes = cellXfsEl ? getNodes(cellXfsEl, 'xf') : [];

    const cellXfStyles = xfNodes.map(xf => {
      const style = {};

      const fontId = parseInt(xf.getAttribute('fontId') || '0');
      const fillId = parseInt(xf.getAttribute('fillId') || '0');
      const borderId = parseInt(xf.getAttribute('borderId') || '0');
      const numFmtId = parseInt(xf.getAttribute('numFmtId') || '0');
      const applyFont = xf.getAttribute('applyFont') !== '0';
      const applyFill = xf.getAttribute('applyFill') !== '0';
      const applyBorder = xf.getAttribute('applyBorder') !== '0';
      const applyAlignment = xf.getAttribute('applyAlignment') !== '0';
      const applyNumberFormat = xf.getAttribute('applyNumberFormat') !== '0';

      if (applyFont && fonts[fontId]) Object.assign(style, fonts[fontId]);
      if (applyFill && fills[fillId] && Object.keys(fills[fillId]).length > 0) Object.assign(style, fills[fillId]);
      if (applyBorder && borders[borderId] && Object.keys(borders[borderId]).length > 0) Object.assign(style, borders[borderId]);
      if (applyNumberFormat && numFmtId && numFmtMap[numFmtId]) style.numberFormat = numFmtMap[numFmtId];

      if (applyAlignment) {
        const align = getFirstNode(xf, 'alignment');
        if (align) {
          const h = align.getAttribute('horizontal');
          if (h && h !== 'general') style.textAlign = h;
          const v = align.getAttribute('vertical');
          if (v) style.verticalAlign = v;
          if (align.getAttribute('wrapText') === '1') style.wrap = true;
        }
      }

      return Object.keys(style).length > 0 ? style : null;
    });

    // --- 7. Build per-sheet style maps using string keys "r,c" ---
    const wbXml = await zip.file('xl/workbook.xml')?.async('string');
    if (!wbXml) return { cellXfStyles, sheetStyleMaps: [] };
    const wbDoc = parser.parseFromString(wbXml, 'text/xml');
    const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
    const relsDoc = relsXml ? parser.parseFromString(relsXml, 'text/xml') : null;

    const sheetPaths = getNodes(wbDoc, 'sheet').map((s, i) => {
      const rId = s.getAttribute('r:id');
      const rels = getNodes(relsDoc || wbDoc, 'Relationship');
      const rel = rels.find(r => r.getAttribute('Id') === rId);
      if (!rel) return `xl/worksheets/sheet${i + 1}.xml`;
      const target = rel.getAttribute('Target');
      return target.startsWith('xl/') ? target : `xl/${target}`;
    });

    const sheetStyleMaps = await Promise.all(sheetPaths.map(async path => {
      const wsXml = await zip.file(path)?.async('string');
      // Map: "row,col" -> xfIndex  (string key avoids hash collisions)
      const map = new Map();
      if (wsXml) {
        // Also capture row-level default styles
        const rowRe = /<[a-z0-9:]*row\s+([^>]*?)(?:\/?>)/gi;
        let rm;
        while ((rm = rowRe.exec(wsXml)) !== null) {
          const attrs = rm[1];
          const rM = attrs.match(/\br="(\d+)"/i);
          const sM = attrs.match(/\bs="(\d+)"/i);
          const cDef = attrs.match(/\bcustomFormat="1"/i);
          if (rM && sM && cDef) {
            // Store row default style: negative col means row-level
            const rowIdx = parseInt(rM[1]) - 1;
            map.set(`${rowIdx},__row__`, parseInt(sM[1]));
          }
        }

        // Cell-level styles (override row defaults)
        const re = /<[a-z0-9:]*c\s+([^>]*?)(?:\/?>)/gi;
        let m;
        while ((m = re.exec(wsXml)) !== null) {
          const attrs = m[1];
          const sM = attrs.match(/\bs="(\d+)"/i);
          const rM = attrs.match(/\br="([A-Z]{1,3})(\d+)"/i);
          if (rM) {
            const xf = sM ? parseInt(sM[1]) : 0;
            const col = colLetterToIndex(rM[1]);
            const row = parseInt(rM[2]) - 1;
            // Only store if there's an explicit style
            if (sM) {
              map.set(`${row},${col}`, xf);
            }
          }
        }
      }
      return map;
    }));

    console.log(`[Parser] Extracted ${cellXfStyles.length} styles, ${borders.length} borders, ${sheetStyleMaps.length} sheet maps.`);
    return { cellXfStyles, sheetStyleMaps };
  } catch (err) {
    console.error('[Parser] Style extraction failed:', err);
    return null;
  }
}

function colLetterToIndex(letters) {
  let idx = 0;
  for (let i = 0; i < letters.length; i++) {
    idx = idx * 26 + (letters.charCodeAt(i) - 64);
  }
  return idx - 1;
}

// ─── SheetJS Workbook Parser ──────────────────────────────────
async function parseWorkbook(buffer, fileName, format) {
  // Step 1: extract styles for XLSX (save buffer copy before transferring)
  let styleData = null;
  const bufferCopy = buffer.slice(0);
  if (format === 'xlsx') {
    styleData = await extractXlsxStyles(bufferCopy);
  }

  // Send extracted styles to main thread
  if (styleData?.cellXfStyles) {
    self.postMessage({ type: 'STYLES', styles: styleData.cellXfStyles });
  }

  // Step 2: parse cell data with SheetJS
  const data = new Uint8Array(buffer);
  const workbook = XLSX.read(data, {
    type: 'array',
    cellDates: true,
    cellNF: true,
    cellStyles: true,
  });

  const totalSheets = workbook.SheetNames.length;

  for (let si = 0; si < totalSheets; si++) {
    const sheetName = workbook.SheetNames[si];
    const ws = workbook.Sheets[sheetName];

    if (!ws || !ws['!ref']) {
      self.postMessage({ type: 'SHEET_EMPTY', sheet: sheetName });
      continue;
    }

    const range = XLSX.utils.decode_range(ws['!ref']);
    const totalRows = range.e.r - range.s.r + 1;

    const meta = extractSheetMetadata(ws, sheetName, range);
    self.postMessage({ type: 'SHEET_META', sheet: sheetName, meta, totalRows, totalCols: range.e.c - range.s.c + 1 });

    const CHUNK_SIZE = 2000;
    const startR = range.s.r;
    const endR = range.e.r;
    const startC = range.s.c; // absolute start column
    const endC = range.e.c;

    const sheetStyleMap = styleData?.sheetStyleMaps?.[si] || null;

    for (let chunkStart = startR; chunkStart <= endR; chunkStart += CHUNK_SIZE) {
      const chunkEnd = Math.min(chunkStart + CHUNK_SIZE - 1, endR);
      const rows = [];
      const styleRows = [];

      for (let r = chunkStart; r <= chunkEnd; r++) {
        const row = [];
        const styleRow = [];
        // Get row-level default style
        const rowDefaultStyle = sheetStyleMap ? (sheetStyleMap.get(`${r},__row__`) || 0) : 0;

        for (let c = startC; c <= endC; c++) {
          const addr = XLSX.utils.encode_cell({ r, c });
          const cell = ws[addr];
          if (cell) {
            row.push(cell.w !== undefined ? cell.w : (cell.v !== undefined ? cell.v : null));
          } else {
            row.push(null);
          }
          // Use cell style, fall back to row default
          let xfIdx = 0;
          if (sheetStyleMap) {
            const cellStyle = sheetStyleMap.get(`${r},${c}`);
            xfIdx = cellStyle !== undefined ? cellStyle : rowDefaultStyle;
          }
          styleRow.push(xfIdx);
        }
        rows.push(row);
        styleRows.push(styleRow);
      }

      // Pass absolute startCol so model can use correct indices
      self.postMessage({
        type: 'ROW_CHUNK',
        sheet: sheetName,
        startRow: chunkStart,
        startCol: startC,  // ← NEW: absolute column start
        rows,
        styleRows,
      });

      const progress = Math.min(1, (chunkEnd - startR + 1) / totalRows);
      self.postMessage({
        type: 'PARSING_PROGRESS',
        sheet: sheetName,
        sheetIndex: si,
        totalSheets,
        progress,
        parsedRows: chunkEnd - startR + 1,
        totalRows,
      });
    }
  }

  self.postMessage({ type: 'PARSING_COMPLETE' });
}

// ─── Sheet Metadata Extraction ────────────────────────────────
function extractSheetMetadata(ws, sheetName, range) {
  const meta = {
    merges: [], colWidths: [], rowHeights: [],
    hiddenRows: [], hiddenCols: [], images: [], charts: [],
    startCol: range.s.c,  // pass through absolute start col
  };

  if (ws['!merges']) {
    meta.merges = ws['!merges'].map(m => ({
      s: { r: m.s.r, c: m.s.c },
      e: { r: m.e.r, c: m.e.c },
    }));
  }

  if (ws['!cols']) {
    for (let c = 0; c < ws['!cols'].length; c++) {
      const col = ws['!cols'][c];
      if (col) {
        // wpx = width in pixels, wch = width in characters
        if (col.wpx) {
          meta.colWidths[c] = Math.round(col.wpx * SCALE_FACTOR);
        } else if (col.wch) {
          // Excel character width → approximate pixels (7px per char + 5px padding)
          meta.colWidths[c] = Math.round((col.wch * 7 + 5) * SCALE_FACTOR);
        }
        if (col.hidden) meta.hiddenCols.push(c);
      }
    }
  }

  if (ws['!rows']) {
    for (let r = 0; r < ws['!rows'].length; r++) {
      const row = ws['!rows'][r];
      if (row) {
        if (row.hpx) {
          meta.rowHeights[r] = Math.round(row.hpx * SCALE_FACTOR);
        } else if (row.hpt) {
          // points → pixels (1pt ≈ 1.333px)
          meta.rowHeights[r] = Math.round(row.hpt * 1.333 * SCALE_FACTOR);
        }
        if (row.hidden) meta.hiddenRows.push(r);
      }
    }
  }

  return meta;
}

// ─── CSV / TSV Fast Parser ────────────────────────────────────
function parseDelimited(buffer, delimiter, fileName) {
  const decoder = new TextDecoder('utf-8');
  const text = decoder.decode(buffer);
  const lines = text.split(/\r?\n/);

  const totalRows = lines.length;
  const CHUNK_SIZE = 5000;
  const sheetName = fileName || 'Sheet1';

  const firstRow = parseDelimitedLine(lines[0] || '', delimiter);
  const totalCols = firstRow.length;

  self.postMessage({
    type: 'SHEET_META', sheet: sheetName,
    meta: { merges: [], colWidths: [], rowHeights: [], hiddenRows: [], hiddenCols: [], images: [], charts: [], startCol: 0 },
    totalRows, totalCols,
  });

  for (let chunkStart = 0; chunkStart < totalRows; chunkStart += CHUNK_SIZE) {
    const chunkEnd = Math.min(chunkStart + CHUNK_SIZE, totalRows);
    const rows = [];

    for (let r = chunkStart; r < chunkEnd; r++) {
      if (!lines[r] && r === totalRows - 1) continue;
      rows.push(parseDelimitedLine(lines[r] || '', delimiter));
    }

    if (rows.length > 0) {
      self.postMessage({
        type: 'ROW_CHUNK', sheet: sheetName,
        startRow: chunkStart, startCol: 0, rows, styleRows: null,
      });
    }

    self.postMessage({
      type: 'PARSING_PROGRESS', sheet: sheetName,
      sheetIndex: 0, totalSheets: 1,
      progress: Math.min(1, chunkEnd / totalRows),
      parsedRows: chunkEnd, totalRows,
    });
  }

  self.postMessage({ type: 'PARSING_COMPLETE' });
}

function parseDelimitedLine(line, delimiter) {
  const result = [];
  let i = 0;
  const len = line.length;

  while (i <= len) {
    if (i >= len) { result.push(''); break; }

    if (line[i] === '"') {
      let val = '';
      i++;
      while (i < len) {
        if (line[i] === '"') {
          if (i + 1 < len && line[i + 1] === '"') { val += '"'; i += 2; }
          else { i++; break; }
        } else { val += line[i]; i++; }
      }
      result.push(val);
      if (i < len && line[i] === delimiter) i++;
    } else {
      let end = line.indexOf(delimiter, i);
      if (end === -1) end = len;
      const val = line.substring(i, end);
      const num = Number(val);
      result.push(val === '' ? null : isNaN(num) ? val : num);
      i = end + 1;
    }
  }
  return result;
}
