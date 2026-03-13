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

      // Use direct child scan for boolean font flags to avoid false positives
      // from xmldom's getElementsByTagNameNS traversing nested elements.
      const childNames = new Set();
      const childVals = {};
      const children = Array.from(el.childNodes).filter(n => n.nodeType === 1);
      for (const child of children) {
        const tag = (child.localName || child.tagName || '').toLowerCase();
        childNames.add(tag);
        const v = child.getAttribute('val');
        if (v !== null) childVals[tag] = v;
      }

      // Bold: <b/> present AND val != "0"
      if (childNames.has('b') && childVals['b'] !== '0') f.fontWeight = 'bold';
      // Italic: <i/> present AND val != "0"
      if (childNames.has('i') && childVals['i'] !== '0') f.fontStyle = 'italic';

      // Underline
      if (childNames.has('u') && childVals['u'] !== 'none') f.textDecoration = 'underline';

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

      // Removed check for explicit 'applyFoo="1"' because Excel often omits it
      if (fonts[fontId]) Object.assign(style, fonts[fontId]);
      if (fills[fillId] && Object.keys(fills[fillId]).length > 0) Object.assign(style, fills[fillId]);
      if (borders[borderId] && Object.keys(borders[borderId]).length > 0) Object.assign(style, borders[borderId]);
      if (numFmtId && numFmtMap[numFmtId]) style.numberFormat = numFmtMap[numFmtId];

      // Read alignment separately
      const align = getFirstNode(xf, 'alignment');
      if (align) {
        const h = align.getAttribute('horizontal');
        if (h && h !== 'general') style.textAlign = h;
        const v = align.getAttribute('vertical');
        if (v) style.verticalAlign = v;
        if (align.getAttribute('wrapText') === '1') style.wrap = true;
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

    // --- 8. Extract embedded images per sheet ---
    const sheetImageArrays = await extractXlsxImages(zip, sheetPaths, parser);

    console.log(`[Parser] Extracted ${cellXfStyles.length} styles, ${borders.length} borders, ${sheetStyleMaps.length} sheet maps.`);
    return { cellXfStyles, sheetStyleMaps, sheetImageArrays };
  } catch (err) {
    console.error('[Parser] Style extraction failed:', err);
    return null;
  }
}

// ─── XLSX Image Extraction ────────────────────────────────────
/**
 * Extract embedded images from xl/drawings/drawingN.xml.
 * Returns an array of per-sheet image arrays:
 *   [ [{row, col, width, height, src}], ... ]
 * where src is a data:image/* URL.
 */
async function extractXlsxImages(zip, sheetPaths, parser) {
  const getNodes = (p, t) => Array.from(p.getElementsByTagNameNS('*', t));
  const getFirstNode = (p, t) => p.getElementsByTagNameNS('*', t)[0] || null;
  const EMU_PER_PX = 9525; // 914400 EMU/inch ÷ 96 dpi

  // Build sheet → drawing path map via sheet .rels
  const sheetImages = await Promise.all(sheetPaths.map(async (sheetPath, _si) => {
    const images = [];
    try {
      // Get sheet rels
      const parts = sheetPath.split('/');
      const relsPath = parts.slice(0, -1).join('/') + '/_rels/' + parts[parts.length - 1] + '.rels';
      const relsXml = await zip.file(relsPath)?.async('string');
      if (!relsXml) return images;

      const relsDoc = parser.parseFromString(relsXml, 'text/xml');
      const relationships = getNodes(relsDoc, 'Relationship');

      for (const rel of relationships) {
        const type = rel.getAttribute('Type') || '';
        if (!type.includes('drawing')) continue;

        let drawingTarget = rel.getAttribute('Target') || '';
        // Resolve path relative to sheet directory
        if (drawingTarget.startsWith('../')) {
          drawingTarget = 'xl/' + drawingTarget.replace('../', '');
        } else if (!drawingTarget.startsWith('xl/')) {
          const dir = parts.slice(0, -1).join('/');
          drawingTarget = dir + '/' + drawingTarget;
        }

        const drawingXml = await zip.file(drawingTarget)?.async('string');
        if (!drawingXml) continue;

        const drawingDoc = parser.parseFromString(drawingXml, 'text/xml');

        // Build drawing rels to resolve image paths
        const drawingParts = drawingTarget.split('/');
        const drawingRelsPath = drawingParts.slice(0, -1).join('/') + '/_rels/' + drawingParts[drawingParts.length - 1] + '.rels';
        const drawingRelsXml = await zip.file(drawingRelsPath)?.async('string');
        const drawingRels = drawingRelsXml ? parser.parseFromString(drawingRelsXml, 'text/xml') : null;
        const drawingRelMap = new Map();
        if (drawingRels) {
          for (const r of getNodes(drawingRels, 'Relationship')) {
            drawingRelMap.set(r.getAttribute('Id'), r.getAttribute('Target'));
          }
        }

        // Parse oneCellAnchor and twoCellAnchor elements
        for (const anchorTag of ['oneCellAnchor', 'twoCellAnchor']) {
          const anchors = getNodes(drawingDoc, anchorTag);
          for (const anchor of anchors) {
            try {
              // From cell (top-left)
              const fromEl = getFirstNode(anchor, 'from');
              if (!fromEl) continue;
              const rowEl = getFirstNode(fromEl, 'row');
              const colEl = getFirstNode(fromEl, 'col');
              const row = parseInt(rowEl?.textContent || '0');
              const col = parseInt(colEl?.textContent || '0');

              // Size in EMU
              let widthPx = null, heightPx = null;
              const extEl = getFirstNode(anchor, 'ext');
              if (extEl) {
                const cx = parseInt(extEl.getAttribute('cx') || '0');
                const cy = parseInt(extEl.getAttribute('cy') || '0');
                widthPx = Math.round(cx / EMU_PER_PX);
                heightPx = Math.round(cy / EMU_PER_PX);
              }

              // Find the blipFill → blip → r:embed
              const blip = getFirstNode(anchor, 'blip');
              if (!blip) continue;
              const rEmbed = blip.getAttribute('r:embed') || blip.getAttribute('embed');
              if (!rEmbed) continue;

              let mediaTarget = drawingRelMap.get(rEmbed);
              if (!mediaTarget) continue;
              // Resolve relative to drawing dir
              if (mediaTarget.startsWith('../')) {
                mediaTarget = 'xl/' + mediaTarget.replace('../', '');
              } else if (!mediaTarget.startsWith('xl/')) {
                const drawingDir = drawingParts.slice(0, -1).join('/');
                mediaTarget = drawingDir + '/' + mediaTarget;
              }

              const mediaData = await zip.file(mediaTarget)?.async('base64');
              if (!mediaData) continue;

              const ext = (mediaTarget.split('.').pop() || 'png').toLowerCase();
              const mimeMap = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', tiff: 'image/tiff', wmf: 'image/wmf', emf: 'image/emf' };
              const mime = mimeMap[ext] || 'image/png';
              const src = `data:${mime};base64,${mediaData}`;

              images.push({ row, col, width: widthPx, height: heightPx, src });
            } catch (anchorErr) {
              console.warn('[Parser] Skipping anchor:', anchorErr.message);
            }
          }
        }
      }
    } catch (err) {
      console.warn('[Parser] Image extraction failed for sheet:', err.message);
    }
    return images;
  }));

  return sheetImages;
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

    // Inject images extracted from drawings XML
    const sheetImgs = styleData?.sheetImageArrays?.[si];
    if (sheetImgs && sheetImgs.length > 0) {
      meta.images = sheetImgs;
    }

    self.postMessage({ type: 'SHEET_META', sheet: sheetName, meta, totalRows, totalCols: range.e.c - range.s.c + 1 });

    const CHUNK_SIZE = 100;
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

      // Yield execution to the event loop so the main thread can process and render the chunk
      await new Promise(resolve => setTimeout(resolve, 0));
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
