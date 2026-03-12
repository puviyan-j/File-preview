/**
 * Workbook Data Model
 *
 * Central data model for a parsed workbook. Holds sheets,
 * style pool, merge ranges, row/col sizing, and hidden rows/cols.
 */

import { createSheetStore, createStringPool } from './columnStore.js';

// ─── Defaults ─────────────────────────────────────────────────
// Change this value to scale defaults (e.g., 1, 1.25, 1.5)
export const SCALE_FACTOR = 1.15;

export const DEFAULT_ROW_HEIGHT = 24 * SCALE_FACTOR;
export const DEFAULT_COL_WIDTH = 120 * SCALE_FACTOR;
const INITIAL_CAPACITY = 8192;

// ─── Style Pool ───────────────────────────────────────────────

function createStylePool() {
  // Index 0 = default style (lean — only what renderEngine needs as fallback)
  const defaultStyle = {
    fontFamily: 'Calibri',
    fontSize: 11,
    fontWeight: 'normal',
    fontStyle: 'normal',
    textDecoration: 'none',
    color: '#000000',
    backgroundColor: null,
    textAlign: 'left',
    verticalAlign: 'bottom',
    wrap: false,
    numberFormat: null,
    borderLeft: null,
    borderRight: null,
    borderTop: null,
    borderBottom: null,
  };

  const styles = [defaultStyle];
  const keyMap = new Map();
  keyMap.set(JSON.stringify(defaultStyle), 0);

  return {
    get defaultStyle() { return defaultStyle; },

    /**
     * Intern a style object and return its index.
     * Stores ONLY the provided fields (no merging with defaults),
     * so distinct styles stay distinct.
     */
    intern(style) {
      if (!style) return 0;
      // Build a clean object with only non-null/non-default meaningful fields
      const clean = {};
      for (const [k, v] of Object.entries(style)) {
        if (v !== null && v !== undefined) {
          clean[k] = v;
        }
      }
      if (Object.keys(clean).length === 0) return 0;

      const key = JSON.stringify(clean);
      let id = keyMap.get(key);
      if (id === undefined) {
        id = styles.length;
        // Store the lean style — renderer will fall back to defaults for missing fields
        styles.push(clean);
        keyMap.set(key, id);
      }
      return id;
    },

    /** Get style by index — always returns an object (may be sparse) */
    get(index) {
      return styles[index] || defaultStyle;
    },

    get size() {
      return styles.length;
    },
  };
}

// ─── Sheet Model ──────────────────────────────────────────────

function createSheetModel(name) {
  const store = createSheetStore();

  // Row heights / col widths — Uint16Array, lazily grown
  let rowHeights = new Uint16Array(INITIAL_CAPACITY);
  let colWidths = new Uint16Array(256);
  rowHeights.fill(DEFAULT_ROW_HEIGHT);
  colWidths.fill(DEFAULT_COL_WIDTH);

  // Merge ranges
  const merges = [];
  const mergeIndex = new Map(); // "row,col" → merge object

  // Hidden rows/columns
  const hiddenRows = new Set();
  const hiddenCols = new Set();

  // Images and charts
  const images = [];
  const charts = [];

  // Cumulative position caches (invalidated on resize/hidden change)
  let rowPositionCache = null;
  let colPositionCache = null;

  const sheet = {
    name,
    store,
    get rowCount() { return store.rowCount; },
    get colCount() { return store.colCount; },

    // ── Row/Col sizing ─────────────────────────────────
    rowHeights,
    colWidths,
    hiddenRows,
    hiddenCols,

    setRowHeight(row, h) {
      if (row >= rowHeights.length) sheet._growRowHeights(row + 1);
      rowHeights[row] = h;
      rowPositionCache = null;
    },
    setColWidth(col, w) {
      if (col >= colWidths.length) sheet._growColWidths(col + 1);
      colWidths[col] = w;
      colPositionCache = null;
    },
    getRowHeight(row) {
      if (row >= rowHeights.length) return DEFAULT_ROW_HEIGHT;
      return rowHeights[row] || DEFAULT_ROW_HEIGHT;
    },
    getColWidth(col) {
      if (col >= colWidths.length) return DEFAULT_COL_WIDTH;
      return colWidths[col] || DEFAULT_COL_WIDTH;
    },

    _growRowHeights(minSize) {
      const newSize = Math.max(minSize, rowHeights.length * 2);
      const newArr = new Uint16Array(newSize);
      newArr.fill(DEFAULT_ROW_HEIGHT);
      newArr.set(rowHeights);
      rowHeights = newArr;
      sheet.rowHeights = rowHeights;
    },
    _growColWidths(minSize) {
      const newSize = Math.max(minSize, colWidths.length * 2);
      const newArr = new Uint16Array(newSize);
      newArr.fill(DEFAULT_COL_WIDTH);
      newArr.set(colWidths);
      colWidths = newArr;
      sheet.colWidths = colWidths;
    },

    // ── Cumulative positions ───────────────────────────
    getRowPositions() {
      if (rowPositionCache) return rowPositionCache;
      const count = store.rowCount;
      const positions = new Float64Array(count + 1);
      let y = 0;
      for (let r = 0; r < count; r++) {
        positions[r] = y;
        if (!hiddenRows.has(r)) {
          y += (r < rowHeights.length ? rowHeights[r] : DEFAULT_ROW_HEIGHT);
        }
      }
      positions[count] = y;
      rowPositionCache = positions;
      return positions;
    },
    getColPositions() {
      if (colPositionCache) return colPositionCache;
      const count = store.colCount;
      const positions = new Float64Array(count + 1);
      let x = 0;
      for (let c = 0; c < count; c++) {
        positions[c] = x;
        if (!hiddenCols.has(c)) {
          x += (c < colWidths.length ? colWidths[c] : DEFAULT_COL_WIDTH);
        }
      }
      positions[count] = x;
      colPositionCache = positions;
      return positions;
    },
    invalidatePositionCache() {
      rowPositionCache = null;
      colPositionCache = null;
    },

    getTotalHeight() {
      const pos = sheet.getRowPositions();
      return pos[pos.length - 1] || 0;
    },
    getTotalWidth() {
      const pos = sheet.getColPositions();
      return pos[pos.length - 1] || 0;
    },

    // ── Merges ─────────────────────────────────────────
    merges,
    mergeIndex,
    addMerge(startRow, startCol, endRow, endCol) {
      const merge = { startRow, startCol, endRow, endCol };
      merges.push(merge);
      // Index every cell in the merge range → the merge object
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          mergeIndex.set(`${r},${c}`, merge);
        }
      }
    },
    getMerge(row, col) {
      return mergeIndex.get(`${row},${col}`) || null;
    },

    // ── Images & Charts ────────────────────────────────
    images,
    charts,
    addImage(img) { images.push(img); },
    addChart(chart) { charts.push(chart); },
  };

  return sheet;
}

// ─── Workbook Model ───────────────────────────────────────────

export function createWorkbookModel() {
  const sheets = [];
  const sheetMap = new Map(); // name → sheet
  const stylePool = createStylePool();
  const stringPool = createStringPool();

  let parsingComplete = false;
  let totalParsedRows = 0;
  // Maps parser xfIndex → our internal stylePool index
  let styleMapping = new Uint32Array(512);

  return {
    sheets,
    stylePool,
    stringPool,
    get sheetCount() { return sheets.length; },
    get parsingComplete() { return parsingComplete; },
    get totalParsedRows() { return totalParsedRows; },

    /** Get or create a sheet by name */
    getOrCreateSheet(name) {
      let sheet = sheetMap.get(name);
      if (!sheet) {
        sheet = createSheetModel(name);
        sheets.push(sheet);
        sheetMap.set(name, sheet);
      }
      return sheet;
    },

    getSheet(index) {
      return sheets[index] || null;
    },

    getSheetByName(name) {
      return sheetMap.get(name) || null;
    },

    /**
     * Process a parsed row chunk into the model.
     * @param {string} sheetName
     * @param {number} startRow
     * @param {number} startCol  - absolute column index where chunk begins
     * @param {Array<Array>} rows - array of row arrays (values, relative to startCol)
     * @param {Array<Array<number>>} [styleRows] - parallel style xf indices
     */
    addChunk(sheetName, startRow, startCol, rows, styleRows) {
      const sheet = this.getOrCreateSheet(sheetName);
      const endRow = startRow + rows.length;
      const absStartCol = startCol || 0;

      // Find max col count (absolute)
      let maxAbsCols = sheet.store.colCount;
      for (const row of rows) {
        if (row && absStartCol + row.length > maxAbsCols) maxAbsCols = absStartCol + row.length;
      }

      sheet.store.ensureSize(endRow, maxAbsCols);

      // Ensure row/col arrays are large enough
      if (endRow > sheet.rowHeights.length) sheet._growRowHeights(endRow);
      if (maxAbsCols > sheet.colWidths.length) sheet._growColWidths(maxAbsCols);

      // Write cells
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;
        const rowIdx = startRow + r;

        for (let ci = 0; ci < row.length; ci++) {
          const absCol = absStartCol + ci;
          const val = row[ci];
          const styleId = styleRows ? (styleRows[r]?.[ci] || 0) : 0;
          const mappedStyleId = styleId < styleMapping.length ? styleMapping[styleId] : 0;

          if (val === null || val === undefined || val === '') {
            // Even empty cells may have a style (background, border)
            if (mappedStyleId) {
              sheet.store.setCellStyle(absCol, rowIdx, mappedStyleId);
            }
            continue;
          }

          if (typeof val === 'number') {
            sheet.store.setCell(absCol, rowIdx, val, mappedStyleId);
          } else {
            const sid = stringPool.intern(String(val));
            sheet.store.setCell(absCol, rowIdx, sid, mappedStyleId);
          }
        }
      }

      // Invalidate position caches since row/col counts may have changed
      sheet.invalidatePositionCache();
      totalParsedRows += rows.length;
    },

    /** Apply sheet metadata (merges, col widths, row heights, etc.) */
    applySheetMetadata(sheetName, meta) {
      const sheet = this.getOrCreateSheet(sheetName);

      if (meta.colWidths) {
        for (let c = 0; c < meta.colWidths.length; c++) {
          if (meta.colWidths[c]) sheet.setColWidth(c, meta.colWidths[c]);
        }
      }

      if (meta.rowHeights) {
        for (let r = 0; r < meta.rowHeights.length; r++) {
          if (meta.rowHeights[r]) sheet.setRowHeight(r, meta.rowHeights[r]);
        }
      }

      if (meta.merges) {
        for (const m of meta.merges) {
          sheet.addMerge(m.s.r, m.s.c, m.e.r, m.e.c);
        }
      }

      if (meta.hiddenRows) {
        for (const r of meta.hiddenRows) sheet.hiddenRows.add(r);
      }

      if (meta.hiddenCols) {
        for (const c of meta.hiddenCols) sheet.hiddenCols.add(c);
      }

      if (meta.images) {
        for (const img of meta.images) sheet.addImage(img);
      }

      if (meta.charts) {
        for (const chart of meta.charts) sheet.addChart(chart);
      }
    },

    /** Apply styles from the parser — maps xfIndex → internal stylePool index */
    applyStyles(parsedStyles) {
      if (!parsedStyles) return;
      if (parsedStyles.length > styleMapping.length) {
        styleMapping = new Uint32Array(parsedStyles.length + 128);
      } else {
        styleMapping.fill(0);
      }
      for (let i = 0; i < parsedStyles.length; i++) {
        styleMapping[i] = stylePool.intern(parsedStyles[i]);
      }
      return styleMapping;
    },

    setParsingComplete() {
      parsingComplete = true;
    },

    /** Resolve a cell's display string */
    resolveCellValue(sheetIndex, row, col) {
      const sheet = sheets[sheetIndex];
      if (!sheet) return '';
      const raw = sheet.store.getCell(col, row);
      if (raw === undefined || raw === null) return '';
      if (typeof raw === 'number') {
        const str = stringPool.get(raw);
        if (str && isNaN(raw)) return str;
        return raw;
      }
      return String(raw);
    },
  };
}
