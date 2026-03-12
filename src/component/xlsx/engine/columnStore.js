/**
 * Column Storage Engine
 *
 * Column-oriented storage with TypedArrays for numeric data,
 * string pooling for text, and hybrid dense/sparse modes.
 */

// ─── String Pool ──────────────────────────────────────────────
export function createStringPool() {
  const strings = [];
  const indexMap = new Map(); // string → id

  return {
    /** Intern a string and return its ID */
    intern(str) {
      if (str === null || str === undefined) return -1;
      const s = String(str);
      let id = indexMap.get(s);
      if (id === undefined) {
        id = strings.length;
        strings.push(s);
        indexMap.set(s, id);
      }
      return id;
    },
    /** Retrieve string by ID */
    get(id) {
      return id < 0 ? '' : (strings[id] ?? '');
    },
    get size() {
      return strings.length;
    },
  };
}

// ─── Column Types ─────────────────────────────────────────────

const SPARSE_THRESHOLD = 0.3; // switch to sparse below 30% fill

/**
 * Create a dense numeric column backed by a TypedArray.
 */
function createDenseNumericColumn(numType, capacity) {
  const ArrayCtor = numType === 'int' ? Int32Array : Float64Array;
  let values = new ArrayCtor(capacity);
  let styleIndex = new Uint32Array(capacity);
  let filled = new Uint8Array(capacity); // 1 = has value

  return {
    type: numType,
    storage: 'dense',
    get(row) {
      return filled[row] ? values[row] : undefined;
    },
    set(row, val, styleId = 0) {
      if (row >= values.length) return;
      values[row] = val;
      filled[row] = 1;
      styleIndex[row] = styleId;
    },
    setStyle(row, styleId) {
      if (row >= styleIndex.length) return;
      styleIndex[row] = styleId;
    },
    getStyle(row) {
      return styleIndex[row] || 0;
    },
    get capacity() {
      return capacity;
    },
    grow(newCapacity) {
      if (newCapacity <= capacity) return;
      const newValues = new ArrayCtor(newCapacity);
      newValues.set(values);
      const newStyle = new Uint32Array(newCapacity);
      newStyle.set(styleIndex);
      const newFilled = new Uint8Array(newCapacity);
      newFilled.set(filled);
      values = newValues;
      styleIndex = newStyle;
      filled = newFilled;
      capacity = newCapacity;
    },
  };
}

/**
 * Create a dense string column that stores string-pool IDs.
 */
function createDenseStringColumn(capacity) {
  let ids = new Int32Array(capacity).fill(-1);
  let styleIndex = new Uint32Array(capacity);

  return {
    type: 'string',
    storage: 'dense',
    get(row) {
      return ids[row] >= 0 ? ids[row] : undefined;
    },
    set(row, stringId, styleId = 0) {
      if (row >= ids.length) return;
      ids[row] = stringId;
      styleIndex[row] = styleId;
    },
    setStyle(row, styleId) {
      if (row >= styleIndex.length) return;
      styleIndex[row] = styleId;
    },
    getStyle(row) {
      return styleIndex[row] || 0;
    },
    get capacity() {
      return ids.length;
    },
    grow(newCapacity) {
      if (newCapacity <= ids.length) return;
      const newIds = new Int32Array(newCapacity).fill(-1);
      newIds.set(ids);
      const newStyle = new Uint32Array(newCapacity);
      newStyle.set(styleIndex);
      ids = newIds;
      styleIndex = newStyle;
    },
  };
}

/**
 * Create a sparse column (for any type) backed by Maps.
 * Also handles style-only entries (cells with style but no value).
 */
function createSparseColumn() {
  const values = new Map();
  const styles = new Map();

  return {
    type: 'mixed',
    storage: 'sparse',
    get(row) {
      return values.get(row);
    },
    set(row, val, styleId = 0) {
      values.set(row, val);
      if (styleId) styles.set(row, styleId);
    },
    setStyle(row, styleId) {
      if (styleId) styles.set(row, styleId);
    },
    getStyle(row) {
      return styles.get(row) || 0;
    },
    get size() {
      return values.size;
    },
    grow() {
      // No-op for sparse
    },
  };
}

// ─── Column Factory ───────────────────────────────────────────

export function detectColumnType(sampleValues, totalRows, filledCount) {
  if (totalRows > 0 && filledCount / totalRows < SPARSE_THRESHOLD) {
    return 'sparse';
  }

  let allInt = true;
  let allNum = true;
  for (const v of sampleValues) {
    if (v === null || v === undefined || v === '') continue;
    if (typeof v !== 'number') { allInt = false; allNum = false; break; }
    if (!Number.isInteger(v)) allInt = false;
  }

  if (allInt && allNum) return 'int';
  if (allNum) return 'float';
  return 'string';
}

export function createColumn(colType, capacity) {
  switch (colType) {
    case 'int':   return createDenseNumericColumn('int', capacity);
    case 'float': return createDenseNumericColumn('float', capacity);
    case 'string': return createDenseStringColumn(capacity);
    default:       return createSparseColumn();
  }
}

// ─── Sheet Column Store ───────────────────────────────────────

/**
 * Create a columnar store for one sheet.
 * Columns are lazily created as data arrives.
 */
export function createSheetStore() {
  const columns = [];  // colIndex → Column
  const colTypes = [];
  let rowCount = 0;
  let colCount = 0;

  function ensureColumn(colIdx) {
    while (columns.length <= colIdx) columns.push(null);
    if (!columns[colIdx]) {
      columns[colIdx] = createSparseColumn();
      colTypes[colIdx] = 'sparse';
    }
    return columns[colIdx];
  }

  return {
    get rowCount() { return rowCount; },
    get colCount() { return colCount; },
    get columns() { return columns; },

    ensureSize(rows, cols) {
      if (rows > rowCount) rowCount = rows;
      if (cols > colCount) colCount = cols;
      for (let c = 0; c < columns.length; c++) {
        const col = columns[c];
        if (col && col.grow && col.capacity !== undefined && col.capacity < rowCount) {
          col.grow(rowCount);
        }
      }
    },

    getColumn(colIdx) {
      return ensureColumn(colIdx);
    },

    upgradeColumn(colIdx, newType, capacity) {
      if (colTypes[colIdx] === newType) return;
      columns[colIdx] = createColumn(newType, capacity || rowCount);
      colTypes[colIdx] = newType;
    },

    /** Set a cell value (and optional style) */
    setCell(colIdx, rowIdx, value, styleId = 0) {
      const col = ensureColumn(colIdx);
      col.set(rowIdx, value, styleId);
    },

    /**
     * Set style on a cell that has no value.
     * Used for empty cells that still have background/border styles.
     */
    setCellStyle(colIdx, rowIdx, styleId) {
      if (!styleId) return;
      const col = ensureColumn(colIdx);
      col.setStyle(rowIdx, styleId);
    },

    /** Get a cell value */
    getCell(colIdx, rowIdx) {
      const col = columns[colIdx];
      if (!col) return undefined;
      return col.get(rowIdx);
    },

    /** Get a cell's style index */
    getCellStyle(colIdx, rowIdx) {
      const col = columns[colIdx];
      if (!col) return 0;
      return col.getStyle(rowIdx) || 0;
    },
  };
}
