/**
 * Virtualization Engine
 *
 * Computes visible row/column ranges from scroll position,
 * accounting for variable row heights, column widths, and hidden rows/cols.
 */

import { DEFAULT_ROW_HEIGHT, DEFAULT_COL_WIDTH } from './workbookModel.js';

/**
 * Binary search for the first row whose cumulative position is ≥ target.
 * @param {Float64Array} positions - cumulative row positions
 * @param {number} target - scroll position (px)
 * @returns {number} row index
 */
function binarySearchPosition(positions, target) {
  let lo = 0;
  let hi = positions.length - 1;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (positions[mid] < target) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  return Math.max(0, lo - 1);
}

/**
 * Get the visible range of rows and columns.
 *
 * @param {object} params
 * @param {number} params.scrollTop
 * @param {number} params.scrollLeft
 * @param {number} params.viewportWidth
 * @param {number} params.viewportHeight
 * @param {object} params.sheet - sheet model with getRowPositions(), getColPositions()
 * @param {number} [params.overscan=5] - extra rows/cols to render outside viewport
 * @returns {{ startRow, endRow, startCol, endCol, rowPositions, colPositions }}
 */
export function getVisibleRange({
  scrollTop,
  scrollLeft,
  viewportWidth,
  viewportHeight,
  sheet,
  overscan = 5,
}) {
  if (!sheet || sheet.rowCount === 0 || sheet.colCount === 0) {
    return { startRow: 0, endRow: 0, startCol: 0, endCol: 0, rowPositions: null, colPositions: null };
  }

  const rowPositions = sheet.getRowPositions();
  const colPositions = sheet.getColPositions();

  // Rows
  let startRow = binarySearchPosition(rowPositions, scrollTop);
  startRow = Math.max(0, startRow - overscan);

  // Find end row
  const bottomEdge = scrollTop + viewportHeight;
  let endRow = startRow;
  while (endRow < sheet.rowCount && rowPositions[endRow] < bottomEdge) {
    endRow++;
  }
  endRow = Math.min(sheet.rowCount, endRow + overscan);

  // Columns
  let startCol = binarySearchPosition(colPositions, scrollLeft);
  startCol = Math.max(0, startCol - overscan);

  const rightEdge = scrollLeft + viewportWidth;
  let endCol = startCol;
  while (endCol < sheet.colCount && colPositions[endCol] < rightEdge) {
    endCol++;
  }
  endCol = Math.min(sheet.colCount, endCol + overscan);

  return {
    startRow,
    endRow,
    startCol,
    endCol,
    rowPositions,
    colPositions,
  };
}

/**
 * Compute the total scrollable content dimensions.
 */
export function getContentSize(sheet) {
  if (!sheet) return { width: 0, height: 0 };
  return {
    width: sheet.getTotalWidth(),
    height: sheet.getTotalHeight(),
  };
}
