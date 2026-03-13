/**
 * Formula Engine — HyperFormula integration
 *
 * Evaluates Excel formulas in a parsed sheet's cell data.
 * Accepts a 2D array of raw cell values (strings, numbers, nulls, or
 * formula strings starting with '=') and returns a 2D array of resolved
 * display values, ready to be stored in the workbook model.
 */

import { HyperFormula } from 'hyperformula';

// HyperFormula options for Excel-compatible evaluation
const HF_OPTIONS = {
  licenseKey: 'gpl-v3',
  // Use English locale for function names like SUM, IF, VLOOKUP etc.
  language: 'enGB',
};

let _hf = null;

/** Get (or lazily create) a reusable HyperFormula instance */
function getHF() {
  if (!_hf) {
    _hf = HyperFormula.buildEmpty(HF_OPTIONS);
  }
  return _hf;
}

/**
 * Evaluate all formulas in a single sheet's cell data.
 *
 * @param {Array<Array<any>>} sheetData  2-D array [row][col] of raw values
 * @param {number}            startRow   Absolute row index of sheetData[0]
 * @param {number}            startCol   Absolute col index of sheetData[0][0]
 * @returns {Array<Array<any>>}          Same shape as sheetData; formula cells replaced by computed values
 */
export function evaluateFormulas(sheetData, startRow = 0, startCol = 0) {
  if (!sheetData || sheetData.length === 0) return sheetData;

  // Check if any cell is a formula to avoid HF overhead when not needed
  let hasFormula = false;
  for (const row of sheetData) {
    if (!row) continue;
    for (const cell of row) {
      if (typeof cell === 'string' && cell.startsWith('=')) {
        hasFormula = true;
        break;
      }
    }
    if (hasFormula) break;
  }
  if (!hasFormula) return sheetData;

  const hf = getHF();
  let sheetId;

  try {
    // Build HF-format data: prepend empty rows/cols for the offset
    // so cell references like A1 map to absolute row 0, col 0.
    const hfData = [];
    // Pad rows before startRow
    for (let r = 0; r < startRow; r++) hfData.push([]);
    for (const row of sheetData) {
      const hfRow = [];
      // Pad cols before startCol
      for (let c = 0; c < startCol; c++) hfRow.push(null);
      if (row) {
        for (const cell of row) {
          hfRow.push(cell === undefined ? null : cell);
        }
      }
      hfData.push(hfRow);
    }

    sheetId = hf.addSheet('eval');
    hf.setSheetContent(sheetId, hfData);

    // Read back evaluated values for the data region
    const result = [];
    for (let ri = 0; ri < sheetData.length; ri++) {
      const srcRow = sheetData[ri];
      if (!srcRow) { result.push(srcRow); continue; }
      const outRow = [];
      for (let ci = 0; ci < srcRow.length; ci++) {
        const cell = srcRow[ci];
        if (typeof cell === 'string' && cell.startsWith('=')) {
          try {
            const val = hf.getCellValue({ sheet: sheetId, row: startRow + ri, col: startCol + ci });
            // HyperFormula returns error objects for failed formulas
            if (val !== null && typeof val === 'object' && val.type) {
              outRow.push(`#${val.type}`);
            } else {
              outRow.push(val ?? '');
            }
          } catch {
            outRow.push(cell); // fallback: show raw formula
          }
        } else {
          outRow.push(cell);
        }
      }
      result.push(outRow);
    }
    return result;
  } catch (err) {
    console.warn('[FormulaEngine] Evaluation error:', err.message);
    return sheetData;
  } finally {
    // Clean up the temporary sheet to avoid memory leaks
    if (sheetId !== undefined) {
      try { hf.removeSheet(sheetId); } catch { /* ignore */ }
    }
  }
}

/** Dispose the shared HyperFormula instance (call on workbook close) */
export function destroyFormulaEngine() {
  if (_hf) {
    _hf.destroy();
    _hf = null;
  }
}
